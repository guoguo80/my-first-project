"""
异常数据分析模块
增加自动识别上传频率功能，用于正确合并连续异常时间段
首张工作表为文件汇总，无异常不建空表，地埋设备异常列拆分
采用先收集后写入方式保证汇总表在第一个
"""

import pandas as pd
import os
import glob
from datetime import datetime


def determine_freq_from_filename(filename):
    """根据文件名判断上传频率"""
    if "状态" in filename:
        return 10
    elif "SKY2" in filename and "状态" not in filename:
        return 5
    else:
        return 1  # 默认1分钟


def analyze_abnormal_data(df, freq_min=1):
    """分析单个DataFrame的异常数据，freq_min为上传频率（分钟）"""
    try:
        if '上传日期' not in df.columns:
            raise ValueError("数据中缺少必须字段：上传日期")

        # 转换时间格式并排序
        df['上传日期'] = pd.to_datetime(df['上传日期'])
        df = df.sort_values('上传日期').reset_index(drop=True)

        # 判断数据是否包含地埋设备特有列（用于内部记录）
        has_water_level = '水位(mm)' in df.columns
        has_flood_state = '水浸状态' in df.columns
        is_buried = has_water_level and has_flood_state

        # 收集原始异常记录
        abnormal_records = []  # 每个元素为字典

        for idx, row in df.iterrows():
            upload_time = row['上传日期']
            reasons = []
            water_val = None
            flood_val = None

            # 异常1：所有数值型字段出现9999或9998
            numeric_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
            for col in numeric_cols:
                val = row[col]
                if pd.notna(val) and val in [9999, 9998]:
                    reasons.append(f"{col}异常值({val})")

            # 异常2：降雨量或雨量数据本分钟数据小于上一分钟数据
            if idx > 0:
                if '降雨量' in df.columns:
                    prev_rain = df.loc[idx - 1, '降雨量']
                    curr_rain = row['降雨量']
                    if pd.notna(prev_rain) and pd.notna(curr_rain) and curr_rain < prev_rain:
                        reasons.append(f"降雨量({curr_rain})<前一分钟降雨量({prev_rain})")

                if '雨量' in df.columns:
                    prev_rain = df.loc[idx - 1, '雨量']
                    curr_rain = row['雨量']
                    if pd.notna(prev_rain) and pd.notna(curr_rain) and curr_rain < prev_rain:
                        reasons.append(f"雨量({curr_rain})<前一分钟雨量({prev_rain})")

            # 异常3：水浸状态出现0、1以外的数据
            if has_flood_state:
                flood_state = row['水浸状态']
                if pd.notna(flood_state) and flood_state not in [0, 1]:
                    reasons.append(f"水浸状态({flood_state})应为0或1")
                if is_buried:
                    flood_val = flood_state if pd.notna(flood_state) else None

            # 异常4：水位非0但水浸状态为0（针对地埋设备数据）
            if is_buried:
                water = row['水位(mm)']
                state = row['水浸状态']
                if pd.notna(water) and pd.notna(state) and water != 0 and state == 0:
                    reasons.append(f"水位({water}mm)非0但水浸状态为0")
                water_val = water if pd.notna(water) else None

            if reasons:
                record = {
                    '时间': upload_time,
                    '原因列表': reasons,
                }
                if is_buried:
                    record['水位'] = water_val
                    record['水浸状态'] = flood_val
                abnormal_records.append(record)

        # 合并连续异常+计算异常时长
        merged_records = []
        if abnormal_records:
            # 按时间排序
            abnormal_records.sort(key=lambda x: x['时间'])
            current = abnormal_records[0].copy()
            current['开始时间'] = current['时间']
            current['结束时间'] = current['时间']
            del current['时间']

            for record in abnormal_records[1:]:
                prev_time = current['结束时间']
                curr_time = record['时间']
                time_diff = (curr_time - prev_time).total_seconds()

                # 判断是否连续且原因相同
                reasons_same = (current['原因列表'] == record['原因列表'])
                if is_buried:
                    # 地埋设备还需水位和水浸状态相同才能合并
                    reasons_same = reasons_same and (current.get('水位') == record.get('水位')) and \
                                   (current.get('水浸状态') == record.get('水浸状态'))

                if time_diff == freq_min * 60 and reasons_same:
                    current['结束时间'] = curr_time
                else:
                    # 计算当前段的时长
                    duration_min = int((current['结束时间'] - current['开始时间']).total_seconds() / 60) + 1
                    current['异常时长(分钟)'] = duration_min
                    merged_records.append(current)

                    # 开始新段
                    current = record.copy()
                    current['开始时间'] = current['时间']
                    current['结束时间'] = current['时间']
                    del current['时间']

            # 处理最后一段
            duration_min = int((current['结束时间'] - current['开始时间']).total_seconds() / 60) + 1
            current['异常时长(分钟)'] = duration_min
            merged_records.append(current)

        # 转换为DataFrame
        if merged_records:
            df_list = []
            for rec in merged_records:
                row_dict = {
                    '异常开始时间': rec['开始时间'].strftime('%Y-%m-%d %H:%M:%S'),
                    '异常结束时间': rec['结束时间'].strftime('%Y-%m-%d %H:%M:%S'),
                    '异常时长(分钟)': rec['异常时长(分钟)'],
                    '异常说明': '; '.join(rec['原因列表']),
                }
                if is_buried:
                    row_dict['水位(mm)'] = rec.get('水位')
                    row_dict['水浸状态'] = rec.get('水浸状态')
                df_list.append(row_dict)
            abnormal_df = pd.DataFrame(df_list)
            if is_buried:
                # 重新排列列顺序
                cols = ['异常开始时间', '异常结束时间', '水位(mm)', '水浸状态', '异常说明', '异常时长(分钟)']
                abnormal_df = abnormal_df[cols]
        else:
            abnormal_df = pd.DataFrame(columns=['异常开始时间', '异常结束时间', '异常说明', '异常时长(分钟)'])

        return abnormal_df

    except Exception as e:
        raise Exception(f"异常数据分析失败: {str(e)}")


def batch_analyze_abnormal(configs, progress_callback=None):
    """批量分析异常数据，自动识别每个文件的上传频率"""
    results = []

    for config in configs:
        input_dir = config["input_dir"]
        output_file = config["output_file"]

        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        # 获取路径下所有Excel文件
        excel_files = glob.glob(os.path.join(input_dir, '*.xls*'))
        if not excel_files:
            msg = f"路径[{input_dir}]未找到Excel文件 → 跳过"
            results.append({"status": "warning", "message": msg})
            if progress_callback:
                progress_callback(msg)
            continue

        # 用于暂存异常DataFrame和汇总信息
        abnormal_dfs = {}
        summary_data = []
        total_abnormal_count = 0
        processed_count = 0

        # 第一步：遍历所有文件，收集数据和统计信息
        for file_path in excel_files:
            file_name = os.path.basename(file_path)
            sheet_name = os.path.splitext(file_name)[0]

            try:
                # 根据文件名判断上传频率
                freq_min = determine_freq_from_filename(file_name)

                # 读取目标sheet
                df = pd.read_excel(file_path, sheet_name='数据')

                # 分析异常数据
                abnormal_df = analyze_abnormal_data(df, freq_min)

                abnormal_count = len(abnormal_df)
                total_abnormal_count += abnormal_count
                processed_count += 1

                # 记录汇总信息
                summary_data.append({
                    '文件名': file_name,
                    '是否有异常': '是' if abnormal_count > 0 else '否',
                    '异常段数': abnormal_count
                })

                # 仅当有异常时才保存DataFrame
                if abnormal_count > 0:
                    # 判断是否为地埋设备（根据文件名）
                    is_buried = ('KTS-04' in file_name or 'KTS-04-P' in file_name) and '状态' not in file_name

                    # 对于非地埋设备，移除水位和水浸状态列（如果存在）
                    if not is_buried:
                        for col in ['水位(mm)', '水浸状态']:
                            if col in abnormal_df.columns:
                                abnormal_df = abnormal_df.drop(columns=[col])

                    abnormal_dfs[sheet_name] = abnormal_df

                    msg = f"✅ {file_name} (频率:{freq_min}分钟) → 异常分析完成（异常段数：{abnormal_count}）"
                else:
                    msg = f"ℹ️ {file_name} (频率:{freq_min}分钟) → 无异常数据，跳过创建工作表"

                results.append({"status": "success", "message": msg})
                if progress_callback:
                    progress_callback(msg)

            except ValueError as e:
                if "No sheet named '数据'" in str(e):
                    msg = f"❌ {file_name} → 跳过：缺少「数据」sheet"
                elif "数据中缺少必须字段" in str(e):
                    msg = f"❌ {file_name} → 跳过：{str(e)}"
                else:
                    msg = f"❌ {file_name} → 跳过：数据格式错误 - {str(e)}"
                results.append({"status": "error", "message": msg})
                if progress_callback:
                    progress_callback(msg)
                # 即使处理失败，也记录汇总信息（异常段数为0）
                summary_data.append({
                    '文件名': file_name,
                    '是否有异常': '否',
                    '异常段数': 0
                })
                processed_count += 1
            except Exception as e:
                msg = f"❌ {file_name} → 处理失败：{str(e)}"
                results.append({"status": "error", "message": msg})
                if progress_callback:
                    progress_callback(msg)
                summary_data.append({
                    '文件名': file_name,
                    '是否有异常': '否',
                    '异常段数': 0
                })
                processed_count += 1

        # 第二步：打开ExcelWriter，按顺序写入工作表
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 1. 先写入汇总表
                if summary_data:
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='文件汇总', index=False)
                    # 调整汇总表列宽
                    ws_summary = writer.sheets['文件汇总']
                    for column in ws_summary.columns:
                        max_width = max(len(str(cell.value)) for cell in column if cell.value is not None)
                        ws_summary.column_dimensions[column[0].column_letter].width = min(max_width + 2, 50)

                # 2. 再写入各设备异常表
                for sheet_name, abnormal_df in abnormal_dfs.items():
                    abnormal_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    # 调整列宽
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_width = max(len(str(cell.value)) for cell in column if cell.value is not None)
                        worksheet.column_dimensions[column[0].column_letter].width = min(max_width + 2, 50)

            msg = f"\n📊 统计：共处理{processed_count}个文件，发现{total_abnormal_count}个异常时间段"
            results.append({"status": "info", "message": msg})
            msg = f"\n🎉 异常数据分析完成 → {os.path.abspath(output_file)}"
            results.append({"status": "success", "message": msg})
            if progress_callback:
                progress_callback(msg)

        except Exception as e:
            msg = f"❌ 文件写入失败：{str(e)}"
            results.append({"status": "error", "message": msg})
            if progress_callback:
                progress_callback(msg)

    return results