"""
到报率分析模块（优化版） - 增加公共缺失时间段分析，支持频率模式选择
修复：所有 numpy.timedelta64 对象的 total_seconds 调用问题，过滤临时文件
"""

import pandas as pd
import os
import glob
from datetime import datetime, timedelta
import logging
import numpy as np

logger = logging.getLogger(__name__)


def analyze_dataframe(df, freq_min, progress_callback=None, tolerance_seconds=60):
    """
    分析单个DataFrame，支持自定义上传频率，去重后计算到报率
    增加 tolerance_seconds：允许实际时间与标准时间点的最大偏差（秒）
    """
    try:
        if progress_callback:
            progress_callback("开始分析数据...")

        # 先复制原始数据并处理时间格式
        df_raw = df.copy()
        df_raw['上传日期'] = pd.to_datetime(df_raw['上传日期'])
        df_raw = df_raw.sort_values('上传日期').reset_index(drop=True)

        # ---------- 重复数据统计（基于原始数据） ----------
        freq_str = f"{freq_min}min"
        df_raw['周期时间'] = df_raw['上传日期'].dt.floor(freq_str)
        period_counts = df_raw['周期时间'].value_counts()
        duplicate_periods = period_counts[period_counts > 1]

        total_dup_periods = len(duplicate_periods)
        total_dup_records = (duplicate_periods - 1).sum()

        if progress_callback and total_dup_periods > 0:
            progress_callback(f"发现重复周期: {total_dup_periods}个, 重复记录: {total_dup_records}条")

        # 构建重复详情DataFrame
        duplicate_details = []
        for period_time, count in duplicate_periods.items():
            duplicate_details.append({
                '周期时间': period_time.strftime('%Y-%m-%d %H:%M:00'),
                '该周期总记录数': count,
                '重复记录数': count - 1
            })
        duplicate_df = pd.DataFrame(duplicate_details) if duplicate_details else pd.DataFrame(
            columns=['周期时间', '该周期总记录数', '重复记录数'])

        # ---------- 去重：按上传日期去除完全相同的记录 ----------
        df_dedup = df_raw.drop_duplicates(subset=['上传日期']).reset_index(drop=True)
        if progress_callback:
            progress_callback(f"原始数据 {len(df_raw)} 条，去重后 {len(df_dedup)} 条")

        # ---------- 使用去重后的数据计算到报率和缺失 ----------
        times = df_dedup['上传日期'].values  # numpy datetime64
        if len(times) < 2:
            # 只有一个点，无法计算缺失段
            start_time = times[0]
            end_time = times[0]
            total_seconds = 0
            total_minutes = 0
            expected_count = 1
            actual_count = len(times)
            report_rate = 1.0 if expected_count > 0 else 0.0

            missing_df = pd.DataFrame(columns=['缺失开始时间', '缺失结束时间', '缺失条数'])
            deviation_count = 0

            summary_data = {
                '指标': [
                    '开始时间', '结束时间', '总时长(分钟)', '上传频率(分钟)', '应报数据条数', '实报数据条数(去重后)', '到报率',
                    '重复周期数', '总重复记录数', '偏差次数(±1分钟)'
                ],
                '值': [
                    pd.Timestamp(start_time).strftime('%Y-%m-%d %H:%M:%S'),
                    pd.Timestamp(end_time).strftime('%Y-%m-%d %H:%M:%S'),
                    round(total_minutes, 2), freq_min, expected_count, actual_count, f"{report_rate:.2%}",
                    total_dup_periods, total_dup_records, deviation_count
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            return summary_df, missing_df, duplicate_df, pd.Timestamp(start_time), pd.Timestamp(end_time)

        start_time = times[0]
        end_time = times[-1]
        total_seconds = (end_time - start_time) / np.timedelta64(1, 's')
        total_minutes = total_seconds / 60.0

        # 生成标准时间轴：以第一个数据点时间为起点，频率 freq_min 分钟
        standard_times = pd.date_range(start=pd.Timestamp(start_time), end=pd.Timestamp(end_time), freq=f'{freq_min}min')
        std_times = standard_times.values  # numpy datetime64
        num_std = len(std_times)

        # 初始化匹配标记数组
        matched = np.zeros(num_std, dtype=bool)
        best_match_idx = {}  # key: 标准点索引, value: (实际点索引, 时间差秒)

        # 使用 searchsorted 为每个实际点找到最近的左右邻居
        indices = np.searchsorted(std_times, times)

        for i, t in enumerate(times):
            pos = indices[i]
            # 候选标准点索引
            candidates = []
            if pos > 0:
                candidates.append(pos - 1)
            if pos < num_std:
                candidates.append(pos)
            if pos + 1 < num_std:
                candidates.append(pos + 1)

            best_candidate = None
            best_diff = float('inf')
            for idx in candidates:
                diff = abs(t - std_times[idx]) / np.timedelta64(1, 's')
                if diff < best_diff:
                    best_diff = diff
                    best_candidate = idx

            if best_candidate is not None:
                if best_candidate in best_match_idx:
                    prev_i, prev_diff = best_match_idx[best_candidate]
                    if best_diff < prev_diff:
                        best_match_idx[best_candidate] = (i, best_diff)
                else:
                    best_match_idx[best_candidate] = (i, best_diff)

        # 根据最佳匹配标记标准点，并累计偏差次数（恰好±60秒）
        deviation_count = 0
        for std_idx, (real_i, diff) in best_match_idx.items():
            matched[std_idx] = True
            if abs(diff - 60) < 1e-6:
                deviation_count += 1

        # 收集缺失的标准时间点
        missing_times = [std_times[i] for i in range(num_std) if not matched[i]]

        # 合并连续缺失时间段（使用 numpy 方式计算时间差）
        missing_periods = []
        if missing_times:
            if progress_callback:
                progress_callback(f"发现缺失时间点: {len(missing_times)}个")

            current_start, current_end, count = missing_times[0], missing_times[0], 1
            for t in missing_times[1:]:
                diff_sec = (t - current_end) / np.timedelta64(1, 's')
                if diff_sec <= freq_min * 60:
                    current_end, count = t, count + 1
                else:
                    missing_periods.append({
                        '缺失开始时间': pd.Timestamp(current_start).strftime('%Y-%m-%d %H:%M:%S'),
                        '缺失结束时间': pd.Timestamp(current_end).strftime('%Y-%m-%d %H:%M:%S'),
                        '缺失条数': count
                    })
                    current_start, current_end, count = t, t, 1
            missing_periods.append({
                '缺失开始时间': pd.Timestamp(current_start).strftime('%Y-%m-%d %H:%M:%S'),
                '缺失结束时间': pd.Timestamp(current_end).strftime('%Y-%m-%d %H:%M:%S'),
                '缺失条数': count
            })

            if progress_callback:
                progress_callback(f"合并为连续缺失段: {len(missing_periods)}个")

        missing_df = pd.DataFrame(missing_periods) if missing_periods else pd.DataFrame(
            columns=['缺失开始时间', '缺失结束时间', '缺失条数'])

        expected_count = num_std
        actual_count = len(times)
        report_rate = actual_count / expected_count if expected_count > 0 else 0.0

        if progress_callback:
            progress_callback(f"应报数据: {expected_count}条, 实报数据(去重后): {actual_count}条, 到报率: {report_rate:.2%}")

        summary_data = {
            '指标': [
                '开始时间', '结束时间', '总时长(分钟)', '上传频率(分钟)', '应报数据条数', '实报数据条数(去重后)', '到报率',
                '重复周期数', '总重复记录数', '偏差次数(±1分钟)'
            ],
            '值': [
                pd.Timestamp(start_time).strftime('%Y-%m-%d %H:%M:%S'),
                pd.Timestamp(end_time).strftime('%Y-%m-%d %H:%M:%S'),
                round(total_minutes, 2), freq_min, expected_count, actual_count, f"{report_rate:.2%}",
                total_dup_periods, total_dup_records, deviation_count
            ]
        }
        summary_df = pd.DataFrame(summary_data)

        if progress_callback:
            progress_callback("数据分析完成")

        return summary_df, missing_df, duplicate_df, pd.Timestamp(start_time), pd.Timestamp(end_time)

    except Exception as e:
        error_msg = f"数据分析失败: {str(e)}"
        if progress_callback:
            progress_callback(error_msg)
        raise Exception(error_msg)


def find_common_time_range(devices_info):
    """查找所有设备的公共时间范围"""
    if not devices_info:
        return None, None

    all_start_times = []
    all_end_times = []

    for device_info in devices_info:
        if device_info['start_time'] and device_info['end_time']:
            all_start_times.append(device_info['start_time'])
            all_end_times.append(device_info['end_time'])

    if not all_start_times or not all_end_times:
        return None, None

    common_start = max(all_start_times)
    common_end = min(all_end_times)

    if common_start >= common_end:
        return None, None

    return common_start, common_end


def analyze_common_missing_periods(devices_info, freq_min, progress_callback=None):
    """分析所有设备的公共缺失时间段"""
    try:
        if progress_callback:
            progress_callback("开始分析公共缺失时间段...")

        common_start, common_end = find_common_time_range(devices_info)

        if not common_start or not common_end:
            if progress_callback:
                progress_callback("无法确定公共时间范围，跳过公共缺失分析")
            return pd.DataFrame(columns=['公共缺失开始时间', '公共缺失结束时间', '涉及设备数', '影响设备列表'])

        if progress_callback:
            progress_callback(f"公共时间范围: {common_start} 到 {common_end}")

        freq_str = f"{freq_min}min"
        common_timeline = pd.date_range(start=common_start, end=common_end, freq=freq_str)

        if progress_callback:
            progress_callback(f"公共时间轴点数: {len(common_timeline)}")

        device_count = len(devices_info)
        time_count = len(common_timeline)

        missing_matrix = []
        device_names = []

        for device_info in devices_info:
            device_name = device_info['device_name']
            missing_df = device_info['missing_df']

            device_missing = [0] * time_count

            if not missing_df.empty:
                for _, row in missing_df.iterrows():
                    missing_start = pd.to_datetime(row['缺失开始时间'])
                    missing_end = pd.to_datetime(row['缺失结束时间'])
                    mask = (common_timeline >= missing_start) & (common_timeline <= missing_end)
                    if mask.any():
                        for i in range(time_count):
                            if mask[i]:
                                device_missing[i] = 1

            missing_matrix.append(device_missing)
            device_names.append(device_name)

        common_missing_times = []

        for i in range(time_count):
            all_missing = True
            for j in range(device_count):
                if missing_matrix[j][i] == 0:
                    all_missing = False
                    break
            if all_missing:
                common_missing_times.append(common_timeline[i])

        if progress_callback:
            progress_callback(f"发现公共缺失时间点: {len(common_missing_times)}个")

        common_missing_periods = []

        if common_missing_times:
            current_start = common_missing_times[0]
            current_end = common_missing_times[0]

            for i in range(1, len(common_missing_times)):
                diff_sec = (common_missing_times[i] - current_end) / np.timedelta64(1, 's')
                if diff_sec == freq_min * 60:
                    current_end = common_missing_times[i]
                else:
                    duration_min = int((current_end - current_start) / np.timedelta64(1, 'm')) + 1

                    affected_devices = []
                    for j, device_name in enumerate(device_names):
                        mask = (common_timeline >= current_start) & (common_timeline <= current_end)
                        device_missing_in_period = any(missing_matrix[j][k] == 1 for k in range(time_count) if mask[k])
                        if device_missing_in_period:
                            affected_devices.append(device_name)

                    common_missing_periods.append({
                        '公共缺失开始时间': current_start.strftime('%Y-%m-%d %H:%M:%S'),
                        '公共缺失结束时间': current_end.strftime('%Y-%m-%d %H:%M:%S'),
                        '涉及设备数': len(affected_devices),
                        '影响设备列表': '; '.join(affected_devices),
                        '持续时间(分钟)': duration_min
                    })

                    current_start = common_missing_times[i]
                    current_end = common_missing_times[i]

            # 处理最后一个时间段
            duration_min = int((current_end - current_start) / np.timedelta64(1, 'm')) + 1
            affected_devices = []
            for j, device_name in enumerate(device_names):
                mask = (common_timeline >= current_start) & (common_timeline <= current_end)
                device_missing_in_period = any(missing_matrix[j][k] == 1 for k in range(time_count) if mask[k])
                if device_missing_in_period:
                    affected_devices.append(device_name)

            common_missing_periods.append({
                '公共缺失开始时间': current_start.strftime('%Y-%m-%d %H:%M:%S'),
                '公共缺失结束时间': current_end.strftime('%Y-%m-%d %H:%M:%S'),
                '涉及设备数': len(affected_devices),
                '影响设备列表': '; '.join(affected_devices),
                '持续时间(分钟)': duration_min
            })

        if progress_callback:
            progress_callback(f"合并为公共缺失时间段: {len(common_missing_periods)}个")

        common_missing_df = pd.DataFrame(common_missing_periods) if common_missing_periods else pd.DataFrame(
            columns=['公共缺失开始时间', '公共缺失结束时间', '涉及设备数', '影响设备列表', '持续时间(分钟)'])

        return common_missing_df

    except Exception as e:
        error_msg = f"公共缺失分析失败: {str(e)}"
        if progress_callback:
            progress_callback(error_msg)
        return pd.DataFrame(
            columns=['公共缺失开始时间', '公共缺失结束时间', '涉及设备数', '影响设备列表', '持续时间(分钟)'])


def determine_freq_from_filename(filename):
    """根据文件名判断上传频率"""
    if "状态" in filename:
        return 10
    elif "SKY2" in filename and "状态" not in filename:
        return 5
    else:
        return 1  # 默认1分钟


def is_temp_file(filename):
    """判断是否为Excel临时文件（以~$开头）"""
    return filename.startswith('~$')


def batch_analyze_by_config(configs, progress_callback=None):
    """按配置批量分析不同路径的数据，每个文件的频率根据配置中的模式决定"""
    results = []

    for config_idx, config in enumerate(configs):
        input_dir = config["input_dir"]
        output_file = config["output_file"]
        common_freq = config.get("freq_min", 1)
        freq_mode = config.get("freq_mode", "manual")

        if progress_callback:
            progress_callback(f"\n处理配置: {config.get('name', '未命名')}")
            progress_callback(f"输入目录: {input_dir}")
            progress_callback(f"输出文件: {output_file}")
            progress_callback(f"公共缺失分析基准频率: {common_freq}分钟")
            progress_callback(f"频率模式: {'自动判断' if freq_mode == 'auto' else '手动设置'}")

        os.makedirs(os.path.dirname(output_file) if os.path.dirname(output_file) else '.', exist_ok=True)

        # 获取路径下所有Excel文件，过滤掉临时文件
        all_files = glob.glob(os.path.join(input_dir, '*.xls*'))
        excel_files = [f for f in all_files if not is_temp_file(os.path.basename(f))]

        if not excel_files:
            msg = f"未找到Excel文件 → 目录: {input_dir}"
            results.append({"status": "warning", "message": msg})
            if progress_callback:
                progress_callback(f"⚠️ {msg}")
            continue

        if progress_callback:
            progress_callback(f"找到 {len(excel_files)} 个Excel文件")

        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                devices_info = []

                for file_idx, file_path in enumerate(excel_files, 1):
                    file_name = os.path.basename(file_path)
                    sheet_name = os.path.splitext(file_name)[0]

                    if freq_mode == "auto":
                        freq_min = determine_freq_from_filename(file_name)
                    else:
                        freq_min = common_freq

                    if progress_callback:
                        progress_callback(f"\n处理文件 [{file_idx}/{len(excel_files)}]: {file_name} (频率: {freq_min}分钟)")

                    try:
                        df = pd.read_excel(file_path, sheet_name='数据')

                        if progress_callback:
                            progress_callback(f"读取数据成功，共 {len(df)} 条记录")

                        summary_df, missing_df, duplicate_df, start_time, end_time = analyze_dataframe(
                            df, freq_min, progress_callback
                        )

                        if freq_min == 1:
                            devices_info.append({
                                'device_name': sheet_name,
                                'start_time': start_time,
                                'end_time': end_time,
                                'missing_df': missing_df
                            })

                        worksheet = writer.book.create_sheet(title=sheet_name)
                        writer.sheets[sheet_name] = worksheet

                        summary_df.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)

                        if not missing_df.empty:
                            worksheet.cell(row=len(summary_df) + 2, column=1, value="=== 缺失时间段详情 ===")
                            missing_df.to_excel(writer, sheet_name=sheet_name,
                                                startrow=len(summary_df) + 3, index=False)

                        if not duplicate_df.empty:
                            skip_rows = len(summary_df) + 3 + len(missing_df) + (2 if not missing_df.empty else 0)
                            worksheet.cell(row=skip_rows, column=1, value="=== 重复数据详情 ===")
                            duplicate_df.to_excel(writer, sheet_name=sheet_name,
                                                  startrow=skip_rows + 1, index=False)

                        for column in worksheet.columns:
                            max_width = max(len(str(cell.value)) for cell in column if cell.value is not None)
                            worksheet.column_dimensions[column[0].column_letter].width = min(max_width + 2, 50)

                        msg = f"✅ 成功处理: {file_name}"
                        results.append({"status": "success", "message": msg})
                        if progress_callback:
                            progress_callback(msg)

                    except ValueError as e:
                        if "No sheet named '数据'" in str(e):
                            msg = f"❌ 跳过: {file_name}（缺少「数据」工作表）"
                        elif "数据中缺少必须字段" in str(e):
                            msg = f"❌ 跳过: {file_name}（{str(e)}）"
                        else:
                            msg = f"❌ 处理失败: {file_name}（数据格式错误: {e}）"
                        results.append({"status": "error", "message": msg})
                        if progress_callback:
                            progress_callback(msg)
                    except Exception as e:
                        msg = f"❌ 处理失败: {file_name}（{str(e)}）"
                        results.append({"status": "error", "message": msg})
                        if progress_callback:
                            progress_callback(msg)

                # 分析公共缺失时间段
                if len(devices_info) >= 2:
                    if progress_callback:
                        progress_callback(f"\n🔍 开始分析公共缺失时间段（共{len(devices_info)}个设备，基准频率{common_freq}分钟）...")

                    common_missing_df = analyze_common_missing_periods(devices_info, common_freq, progress_callback)

                    if not common_missing_df.empty:
                        common_missing_sheet = writer.book.create_sheet(title="公共缺失时间段")
                        writer.sheets["公共缺失时间段"] = common_missing_sheet

                        common_missing_sheet.cell(row=1, column=1, value="=== 所有设备公共缺失时间段 ===")
                        common_missing_sheet.cell(row=2, column=1, value=f"分析设备数: {len(devices_info)}")
                        common_missing_sheet.cell(row=3, column=1, value=f"基准上传频率: {common_freq}分钟")

                        common_missing_df.to_excel(writer, sheet_name="公共缺失时间段",
                                                   startrow=5, index=False)

                        for column in common_missing_sheet.columns:
                            max_width = max(len(str(cell.value)) for cell in column if cell.value is not None)
                            common_missing_sheet.column_dimensions[column[0].column_letter].width = min(max_width + 2,
                                                                                                        50)

                        msg = f"✅ 发现 {len(common_missing_df)} 个公共缺失时间段"
                        results.append({"status": "success", "message": msg})
                        if progress_callback:
                            progress_callback(msg)
                    else:
                        msg = f"✅ 未发现公共缺失时间段"
                        results.append({"status": "info", "message": msg})
                        if progress_callback:
                            progress_callback(msg)
                else:
                    msg = f"⚠️ 设备数量不足（{len(devices_info)}个），跳过公共缺失分析"
                    results.append({"status": "warning", "message": msg})
                    if progress_callback:
                        progress_callback(msg)

                msg = f"\n🎉 处理完成 → 结果文件: {os.path.abspath(output_file)}"
                results.append({"status": "success", "message": msg})
                if progress_callback:
                    progress_callback(msg)

        except Exception as e:
            msg = f"❌ 文件写入失败: {str(e)}"
            results.append({"status": "error", "message": msg})
            if progress_callback:
                progress_callback(msg)

    return results