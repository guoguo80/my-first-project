"""
气象水文数据分析软件 - GUI界面（优化版本）
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import os
import sys
import time
import logging
from datetime import datetime

# 导入功能模块
from modules.data_reporting import batch_analyze_by_config
from modules.abnormal_analysis import batch_analyze_abnormal
from modules.rainfall_analysis import analyze_rainfall_batch
from modules.weather_analysis import process_multiple_devices
from modules.deviation_analysis import read_rainfall_data, analyze_rainfall_devices


class WeatherHydroAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("气象水文数据分析系统 v1.0.0")

        # 设置窗口图标
        self.set_window_icon()

        # 设置日志系统
        self.setup_logging()

        # 当前分析任务
        self.current_task = None
        self.is_running = False
        self.task_thread = None

        # 优化窗口显示
        self.setup_window_geometry()

        # 创建GUI组件
        self.create_widgets()

        # 绑定窗口事件
        self.bind_window_events()

        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def set_window_icon(self):
        """设置窗口图标"""
        try:
            # 尝试从多个可能的路径加载图标
            icon_paths = ["icon.ico", "./icon.ico", "气象水文数据分析系统.ico"]
            for path in icon_paths:
                if os.path.exists(path):
                    self.root.iconbitmap(path)
                    self.log_info(f"已设置窗口图标: {path}")
                    break
        except:
            pass  # 图标文件不存在，使用默认图标

    def setup_window_geometry(self):
        """设置窗口几何属性，优化显示效果"""
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 根据屏幕尺寸动态计算窗口大小
        if screen_width >= 1920 and screen_height >= 1080:
            # 大屏幕 (1920x1080 或更高)
            window_width = 1400
            window_height = 780
        elif screen_width >= 1600 and screen_height >= 900:
            # 中等屏幕
            window_width = 1280
            window_height = 720
        else:
            # 小屏幕
            window_width = 1150
            window_height = 680

        # 确保窗口不会超出屏幕边界
        window_width = min(window_width, screen_width - 50)
        window_height = min(window_height, screen_height - 100)

        # 计算窗口位置（居中偏上）
        x = (screen_width - window_width) // 2
        y = max(10, (screen_height - window_height) // 2 - 20)

        # 设置窗口几何属性
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 设置窗口最小尺寸
        self.root.minsize(1000, 600)

        # 设置窗口最大尺寸为屏幕尺寸（允许最大化占满屏幕）
        self.root.maxsize(screen_width, screen_height)

        # 初始窗口状态为正常（非最大化）
        self.root.state('normal')

        self.log_info(f"窗口大小: {window_width}x{window_height}, 位置: ({x}, {y})")
        self.log_info(f"屏幕尺寸: {screen_width}x{screen_height}")

    def bind_window_events(self):
        """绑定窗口事件"""
        # 绑定窗口大小变化事件
        self.root.bind('<Configure>', self.on_window_configure)

        # 绑定窗口状态变化事件
        self.root.bind('<Map>', self.on_window_map)  # 窗口显示时

    def on_window_configure(self, event):
        """窗口大小或位置变化时的处理"""
        if event.widget == self.root:
            # 获取当前窗口位置和大小
            x = self.root.winfo_x()
            y = self.root.winfo_y()
            width = self.root.winfo_width()
            height = self.root.winfo_height()

            # 获取屏幕尺寸
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()

            # 检查窗口是否移出屏幕右侧
            if x + width > screen_width:
                new_x = max(0, screen_width - width)
                self.root.geometry(f"{width}x{height}+{new_x}+{y}")

            # 检查窗口是否移出屏幕底部（考虑任务栏）
            taskbar_height = 40  # 假设任务栏高度
            if y + height > screen_height - taskbar_height:
                new_y = max(0, screen_height - height - taskbar_height)
                self.root.geometry(f"{width}x{height}+{x}+{new_y}")

    def on_window_map(self, event):
        """窗口显示时的处理"""
        # 确保窗口可见
        self.root.lift()
        self.root.focus_force()

    def setup_logging(self):
        """设置日志系统"""
        # 创建日志目录
        log_dir = "运行日志"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)

        # 创建日志文件名（带时间戳）
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"运行日志_{timestamp}.log")

        # 配置日志记录器
        self.logger = logging.getLogger('气象水文分析系统')
        self.logger.setLevel(logging.INFO)

        # 清除之前可能存在的处理器
        self.logger.handlers = []

        # 文件处理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)

        # 设置日志格式
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s',
                                      datefmt='%Y-%m-%d %H:%M:%S')
        file_handler.setFormatter(formatter)

        # 添加文件处理器
        self.logger.addHandler(file_handler)

        self.log_file_path = log_file
        self.logger.info(f"日志系统初始化完成，日志文件: {log_file}")

    def create_widgets(self):
        """创建GUI组件 - 优化显示区域"""
        # 创建主容器，减少内边距以充分利用空间
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # 配置主容器的网格权重
        self.main_container.columnconfigure(0, weight=1)
        self.main_container.rowconfigure(0, weight=1)

        # 创建带滚动条的Canvas（初始隐藏滚动条）
        self.canvas = tk.Canvas(self.main_container, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 添加垂直滚动条（初始隐藏）
        self.v_scrollbar = ttk.Scrollbar(self.main_container, orient=tk.VERTICAL, command=self.canvas.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.v_scrollbar.grid_remove()  # 初始隐藏

        # 添加水平滚动条（初始隐藏）
        self.h_scrollbar = ttk.Scrollbar(self.main_container, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.h_scrollbar.grid_remove()  # 初始隐藏

        # 配置Canvas的滚动
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        # 创建内部框架（实际内容容器）
        self.content_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor=tk.NW)

        # 绑定配置事件以更新滚动区域
        self.content_frame.bind('<Configure>', self.on_content_configure)
        self.canvas.bind('<Configure>', self.on_canvas_configure)

        # 配置内容框架的网格权重 - 优化显示比例
        self.content_frame.columnconfigure(0, weight=3)  # 左侧区域权重
        self.content_frame.columnconfigure(1, weight=2)  # 右侧日志区域权重
        self.content_frame.rowconfigure(2, weight=1)  # 标签页和日志区域可扩展

        # 标题栏
        self.create_title_bar()

        # 创建标签页区域
        self.create_tabs_area()

        # 创建日志区域
        self.create_log_area()

        # 绑定鼠标滚轮事件
        self.bind_mouse_scroll()

    def on_content_configure(self, event):
        """内容框架大小变化时更新Canvas滚动区域"""
        # 更新Canvas的滚动区域
        bbox = self.canvas.bbox("all")
        self.canvas.configure(scrollregion=bbox)

        # 动态显示/隐藏滚动条
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # 如果需要垂直滚动条
        if bbox[3] > canvas_height:
            self.v_scrollbar.grid()
        else:
            self.v_scrollbar.grid_remove()

        # 如果需要水平滚动条
        if bbox[2] > canvas_width:
            self.h_scrollbar.grid()
        else:
            self.h_scrollbar.grid_remove()

    def on_canvas_configure(self, event):
        """Canvas大小变化时调整内容框架宽度"""
        # 设置内容框架的宽度与Canvas相同
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def bind_mouse_scroll(self):
        """绑定鼠标滚轮事件"""
        # 在Canvas上绑定鼠标滚轮
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind_all("<Button-4>", self.on_mousewheel)  # Linux向上滚动
        self.canvas.bind_all("<Button-5>", self.on_mousewheel)  # Linux向下滚动

        # 在内容框架上绑定鼠标滚轮
        self.content_frame.bind_all("<MouseWheel>", self.on_mousewheel)

    def on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        # Windows和macOS
        if event.delta:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # Linux
        elif event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def create_title_bar(self):
        """创建标题栏 - 简洁的居中方案"""
        title_frame = ttk.Frame(self.content_frame)
        title_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(2, 5))

        # 直接使用pack居中
        title_label = ttk.Label(title_frame, text="气象水文数据分析系统",
                                font=('微软雅黑', 18, 'bold'), foreground='#2c3e50')
        title_label.pack()

        version_label = ttk.Label(title_frame, text="v1.0.0",
                                  font=('微软雅黑', 9), foreground='#7f8c8d')
        version_label.pack()

        # 分隔线
        separator = ttk.Separator(self.content_frame, orient='horizontal')
        separator.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))

    def create_tabs_area(self):
        """创建标签页区域 - 充分利用垂直空间"""
        tabs_frame = ttk.Frame(self.content_frame)
        tabs_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tabs_frame.columnconfigure(0, weight=1)
        tabs_frame.rowconfigure(0, weight=1)

        # 创建Notebook
        self.notebook = ttk.Notebook(tabs_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 创建各个标签页
        self.create_data_reporting_tab()
        self.create_abnormal_analysis_tab()
        self.create_rainfall_analysis_tab()
        self.create_weather_analysis_tab()
        self.create_deviation_analysis_tab()

        # 绑定标签页切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def create_data_reporting_tab(self):
        """创建到报率分析标签页 - 优化布局"""
        tab = ttk.Frame(self.notebook, padding="8")
        self.notebook.add(tab, text="数据到报率分析")

        # 标签页内容
        content_frame = ttk.Frame(tab)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 使用Frame和自动布局，只在需要时添加滚动条
        main_content = ttk.Frame(content_frame)
        main_content.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # 标题
        ttk.Label(main_content, text="数据到报率分析",
                  font=('微软雅黑', 13, 'bold')).grid(row=0, column=0, columnspan=3,
                                                      pady=(0, 10), sticky=tk.W)

        # 输入目录选择
        ttk.Label(main_content, text="选择输入目录:").grid(row=1, column=0, sticky=tk.W,
                                                           padx=5, pady=6)
        self.data_reporting_input_var = tk.StringVar()
        ttk.Entry(main_content, textvariable=self.data_reporting_input_var,
                  width=45).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_data_reporting_input,
                   width=10).grid(row=1, column=2, padx=5, pady=6)

        # 输出文件选择
        ttk.Label(main_content, text="输出文件:").grid(row=2, column=0, sticky=tk.W,
                                                       padx=5, pady=6)
        self.data_reporting_output_var = tk.StringVar(value="./分析结果/到报率计算结果/数据到报率分析结果.xlsx")
        ttk.Entry(main_content, textvariable=self.data_reporting_output_var,
                  width=45).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_data_reporting_output,
                   width=10).grid(row=2, column=2, padx=5, pady=6)

        # 新增：频率模式选择
        mode_frame = ttk.Frame(main_content)
        mode_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5, 2))

        self.freq_mode_var = tk.StringVar(value="auto")  # 默认自动判断
        ttk.Radiobutton(mode_frame, text="根据文件名自动判断频率",
                        variable=self.freq_mode_var, value="auto",
                        command=self.toggle_freq_selection).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="手动设置统一频率",
                        variable=self.freq_mode_var, value="manual",
                        command=self.toggle_freq_selection).pack(side=tk.LEFT, padx=5)

        # 上传频率选择（原代码位置略有调整，现在行号增加）
        ttk.Label(main_content, text="上传频率:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=6)

        # 创建频率选择框架，保存为实例属性以便控制状态
        self.freq_frame = ttk.Frame(main_content)
        self.freq_frame.grid(row=4, column=1, sticky=tk.W, padx=5, pady=6, columnspan=2)

        # 预定义频率
        self.data_reporting_freq_var = tk.StringVar(value="1分钟")
        predefined_freqs = ["1分钟", "5分钟", "10分钟", "30分钟", "60分钟"]
        self.freq_radio_btns = []
        for i, freq in enumerate(predefined_freqs):
            rb = ttk.Radiobutton(self.freq_frame, text=freq,
                                 variable=self.data_reporting_freq_var, value=freq)
            rb.grid(row=0, column=i, padx=2)
            self.freq_radio_btns.append(rb)

        # 自定义频率框架
        self.custom_frame = ttk.Frame(main_content)
        self.custom_frame.grid(row=5, column=1, sticky=tk.W, padx=5, pady=6, columnspan=2)

        self.custom_rb = ttk.Radiobutton(self.custom_frame, text="其他:",
                                          variable=self.data_reporting_freq_var, value="custom")
        self.custom_rb.grid(row=0, column=0, padx=2)

        self.data_reporting_custom_freq_var = tk.StringVar(value="15")
        self.custom_entry = ttk.Entry(self.custom_frame,
                                       textvariable=self.data_reporting_custom_freq_var, width=8)
        self.custom_entry.grid(row=0, column=1, padx=2)
        self.custom_label = ttk.Label(self.custom_frame, text="分钟")
        self.custom_label.grid(row=0, column=2, padx=2)

        # 初始状态：根据默认模式设置频率控件可用性
        self.toggle_freq_selection()

        # 新增：公共缺失分析选项
        self.data_reporting_common_missing_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_content, text="分析公共缺失时间段（需要至少2个设备）",
                        variable=self.data_reporting_common_missing_var).grid(row=6, column=1, columnspan=2,
                                                                              sticky=tk.W, padx=5, pady=6)

        # 分析按钮框架
        analyze_frame = ttk.Frame(main_content)
        analyze_frame.grid(row=7, column=0, columnspan=3, pady=12)

        self.data_reporting_start_btn = ttk.Button(analyze_frame, text="开始分析",
                                                   command=self.start_data_reporting_analysis,
                                                   width=12)
        self.data_reporting_start_btn.pack(side=tk.LEFT, padx=5)

        self.data_reporting_stop_btn = ttk.Button(analyze_frame, text="停止分析",
                                                  command=self.stop_analysis,
                                                  width=12, state=tk.DISABLED)
        self.data_reporting_stop_btn.pack(side=tk.LEFT, padx=5)

        # 说明文字 - 使用更紧凑的布局
        info_frame = ttk.LabelFrame(main_content, text="功能说明", padding="6")
        info_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))

        info_text = """数据到报率分析功能用于检查设备数据上传的完整性和及时性。

    主要功能：
    1. 分析指定文件夹下的所有Excel文件（需包含"数据"工作表）
    2. 计算数据到报率（实报数据条数/应报数据条数）
    3. 识别数据缺失时间段
    4. 检测重复数据记录
    5. 支持不同上传频率（1分钟、5分钟、10分钟、30分钟、60分钟等）
    6. 新增：分析多个设备间的公共缺失时间段

    新增功能说明：
    • 公共缺失时间段分析：识别多个设备在同一时间段内都出现数据缺失的情况
    • 需要至少2个设备才能进行公共缺失分析
    • 公共缺失时间段会输出到单独的Excel工作表中
    • 显示每个公共缺失时间段影响的设备列表

    频率模式：
    • 自动判断：根据文件名中的关键字自动识别频率（状态文件10分钟，SKY2文件5分钟，其他1分钟）
    • 手动设置：所有文件使用用户指定的统一频率，公共缺失分析也使用此频率

    使用方法：
    1. 选择包含Excel文件的输入目录
    2. 设置输出文件路径（默认输出到"./到报率计算结果"文件夹）
    3. 选择频率模式
    4. 选择是否分析公共缺失时间段
    5. 点击"开始分析"按钮"""
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT, font=('微软雅黑', 9)).pack(anchor=tk.W)

        # 配置网格列权重，使Entry控件可以扩展
        main_content.columnconfigure(1, weight=1)

    def toggle_freq_selection(self):
        """根据频率模式启用/禁用频率选择控件"""
        mode = self.freq_mode_var.get()
        if mode == "manual":
            # 手动模式：启用所有频率控件
            state = tk.NORMAL
        else:
            # 自动模式：禁用所有频率控件，但仍显示，公共缺失分析仍使用此频率
            state = tk.DISABLED

        # 设置频率选择框架内所有子控件的状态
        for child in self.freq_frame.winfo_children():
            try:
                child.config(state=state)
            except:
                pass
        for child in self.custom_frame.winfo_children():
            try:
                child.config(state=state)
            except:
                pass
        # 单选按钮本身可能不在上述框架中，单独处理
        for rb in self.freq_radio_btns:
            rb.config(state=state)
        self.custom_rb.config(state=state)
        self.custom_entry.config(state=state)
        self.custom_label.config(state=state)  # 标签也可禁用，但通常不需修改

    # 其余标签页创建方法保持不变（略）
    def create_abnormal_analysis_tab(self):
        """创建异常数据分析标签页"""
        tab = ttk.Frame(self.notebook, padding="8")
        self.notebook.add(tab, text="异常数据分析")

        # 标签页内容
        content_frame = ttk.Frame(tab)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 使用Frame和自动布局
        main_content = ttk.Frame(content_frame)
        main_content.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # 标题
        ttk.Label(main_content, text="异常数据分析",
                  font=('微软雅黑', 13, 'bold')).grid(row=0, column=0, columnspan=3,
                                                      pady=(0, 10), sticky=tk.W)

        # 输入目录选择
        ttk.Label(main_content, text="选择输入目录:").grid(row=1, column=0, sticky=tk.W,
                                                           padx=5, pady=6)
        self.abnormal_input_var = tk.StringVar()
        ttk.Entry(main_content, textvariable=self.abnormal_input_var,
                  width=45).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_abnormal_input,
                   width=10).grid(row=1, column=2, padx=5, pady=6)

        # 输出文件选择
        ttk.Label(main_content, text="输出文件:").grid(row=2, column=0, sticky=tk.W,
                                                       padx=5, pady=6)
        self.abnormal_output_var = tk.StringVar(value="./分析结果/到报率计算结果/异常数据分析结果.xlsx")
        ttk.Entry(main_content, textvariable=self.abnormal_output_var,
                  width=45).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_abnormal_output,
                   width=10).grid(row=2, column=2, padx=5, pady=6)

        # 异常类型选择
        ttk.Label(main_content, text="检测异常类型:").grid(row=3, column=0, sticky=tk.W,
                                                           padx=5, pady=6)

        types_frame = ttk.Frame(main_content)
        types_frame.grid(row=3, column=1, sticky=tk.W, padx=5, pady=6, columnspan=2)

        self.abnormal_type_9999_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(types_frame, text="9999/9998异常值",
                        variable=self.abnormal_type_9999_var).grid(row=0, column=0, sticky=tk.W, padx=2)

        self.abnormal_type_rainfall_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(types_frame, text="降雨量递减",
                        variable=self.abnormal_type_rainfall_var).grid(row=0, column=1, sticky=tk.W, padx=2)

        self.abnormal_type_flood_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(types_frame, text="水浸状态异常",
                        variable=self.abnormal_type_flood_var).grid(row=0, column=2, sticky=tk.W, padx=2)

        # 分析按钮框架
        analyze_frame = ttk.Frame(main_content)
        analyze_frame.grid(row=4, column=0, columnspan=3, pady=12)

        self.abnormal_start_btn = ttk.Button(analyze_frame, text="开始分析",
                                             command=self.start_abnormal_analysis,
                                             width=12)
        self.abnormal_start_btn.pack(side=tk.LEFT, padx=5)

        self.abnormal_stop_btn = ttk.Button(analyze_frame, text="停止分析",
                                            command=self.stop_analysis,
                                            width=12, state=tk.DISABLED)
        self.abnormal_stop_btn.pack(side=tk.LEFT, padx=5)

        # 说明文字
        info_frame = ttk.LabelFrame(main_content, text="功能说明", padding="6")
        info_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))

        info_text = """异常数据分析功能用于检测数据中的异常情况。

主要功能：
1. 检测数值型字段中的9999/9998异常值
2. 检测降雨量/雨量数据递减异常
3. 检测水浸状态非0/1异常
4. 自动合并连续异常时间段
5. 计算异常持续时间
6. 生成异常统计报告

使用方法：
1. 选择包含Excel文件的输入目录
2. 设置输出文件路径（默认输出到"./到报率计算结果"文件夹）
3. 选择需要检测的异常类型
4. 点击"开始分析"按钮"""
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT, font=('微软雅黑', 9)).pack(anchor=tk.W)

        # 配置网格列权重
        main_content.columnconfigure(1, weight=1)

    def create_rainfall_analysis_tab(self):
        """创建降雨数据分析标签页"""
        tab = ttk.Frame(self.notebook, padding="8")
        self.notebook.add(tab, text="降雨数据分析")

        # 标签页内容
        content_frame = ttk.Frame(tab)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 使用Frame和自动布局
        main_content = ttk.Frame(content_frame)
        main_content.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # 标题
        ttk.Label(main_content, text="降雨数据分析",
                  font=('微软雅黑', 13, 'bold')).grid(row=0, column=0, columnspan=3,
                                                      pady=(0, 10), sticky=tk.W)

        # 输入目录选择
        ttk.Label(main_content, text="选择输入目录:").grid(row=1, column=0, sticky=tk.W,
                                                           padx=5, pady=6)
        self.rainfall_input_var = tk.StringVar()
        ttk.Entry(main_content, textvariable=self.rainfall_input_var,
                  width=45).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_rainfall_input,
                   width=10).grid(row=1, column=2, padx=5, pady=6)

        # 输出文件选择
        ttk.Label(main_content, text="输出文件:").grid(row=2, column=0, sticky=tk.W,
                                                       padx=5, pady=6)
        self.rainfall_output_var = tk.StringVar(value="./分析结果/降雨数据分析结果/降雨数据分析结果.xlsx")
        ttk.Entry(main_content, textvariable=self.rainfall_output_var,
                  width=45).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_rainfall_output,
                   width=10).grid(row=2, column=2, padx=5, pady=6)

        # 计算选项
        calc_frame = ttk.Frame(main_content)
        calc_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, padx=5, pady=6)

        ttk.Label(calc_frame, text="计算类型:").grid(row=0, column=0, sticky=tk.W, padx=5)

        self.rainfall_calc_type_var = tk.StringVar(value="全部")
        ttk.Radiobutton(calc_frame, text="分钟降雨强度",
                        variable=self.rainfall_calc_type_var,
                        value="分钟").grid(row=0, column=1, padx=2)
        ttk.Radiobutton(calc_frame, text="小时降雨量",
                        variable=self.rainfall_calc_type_var,
                        value="小时").grid(row=0, column=2, padx=2)
        ttk.Radiobutton(calc_frame, text="日降雨量",
                        variable=self.rainfall_calc_type_var,
                        value="日").grid(row=0, column=3, padx=2)
        ttk.Radiobutton(calc_frame, text="全部",
                        variable=self.rainfall_calc_type_var,
                        value="全部").grid(row=0, column=4, padx=2)

        # 分析按钮框架
        analyze_frame = ttk.Frame(main_content)
        analyze_frame.grid(row=4, column=0, columnspan=3, pady=12)

        self.rainfall_start_btn = ttk.Button(analyze_frame, text="开始分析",
                                             command=self.start_rainfall_analysis,
                                             width=12)
        self.rainfall_start_btn.pack(side=tk.LEFT, padx=5)

        self.rainfall_stop_btn = ttk.Button(analyze_frame, text="停止分析",
                                            command=self.stop_analysis,
                                            width=12, state=tk.DISABLED)
        self.rainfall_stop_btn.pack(side=tk.LEFT, padx=5)

        # 说明文字
        info_frame = ttk.LabelFrame(main_content, text="功能说明", padding="6")
        info_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))

        info_text = """降雨数据分析功能用于分析降雨数据的统计特征。

主要功能：
1. 计算累计降雨总量
2. 统计降雨天数
3. 分析最大日降雨量、小时降雨量和分钟降雨强度
4. 生成小时降雨量横向排列表格
5. 自动标记异常值（9998、9999）
6. 生成降雨数据统计报告

使用方法：
1. 选择包含降雨数据Excel文件的输入目录
2. 设置输出文件路径（默认输出到"./降雨数据分析结果"文件夹）
3. 选择需要计算的降雨强度类型
4. 点击"开始分析"按钮"""
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT, font=('微软雅黑', 9)).pack(anchor=tk.W)

        # 配置网格列权重
        main_content.columnconfigure(1, weight=1)

    def create_weather_analysis_tab(self):
        """创建气象数据分析标签页"""
        tab = ttk.Frame(self.notebook, padding="8")
        self.notebook.add(tab, text="气象数据分析")

        # 标签页内容
        content_frame = ttk.Frame(tab)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 使用Frame和自动布局
        main_content = ttk.Frame(content_frame)
        main_content.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # 标题
        ttk.Label(main_content, text="气象数据分析",
                  font=('微软雅黑', 13, 'bold')).grid(row=0, column=0, columnspan=3,
                                                      pady=(0, 10), sticky=tk.W)

        # 输入目录选择
        ttk.Label(main_content, text="选择输入目录:").grid(row=1, column=0, sticky=tk.W,
                                                           padx=5, pady=6)
        self.weather_input_var = tk.StringVar()
        ttk.Entry(main_content, textvariable=self.weather_input_var,
                  width=45).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_weather_input,
                   width=10).grid(row=1, column=2, padx=5, pady=6)

        # 输出文件选择
        ttk.Label(main_content, text="输出文件:").grid(row=2, column=0, sticky=tk.W,
                                                       padx=5, pady=6)
        self.weather_output_var = tk.StringVar(value="./分析结果/气象数据分析结果/气象数据分析结果.xlsx")
        ttk.Entry(main_content, textvariable=self.weather_output_var,
                  width=45).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_weather_output,
                   width=10).grid(row=2, column=2, padx=5, pady=6)

        # 图表目录选择
        ttk.Label(main_content, text="图表保存目录:").grid(row=3, column=0, sticky=tk.W,
                                                           padx=5, pady=6)
        self.weather_image_var = tk.StringVar(value="./分析结果/气象数据分析结果/图表")
        ttk.Entry(main_content, textvariable=self.weather_image_var,
                  width=45).grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_weather_image,
                   width=10).grid(row=3, column=2, padx=5, pady=6)

        # 分析参数选择
        ttk.Label(main_content, text="分析参数:").grid(row=4, column=0, sticky=tk.W,
                                                       padx=5, pady=6)

        params_frame = ttk.Frame(main_content)
        params_frame.grid(row=4, column=1, sticky=tk.W, padx=5, pady=6, columnspan=2)

        self.weather_param_temp_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="温度",
                        variable=self.weather_param_temp_var).grid(row=0, column=0,
                                                                   sticky=tk.W, padx=2)

        self.weather_param_humidity_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="湿度",
                        variable=self.weather_param_humidity_var).grid(row=0, column=1,
                                                                       sticky=tk.W, padx=2)

        self.weather_param_pressure_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="气压",
                        variable=self.weather_param_pressure_var).grid(row=0, column=2,
                                                                       sticky=tk.W, padx=2)

        self.weather_param_wind_speed_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="风速",
                        variable=self.weather_param_wind_speed_var).grid(row=0, column=3,
                                                                         sticky=tk.W, padx=2)

        self.weather_param_wind_dir_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="风向",
                        variable=self.weather_param_wind_dir_var).grid(row=0, column=4,
                                                                       sticky=tk.W, padx=2)

        # 分析按钮框架
        analyze_frame = ttk.Frame(main_content)
        analyze_frame.grid(row=5, column=0, columnspan=3, pady=12)

        self.weather_start_btn = ttk.Button(analyze_frame, text="开始分析",
                                            command=self.start_weather_analysis,
                                            width=12)
        self.weather_start_btn.pack(side=tk.LEFT, padx=5)

        self.weather_stop_btn = ttk.Button(analyze_frame, text="停止分析",
                                           command=self.stop_analysis,
                                           width=12, state=tk.DISABLED)
        self.weather_stop_btn.pack(side=tk.LEFT, padx=5)

        # 说明文字
        info_frame = ttk.LabelFrame(main_content, text="功能说明", padding="6")
        info_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))

        info_text = """气象数据分析功能用于分析多设备气象数据。

主要功能：
1. 计算各设备的日平均温度、湿度、气压
2. 生成多设备对比趋势图
3. 异常数据处理（9998、9999自动排除）
4. 数据频率统计
5. 生成风速、风向对比图表
6. 标记异常数据超过100条的日期

使用方法：
1. 选择包含气象数据Excel文件的输入目录
2. 设置输出文件路径（默认输出到"./气象数据分析结果"文件夹）
3. 设置图表保存目录
4. 选择需要分析的气象参数
5. 点击"开始分析"按钮"""
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT, font=('微软雅黑', 9)).pack(anchor=tk.W)

        # 配置网格列权重
        main_content.columnconfigure(1, weight=1)

    def create_deviation_analysis_tab(self):
        """创建设备偏差分析标签页"""
        tab = ttk.Frame(self.notebook, padding="8")
        self.notebook.add(tab, text="雨量设备偏差分析")

        # 标签页内容
        content_frame = ttk.Frame(tab)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 使用Frame和自动布局
        main_content = ttk.Frame(content_frame)
        main_content.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # 标题
        ttk.Label(main_content, text="雨量设备偏差分析",
                  font=('微软雅黑', 13, 'bold')).grid(row=0, column=0, columnspan=3,
                                                      pady=(0, 10), sticky=tk.W)

        # 输入文件选择
        ttk.Label(main_content, text="选择输入文件:").grid(row=1, column=0, sticky=tk.W,
                                                           padx=5, pady=6)
        self.deviation_input_var = tk.StringVar()
        ttk.Entry(main_content, textvariable=self.deviation_input_var,
                  width=45).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_deviation_input,
                   width=10).grid(row=1, column=2, padx=5, pady=6)

        # 输出目录选择
        ttk.Label(main_content, text="输出目录:").grid(row=2, column=0, sticky=tk.W,
                                                       padx=5, pady=6)
        self.deviation_output_var = tk.StringVar(value="./分析结果/降雨数据分析结果/设备偏差分析")
        ttk.Entry(main_content, textvariable=self.deviation_output_var,
                  width=45).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=6)
        ttk.Button(main_content, text="浏览", command=self.browse_deviation_output,
                   width=10).grid(row=2, column=2, padx=5, pady=6)

        # 标准设备输入
        ttk.Label(main_content, text="标准设备名称:").grid(row=3, column=0, sticky=tk.W,
                                                           padx=5, pady=6)
        self.deviation_standard_var = tk.StringVar(value="华云")
        ttk.Entry(main_content, textvariable=self.deviation_standard_var,
                  width=25).grid(row=3, column=1, sticky=tk.W, padx=5, pady=6)

        # 分析按钮框架
        analyze_frame = ttk.Frame(main_content)
        analyze_frame.grid(row=4, column=0, columnspan=3, pady=12)

        self.deviation_start_btn = ttk.Button(analyze_frame, text="开始分析",
                                              command=self.start_deviation_analysis,
                                              width=12)
        self.deviation_start_btn.pack(side=tk.LEFT, padx=5)

        self.deviation_stop_btn = ttk.Button(analyze_frame, text="停止分析",
                                             command=self.stop_analysis,
                                             width=12, state=tk.DISABLED)
        self.deviation_stop_btn.pack(side=tk.LEFT, padx=5)

        # 说明文字
        info_frame = ttk.LabelFrame(main_content, text="功能说明", padding="6")
        info_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))

        info_text = """设备偏差分析功能用于比较不同设备相对于标准设备的计量偏差。

主要功能：
1. 分析各设备与标准设备的一致性
2. 计算相关系数和回归方程
3. 生成散点图、趋势图等可视化图表
4. 创建详细的分析报告
5. 评估设备计量准确度等级
6. 提供设备校准建议

使用方法：
1. 选择包含降雨分析结果的Excel文件
2. 设置输出目录（默认输出到"./降雨数据分析结果"文件夹）
3. 输入标准设备名称
4. 点击"开始分析"按钮

注意：输入文件应为降雨数据分析模块生成的Excel文件，包含各设备的降雨数据。"""
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT, font=('微软雅黑', 9)).pack(anchor=tk.W)

        # 配置网格列权重
        main_content.columnconfigure(1, weight=1)

    def create_log_area(self):
        """创建日志输出区域（右侧）- 充分利用垂直空间"""
        log_frame = ttk.LabelFrame(self.content_frame, text="运行日志", padding="6")
        log_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(8, 0))

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)  # 日志文本框可扩展

        # 日志文本框 - 增大高度
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 9), width=45, height=30)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置文本标签颜色
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("DEBUG", foreground="gray")
        self.log_text.tag_config("WARNING", foreground="orange")
        self.log_text.tag_config("ERROR", foreground="red")
        self.log_text.tag_config("SUCCESS", foreground="green")

        # 日志控制按钮框架
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.grid(row=1, column=0, sticky=tk.E, pady=(6, 0))

        ttk.Button(log_control_frame, text="清空日志", command=self.clear_log,
                   width=10).pack(side=tk.LEFT, padx=3)

        ttk.Button(log_control_frame, text="打开日志", command=self.open_log_file,
                   width=10).pack(side=tk.LEFT, padx=3)

        ttk.Button(log_control_frame, text="保存日志", command=self.save_log,
                   width=10).pack(side=tk.LEFT, padx=3)

    def on_tab_changed(self, event):
        """标签页切换事件"""
        selected_tab = self.notebook.index(self.notebook.select())
        tab_names = {
            0: "数据到报率分析",
            1: "异常数据分析",
            2: "降雨数据分析",
            3: "气象数据分析",
            4: "设备偏差分析"
        }

        current_tab_name = tab_names.get(selected_tab, "未知")
        self.log_info(f"切换到功能: {current_tab_name}")

    # 文件浏览方法（保持原样，仅修改日志显示）
    def browse_data_reporting_input(self):
        directory = filedialog.askdirectory(title="选择到报率分析输入目录")
        if directory:
            self.data_reporting_input_var.set(directory)
            self.log_info(f"已选择到报率分析输入目录: {directory}")

    def browse_data_reporting_output(self):
        filename = filedialog.asksaveasfilename(
            title="选择到报率分析输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir="./到报率计算结果",
            initialfile="数据到报率分析结果.xlsx"
        )
        if filename:
            self.data_reporting_output_var.set(filename)
            self.log_info(f"已设置到报率分析输出文件: {filename}")

    def browse_abnormal_input(self):
        directory = filedialog.askdirectory(title="选择异常分析输入目录")
        if directory:
            self.abnormal_input_var.set(directory)
            self.log_info(f"已选择异常分析输入目录: {directory}")

    def browse_abnormal_output(self):
        filename = filedialog.asksaveasfilename(
            title="选择异常分析输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir="./到报率计算结果",
            initialfile="异常数据分析结果.xlsx"
        )
        if filename:
            self.abnormal_output_var.set(filename)
            self.log_info(f"已设置异常分析输出文件: {filename}")

    def browse_rainfall_input(self):
        directory = filedialog.askdirectory(title="选择降雨分析输入目录")
        if directory:
            self.rainfall_input_var.set(directory)
            self.log_info(f"已选择降雨分析输入目录: {directory}")

    def browse_rainfall_output(self):
        filename = filedialog.asksaveasfilename(
            title="选择降雨分析输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir="./降雨数据分析结果",
            initialfile="降雨数据分析结果.xlsx"
        )
        if filename:
            self.rainfall_output_var.set(filename)
            self.log_info(f"已设置降雨分析输出文件: {filename}")

    def browse_weather_input(self):
        directory = filedialog.askdirectory(title="选择气象分析输入目录")
        if directory:
            self.weather_input_var.set(directory)
            self.log_info(f"已选择气象分析输入目录: {directory}")

    def browse_weather_output(self):
        filename = filedialog.asksaveasfilename(
            title="选择气象分析输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir="./气象数据分析结果",
            initialfile="气象数据分析结果.xlsx"
        )
        if filename:
            self.weather_output_var.set(filename)
            self.log_info(f"已设置气象分析输出文件: {filename}")

    def browse_weather_image(self):
        directory = filedialog.askdirectory(title="选择气象分析图表目录")
        if directory:
            self.weather_image_var.set(directory)
            self.log_info(f"已选择气象分析图表目录: {directory}")

    def browse_deviation_input(self):
        filename = filedialog.askopenfilename(
            title="选择设备偏差分析输入文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
            initialdir="./降雨数据分析结果"
        )
        if filename:
            self.deviation_input_var.set(filename)
            self.log_info(f"已选择设备偏差分析输入文件: {filename}")

    def browse_deviation_output(self):
        directory = filedialog.askdirectory(title="选择设备偏差分析输出目录")
        if directory:
            self.deviation_output_var.set(directory)
            self.log_info(f"已选择设备偏差分析输出目录: {directory}")

    # 分析功能方法
    def start_data_reporting_analysis(self):
        """开始到报率分析"""
        if self.is_running:
            messagebox.showwarning("警告", "已有分析任务正在运行")
            return

        # 检查输入目录
        input_dir = self.data_reporting_input_var.get()
        if not input_dir or not os.path.exists(input_dir):
            messagebox.showerror("错误", "请选择有效的输入目录")
            return

        # 检查输出文件路径
        output_file = self.data_reporting_output_var.get()
        if not output_file:
            messagebox.showerror("错误", "请设置输出文件路径")
            return

        # 获取频率模式
        freq_mode = self.freq_mode_var.get()

        # 获取频率值（无论哪种模式，都获取用户选择的数值，用于公共缺失分析）
        freq_str = self.data_reporting_freq_var.get()
        if freq_str == "custom":
            try:
                freq_min = int(self.data_reporting_custom_freq_var.get())
            except ValueError:
                messagebox.showerror("错误", "自定义频率必须是整数")
                return
        else:
            freq_min = int(freq_str.replace("分钟", ""))

        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_info(f"创建输出目录: {output_dir}")

        # 创建配置
        config = {
            "name": f"到报率分析_{freq_min}分钟",
            "input_dir": input_dir,
            "freq_min": freq_min,
            "freq_mode": freq_mode,          # 新增
            "output_file": output_file
        }

        self.is_running = True
        self.current_task = "data_reporting"
        self.update_button_state(True)
        self.log_info("=" * 60)
        self.log_info(f"开始数据到报率分析")
        self.log_info(f"输入目录: {input_dir}")
        self.log_info(f"输出文件: {output_file}")
        self.log_info(f"频率模式: {'自动判断' if freq_mode == 'auto' else '手动设置'}")
        self.log_info(f"基准频率: {freq_min}分钟")
        self.log_info("=" * 60)

        # 在新线程中运行分析
        self.task_thread = threading.Thread(
            target=self.run_data_reporting_analysis,
            args=([config],)
        )
        self.task_thread.daemon = True
        self.task_thread.start()

    def run_data_reporting_analysis(self, configs):
        """运行到报率分析"""
        try:
            results = batch_analyze_by_config(configs, self.log_progress)

            # 显示结果摘要
            success_count = sum(1 for r in results if r["status"] == "success")
            error_count = sum(1 for r in results if r["status"] == "error")
            warning_count = sum(1 for r in results if r["status"] == "warning")

            self.log_info("\n" + "=" * 60)
            self.log_info("数据到报率分析完成!")
            self.log_info(f"成功: {success_count}, 错误: {error_count}, 警告: {warning_count}")

            if error_count == 0:
                self.log_success("所有文件处理成功!")
            else:
                self.log_warning(f"有{error_count}个文件处理失败，请检查日志")

            # 显示完成提示
            self.show_completion_message("数据到报率分析")

        except Exception as e:
            self.log_error(f"分析过程中出现错误: {str(e)}")
            self.show_error_message("数据到报率分析", str(e))
        finally:
            self.analysis_complete()

    def start_abnormal_analysis(self):
        """开始异常数据分析"""
        if self.is_running:
            messagebox.showwarning("警告", "已有分析任务正在运行")
            return

        input_dir = self.abnormal_input_var.get()
        if not input_dir or not os.path.exists(input_dir):
            messagebox.showerror("错误", "请选择有效的输入目录")
            return

        output_file = self.abnormal_output_var.get()
        if not output_file:
            messagebox.showerror("错误", "请设置输出文件路径")
            return

        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_info(f"创建输出目录: {output_dir}")

        # 创建配置
        config = {
            "name": "异常数据分析",
            "input_dir": input_dir,
            "output_file": output_file
        }

        self.is_running = True
        self.current_task = "abnormal"
        self.update_button_state(True)
        self.log_info("=" * 60)
        self.log_info(f"开始异常数据分析")
        self.log_info(f"输入目录: {input_dir}")
        self.log_info(f"输出文件: {output_file}")
        self.log_info("=" * 60)

        self.task_thread = threading.Thread(
            target=self.run_abnormal_analysis,
            args=([config],)
        )
        self.task_thread.daemon = True
        self.task_thread.start()

    def run_abnormal_analysis(self, configs):
        """运行异常数据分析"""
        try:
            results = batch_analyze_abnormal(configs, self.log_progress)

            success_count = sum(1 for r in results if r["status"] == "success")
            error_count = sum(1 for r in results if r["status"] == "error")
            warning_count = sum(1 for r in results if r["status"] == "warning")

            self.log_info("\n" + "=" * 60)
            self.log_info("异常数据分析完成!")
            self.log_info(f"成功: {success_count}, 错误: {error_count}, 警告: {warning_count}")

            if error_count == 0:
                self.log_success("所有文件处理成功!")

            # 显示完成提示
            self.show_completion_message("异常数据分析")

        except Exception as e:
            self.log_error(f"分析过程中出现错误: {str(e)}")
            self.show_error_message("异常数据分析", str(e))
        finally:
            self.analysis_complete()

    def start_rainfall_analysis(self):
        """开始降雨数据分析"""
        if self.is_running:
            messagebox.showwarning("警告", "已有分析任务正在运行")
            return

        input_dir = self.rainfall_input_var.get()
        if not input_dir or not os.path.exists(input_dir):
            messagebox.showerror("错误", "请选择有效的输入目录")
            return

        output_file = self.rainfall_output_var.get()
        if not output_file:
            messagebox.showerror("错误", "请设置输出文件路径")
            return

        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_info(f"创建输出目录: {output_dir}")

        self.is_running = True
        self.current_task = "rainfall"
        self.update_button_state(True)
        self.log_info("=" * 60)
        self.log_info(f"开始降雨数据分析")
        self.log_info(f"输入目录: {input_dir}")
        self.log_info(f"输出文件: {output_file}")
        self.log_info("=" * 60)

        self.task_thread = threading.Thread(
            target=self.run_rainfall_analysis,
            args=(input_dir, output_file)
        )
        self.task_thread.daemon = True
        self.task_thread.start()

    def run_rainfall_analysis(self, input_dir, output_file):
        """运行降雨数据分析"""
        try:
            results = analyze_rainfall_batch(input_dir, output_file, self.log_progress)

            success_count = sum(1 for r in results if r["status"] == "success")
            error_count = sum(1 for r in results if r["status"] == "error")
            warning_count = sum(1 for r in results if r["status"] == "warning")

            self.log_info("\n" + "=" * 60)
            self.log_info("降雨数据分析完成!")
            self.log_info(f"成功: {success_count}, 错误: {error_count}, 警告: {warning_count}")

            if error_count == 0:
                self.log_success("所有文件处理成功!")

            # 显示完成提示
            self.show_completion_message("降雨数据分析")

        except Exception as e:
            self.log_error(f"分析过程中出现错误: {str(e)}")
            self.show_error_message("降雨数据分析", str(e))
        finally:
            self.analysis_complete()

    def start_weather_analysis(self):
        """开始气象数据分析"""
        if self.is_running:
            messagebox.showwarning("警告", "已有分析任务正在运行")
            return

        input_dir = self.weather_input_var.get()
        if not input_dir or not os.path.exists(input_dir):
            messagebox.showerror("错误", "请选择有效的输入目录")
            return

        output_file = self.weather_output_var.get()
        if not output_file:
            messagebox.showerror("错误", "请设置输出文件路径")
            return

        image_dir = self.weather_image_var.get()
        if not image_dir:
            image_dir = "./气象数据分析结果/图表"

        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_info(f"创建输出目录: {output_dir}")

        # 确保图表目录存在
        if not os.path.exists(image_dir):
            os.makedirs(image_dir)
            self.log_info(f"创建图表目录: {image_dir}")

        self.is_running = True
        self.current_task = "weather"
        self.update_button_state(True)
        self.log_info("=" * 60)
        self.log_info(f"开始气象数据分析")
        self.log_info(f"输入目录: {input_dir}")
        self.log_info(f"输出文件: {output_file}")
        self.log_info(f"图表目录: {image_dir}")
        self.log_info("=" * 60)

        self.task_thread = threading.Thread(
            target=self.run_weather_analysis,
            args=(input_dir, output_file, image_dir)
        )
        self.task_thread.daemon = True
        self.task_thread.start()

    def run_weather_analysis(self, input_dir, output_file, image_dir):
        """运行气象数据分析"""
        try:
            results = process_multiple_devices(input_dir, output_file, image_dir, self.log_progress)

            success_count = sum(1 for r in results if r["status"] == "success")
            error_count = sum(1 for r in results if r["status"] == "error")
            warning_count = sum(1 for r in results if r["status"] == "warning")

            self.log_info("\n" + "=" * 60)
            self.log_info("气象数据分析完成!")
            self.log_info(f"成功: {success_count}, 错误: {error_count}, 警告: {warning_count}")

            if error_count == 0:
                self.log_success("所有文件处理成功!")

            # 显示完成提示
            self.show_completion_message("气象数据分析")

        except Exception as e:
            self.log_error(f"分析过程中出现错误: {str(e)}")
            self.show_error_message("气象数据分析", str(e))
        finally:
            self.analysis_complete()

    def start_deviation_analysis(self):
        """开始设备偏差分析"""
        if self.is_running:
            messagebox.showwarning("警告", "已有分析任务正在运行")
            return

        input_file = self.deviation_input_var.get()
        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("错误", "请选择有效的输入文件")
            return

        output_dir = self.deviation_output_var.get()
        if not output_dir:
            output_dir = "./降雨数据分析结果/设备偏差分析"

        standard_device = self.deviation_standard_var.get()
        if not standard_device:
            messagebox.showerror("错误", "请输入标准设备名称")
            return

        # 确保输出目录存在
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_info(f"创建输出目录: {output_dir}")

        self.is_running = True
        self.current_task = "deviation"
        self.update_button_state(True)
        self.log_info("=" * 60)
        self.log_info(f"开始设备偏差分析")
        self.log_info(f"输入文件: {input_file}")
        self.log_info(f"输出目录: {output_dir}")
        self.log_info(f"标准设备: {standard_device}")
        self.log_info("=" * 60)

        self.task_thread = threading.Thread(
            target=self.run_deviation_analysis,
            args=(input_file, output_dir, standard_device)
        )
        self.task_thread.daemon = True
        self.task_thread.start()

    def run_deviation_analysis(self, input_file, output_dir, standard_device):
        """运行设备偏差分析"""
        try:
            # 读取数据
            data_dict, read_results = read_rainfall_data(input_file)

            for result in read_results:
                if result["status"] == "success":
                    self.log_info(result["message"])
                elif result["status"] == "warning":
                    self.log_warning(result["message"])
                elif result["status"] == "error":
                    self.log_error(result["message"])

            if data_dict is None:
                self.log_error("无法读取数据，分析终止")
                return

            # 分析数据
            analysis_results = analyze_rainfall_devices(
                data_dict, standard_device, output_dir, self.log_progress
            )

            for result in analysis_results:
                if result["status"] == "success":
                    self.log_info(result["message"])
                elif result["status"] == "warning":
                    self.log_warning(result["message"])
                elif result["status"] == "error":
                    self.log_error(result["message"])

            self.log_info("\n" + "=" * 60)
            self.log_info("设备偏差分析完成!")

            # 检查是否有错误
            error_count = sum(1 for r in analysis_results if r["status"] == "error")
            if error_count == 0:
                self.log_success("分析处理成功!")

            # 显示完成提示
            self.show_completion_message("设备偏差分析")

        except Exception as e:
            self.log_error(f"分析过程中出现错误: {str(e)}")
            self.show_error_message("设备偏差分析", str(e))
        finally:
            self.analysis_complete()

    def show_completion_message(self, task_name):
        """显示分析完成消息"""
        self.root.after(100, lambda: messagebox.showinfo(
            "分析完成",
            f"{task_name}已完成！\n\n"
            f"分析结果已保存到相应目录。\n"
            f"详细日志请查看运行日志文件夹。"
        ))

    def show_error_message(self, task_name, error_msg):
        """显示分析错误消息"""
        self.root.after(100, lambda: messagebox.showerror(
            "分析错误",
            f"{task_name}过程中出现错误：\n\n{error_msg}\n\n"
            f"请查看运行日志获取详细信息。"
        ))

    def stop_analysis(self):
        """停止分析"""
        if self.is_running:
            self.is_running = False
            self.log_warning("分析正在停止，请等待...")

    def analysis_complete(self):
        """分析完成后的清理工作"""
        self.is_running = False
        self.update_button_state(False)
        self.log_info("分析任务已完成")

    def update_button_state(self, is_running):
        """更新按钮状态"""
        state = tk.DISABLED if is_running else tk.NORMAL

        # 更新各个标签页的开始按钮
        self.data_reporting_start_btn.config(state=state)
        self.abnormal_start_btn.config(state=state)
        self.rainfall_start_btn.config(state=state)
        self.weather_start_btn.config(state=state)
        self.deviation_start_btn.config(state=state)

        # 更新各个标签页的停止按钮
        stop_state = tk.NORMAL if is_running else tk.DISABLED
        self.data_reporting_stop_btn.config(state=stop_state)
        self.abnormal_stop_btn.config(state=stop_state)
        self.rainfall_stop_btn.config(state=stop_state)
        self.weather_stop_btn.config(state=stop_state)
        self.deviation_stop_btn.config(state=stop_state)

    def log_progress(self, message):
        """记录进度日志"""
        self.log_info(message)

    def log_info(self, message):
        """记录信息日志"""
        # 记录到日志文件
        self.logger.info(message)
        # 更新GUI显示
        self._update_log_gui(message, "INFO")

    def log_warning(self, message):
        """记录警告日志"""
        # 记录到日志文件
        self.logger.warning(message)
        # 更新GUI显示
        self._update_log_gui(f"警告: {message}", "WARNING")

    def log_error(self, message):
        """记录错误日志"""
        # 记录到日志文件
        self.logger.error(message)
        # 更新GUI显示
        self._update_log_gui(f"错误: {message}", "ERROR")

    def log_success(self, message):
        """记录成功日志"""
        # 记录到日志文件
        self.logger.info(message)
        # 更新GUI显示
        self._update_log_gui(f"成功: {message}", "SUCCESS")

    def _update_log_gui(self, message, level="INFO"):
        """更新GUI日志显示"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        log_entry = f"[{timestamp}] {message}\n"

        # 在主线程中更新GUI
        self.root.after(0, self._update_log_text, log_entry, level)

    def _update_log_text(self, message, level):
        """更新日志文本框"""
        self.log_text.insert(tk.END, message, level)
        self.log_text.see(tk.END)
        self.log_text.update()

    def clear_log(self):
        """清空日志显示"""
        self.log_text.delete(1.0, tk.END)
        self.log_info("日志显示已清空")

    def open_log_file(self):
        """打开日志文件"""
        try:
            import subprocess
            import os

            if os.path.exists(self.log_file_path):
                if os.name == 'nt':  # Windows
                    os.startfile(self.log_file_path)
                elif os.name == 'posix':  # Linux or Mac
                    subprocess.call(['xdg-open', self.log_file_path])
                self.log_info(f"已打开日志文件: {self.log_file_path}")
            else:
                self.log_warning("日志文件不存在")
        except Exception as e:
            self.log_error(f"无法打开日志文件: {str(e)}")

    def save_log(self):
        """保存当前日志显示"""
        filename = filedialog.asksaveasfilename(
            title="保存日志内容",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )

        if filename:
            try:
                log_content = self.log_text.get(1.0, tk.END)
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                self.log_info(f"日志内容已保存到: {filename}")
            except Exception as e:
                self.log_error(f"保存日志失败: {str(e)}")

    def on_closing(self):
        """关闭窗口时的处理"""
        if self.is_running:
            if messagebox.askyesno("确认", "分析任务正在运行，确定要退出吗？"):
                # 取消绑定鼠标滚轮事件
                self.canvas.unbind_all("<MouseWheel>")
                self.canvas.unbind_all("<Button-4>")
                self.canvas.unbind_all("<Button-5>")
                self.root.destroy()
        else:
            # 取消绑定鼠标滚轮事件
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
            self.root.destroy()