# build_exe.py
import PyInstaller.__main__
import os
import shutil
import sys


def build_exe():
    """打包气象水文数据分析系统为单个exe文件"""

    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 检查icon文件
    icon_path = os.path.join(current_dir, "icon.ico")
    if not os.path.exists(icon_path):
        # 如果没有icon.ico，可以创建一个简单的替代品或提示
        print("警告: 未找到 icon.ico 文件，将使用默认图标")
        icon_path = None

    # PyInstaller配置参数
    params = [
        'main.py',  # 主程序文件
        '--name=气象水文数据分析系统',  # 生成的exe文件名
        '--onefile',  # 打包为单个exe文件
        '--windowed',  # 隐藏控制台窗口
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不询问确认
        '--distpath=./dist',  # 输出目录
        '--workpath=./build',  # 临时工作目录
        '--specpath=./',  # spec文件位置
    ]

    # 添加图标（如果存在）
    if icon_path and os.path.exists(icon_path):
        params.append(f'--icon={icon_path}')

    # 添加数据文件（非Python文件）
    # 添加可能需要的配置文件
    if os.path.exists("config.txt"):
        params.append('--add-data=config.txt;.')

    # 添加隐藏导入（解决模块导入问题）
    hidden_imports = [
        'pandas',
        'numpy',
        'openpyxl',
        'matplotlib',
        'matplotlib.pyplot',
        'tkinter',
        'tkinter.ttk',
        'scipy',
        'scipy.stats',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'logging',
        'threading',
        'datetime',
        'glob',
        'warnings',
        'json',
        'io',
        'sys',
        'os',
        'time',
    ]

    for module in hidden_imports:
        params.append(f'--hidden-import={module}')

    # 运行PyInstaller
    print("开始打包...")
    print(f"参数: {params}")
    PyInstaller.__main__.run(params)

    # 复制必要的依赖文件到dist目录
    dist_dir = os.path.join(current_dir, "dist")
    exe_path = os.path.join(dist_dir, "气象水文数据分析系统.exe")

    if os.path.exists(exe_path):
        print(f"\n✅ 打包完成!")
        print(f"exe文件位置: {exe_path}")
        print(f"文件大小: {os.path.getsize(exe_path) / (1024 * 1024):.2f} MB")

        # 创建说明文件
        readme_path = os.path.join(dist_dir, "使用说明.txt")
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write("气象水文数据分析系统 - 使用说明\n")
            f.write("=" * 60 + "\n\n")
            f.write("1. 系统要求\n")
            f.write("   - Windows 7/8/10/11 操作系统\n")
            f.write("   - 需要.NET Framework 4.5或更高版本\n")
            f.write("   - 至少4GB可用内存\n")
            f.write("   - 建议屏幕分辨率：1366×768或更高\n\n")
            f.write("2. 使用方法\n")
            f.write("   - 双击\"气象水文数据分析系统.exe\"即可运行\n")
            f.write("   - 不需要安装Python或其他依赖库\n")
            f.write("   - 首次运行可能会稍慢，请耐心等待\n\n")
            f.write("3. 注意事项\n")
            f.write("   - 请确保有足够的磁盘空间存放分析结果\n")
            f.write("   - 分析大量数据时可能需要较长时间\n")
            f.write("   - 如果遇到杀毒软件误报，请添加信任\n\n")
            f.write("4. 文件夹结构\n")
            f.write("   - 运行日志: 保存程序运行日志\n")
            f.write("   - 分析结果: 保存各种分析结果\n")
            f.write("   - 数据文件: 存放待分析的Excel文件\n\n")
            f.write("=" * 60 + "\n")

        print(f"使用说明已保存到: {readme_path}")

        # 提供复制命令
        print("\n📋 复制exe文件到其他位置：")
        print(f"只需复制这个文件即可：\n{exe_path}")

    else:
        print("❌ 打包失败，exe文件未生成")


if __name__ == "__main__":
    build_exe()