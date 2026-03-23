"""
气象水文数据分析系统 v1.0.0
"""

import tkinter as tk
from gui import WeatherHydroAnalysisApp
import os
import sys


def check_dependencies():
    """检查必要的依赖包"""
    required_packages = [
        'pandas',
        'openpyxl',
        'matplotlib',
        'numpy'
    ]

    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    return missing_packages


def create_directory_structure():
    """创建必要的目录结构"""
    directories = [
        "运行日志",
        "分析结果"
    ]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"创建目录: {directory}")


def main():
    """主函数"""
    print("=" * 60)
    print("气象水文数据分析系统 v1.0.0")
    print("=" * 60)

    # 检查依赖
    missing_packages = check_dependencies()
    if missing_packages:
        print("缺少必要的Python包:")
        for package in missing_packages:
            print(f"  - {package}")
        print("\n请使用以下命令安装:")
        print(f"pip install {' '.join(missing_packages)}")

        # 询问是否自动安装
        try:
            input_str = input("\n是否自动安装缺失的包？(y/n): ")
            if input_str.lower() == 'y':
                import subprocess
                for package in missing_packages:
                    print(f"正在安装 {package}...")
                    try:
                        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                        print(f"  {package} 安装成功")
                    except subprocess.CalledProcessError:
                        print(f"  {package} 安装失败")
                print("\n依赖包安装完成！")
            else:
                return
        except Exception as e:
            print(f"安装依赖包时出错: {e}")
            return

    # 创建目录结构
    create_directory_structure()

    # 启动GUI
    try:
        root = tk.Tk()

        # 设置窗口图标（如果有）
        try:
            root.iconbitmap("icon.ico")
        except:
            pass

        app = WeatherHydroAnalysisApp(root)

        # 启动消息循环
        print("启动GUI界面...")
        print("日志文件将保存在 '运行日志' 文件夹中")
        print("分析结果将保存在 '分析结果' 文件夹中")
        print("-" * 60)

        root.mainloop()

    except Exception as e:
        print(f"启动GUI时出错: {e}")
        print("请确保已正确安装所有依赖包。")
        input("按Enter键退出...")


if __name__ == "__main__":
    main()