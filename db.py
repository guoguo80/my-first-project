"""
图标修复与打包脚本
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path


def verify_and_fix_icon():
    """验证并修复图标文件"""
    icon_file = "icon.ico"

    print("🔍 检查图标文件...")

    if not os.path.exists(icon_file):
        print(f"❌ 图标文件不存在: {icon_file}")
        return create_default_icon()

    # 检查文件大小
    file_size = os.path.getsize(icon_file)
    if file_size < 1024:  # 小于1KB可能是损坏的文件
        print(f"⚠️  图标文件过小 ({file_size} bytes)，可能已损坏")
        return create_default_icon()

    # 尝试用PIL检查图标
    try:
        from PIL import Image
        img = Image.open(icon_file)

        if img.format != 'ICO':
            print(f"⚠️  文件格式不是ICO: {img.format}")
            # 尝试转换
            return convert_to_ico(icon_file, img)

        print(f"✅ 图标验证通过:")
        print(f"   尺寸: {img.size}")
        print(f"   格式: {img.format}")
        print(f"   大小: {file_size} bytes")

        return icon_file

    except Exception as e:
        print(f"❌ 图标文件损坏: {e}")
        return create_default_icon()


def convert_to_ico(src_file, img):
    """将其他格式转换为ICO"""
    try:
        print(f"🔄 正在转换 {src_file} 为ICO格式...")

        # 创建备份
        backup_file = f"{src_file}.backup"
        shutil.copy2(src_file, backup_file)
        print(f"   已备份原文件: {backup_file}")

        # 保存为ICO
        ico_file = "icon_fixed.ico"
        img.save(ico_file, format='ICO',
                 sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])

        print(f"✅ 已创建新的ICO文件: {ico_file}")
        return ico_file

    except Exception as e:
        print(f"❌ 转换失败: {e}")
        return create_default_icon()


def create_default_icon():
    """创建默认图标"""
    try:
        from PIL import Image, ImageDraw, ImageFont

        print("🔄 正在创建默认图标...")

        # 创建256x256的图像
        img = Image.new('RGBA', (256, 256), color=(70, 130, 180, 255))
        draw = ImageDraw.Draw(img)

        # 背景圆形
        draw.ellipse([40, 40, 216, 216], fill=(30, 144, 255, 255))

        # 云朵形状
        draw.ellipse([70, 90, 186, 150], fill=(255, 255, 255, 230))

        # 雨滴
        for i in range(3):
            x = 100 + i * 30
            draw.line([x, 160, x, 190], fill=(30, 144, 255, 255), width=10)
            draw.ellipse([x - 5, 185, x + 5, 195], fill=(30, 144, 255, 255))

        # 文字
        try:
            font = ImageFont.truetype("arial.ttf", 40)
        except:
            font = ImageFont.load_default()

        draw.text((128, 128), "W", fill=(255, 255, 255, 255),
                  anchor="mm", font=font)

        # 保存为ICO
        ico_file = "icon_default.ico"
        img.save(ico_file, format='ICO',
                 sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])

        print(f"✅ 已创建默认图标: {ico_file}")
        return ico_file

    except ImportError:
        print("❌ 需要Pillow库来创建图标")
        print("请运行: pip install pillow")
        return None


def clean_build_dirs():
    """清理构建目录"""
    print("\n🧹 清理构建目录...")

    dirs_to_remove = ['build', 'dist']
    files_to_remove = ['气象水文数据分析系统.spec']

    for directory in dirs_to_remove:
        if os.path.exists(directory):
            shutil.rmtree(directory)
            print(f"✅ 删除目录: {directory}")

    for file in files_to_remove:
        if os.path.exists(file):
            os.remove(file)
            print(f"✅ 删除文件: {file}")


def build_with_icon(icon_path):
    """使用图标进行构建"""
    print(f"\n🚀 开始打包，使用图标: {icon_path}")

    # 确保图标路径是绝对路径
    icon_abs_path = os.path.abspath(icon_path)

    # 构建命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--clean",
        f"--name=气象水文数据分析系统",
        f"--icon={icon_abs_path}",
        "--add-data=config.txt;.",
        "--add-data=modules;modules",
        "--hidden-import=pandas",
        "--hidden-import=numpy",
        "--hidden-import=openpyxl",
        "--hidden-import=matplotlib",
        "--hidden-import=matplotlib.backends.backend_tkagg",
        "--exclude-module=test",
        "--exclude-module=unittest",
        "main.py"
    ]

    print(f"📦 打包命令:\n{' '.join(cmd)}")

    # 执行打包
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            print("✅ 打包成功！")

            # 验证输出文件
            exe_path = "dist/气象水文数据分析系统.exe"
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"📁 生成文件: {exe_path}")
                print(f"📊 文件大小: {size_mb:.1f} MB")

                # 验证图标是否嵌入
                verify_exe_icon(exe_path)
                return True
            else:
                print("❌ 打包失败：未生成exe文件")
                return False
        else:
            print(f"❌ 打包失败，错误代码: {result.returncode}")
            print(f"错误输出:\n{result.stderr}")
            return False

    except Exception as e:
        print(f"❌ 打包过程出错: {e}")
        return False


def verify_exe_icon(exe_path):
    """验证exe文件中的图标"""
    print("\n🔍 验证exe文件图标...")

    if not os.path.exists(exe_path):
        print("❌ exe文件不存在")
        return

    # 方法1：检查文件属性
    print("✅ exe文件已生成")
    print(f"   路径: {exe_path}")
    print(f"   大小: {os.path.getsize(exe_path) / 1024 / 1024:.1f} MB")

    # 方法2：尝试提取图标
    try:
        import win32api
        import win32con
        import win32gui

        # 获取exe文件的图标
        large_icons, small_icons = win32gui.ExtractIconEx(exe_path, 0)
        print(f"✅ exe包含图标:")
        print(f"   大图标数量: {len(large_icons)}")
        print(f"   小图标数量: {len(small_icons)}")

        # 清理句柄
        for icon in large_icons + small_icons:
            win32gui.DestroyIcon(icon)

    except ImportError:
        print("ℹ️  如需详细图标信息，请安装pywin32: pip install pywin32")
    except Exception as e:
        print(f"⚠️  无法提取图标信息: {e}")


def main():
    """主函数"""
    print("=" * 60)
    print("气象水文数据分析系统 - 图标修复与打包工具")
    print("=" * 60)

    # 检查必要文件
    if not os.path.exists("main.py"):
        print("❌ 错误: 找不到main.py文件")
        print("请确保在项目根目录运行此脚本")
        input("按Enter键退出...")
        return

    # 步骤1：验证和修复图标
    icon_path = verify_and_fix_icon()
    if not icon_path:
        print("❌ 无法获取有效的图标文件")
        input("按Enter键退出...")
        return

    # 步骤2：清理构建目录
    clean_build_dirs()

    # 步骤3：开始打包
    if build_with_icon(icon_path):
        print("\n" + "=" * 60)
        print("🎉 打包完成！")
        print("\n💡 后续操作:")
        print("1. 右键点击 'dist/气象水文数据分析系统.exe'")
        print("2. 选择 '属性'")
        print("3. 查看 '详细信息' 标签页是否显示图标")
        print("\n⚠️  如果仍不显示图标，请尝试:")
        print("   - 刷新文件资源管理器 (F5)")
        print("   - 重启电脑")
        print("   - 检查图标文件格式是否正确")
        print("=" * 60)
    else:
        print("\n❌ 打包失败")

    input("\n按Enter键退出...")


if __name__ == "__main__":
    main()