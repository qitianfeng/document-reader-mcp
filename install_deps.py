#!/usr/bin/env python3
"""
安装文档阅读器所需的依赖
"""

import subprocess
import sys

def run_command(command, description):
    """运行命令并显示结果"""
    print(f"🔧 {description}...")
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"✅ {description} 成功")
            return True
        else:
            print(f"❌ {description} 失败:")
            print(result.stderr)
            return False
    except Exception as e:
        print(f"❌ {description} 出错: {str(e)}")
        return False

def install_packages():
    """安装Python包"""
    packages = [
        "mcp",
        "python-docx",
        "PyPDF2", 
        "striprtf",
        "Pillow",
        "requests",
        "opencv-python",
        "numpy"
    ]
    
    for package in packages:
        run_command(f"pip install {package}", f"安装 {package}")

def main():
    print("🚀 安装文档阅读器依赖...")
    install_packages()
    print("🎉 安装完成!")
    print("\n💡 使用方法:")
    print("1. 启动MCP服务器: python server.py")
    print("2. 测试图表分析: python simple_diagram_reader.py")

if __name__ == "__main__":
    main()