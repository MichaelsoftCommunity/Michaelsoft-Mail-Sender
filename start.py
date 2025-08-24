#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Michaelsoft Mail Toolbox - 启动脚本
版本：v 0.0.0
"""

import os
import sys
import subprocess

def main():
    """启动邮件发送器"""
    print("正在启动 Michaelsoft Mail Sender...")
    
    # 获取当前脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 构建主程序路径
    main_script = os.path.join(script_dir, "mail_sender.py")
    
    # 检查主程序是否存在
    if not os.path.exists(main_script):
        print("错误: 找不到主程序文件 mail_sender.py")
        print(f"请确保该文件位于 {script_dir} 目录中")
        input("按回车键退出...")
        sys.exit(1)
    
    try:
        # 启动主程序
        subprocess.run([sys.executable, main_script], check=True)
    except subprocess.CalledProcessError:
        print("程序异常退出")
        input("按回车键退出...")
    except KeyboardInterrupt:
        print("程序被用户中断")
    except Exception as e:
        print(f"启动失败: {str(e)}")
        input("按回车键退出...")

if __name__ == "__main__":
    main()