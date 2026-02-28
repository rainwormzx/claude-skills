#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""快速启动录制工具"""

from browser_recorder import BrowserRecorder

if __name__ == "__main__":
    print("=" * 60)
    print("浏览器操作录制工具")
    print("=" * 60)
    print("\n使用默认URL: https://pxxt.zju.edu.cn")
    print("\n请在浏览器中手动操作一遍流程：")
    print("  1. 登录系统")
    print("  2. 搜索设备")
    print("  3. 点击编辑")
    print("  4. 修改存放地")
    print("  5. 点击保存")
    print("\n操作完成后，按 Ctrl+C 结束录制")
    print("=" * 60 + "\n")

    recorder = BrowserRecorder(url="https://pxxt.zju.edu.cn")

    try:
        recorder.start()
    except KeyboardInterrupt:
        print("\n录制已停止")
        recorder.stop()
