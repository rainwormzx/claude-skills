#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""直接运行测试模式"""

from update_device_location import DeviceLocationUpdater

if __name__ == "__main__":
    print("浙江大学设备存放地批量更新脚本 - 测试模式")
    print("=" * 50)
    print("正在运行测试模式（处理前3条记录）...")
    print("=" * 50)

    updater = DeviceLocationUpdater()

    try:
        # 运行测试模式：只处理前3条
        updater.run(start_index=0, end_index=3)
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
    except Exception as e:
        print(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()
    finally:
        updater.close()
