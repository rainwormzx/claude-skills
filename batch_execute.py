#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量执行录制的操作
将录制的操作模板与Excel数据结合，批量处理所有记录
"""

import json
import time
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ==================== 配置区域 ====================

# Excel文件路径
EXCEL_FILE = "存放地测试.xlsx"

# 录制的操作JSON文件路径
RECORDED_ACTIONS_FILE = "recorded_actions.json"  # TODO: 修改为实际的录制文件名

# 数据映射：将Excel列名映射到操作中的占位符
DATA_MAPPING = {
    '资产编号': 'ASSET_NUMBER',      # Excel中的"资产编号"列 → 操作中的{{ASSET_NUMBER}}占位符
    '学院存放地': 'NEW_LOCATION',    # Excel中的"学院存放地"列 → 操作中的{{NEW_LOCATION}}占位符
}

# 系统URL
BASE_URL = "https://pxxt.zju.edu.cn"

# Cookie配置（可选，用于免登录）
COOKIES_CONFIG = []

# 每条记录处理后的等待时间（秒）
RECORD_DELAY = 2


# ==================== 日志配置 ====================

LOG_FILE = f"batch_execute_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


# ==================== 批量执行器 ====================

class BatchExecutor:
    """批量执行器 - 使用录制的操作模板处理Excel数据"""

    def __init__(self):
        self.driver = None
        self.wait = None
        self.actions_template = None
        self.data_df = None

    def load_actions(self, actions_file):
        """加载录制的操作模板"""
        try:
            with open(actions_file, 'r', encoding='utf-8') as f:
                self.actions_template = json.load(f)
            logger.info(f"成功加载操作模板: {len(self.actions_template)} 个操作")
            return True
        except Exception as e:
            logger.error(f"加载操作模板失败: {e}")
            return False

    def load_excel(self, excel_file):
        """加载Excel数据"""
        try:
            self.data_df = pd.read_excel(excel_file)

            # 过滤有效数据
            for col in DATA_MAPPING.keys():
                if col in self.data_df.columns:
                    self.data_df = self.data_df.dropna(subset=[col])

            total_records = len(self.data_df)
            logger.info(f"成功读取Excel文件，共 {total_records} 条有效记录")
            logger.info(f"列名: {list(self.data_df.columns)}")

            return True
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
            return False

    def init_driver(self):
        """初始化浏览器"""
        try:
            chrome_options = Options()
            chrome_options.add_argument('--window-size=1920,1080')
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 30)

            logger.info("浏览器启动成功")
            return True
        except Exception as e:
            logger.error(f"浏览器启动失败: {e}")
            return False

    def find_element(self, locator_type, locator_value, description=""):
        """查找页面元素"""
        by_mapping = {
            'id': By.ID,
            'name': By.NAME,
            'class': By.CLASS_NAME,
            'xpath': By.XPATH,
        }

        by = by_mapping.get(locator_type, By.XPATH)

        try:
            element = self.wait.until(
                EC.presence_of_element_located((by, locator_value))
            )
            return element
        except Exception as e:
            logger.error(f"找不到元素: {locator_type}={locator_value} ({description})")
            return None

    def execute_action(self, action, data_dict=None):
        """执行单个操作"""
        action_type = action['type']
        data = action['data']

        try:
            if action_type == 'click':
                # 点击操作
                locator = action.get('locator', {})
                if locator:
                    elem = self.find_element(locator['type'], locator['value'], f"点击操作")
                    if elem:
                        elem.click()
                        logger.debug(f"✓ 点击: {locator['type']}={locator['value']}")
                        return True

            elif action_type == 'input':
                # 输入操作
                value = data.get('value', '')

                # 如果值包含占位符，进行替换
                if data_dict and value:
                    for excel_col, placeholder in DATA_MAPPING.items():
                        if placeholder in value:
                            value = value.replace(f'{{{{{placeholder}}}}}', str(data_dict.get(excel_col, '')))

                locator = action.get('locator', {})
                if locator:
                    elem = self.find_element(locator['type'], locator['value'], f"输入操作")
                    if elem:
                        elem.clear()
                        elem.send_keys(value)
                        logger.debug(f"✓ 输入: {locator['type']}={locator['value']}, 值={value}")
                        return True

            elif action_type == 'navigate':
                # 页面导航（通常由其他操作触发，无需额外处理）
                pass

            return False

        except Exception as e:
            logger.error(f"执行操作出错: {e}")
            return False

    def execute_actions_for_record(self, record_data):
        """为单条记录执行所有操作"""
        success_count = 0

        for action in self.actions_template:
            if self.execute_action(action, record_data):
                success_count += 1

            # 操作间短暂等待
            time.sleep(0.5)

        logger.info(f"该记录执行完成: {success_count}/{len(self.actions_template)} 个操作成功")
        return success_count == len(self.actions_template)

    def run(self, start_index=0, end_index=None, test_mode=False):
        """批量执行"""
        logger.info("=" * 60)
        logger.info("开始批量执行")
        logger.info("=" * 60)

        # 加载操作模板
        if not self.load_actions(RECORDED_ACTIONS_FILE):
            return False

        # 加载Excel数据
        if not self.load_excel(EXCEL_FILE):
            return False

        # 初始化浏览器
        if not self.init_driver():
            return False

        # 设置处理范围
        if end_index is None:
            end_index = len(self.data_df)

        records_to_process = self.data_df.iloc[start_index:end_index]
        total = len(records_to_process)

        if test_mode:
            logger.info(f"【测试模式】只处理前3条记录")
            records_to_process = records_to_process.head(3)
            total = len(records_to_process)

        logger.info(f"准备处理第 {start_index + 1} 到第 {end_index} 条记录，共 {total} 条")

        # 访问起始页面
        try:
            self.driver.get(BASE_URL)
            time.sleep(2)
            logger.info("已打开系统页面，请手动登录（如果需要）")
            input("登录完成后按回车继续...")
        except Exception as e:
            logger.error(f"访问系统页面失败: {e}")
            return False

        # 统计结果
        success_count = 0
        failed_count = 0
        failed_records = []

        # 遍历处理每条记录
        for idx, row in records_to_process.iterrows():
            record_data = row.to_dict()

            # 获取关键信息用于日志
            asset_number = record_data.get('资产编号', 'N/A')
            new_location = record_data.get('学院存放地', 'N/A')

            logger.info(f"\n[{idx - start_index + 1}/{total}] 处理资产: {asset_number} → {new_location}")

            if self.execute_actions_for_record(record_data):
                success_count += 1
            else:
                failed_count += 1
                failed_records.append({
                    'index': idx - start_index + 1,
                    'asset_number': asset_number,
                    'location': new_location
                })

            # 等待一段时间再处理下一条
            time.sleep(RECORD_DELAY)

        # 输出统计结果
        logger.info("\n" + "=" * 60)
        logger.info("批量执行完成")
        logger.info("=" * 60)
        logger.info(f"总计: {total} 条")
        logger.info(f"成功: {success_count} 条")
        logger.info(f"失败: {failed_count} 条")

        if failed_records:
            logger.info("\n失败记录列表:")
            for record in failed_records:
                logger.info(f"  [{record['index']}] {record['asset_number']} -> {record['location']}")

        print("\n按回车键关闭浏览器...")
        input()

        return True

    def close(self):
        """关闭浏览器"""
        if self.driver:
            self.driver.quit()
            logger.info("浏览器已关闭")


# ==================== 主程序 ====================

def main():
    """主函数"""
    print("=" * 60)
    print("批量执行工具 - 使用录制的操作处理Excel数据")
    print("=" * 60)

    # 检查文件是否存在
    if not Path(EXCEL_FILE).exists():
        print(f"错误: 找不到Excel文件 '{EXCEL_FILE}'")
        return

    if not Path(RECORDED_ACTIONS_FILE).exists():
        print(f"错误: 找不到录制文件 '{RECORDED_ACTIONS_FILE}'")
        print(f"请先运行 browser_recorder.py 录制操作")
        return

    executor = BatchExecutor()

    try:
        print("\n请选择运行模式:")
        print("1. 测试模式（只处理前3条记录）")
        print("2. 批量处理（处理所有记录）")
        print("3. 自定义范围")

        choice = input("请输入选项 (1-3): ").strip()

        if choice == '1':
            executor.run(test_mode=True)
        elif choice == '2':
            confirm = input("确认要处理所有记录？输入 'yes' 继续: ").strip().lower()
            if confirm == 'yes':
                executor.run()
            else:
                print("已取消")
        elif choice == '3':
            try:
                start = int(input("起始索引(从0开始): "))
                end = int(input("结束索引(不包含): "))
                executor.run(start_index=start, end_index=end)
            except ValueError:
                print("输入的索引无效")
        else:
            print("无效的选项")

    except KeyboardInterrupt:
        print("\n\n用户中断操作")
    except Exception as e:
        logger.error(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()
    finally:
        executor.close()


if __name__ == "__main__":
    main()
