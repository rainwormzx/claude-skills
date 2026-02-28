#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SKILL: 批量更新设备存放地
用途：从Excel文件读取设备信息，批量更新内网资产管理系统的"学院存放地"字段
"""

import json
import time
import logging
from datetime import datetime
from pathlib import Path
import sys

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException


# ==================== 配置区域 ====================

EXCEL_FILE = "存放地测试.xlsx"  # Excel文件路径

BASE_URL = "https://pxxt.zju.edu.cn"  # 系统URL

HEADLESS = False  # 是否显示浏览器窗口
MANUAL_LOGIN = True  # 是否需要手动登录
WAIT_TIME = 3  # 每次操作后的等待时间（秒）
PAGE_LOAD_TIMEOUT = 30  # 页面加载超时时间（秒）

# 元素定位配置
ELEMENT_LOCATORS = {
    'search_input': {'by': By.XPATH, 'value': '//*[@id="mc"]'},
    'search_button': {'by': By.XPATH, 'value': '//*[@id="query_id"]'},
    'edit_button': {'by': By.XPATH, 'value': '//*[@id="PrintA"]/tbody/tr/td[3]/a[1]/i'},
    'location_input': {'by': By.XPATH, 'value': '/html/body/div[3]/div/div[2]/form/div[19]/div/input'},
    'save_button': {'by': By.XPATH, 'value': '//*[@id="submitForm"]/div[20]/button'},
    'admin_asset_management': {'by': By.XPATH, 'value': '//*[@id="content-wrapper"]/div/div[2]/div/div/div[1]/a/div/div'},
    'success_message': {'by': By.XPATH, 'value': '//div[contains(text(),"保存成功")]'},
}

# Excel列名配置
COLUMN_NAMES = {
    'asset_number': '资产编号',
    'new_location': '学院新存放地',
}

# 日志配置
LOG_FILE = f"update_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class DeviceLocationUpdater:
    """设备存放地批量更新器"""

    def __init__(self):
        self.driver = None
        self.wait = None
        self.data_df = None

    def init_driver(self):
        """初始化Chrome浏览器驱动"""
        try:
            chrome_options = Options()
            if HEADLESS:
                chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')

            if CHROME_DRIVER_PATH:
                self.driver = webdriver.Chrome(executable_path=CHROME_DRIVER_PATH, options=chrome_options)
            else:
                self.driver = webdriver.Chrome(options=chrome_options)

            self.driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
            self.wait = WebDriverWait(self.driver, PAGE_LOAD_TIMEOUT)
            logger.info("Chrome浏览器启动成功")
            return True
        except Exception as e:
            logger.error(f"Chrome浏览器启动失败: {e}")
            return False

    def find_element(self, element_name):
        """查找页面元素"""
        try:
            locator = ELEMENT_LOCATORS[element_name]
            element = self.wait.until(
                EC.presence_of_element_located((locator['by'], locator['value']))
            )
            # 滚动到元素可见
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.5)
            return element
        except TimeoutException:
            logger.error(f"找不到元素: {element_name}")
            logger.error(f"定位方式: {locator['by']} = {locator['value']}")
            return None
        except Exception as e:
            logger.error(f"查找元素 {element_name} 时出错: {e}")
            return None

    def read_excel(self):
        """读取Excel文件数据"""
        try:
            self.data_df = pd.read_excel(EXCEL_FILE)
            asset_col = COLUMN_NAMES['asset_number']
            location_col = COLUMN_NAMES['new_location']
            self.data_df = self.data_df.dropna(subset=[asset_col, location_col])
            total_records = len(self.data_df)
            logger.info(f"成功读取Excel文件，共 {total_records} 条有效记录")
            logger.info(f"列名: {list(self.data_df.columns)}")
            return True
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
            return False

    def update_device_location(self, asset_number, new_location):
        """更新单条设备的存放地"""
        try:
            logger.info(f"开始处理资产编号: {asset_number}")

            # 1. 输入资产编号并搜索
            search_input = self.find_element('search_input')
            if not search_input:
                return False
            search_input.clear()
            search_input.send_keys(asset_number)
            logger.debug(f"已输入资产编号: {asset_number}")
            time.sleep(WAIT_TIME)

            search_button = self.find_element('search_button')
            if not search_button:
                return False
            search_button.click()
            logger.debug("已点击搜索按钮")
            time.sleep(WAIT_TIME * 2)

            # 2. 点击编辑按钮
            edit_button = self.find_element('edit_button')
            if not edit_button:
                return False
            edit_button.click()
            logger.debug("已点击编辑按钮")
            time.sleep(WAIT_TIME * 2)

            # 处理弹窗/iframe
            try:
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    logger.debug(f"发现 {len(iframes)} 个iframe，尝试切换...")
                    for i, iframe in enumerate(iframes):
                        try:
                            self.driver.switch_to.frame(iframe)
                            time.sleep(0.5)
                            test_locator = ELEMENT_LOCATORS['location_input']
                            test_elem = self.driver.find_elements(test_locator['by'], test_locator['value'])
                            if test_elem:
                                logger.debug(f"在第{i+1}个iframe中找到存放地输入框")
                                break
                            else:
                                self.driver.switch_to.default_content()
                        except:
                            self.driver.switch_to.default_content()
            except:
                pass

            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
            time.sleep(1)

            # 3. 修改学院存放地
            location_input = self.find_element('location_input')
            if not location_input:
                return False

            location_input.click()
            time.sleep(0.5)
            location_input.clear()
            time.sleep(0.5)
            location_input.send_keys(new_location)

            # 使用JavaScript确保值被设置
            self.driver.execute_script("arguments[0].value = arguments[1];", location_input, new_location)
            self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", location_input)

            logger.debug(f"已设置新的存放地: {new_location}")
            time.sleep(WAIT_TIME)

            # 4. 点击保存按钮
            save_button = self.find_element('save_button')
            if not save_button:
                return False
            save_button.click()
            logger.debug("已点击保存按钮")
            time.sleep(WAIT_TIME * 2)

            # 5. 处理弹窗
            try:
                alert = self.driver.switch_to.alert
                alert.accept()
                logger.debug("已接受弹窗")
                time.sleep(1)
            except:
                pass

            logger.info(f"资产编号 {asset_number} 更新完成")
            return True

        except Exception as e:
            logger.error(f"更新资产编号 {asset_number} 时出错: {e}")
            return False

    def run(self, start_index=0, end_index=None):
        """执行批量更新"""
        logger.info("=" * 50)
        logger.info("开始批量更新设备存放地")
        logger.info("=" * 50)

        # 初始化浏览器
        if not self.init_driver():
            return False

        # 手动登录模式
        if MANUAL_LOGIN:
            logger.info("=" * 50)
            logger.info("手动登录模式")
            logger.info("=" * 50)
            logger.info("请在浏览器中完成以下操作：")
            logger.info("  1. 登录系统")
            logger.info(f"脚本将在 {WAIT_TIME * 10} 秒后继续执行...")
            logger.info("=" * 50)

            self.driver.get(BASE_URL)
            time.sleep(WAIT_TIME * 10)

            logger.info("继续执行自动化操作...")

            # 自动点击"管理员资产管理"按钮
            logger.info("正在点击'管理员资产管理'按钮...")
            admin_button = self.find_element('admin_asset_management')
            if admin_button:
                admin_button.click()
                logger.info("已点击'管理员资产管理'")
                time.sleep(WAIT_TIME * 2)
            else:
                logger.warning("未找到'管理员资产管理'按钮，请手动点击")
                time.sleep(WAIT_TIME * 5)

        # 读取数据
        if not self.read_excel():
            return False

        # 设置处理范围
        if end_index is None:
            end_index = len(self.data_df)

        records_to_process = self.data_df.iloc[start_index:end_index]
        total = len(records_to_process)

        logger.info(f"准备处理第 {start_index + 1} 到第 {end_index} 条记录，共 {total} 条")

        # 统计结果
        success_count = 0
        failed_count = 0
        failed_records = []

        # 遍历处理每条记录
        for idx, row in records_to_process.iterrows():
            asset_number = str(row[COLUMN_NAMES['asset_number']]).strip()
            new_location = str(row[COLUMN_NAMES['new_location']]).strip()

            logger.info(f"\n[{idx - start_index + 1}/{total}] 处理资产: {asset_number}")

            if self.update_device_location(asset_number, new_location):
                success_count += 1
            else:
                failed_count += 1
                failed_records.append({
                    'index': idx - start_index + 1,
                    'asset_number': asset_number,
                    'location': new_location
                })

            time.sleep(WAIT_TIME)

        # 输出统计结果
        logger.info("\n" + "=" * 50)
        logger.info("批量更新完成")
        logger.info("=" * 50)
        logger.info(f"总计: {total} 条")
        logger.info(f"成功: {success_count} 条")
        logger.info(f"失败: {failed_count} 条")

        if failed_records:
            logger.info("\n失败记录列表:")
            for record in failed_records:
                logger.info(f"  [{record['index']}] {record['asset_number']} -> {record['location']}")

        logger.info(f"\n日志文件: {LOG_FILE}")

        return True

    def close(self):
        """关闭浏览器"""
        if self.driver:
            self.driver.quit()
            logger.info("浏览器已关闭")


# CHROME_DRIVER_PATH留空使用系统PATH中的chromedriver
CHROME_DRIVER_PATH = ""


def main():
    """主函数"""
    print("=" * 60)
    print("SKILL: 批量更新设备存放地")
    print("=" * 60)
    print(f"\n数据文件: {EXCEL_FILE}")
    print(f"日志文件: {LOG_FILE}")
    print(f"处理模式: 手动登录后自动处理")

    updater = DeviceLocationUpdater()

    try:
        updater.run()
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
    except Exception as e:
        logger.error(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()
    finally:
        updater.close()


if __name__ == "__main__":
    main()
