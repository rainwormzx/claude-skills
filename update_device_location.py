#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
浙江大学设备存放地批量更新脚本
用途：从Excel文件读取设备信息，批量更新内网资产管理系统的"学院存放地"字段
"""

import json
import time
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException


# ==================== 配置区域 ====================

# Excel文件路径
EXCEL_FILE = "存放地测试.xlsx"

# 资产管理系统URL（需要用户填写）
BASE_URL = "https://pxxt.zju.edu.cn"  # TODO: 修改为实际的系统URL

# ChromeDriver路径（留空则使用系统PATH中的chromedriver）
CHROME_DRIVER_PATH = ""

# 是否显示浏览器窗口（调试时设为True，正式运行可设为False）
HEADLESS = False

# 是否需要手动登录（True=等待用户手动登录，False=使用Cookie）
MANUAL_LOGIN = True

# 每次操作后的等待时间（秒），可根据网络情况调整
WAIT_TIME = 3
PAGE_LOAD_TIMEOUT = 30

# ==================== Cookie配置 ====================
# TODO: 用户需要从浏览器中复制Cookie并更新此配置
# 获取Cookie方法：
# 1. 登录系统后，按F12打开开发者工具
# 2. 切换到 Application/存储 -> Cookies
# 3. 复制相关Cookie的 name 和 value
COOKIES_CONFIG = [
    {
        'name': 'iPlanetDirectoryPro',
        'value': '',  # TODO: 填写实际值（登录token）
        'domain': '.zju.edu.cn',
        'path': '/'
    },
    {
        'name': 'JSESSIONID',
        'value': '',  # TODO: 填写实际值（会话ID）
        'domain': 'pxxt.zju.edu.cn',
        'path': '/'
    },
    # TODO: 根据实际情况添加其他必要的Cookie
]

# ==================== 元素定位配置 ====================
# 使用浏览器F12开发者工具获取的元素定位
ELEMENT_LOCATORS = {
    # 搜索框 - 输入资产编号
    'search_input': {
        'by': By.XPATH,
        'value': '//*[@id="mc"]'
    },

    # 搜索按钮
    'search_button': {
        'by': By.XPATH,
        'value': '//*[@id="query_id"]'
    },

    # 编辑/详情按钮 - 点击进入设备编辑页面
    'edit_button': {
        'by': By.XPATH,
        'value': '//*[@id="PrintA"]/tbody/tr/td[3]/a[1]/i'
    },

    # 学院存放地输入框
    'location_input': {
        'by': By.XPATH,
        'value': '/html/body/div[3]/div/div[2]/form/div[19]/div/input'
    },

    # 保存按钮
    'save_button': {
        'by': By.XPATH,
        'value': '//*[@id="submitForm"]/div[20]/button'
    },

    # 成功提示元素（用于判断保存是否成功）
    'success_message': {
        'by': By.XPATH,
        'value': '//div[contains(text(),"保存成功")]'
    },

    # 管理员资产管理入口按钮
    'admin_asset_management': {
        'by': By.XPATH,
        'value': '//*[@id="content-wrapper"]/div/div[2]/div/div/div[1]/a/div/div'
    },
}

# ==================== Excel列名配置 ====================
# 根据Excel文件的实际列名配置
COLUMN_NAMES = {
    'asset_number': '资产编号',      # 资产编号列名
    'new_location': '学院新存放地',   # 新存放地列名（要更新的值）
}

# ==================== 日志配置 ====================
LOG_FILE = f"update_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


# ==================== 类定义 ====================

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

            # 设置页面加载超时（在driver创建后设置）
            self.driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)

            self.wait = WebDriverWait(self.driver, PAGE_LOAD_TIMEOUT)
            logger.info("Chrome浏览器启动成功")
            return True
        except Exception as e:
            logger.error(f"Chrome浏览器启动失败: {e}")
            logger.error("请检查ChromeDriver是否已正确安装")
            return False

    def load_cookies(self):
        """加载Cookie到浏览器"""
        try:
            # 先访问域名以设置cookie domain
            self.driver.get(BASE_URL)
            time.sleep(1)

            for cookie in COOKIES_CONFIG:
                if cookie['value']:  # 只添加有值的cookie
                    self.driver.add_cookie(cookie)
                    logger.debug(f"添加Cookie: {cookie['name']}")

            logger.info("Cookie加载完成")
            return True
        except Exception as e:
            logger.error(f"Cookie加载失败: {e}")
            return False

    def read_excel(self):
        """读取Excel文件数据"""
        try:
            self.data_df = pd.read_excel(EXCEL_FILE)

            # 过滤有效数据
            asset_col = COLUMN_NAMES['asset_number']
            location_col = COLUMN_NAMES['new_location']

            # 去除空值
            self.data_df = self.data_df.dropna(subset=[asset_col, location_col])

            total_records = len(self.data_df)
            logger.info(f"成功读取Excel文件，共 {total_records} 条有效记录")
            logger.info(f"列名: {list(self.data_df.columns)}")

            return True
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
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

            # 检查是否有弹窗/iframe
            try:
                # 尝试切换到iframe（很多弹窗使用iframe）
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    logger.debug(f"发现 {len(iframes)} 个iframe，尝试切换...")
                    for i, iframe in enumerate(iframes):
                        try:
                            self.driver.switch_to.frame(iframe)
                            time.sleep(0.5)
                            # 检查是否能找到存放地输入框
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

            # 滚动弹窗内容
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
            time.sleep(1)

            # 3. 修改学院存放地
            location_input = self.find_element('location_input')
            if not location_input:
                return False

            # 先点击激活输入框
            location_input.click()
            time.sleep(0.5)

            # 清空并输入新值
            location_input.clear()
            time.sleep(0.5)
            location_input.send_keys(new_location)

            # 使用JavaScript确保值被设置（针对特殊输入框）
            self.driver.execute_script("arguments[0].value = arguments[1];", location_input, new_location)

            # 触发change事件（某些系统需要）
            self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", location_input)

            logger.debug(f"已设置新的存放地: {new_location}")

            time.sleep(WAIT_TIME)

            # 4. 点击保存按钮
            save_button = self.find_element('save_button')
            if not save_button:
                return False
            save_button.click()
            logger.debug("已点击保存按钮")

            # 5. 等待保存成功
            time.sleep(WAIT_TIME * 2)

            # 处理可能的弹窗
            try:
                alert = self.driver.switch_to.alert
                alert.accept()
                logger.debug("已接受弹窗")
                time.sleep(1)
            except:
                pass

            # 尝试查找成功提示（可选）
            try:
                success_elem = self.find_element('success_message')
                if success_elem:
                    logger.info(f"资产编号 {asset_number} 更新成功")
            except:
                pass  # 如果没有找到成功提示，也不算失败

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

        # 如果是手动登录模式，等待用户登录
        if MANUAL_LOGIN:
            logger.info("=" * 50)
            logger.info("手动登录模式")
            logger.info("=" * 50)
            logger.info("请在浏览器中完成以下操作：")
            logger.info("  1. 登录系统")
            logger.info(f"脚本将在 {WAIT_TIME * 10} 秒后继续执行...")
            logger.info("=" * 50)

            # 打开系统页面
            self.driver.get(BASE_URL)

            # 等待固定时间让用户登录
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
        else:
            # 加载Cookie
            if not self.load_cookies():
                logger.error("Cookie加载失败，可能需要重新登录")
                return False

        # 读取数据
        if not self.read_excel():
            return False

        # 设置处理范围
        if end_index is None:
            end_index = len(self.data_df)

        records_to_process = self.data_df.iloc[start_index:end_index]
        total = len(records_to_process)

        logger.info(f"准备处理第 {start_index + 1} 到第 {end_index} 条记录，共 {total} 条")

        # 如果不是手动登录模式，访问系统页面
        if not MANUAL_LOGIN:
            try:
                self.driver.get(BASE_URL)
                time.sleep(WAIT_TIME * 2)

                # 点击"管理员资产管理"按钮
                logger.info("正在点击'管理员资产管理'按钮...")
                admin_button = self.find_element('admin_asset_management')
                if admin_button:
                    admin_button.click()
                    logger.info("已点击'管理员资产管理'")
                    time.sleep(WAIT_TIME * 2)
                else:
                    logger.warning("未找到'管理员资产管理'按钮，可能已经在该页面")

            except Exception as e:
                logger.error(f"访问系统页面失败: {e}")
                return False

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

            # 等待一段时间再处理下一条
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

        # 保持浏览器打开一段时间供用户查看
        logger.info(f"\n浏览器将在10秒后关闭...")
        time.sleep(10)

        return True

    def close(self):
        """关闭浏览器"""
        if self.driver:
            self.driver.quit()
            logger.info("浏览器已关闭")


# ==================== 测试元素定位 ====================

class ElementLocatorTester:
    """元素定位测试工具 - 用于帮助用户找到正确的元素定位方式"""

    def __init__(self):
        self.driver = None

    def test_locators(self):
        """测试元素定位是否正确"""
        print("元素定位测试工具")
        print("=" * 50)
        print("请在浏览器中手动登录系统后，运行此测试")

        if not self.init_driver():
            return

        # 让用户手动登录
        print("\n请手动登录系统，登录完成后按回车继续...")
        input()

        # 测试每个元素定位
        for element_name, locator in ELEMENT_LOCATORS.items():
            print(f"\n测试元素: {element_name}")
            print(f"  定位方式: {locator['by']}")
            print(f"  定位值: {locator['value']}")

            try:
                element = self.driver.find_element(locator['by'], locator['value'])
                print(f"  ✓ 找到元素: {element.tag_name}")
            except NoSuchElementException:
                print(f"  ✗ 未找到元素，请检查定位方式")

        self.driver.quit()

    def init_driver(self):
        """初始化Chrome浏览器驱动"""
        try:
            chrome_options = Options()
            chrome_options.add_argument('--window-size=1920,1080')
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.get(BASE_URL)
            return True
        except Exception as e:
            print(f"浏览器启动失败: {e}")
            return False


# ==================== 配置导出工具 ====================

def export_config_template():
    """导出配置模板到JSON文件，方便用户修改"""
    config = {
        "base_url": BASE_URL,
        "cookies": COOKIES_CONFIG,
        "element_locators": {
            k: {
                'by': str(v['by']),
                'value': v['value']
            }
            for k, v in ELEMENT_LOCATORS.items()
        },
        "column_names": COLUMN_NAMES
    }

    with open('config_template.json', 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    print("配置模板已导出到: config_template.json")
    print("你可以修改此文件后，使用 load_config_from_file() 加载配置")


def load_config_from_file(filename='config_template.json'):
    """从JSON文件加载配置"""
    global BASE_URL, COOKIES_CONFIG, ELEMENT_LOCATORS, COLUMN_NAMES

    try:
        with open(filename, 'r', encoding='utf-8') as f:
            config = json.load(f)

        BASE_URL = config['base_url']
        COOKIES_CONFIG = config['cookies']

        # 转换字符串为By对象
        by_mapping = {
            'By.ID': By.ID,
            'By.NAME': By.NAME,
            'By.CLASS_NAME': By.CLASS_NAME,
            'By.XPATH': By.XPATH,
            'By.CSS_SELECTOR': By.CSS_SELECTOR,
            'By.TAG_NAME': By.TAG_NAME,
        }

        for name, locator in config['element_locators'].items():
            ELEMENT_LOCATORS[name] = {
                'by': by_mapping.get(locator['by'], By.ID),
                'value': locator['value']
            }

        COLUMN_NAMES = config['column_names']

        print(f"配置已从 {filename} 加载")
        return True
    except Exception as e:
        print(f"加载配置文件失败: {e}")
        return False


# ==================== 主程序 ====================

def main():
    """主函数"""
    print("浙江大学设备存放地批量更新脚本")
    print("=" * 50)

    # 检查Excel文件是否存在
    if not Path(EXCEL_FILE).exists():
        print(f"错误: 找不到Excel文件 '{EXCEL_FILE}'")
        print(f"请确保文件在当前目录下")
        return

    # 创建更新器实例
    updater = DeviceLocationUpdater()

    try:
        # 询问用户处理范围
        print("\n请选择运行模式:")
        print("1. 测试模式（只处理前3条记录）")
        print("2. 批量处理（处理所有记录）")
        print("3. 自定义范围")
        print("4. 测试元素定位（用于调试）")

        choice = input("请输入选项 (1-4): ").strip()

        if choice == '1':
            # 测试模式：只处理前3条
            updater.run(start_index=0, end_index=3)
        elif choice == '2':
            # 批量处理所有记录
            confirm = input(f"确认要处理所有记录？输入 'yes' 继续: ").strip().lower()
            if confirm == 'yes':
                updater.run()
            else:
                print("已取消")
        elif choice == '3':
            # 自定义范围
            try:
                start = int(input("起始索引(从0开始): "))
                end = int(input("结束索引(不包含): "))
                updater.run(start_index=start, end_index=end)
            except ValueError:
                print("输入的索引无效")
        elif choice == '4':
            # 元素定位测试
            tester = ElementLocatorTester()
            tester.test_locators()
        else:
            print("无效的选项")

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
