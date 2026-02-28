#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
浏览器操作录制工具
记录用户在浏览器中的操作，生成可自动重放的脚本
"""

import json
import time
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class BrowserRecorder:
    """浏览器操作录制器"""

    def __init__(self, url="https://pxxt.zju.edu.cn"):
        self.url = url
        self.driver = None
        self.actions = []
        self.start_time = None

    def start(self):
        """启动浏览器并开始录制"""
        chrome_options = Options()
        chrome_options.add_argument('--window-size=1920,1080')
        self.driver = webdriver.Chrome(options=chrome_options)

        # 注入监听脚本
        self._inject_listener()

        print(f"\n正在打开 {self.url} ...")
        self.driver.get(self.url)

        self.start_time = time.time()
        print("\n" + "=" * 60)
        print("录制已开始")
        print("=" * 60)
        print("\n现在可以开始操作浏览器，脚本会自动记录以下操作：")
        print("  - 点击元素")
        print("  - 输入文本")
        print("  - 页面导航")
        print("  - 等待时间")
        print("\n操作完成后，在命令行按 Ctrl+C 结束录制")
        print("=" * 60 + "\n")

        # 持续监听操作
        self._listen()

    def _inject_listener(self):
        """注入JavaScript监听脚本"""
        listener_script = """
        // 存储操作历史
        window.recordedActions = [];
        window.lastInputTime = 0;

        // 监听点击事件
        document.addEventListener('click', function(e) {
            var elem = e.target;
            var action = {
                type: 'click',
                timestamp: Date.now(),
                tagName: elem.tagName,
                id: elem.id || '',
                name: elem.name || '',
                className: elem.className || '',
                text: elem.textContent ? elem.textContent.trim().substring(0, 50) : '',
                xpath: getXPath(elem),
                value: elem.value || ''
            };

            // 如果是输入框，记录输入前的值
            if (elem.tagName === 'INPUT' || elem.tagName === 'TEXTAREA') {
                action.inputType = elem.type || 'text';
                action.placeholder = elem.placeholder || '';
            }

            window.recordedActions.push(action);
            console.log('录制: 点击', action);
        }, true);

        // 监听输入事件
        document.addEventListener('input', function(e) {
            var elem = e.target;
            var now = Date.now();

            // 防止重复记录（同一个输入框1秒内只记录一次）
            if (elem.tagName === 'INPUT' || elem.tagName === 'TEXTAREA') {
                var lastInput = window.lastInputElements || {};
                var key = elem.id || elem.name || getXPath(elem);

                if (!lastInput[key] || now - lastInput[key] > 1000) {
                    var action = {
                        type: 'input',
                        timestamp: Date.now(),
                        tagName: elem.tagName,
                        id: elem.id || '',
                        name: elem.name || '',
                        className: elem.className || '',
                        value: elem.value || '',
                        xpath: getXPath(elem)
                    };
                    window.recordedActions.push(action);
                    console.log('录制: 输入', action);

                    lastInput[key] = now;
                    window.lastInputElements = lastInput;
                }
            }
        }, true);

        // 监听页面导航
        let lastUrl = location.href;
        new MutationObserver(() => {
            const url = location.href;
            if (url !== lastUrl) {
                var action = {
                    type: 'navigate',
                    timestamp: Date.now(),
                    url: url
                };
                window.recordedActions.push(action);
                console.log('录制: 导航', action);
                lastUrl = url;
            }
        }).observe(document, { subtree: true, childList: true });

        // 获取XPath的工具函数
        function getXPath(element) {
            if (element.id !== '') {
                return "//*[@id='" + element.id + "']";
            }
            if (element === document.body) {
                return '/html/body';
            }

            var ix = 0;
            var siblings = element.parentNode.childNodes;
            for (var i = 0; i < siblings.length; i++) {
                var sibling = siblings[i];
                if (sibling === element) {
                    return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                }
                if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {
                    ix++;
                }
            }
        }
        """
        self.driver.execute_script(listener_script)

    def _listen(self):
        """监听并收集操作"""
        try:
            while True:
                time.sleep(0.5)
                # 定期获取录制的操作
                actions = self.driver.execute_script("return window.recordedActions || [];")
                self.actions = actions
        except KeyboardInterrupt:
            self.stop()

    def stop(self):
        """停止录制并保存"""
        print("\n\n" + "=" * 60)
        print("录制已停止")
        print("=" * 60)

        if self.actions:
            print(f"\n共录制了 {len(self.actions)} 个操作")

            # 处理操作数据
            processed_actions = self._process_actions()

            # 保存为JSON
            filename = f"recorded_actions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(processed_actions, f, ensure_ascii=False, indent=2)
            print(f"\n操作已保存到: {filename}")

            # 生成可执行脚本
            script_filename = self._generate_script(processed_actions)
            print(f"可执行脚本已生成: {script_filename}")

            # 显示操作摘要
            self._show_summary(processed_actions)
        else:
            print("\n没有录制到任何操作")

        self.driver.quit()
        print("\n浏览器已关闭")

    def _process_actions(self):
        """处理录制的操作"""
        if not self.actions:
            return []

        processed = []
        start_time = self.actions[0]['timestamp'] if self.actions else 0

        for i, action in enumerate(self.actions):
            # 计算相对时间（秒）
            relative_time = (action['timestamp'] - start_time) / 1000

            processed_action = {
                'index': i + 1,
                'type': action['type'],
                'time_delay': round(relative_time, 2),
                'data': action
            }

            # 生成定位器
            locator = self._generate_locator(action)
            if locator:
                processed_action['locator'] = locator

            processed.append(processed_action)

        return processed

    def _generate_locator(self, action):
        """生成元素定位器"""
        if action['type'] in ['click', 'input']:
            # 优先级：ID > Name > Class > XPath > Text
            if action.get('id'):
                return {'type': 'id', 'value': action['id']}
            elif action.get('name'):
                return {'type': 'name', 'value': action['name']}
            elif action.get('className'):
                classes = action['className'].split()
                if classes:
                    return {'type': 'class', 'value': classes[0]}
            elif action.get('xpath'):
                return {'type': 'xpath', 'value': action['xpath']}
            elif action.get('text'):
                return {'type': 'text', 'value': action['text']}
        return None

    def _generate_script(self, actions):
        """生成可执行的Python脚本"""
        filename = f"auto_replay_{datetime.now().strftime('%Y%m%d_%H%M%S')}.py"

        script_content = f'''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动操作脚本
生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
录制操作数: {len(actions)}
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


class AutoPlayer:
    """自动操作播放器"""

    def __init__(self, base_url="{self.url}"):
        self.base_url = base_url
        self.driver = None
        self.wait = None

    def start(self):
        """启动浏览器"""
        chrome_options = Options()
        chrome_options.add_argument('--window-size=1920,1080')
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 10)

        print("正在打开页面...")
        self.driver.get(self.base_url)
        time.sleep(2)

        print("开始执行操作...\\n")

    def play(self):
        """执行录制的操作"""
        try:
'''

        # 为每个操作生成代码
        for action in actions:
            action_type = action['type']
            delay = action['time_delay']
            data = action['data']

            script_content += f"\n            # 操作 {action['index']}: {action_type}\n"

            if action_type == 'navigate':
                script_content += f"            # 导航到: {data.get('url', 'N/A')}\n"
                script_content += f"            # 注意: 页面导航由原始操作触发，此处仅记录\n"

            elif action_type == 'click':
                locator = action.get('locator', {})
                script_content += self._generate_click_code(locator, data)
                script_content += f"            time.sleep({delay})\n"

            elif action_type == 'input':
                locator = action.get('locator', {})
                value = data.get('value', '')
                script_content += self._generate_input_code(locator, value, data)
                script_content += f"            time.sleep({delay})\n"

        # 添加结尾
        script_content += '''
            print("\\n所有操作执行完成！")
            print("按回车键关闭浏览器...")
            input()

        except Exception as e:
            print(f"执行出错: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.driver:
                self.driver.quit()

    def _find_element(self, locator_type, locator_value):
        """查找元素"""
        by_mapping = {
            'id': By.ID,
            'name': By.NAME,
            'class': By.CLASS_NAME,
            'xpath': By.XPATH,
            'text': By.XPATH,
        }

        by = by_mapping.get(locator_type, By.XPATH)

        if locator_type == 'text':
            # 使用文本定位
            xpath = f"//*[contains(text(), '{locator_value}')]"
            return self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))

        return self.wait.until(EC.presence_of_element_located((by, locator_value))

    def _click(self, locator_type, locator_value):
        """点击元素"""
        elem = self._find_element(locator_type, locator_value)
        elem.click()
        print(f"  ✓ 点击: {{locator_type}}={{locator_value}}")

    def _input(self, locator_type, locator_value, text):
        """输入文本"""
        elem = self._find_element(locator_type, locator_value)
        elem.clear()
        elem.send_keys(text)
        print(f"  ✓ 输入: {{locator_type}}={{locator_value}}, 值={{text}}")


if __name__ == "__main__":
    print("=" * 50)
    print("自动操作脚本")
    print("=" * 50)

    player = AutoPlayer()
    player.start()
    player.play()
'''

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(script_content)

        return filename

    def _generate_click_code(self, locator, data):
        """生成点击操作代码"""
        if not locator:
            return "            # 无法定位该元素\n"

        loc_type = locator['type']
        loc_value = locator['value']

        return f'''            try:
                elem = self._find_element('{loc_type}', '{loc_value}')
                elem.click()
                print(f"  ✓ 操作 {action['index']}: 点击")
            except:
                print(f"  ✗ 操作 {action['index']}: 点击失败，定位方式={{'{loc_type}'}}={{'{loc_value}'}}")
'''

    def _generate_input_code(self, locator, value, data):
        """生成输入操作代码"""
        if not locator:
            return f"            # 无法定位该输入框，值: {value}\n"

        loc_type = locator['type']
        loc_value = locator['value']

        return f'''            try:
                elem = self._find_element('{loc_type}', '{loc_value}')
                elem.clear()
                elem.send_keys('{value}')
                print(f"  ✓ 操作 {action['index']}: 输入 = {{'{value}'}}")
            except:
                print(f"  ✗ 操作 {action['index']}: 输入失败")
'''

    def _show_summary(self, actions):
        """显示操作摘要"""
        print("\n" + "=" * 60)
        print("操作摘要")
        print("=" * 60)

        click_count = sum(1 for a in actions if a['type'] == 'click')
        input_count = sum(1 for a in actions if a['type'] == 'input')
        navigate_count = sum(1 for a in actions if a['type'] == 'navigate')

        print(f"\n点击操作: {click_count} 次")
        print(f"输入操作: {input_count} 次")
        print(f"页面导航: {navigate_count} 次")
        print(f"总操作数: {len(actions)} 次")

        print("\n主要操作列表:")
        for action in actions[:10]:  # 显示前10个
            print(f"  [{action['index']:2d}] {action['type']:10s} - {action.get('locator', {}).get('type', 'N/A')}")

        if len(actions) > 10:
            print(f"  ... 还有 {len(actions) - 10} 个操作")


class AutoPlayer:
    """自动操作播放器"""

    def __init__(self, base_url):
        self.base_url = base_url
        self.driver = None
        self.wait = None

    def start(self):
        """启动浏览器"""
        chrome_options = Options()
        chrome_options.add_argument('--window-size=1920,1080')
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 10)

        print("正在打开页面...")
        self.driver.get(self.base_url)
        time.sleep(2)
        print("开始执行操作...\n")

    def play(self, actions):
        """执行录制的操作"""
        try:
            for action in actions:
                action_type = action['type']

                if action_type == 'click':
                    self._click(action)
                elif action_type == 'input':
                    self._input(action)

                time.sleep(0.5)

            print("\n所有操作执行完成！")
            print("按回车键关闭浏览器...")
            input()

        except Exception as e:
            print(f"执行出错: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.driver:
                self.driver.quit()

    def _click(self, action):
        """点击元素"""
        # 实现点击逻辑
        pass

    def _input(self, action):
        """输入文本"""
        # 实现输入逻辑
        pass


def main():
    """主函数"""
    print("=" * 60)
    print("浏览器操作录制工具")
    print("=" * 60)

    url = input("\n请输入起始URL (默认: https://pxxt.zju.edu.cn): ").strip()
    if not url:
        url = "https://pxxt.zju.edu.cn"

    recorder = BrowserRecorder(url)
    recorder.start()


if __name__ == "__main__":
    main()
