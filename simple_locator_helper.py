#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化版元素定位获取工具
打开浏览器，监听点击事件，显示元素定位信息
"""

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def main():
    print("=" * 60)
    print("元素定位获取工具")
    print("=" * 60)
    print("\n正在启动浏览器...")

    chrome_options = Options()
    chrome_options.add_argument('--window-size=1920,1080')
    driver = webdriver.Chrome(options=chrome_options)

    # 注入监听脚本
    inject_script = """
    window.clickedElements = [];

    document.addEventListener('click', function(e) {
        var elem = e.target;
        var info = {
            tag: elem.tagName,
            id: elem.id || '',
            name: elem.name || '',
            class: elem.className || '',
            text: elem.textContent ? elem.textContent.trim().substring(0, 30) : '',
            xpath: getXPath(elem),
            timestamp: new Date().toLocaleTimeString()
        };
        window.clickedElements.push(info);
        console.log('========== 元素信息 ==========');
        console.log('Tag:', info.tag);
        console.log('ID:', info.id);
        console.log('Name:', info.name);
        console.log('Class:', info.class);
        console.log('Text:', info.text);
        console.log('XPath:', info.xpath);
        console.log('推荐定位:', getRecommendedLocator(info));
        console.log('============================');
        e.preventDefault();
        e.stopPropagation();
    }, true);

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

    function getRecommendedLocator(info) {
        if (info.id) {
            return "By.ID, '" + info.id + "'";
        } else if (info.name) {
            return "By.NAME, '" + info.name + "'";
        } else if (info.class) {
            var classes = info.class.split(' ');
            return "By.CLASS_NAME, '" + classes[0] + "'";
        } else {
            return "By.XPATH, " + info.xpath;
        }
    }

    console.log('元素定位监听已启动！点击页面上的任意元素查看定位信息');
    """

    driver.get("https://pxxt.zju.edu.cn")

    # 注入脚本
    driver.execute_script(inject_script)

    print("\n" + "=" * 60)
    print("浏览器已打开，监听已启动！")
    print("=" * 60)
    print("\n使用说明：")
    print("1. 手动登录系统")
    print("2. 导航到设备管理页面")
    print("3. 依次点击以下元素：")
    print("   - 搜索框（输入资产编号）")
    print("   - 搜索按钮")
    print("   - 编辑按钮")
    print("   - 存放地输入框")
    print("   - 保存按钮")
    print("\n4. 每次点击后，按F12打开控制台查看元素信息")
    print("5. 完成后，回到这里按Ctrl+C结束")
    print("\n" + "=" * 60)

    # 持续运行，定期检查点击记录
    try:
        while True:
            time.sleep(2)
            elements = driver.execute_script("return window.clickedElements || [];")
            # 可以在这里处理记录
    except KeyboardInterrupt:
        print("\n\n正在获取记录的元素信息...")

        # 获取所有点击的元素
        elements = driver.execute_script("return window.clickedElements || [];")

        if elements:
            print("\n" + "=" * 60)
            print("已记录的元素:")
            print("=" * 60)

            for i, elem in enumerate(elements, 1):
                print(f"\n[{i}] {elem['tag']}")
                if elem['id']:
                    print(f"    ID: {elem['id']}")
                if elem['name']:
                    print(f"    Name: {elem['name']}")
                if elem['class']:
                    print(f"    Class: {elem['class']}")
                if elem['text']:
                    print(f"    Text: {elem['text']}")
                print(f"    XPath: {elem['xpath']}")
                print(f"    Time: {elem['timestamp']}")

            # 保存到文件
            with open('element_locators_captured.txt', 'w', encoding='utf-8') as f:
                f.write("已捕获的元素定位信息:\n\n")
                for i, elem in enumerate(elements, 1):
                    f.write(f"[{i}] {elem['tag']}\n")
                    if elem['id']:
                        f.write(f"    ID: {elem['id']}\n")
                        f.write(f"    推荐: By.ID, '{elem['id']}'\n")
                    elif elem['name']:
                        f.write(f"    Name: {elem['name']}\n")
                        f.write(f"    推荐: By.NAME, '{elem['name']}'\n")
                    elif elem['class']:
                        classes = elem['class'].split()
                        f.write(f"    Class: {elem['class']}\n")
                        f.write(f"    推荐: By.CLASS_NAME, '{classes[0]}'\n")
                    f.write(f"    XPath: {elem['xpath']}\n")
                    f.write("\n")

            print("\n信息已保存到: element_locators_captured.txt")
        else:
            print("\n没有记录到任何元素点击")

    finally:
        print("\n浏览器将关闭...")
        time.sleep(2)
        driver.quit()

if __name__ == "__main__":
    main()
