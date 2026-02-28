#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
元素定位辅助工具
用于帮助用户获取页面元素的定位方式
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By


def print_element_info(element, description):
    """打印元素信息"""
    print(f"\n【{description}】")
    print(f"  TagName: {element.tag_name}")
    print(f"  ID: {element.get_attribute('id')}")
    print(f"  Name: {element.get_attribute('name')}")
    print(f"  Class: {element.get_attribute('class')}")
    print(f"  XPath (相对): {find_xpath_relative(element)}")
    print(f"  Full XPath: {find_xpath_full(element)}")


def find_xpath_relative(element):
    """生成相对XPath（简化版）"""
    tag = element.tag_name
    id_attr = element.get_attribute('id')
    class_attr = element.get_attribute('class')

    if id_attr:
        return f"//*[@id='{id_attr}']"
    elif class_attr:
        classes = class_attr.strip().split()
        if classes:
            return f"//{tag}[@class='{classes[0]}']"
    return f"//{tag}"


def find_xpath_full(element):
    """生成完整XPath（从html根节点开始）"""
    driver = element.parent
    components = []
    current = element

    while current and current.tag_name.lower() != 'html':
        tag = current.tag_name
        id_attr = current.get_attribute('id')

        if id_attr:
            components.append(f"{tag}[@id='{id_attr}']")
            break
        else:
            # 计算同级索引
            siblings = [s for s in current.find_elements(By.XPATH, f"../{tag}")]
            if len(siblings) > 1:
                index = siblings.index(current) + 1
                components.append(f"{tag}[{index}]")
            else:
                components.append(tag)

        current = current.find_element(By.XPATH, "..")

    components.reverse()
    return "//html/" + "/".join(components)


def main():
    """主函数"""
    print("=" * 60)
    print("元素定位辅助工具")
    print("=" * 60)
    print("\n本工具帮助你获取页面元素的定位方式（ID、XPath等）")
    print("\n使用说明：")
    print("1. 脚本会打开Chrome浏览器")
    print("2. 请手动登录系统")
    print("3. 导航到需要操作的页面")
    print("4. 按回车键，然后点击页面上的目标元素")
    print("5. 脚本会显示该元素的各种定位方式")
    print("\n" + "=" * 60)

    url = input("\n请输入系统URL (默认: https://pxxt.zju.edu.cn): ").strip()
    if not url:
        url = "https://pxxt.zju.edu.cn"

    try:
        # 启动浏览器
        chrome_options = Options()
        chrome_options.add_argument('--window-size=1920,1080')
        driver = webdriver.Chrome(options=chrome_options)

        print(f"\n正在打开 {url} ...")
        driver.get(url)

        print("\n" + "=" * 60)
        print("请在浏览器中完成以下操作：")
        print("  1. 手动登录系统")
        print("  2. 导航到设备管理页面")
        print("完成后按回车键继续...")
        print("=" * 60)
        input()

        print("\n\n现在请在浏览器中依次点击以下元素，每次点击后按回车:")

        # 需要获取的元素列表
        elements_to_find = [
            ("搜索框", "用于输入资产编号的输入框"),
            ("搜索按钮", "点击执行搜索的按钮"),
            ("编辑按钮", "点击进入设备编辑页面的按钮"),
            ("存放地输入框", "学院存放地的输入框"),
            ("保存按钮", "点击保存修改的按钮"),
        ]

        locator_config = {}

        for elem_name, description in elements_to_find:
            print(f"\n{'=' * 50}")
            print(f"【{elem_name}】")
            print(f"说明: {description}")
            print(f"请在页面中点击该元素，然后按回车...")
            input()

            try:
                # 使用JavaScript获取当前聚焦的元素或最后点击的元素
                element = driver.execute_script("""
                    var elem = window.lastClickedElement;
                    if (!elem) {
                        elem = document.activeElement;
                    }
                    return elem;
                """)

                if element:
                    # 获取元素属性
                    elem_id = element.get('id', '')
                    elem_name = element.get('name', '')
                    elem_class = element.get('class', '')
                    elem_tag = element.get('tagName', '').lower()

                    print(f"\n✓ 找到元素!")
                    print(f"  TagName: <{elem_tag}>")
                    print(f"  ID: '{elem_id}'")
                    print(f"  Name: '{elem_name}'")
                    print(f"  Class: '{elem_class}'")

                    # 生成推荐的定位方式
                    recommendations = []

                    if elem_id:
                        recommendations.append({
                            'type': 'By.ID',
                            'value': elem_id,
                            'priority': 1,
                            'code': f"By.ID, '{elem_id}'"
                        })

                    if elem_name:
                        recommendations.append({
                            'type': 'By.NAME',
                            'value': elem_name,
                            'priority': 2,
                            'code': f"By.NAME, '{elem_name}'"
                        })

                    if elem_class:
                        first_class = elem_class.split()[0]
                        recommendations.append({
                            'type': 'By.CLASS_NAME',
                            'value': first_class,
                            'priority': 3,
                            'code': f"By.CLASS_NAME, '{first_class}'"
                        })

                    # XPath by ID or text
                    if elem_id:
                        recommendations.append({
                            'type': 'By.XPATH',
                            'value': f"//*[@id='{elem_id}']",
                            'priority': 1,
                            'code': f"By.XPATH, \"//*[@id='{elem_id}']\""
                        })
                    elif elem_class:
                        recommendations.append({
                            'type': 'By.XPATH',
                            'value': f"//{elem_tag}[@class='{elem_class.split()[0]}']",
                            'priority': 4,
                            'code': f"By.XPATH, \"//{elem_tag}[@class='{elem_class.split()[0]}']\""
                        })

                    # 显示推荐定位方式
                    print(f"\n推荐的定位方式（按优先级排序）:")
                    recommendations.sort(key=lambda x: x['priority'])
                    for i, rec in enumerate(recommendations[:3], 1):
                        print(f"  {i}. {rec['code']}")

                    # 保存最高优先级的定位方式
                    if recommendations:
                        best = min(recommendations, key=lambda x: x['priority'])
                        locator_config[elem_name] = {
                            'code': best['code'],
                            'by_type': best['type'],
                            'value': best['value']
                        }

                else:
                    print(f"\n✗ 无法获取元素，请重试")

            except Exception as e:
                print(f"\n✗ 出错: {e}")

        # 注入点击跟踪脚本
        driver.execute_script("""
            document.addEventListener('click', function(e) {
                window.lastClickedElement = e.target;
            }, true);
        """)

        # 输出配置结果
        print("\n\n" + "=" * 60)
        print("配置结果")
        print("=" * 60)

        if locator_config:
            print("\n获取到的元素定位配置：\n")
            print("ELEMENT_LOCATORS = {")
            for elem_name, config in locator_config.items():
                print(f"    '{elem_name}': {{")
                print(f"        'by': {config['by_type']},")
                print(f"        'value': '{config['value']}'")
                print(f"    }},")
            print("}")

            # 保存到文件
            with open('element_locators_output.txt', 'w', encoding='utf-8') as f:
                f.write("ELEMENT_LOCATORS = {\n")
                for elem_name, config in locator_config.items():
                    # 转换为配置格式
                    by_mapping = {
                        'By.ID': 'By.ID',
                        'By.NAME': 'By.NAME',
                        'By.CLASS_NAME': 'By.CLASS_NAME',
                        'By.XPATH': 'By.XPATH',
                    }
                    f.write(f"    '{elem_name}': {{\n")
                    f.write(f"        'by': {by_mapping.get(config['by_type'], config['by_type'])},\n")
                    f.write(f"        'value': '{config['value']}'\n")
                    f.write(f"    }},\n")
                f.write("}\n")

            print("\n配置已保存到: element_locators_output.txt")
            print("请将上述配置复制到 update_device_location.py 中的 ELEMENT_LOCATORS 部分")

        print("\n\n按回车键关闭浏览器...")
        input()
        driver.quit()

    except Exception as e:
        print(f"\n程序出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
