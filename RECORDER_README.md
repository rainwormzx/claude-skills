# 浏览器操作录制工具使用说明

## 功能介绍

录制你在浏览器中的操作，自动生成可重复执行的Python脚本。

## 使用方法

### 第一步：启动录制工具

```bash
python browser_recorder.py
```

### 第二步：输入起始URL

```
请输入起始URL (默认: https://pxxt.zju.edu.cn):
```

按回车使用默认URL，或输入自定义URL。

### 第三步：进行操作

浏览器会自动打开，现在可以开始你的操作：

1. **登录系统** - 输入用户名密码登录
2. **搜索设备** - 输入资产编号并搜索
3. **编辑设备** - 点击编辑按钮
4. **修改存放地** - 输入新的存放位置
5. **保存** - 点击保存按钮

脚本会自动记录以下操作：
- ✓ 点击元素
- ✓ 输入文本
- ✓ 页面导航

### 第四步：结束录制

操作完成后，**在命令行窗口按 Ctrl+C** 结束录制。

## 生成的文件

录制结束后会自动生成两个文件：

| 文件名 | 说明 |
|--------|------|
| `recorded_actions_YYYYMMDD_HHMMSS.json` | 操作记录（JSON格式） |
| `auto_replay_YYYYMMDD_HHMMSS.py` | 可自动执行的Python脚本 |

## 自动执行录制的操作

运行生成的脚本即可自动重放操作：

```bash
python auto_replay_20250228_160000.py
```

## 操作循环执行

如果需要对Excel中的所有记录循环执行，可以使用批量执行脚本。

## 示例

### 录制一次完整操作流程

假设需要录制以下流程：
1. 登录 → 2. 输入资产编号"ZJU001" → 3. 点击搜索 → 4. 点击编辑 → 5. 修改存放地为"实验楼301" → 6. 点击保存

录制完成后，生成的脚本会自动重现这个流程。

### 批量处理多个记录

将生成的脚本修改为循环读取Excel数据：

```python
# 在生成的脚本中添加
import pandas as pd

df = pd.read_excel('存放地测试.xlsx')

for index, row in df.iterrows():
    asset_number = row['资产编号']
    new_location = row['学院存放地']

    # 使用录制的操作模板，替换为实际数据
    player.search_and_update(asset_number, new_location)
```

## 注意事项

1. **首次录制建议走完整流程** - 录制一次完整的操作流程
2. **Cookie有效期** - 如果Cookie过期，需要重新录制登录部分
3. **页面加载时间** - 如果网络较慢，操作之间会自动添加等待时间
4. **元素定位** - 工具会自动选择最佳定位方式（优先ID > Name > Class > XPath）

## 高级用法

### 仅录制关键操作

如果登录信息是固定的，可以：
1. 先录制登录部分，保存为 `login.py`
2. 再录制业务操作部分，保存为 `update_device.py`
3. 在主脚本中组合调用

### 手动调整定位方式

如果自动生成的定位方式不准确，可以打开生成的 `.py` 文件，手动修改元素定位方式：

```python
# 原始生成
elem = self._find_element('class', 'search-btn')

# 手动修改为更精确的定位
elem = self._find_element('xpath', '//button[@type="submit"]')
```

### 与Excel数据结合

在批量更新脚本中，使用录制好的操作模板：

```python
# 使用录制的操作作为模板
template_actions = load_actions_from_json('recorded_actions.json')

for asset_number, new_location in excel_data:
    # 替换操作中的数据
    execute_actions_with_data(template_actions, {
        'asset_number': asset_number,
        'location': new_location
    })
```

## 常见问题

### Q: 录制的操作无法重放？

A: 可能原因：
- 页面结构变化了
- 元素定位方式不准确
- 需要等待页面加载完成

解决方法：
1. 重新录制
2. 手动修改元素定位方式
3. 添加等待时间

### Q: 如何处理动态数据？

A: 在生成的脚本中，将固定值替换为变量：

```python
# 原始
elem.send_keys('ZJU001')

# 修改为
elem.send_keys(asset_number)  # 从Excel读取
```

### Q: Cookie过期怎么办？

A: 重新录制登录部分，或手动更新Cookie配置。
