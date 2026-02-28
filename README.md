# 浙江大学设备存放地批量更新脚本

## 功能说明

从Excel文件读取设备信息，批量更新内网资产管理系统的"学院存放地"字段。

## 文件说明

| 文件名 | 说明 |
|--------|------|
| `update_device_location.py` | 主程序脚本 |
| `设备存放地修改0227.xls` | 数据源文件（42条记录） |
| `config_template.json` | 配置文件模板 |
| `requirements.txt` | Python依赖包列表 |
| `update_log_*.txt` | 运行日志（运行后生成） |

## 安装步骤

### 1. 安装Python依赖

```bash
pip install -r requirements.txt
```

或手动安装：

```bash
pip install selenium pandas openpyxl xlrd
```

### 2. 下载ChromeDriver

1. 查看Chrome版本：地址栏输入 `chrome://version/`
2. 下载对应版本的ChromeDriver：
   - 官方地址：https://chromedriver.chromium.org/
   - 或使用镜像：https://googlechromelabs.github.io/chrome-for-testing/
3. 将下载的 `chromedriver.exe` 放到以下位置之一：
   - 与脚本同目录
   - 系统PATH中的任意目录

### 3. 配置元素定位（重要！）

打开 `update_device_location.py`，找到 `ELEMENT_LOCATORS` 部分，根据实际页面填写元素定位方式。

#### 获取元素定位的方法：

1. 登录系统后，按 **F12** 打开开发者工具
2. 点击左上角的**元素选择器**（或按 Ctrl+Shift+C）
3. 点击页面上的目标元素
4. 在Elements面板中右键点击该元素
5. 选择 **Copy** → **Copy XPath** 或 **Copy selector**

| 元素 | 需要定位的字段 |
|------|---------------|
| 搜索框 | 输入资产编号的输入框 |
| 搜索按钮 | 点击搜索的按钮 |
| 编辑按钮 | 打开设备编辑页面的按钮 |
| 存放地输入框 | "学院存放地"输入框 |
| 保存按钮 | 保存修改的按钮 |
| 成功提示 | 保存成功后的提示信息 |

### 4. 配置Cookie（用于免登录）

打开 `update_device_location.py`，找到 `COOKIES_CONFIG` 部分，填写Cookie值。

#### 获取Cookie的方法：

1. 登录系统后，按 **F12** 打开开发者工具
2. 切换到 **Application** 标签页
3. 左侧找到 **Storage** → **Cookies** → 选择对应域名
4. 找到并复制以下Cookie的值：
   - `iPlanetDirectoryPro` （登录token）
   - `JSESSIONID` （会话ID）
   - 其他必要的Cookie

### 5. 确认Excel列名

打开 `update_device_location.py`，找到 `COLUMN_NAMES` 部分，确认Excel文件的列名是否正确。

## 使用方法

### 运行脚本

```bash
python update_device_location.py
```

### 运行模式

脚本提供4种运行模式：

| 模式 | 说明 |
|------|------|
| 1. 测试模式 | 只处理前3条记录，用于验证配置是否正确 |
| 2. 批量处理 | 处理所有42条记录 |
| 3. 自定义范围 | 指定处理的记录范围（索引从0开始） |
| 4. 元素定位测试 | 帮助调试元素定位是否正确 |

### 建议使用流程

1. **先运行模式1（测试模式）**，验证配置正确
2. 检查日志文件，确认前3条记录处理成功
3. **再运行模式2（批量处理）**，处理所有记录
4. 最后登录系统抽查几条记录，确认修改成功

## 日志文件

运行后会生成 `update_log_YYYYMMDD_HHMMSS.txt` 文件，记录：
- 每条记录的处理状态
- 成功/失败统计
- 失败记录的详细信息

## 常见问题

### Q1: ChromeDriver版本不匹配

错误信息：`This version of ChromeDriver only supports Chrome version XX`

解决方法：下载与Chrome版本匹配的ChromeDriver

### Q2: 找不到元素

错误信息：`找不到元素: xxx`

解决方法：
1. 使用模式4（元素定位测试）检查定位是否正确
2. 确认页面是否完全加载，可增加 `WAIT_TIME` 值
3. 检查元素是否在iframe中，需要切换frame

### Q3: Cookie过期

错误信息：`Cookie加载失败` 或 `未登录`

解决方法：
1. 重新登录系统
2. 获取新的Cookie值
3. 更新 `COOKIES_CONFIG` 配置

### Q4: 页面加载超时

错误信息：`TimeoutException`

解决方法：
1. 增加 `PAGE_LOAD_TIMEOUT` 和 `WAIT_TIME` 的值
2. 检查网络连接

### Q5: Excel读取失败

错误信息：`读取Excel文件失败`

解决方法：
1. 确认文件名正确：`设备存放地修改0227.xls`
2. 确认文件在脚本同目录下
3. 检查列名是否与 `COLUMN_NAMES` 配置一致

## 注意事项

1. **Cookie会过期**，建议在运行前重新获取
2. **建议先测试**，确认无误后再批量处理
3. **处理速度**：默认每条记录间隔2秒，可根据网络情况调整 `WAIT_TIME`
4. **备份重要数据**：虽然只修改"学院存放地"字段，但仍建议先备份
5. **保持浏览器可见**：首次运行建议设置 `HEADLESS = False`，方便观察运行情况

## 高级配置

### 使用配置文件

可以修改 `config_template.json` 文件，然后在脚本中加载：

```python
# 在脚本开头添加
load_config_from_file('config_template.json')
```

### 调整等待时间

在脚本中修改以下参数：

```python
WAIT_TIME = 2  # 每次操作后的等待时间（秒）
PAGE_LOAD_TIMEOUT = 30  # 页面加载超时时间（秒）
```

### 无头模式

正式运行时可设置 `HEADLESS = True`，不显示浏览器窗口。

## 技术支持

如遇问题，请检查：
1. ChromeDriver版本是否匹配
2. 元素定位是否正确
3. Cookie是否有效
4. 网络连接是否正常
5. 日志文件中的错误信息
