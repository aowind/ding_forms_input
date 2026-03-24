# 钉钉表格自动填入工具

将本地 Excel 数据自动填入钉钉在线表格的 Windows 桌面工具。

## 功能

- 🔐 通过 Cookie 登录钉钉文档（无需扫码）
- 📋 自动检测 Sheet 标签页
- 📥 自动下载钉钉表格数据并建立 ID → 行号映射
- 📊 支持动态选择匹配列和填入列
- 🚀 模拟键盘操作精确填入数据
- 📝 实时日志和进度显示

## 环境要求

- Windows 10/11
- Python 3.10+

## 安装

### 1. 克隆仓库

```bash
git clone https://github.com/aowind/ding_forms_input.git
cd ding_forms_input
```

### 2. 安装 Python 依赖

```bash
pip install -r requirements.txt
```

### 3. 安装 Playwright 浏览器

```bash
playwright install chromium
```

> 首次运行时如果未安装浏览器，会提示自动安装。

## 使用方法

### 方式一：直接运行 Python 脚本

```bash
python main.py
```

### 方式二：打包为 exe（可选）

```bash
pip install pyinstaller
pyinstaller build.spec
```

打包后的 exe 在 `dist/钉钉表格自动填入工具/` 目录下。

## 使用步骤

### Step 1: 登录钉钉文档

1. 在浏览器中打开钉钉文档并登录
2. 按 F12 打开开发者工具 → Network 面板
3. 刷新页面，找到任意请求，复制 Request Headers 中的 `Cookie` 值
4. 在工具中粘贴文档 URL 和 Cookie
5. 点击「连接」

> ⚠️ `document.cookie` 无法获取 HttpOnly cookie，请从 Network 面板复制完整 Cookie 请求头。

### Step 2: 选择 Sheet

1. 工具自动检测文档中的 Sheet 标签页
2. 选择目标 Sheet
3. 设置匹配列（默认 D 列，用于定位行）
4. 点击「下载并建立映射」

### Step 3: 选择本地 Excel

1. 选择包含待填数据的 `.xlsx` 文件
2. 选择 Sheet
3. 选择**匹配列**（用于匹配钉钉表格中的行）
4. 选择**填入列**（要填入的数据列，可多选）
5. 点击「确认并准备填入」

### Step 4: 执行填入

1. 查看匹配摘要
2. 点击「开始填入」
3. 观察实时日志和进度
4. 完成后查看统计结果

## 获取 Cookie 方法

### 方法一（推荐）：Network 面板

1. Chrome 打开钉钉文档页面
2. F12 → Network 面板
3. 刷新页面（F5）
4. 点击任意请求
5. 在 Headers 中找到 `Cookie:` 行
6. 复制整个 Cookie 值（`name1=val1; name2=val2; ...`）

### 方法二：Console（不完整）

```javascript
copy(document.cookie)
```

> 注意：此方法**无法获取 HttpOnly 的 cookie**（如 `doc_atoken`），可能导致登录失败。

## 技术原理

钉钉在线表格使用 Canvas 渲染，没有 DOM 表格元素，无法通过常规 API 写入数据。本工具通过以下方式实现：

1. **Cookie 注入**：模拟浏览器登录状态
2. **键盘模拟**：使用 Playwright 发送键盘事件
3. **精确定位**：从 A1 逐行 ArrowDown 导航（不使用 PageDown）
4. **值验证**：Ctrl+C 读取单元格值验证定位准确性
5. **安全编辑**：F2 → Ctrl+A → 输入 → Tab（不使用 Enter）

## 注意事项

- Cookie 有效期有限，过期后需要重新获取
- 只能编辑已有单元格，无法插入新行
- 首次运行需要安装 Chromium 浏览器
- 填入过程中请勿操作浏览器窗口
- 大数据量（500+ 行）填入需要较长时间

## 日志

日志文件保存在 `~/.ding_forms_input/app.log`。

## License

MIT
