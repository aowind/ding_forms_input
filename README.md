# 钉钉表格自动填入工具

将本地 Excel 数据自动填入钉钉在线表格的 Windows 桌面工具。

## 功能

- 🔐 扫码登录钉钉文档（无需手动复制 Cookie）
- 📊 从钉钉表格导出的 Excel 建立行映射
- 📁 选择本地 Excel 数据文件，动态选择匹配列和填入列
- 👁️ 映射预览：填入前可查看匹配/未匹配数据
- 🚀 自动填入：模拟键盘操作精确填入钉钉表格
- 🖥️ 填入时自动锁定浏览器窗口，防止误操作

## 作为用户直接使用

### 方式一：从 exe 运行（推荐）

1. 解压分发的 zip 文件
2. 双击 `钉钉表格自动填入工具.exe`
3. 按照界面引导操作即可

### 方式二：从源码运行

1. 安装 Python 3.9+（[下载地址](https://www.python.org/downloads/)）
2. 安装依赖：
   ```
   pip install -r requirements.txt
   playwright install chromium
   ```
3. 运行：
   ```
   python main.py
   ```

## 使用步骤

### Step 1: 登录
1. 输入钉钉文档 URL（如 `https://alidocs.dingtalk.com/i/nodes/...`）
2. 点击「打开浏览器」，会弹出 Chrome
3. 在弹出的浏览器中扫码登录
4. 登录成功看到表格后，点击「我已登录」

### Step 2: 表格信息
自动检测文档中的 Sheet 标签，确认后进入下一步。

### Step 3: 文件配置

**第一步 - 上传钉钉导出文件（建立映射）：**
1. 在已打开的钉钉表格页面，点击顶部工具栏「表格」→「下载」→「Excel」
2. 将下载的 .xlsx 文件选择到工具中
3. 选择 Sheet 和映射匹配列（用于定位行的 ID 列）

**第二步 - 选择本地数据文件：**
1. 选择本地 Excel 数据文件
2. 选择 Sheet
3. 选择数据匹配列（与导出文件的匹配列对应）
4. 勾选要填入的列

**确认后查看映射预览**，确认无误后点击「开始执行填入」。

### Step 4: 执行填入
点击「开始填入」，工具会自动完成填入。填入过程中浏览器窗口会被移到屏幕外，防止误操作。填入完成后浏览器窗口会自动恢复。

## 从源码构建 exe

在 **Windows** 上执行：

```bash
# 一键构建（自动安装依赖 + 打包）
python build.py
```

构建完成后，输出在 `dist/钉钉表格自动填入工具/` 目录。

**手动构建：**
```bash
pip install -r requirements.txt
pip install pyinstaller
playwright install chromium
pyinstaller build.spec --clean --noconfirm
```

### 发布分发

将整个 `dist/钉钉表格自动填入工具/` 文件夹压缩为 zip 发给用户。用户解压后双击 exe 即可运行，无需安装任何东西。

## 技术栈

- **Python 3.9+**
- **customtkinter** — 现代 GUI 界面
- **Playwright** — 浏览器自动化（非 headless，需要可见窗口）
- **openpyxl** — Excel 读写
- **PyInstaller** — 打包为 exe

## 注意事项

- 填入时请不要关闭弹出的浏览器窗口
- 钉钉文档的 Cookie/登录态有效期为几个小时，超时需要重新登录
- 每次键盘操作间隔 ~450ms，500 行数据约需 4-5 分钟
- 不支持新增行，只能编辑已有单元格

## 许可

MIT
