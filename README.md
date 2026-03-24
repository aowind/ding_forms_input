# 钉钉表格自动填入工具

<p align="center">
  <b>将本地 Excel 数据自动填入钉钉在线表格的 Windows 桌面工具</b>
</p>

---

## ✨ 功能特性

- 📱 **扫码登录** — 弹出浏览器扫码，无需手动复制 Cookie
- 📋 **智能映射** — 上传钉钉导出的 Excel 自动建立「ID → 行号」映射
- 🎯 **动态列选择** — 匹配列、填入列均由用户选择，不硬编码
- 👁️ **映射预览** — 填入前查看匹配/未匹配数据，确认无误再执行
- 🖥️ **自动锁定** — 填入时浏览器窗口移到屏幕外，防止误操作
- 📦 **轻量发布** — 单个 exe 文件，首次运行自动下载浏览器（~150MB，只需一次）

## 🚀 快速使用

1. 双击 `钉钉表格自动填入工具.exe`
2. 首次运行会自动下载 Chromium 浏览器（约 1-2 分钟，后续无需重复）
3. 按照界面引导完成 4 步操作

## 📖 使用教程

### 第一步：登录钉钉文档

1. 输入钉钉文档 URL
2. 点击「打开浏览器」，弹出 Chrome
3. 扫码登录
4. 看到表格后点击「我已登录」

### 第二步：表格信息

自动检测 Sheet 标签，确认后进入下一步。

### 第三步：文件配置

**上传钉钉导出文件（建立映射）：**
1. 在钉钉表格页面：顶部工具栏 → 「表格」→「下载」→「Excel」
2. 选择下载的文件，选择 Sheet 和映射匹配列

**选择本地数据文件：**
1. 选择本地 Excel
2. 选择 Sheet、数据匹配列、勾选填入列
3. 确认后查看映射预览

### 第四步：执行填入

点击「开始填入」，浏览器自动移到屏幕外，填入完成后自动恢复。

## 🏗️ 构建 exe

在 **Windows** 上：

```bash
git clone https://github.com/aowind/ding_forms_input.git
cd ding_forms_input
python build.py
```

输出 `dist/钉钉表格自动填入工具.exe`，直接分发即可。

## 📁 项目结构

```
ding_forms_input/
├── main.py              # 入口
├── build.py             # 一键构建
├── build.spec           # PyInstaller 配置
├── requirements.txt
├── core/
│   ├── browser.py       # 浏览器管理（启动/隐藏/自动下载）
│   ├── cookie.py        # Cookie 解析
│   ├── excel_reader.py  # Excel 读写
│   ├── filler.py        # 核心填入逻辑
│   └── table_downloader.py
└── ui/
    ├── app.py           # 主窗口（4 步向导）
    ├── step_login.py    # Step 1: 扫码登录
    ├── step_sheet.py    # Step 2: 检测 Sheet
    ├── step_excel.py    # Step 3: 文件配置 + 映射预览
    └── step_execute.py  # Step 4: 执行填入
```

## 🔧 技术栈

| 组件 | 技术 | 用途 |
|------|------|------|
| GUI | customtkinter | 现代深色主题界面 |
| 浏览器自动化 | Playwright | 控制 Chromium 填入数据 |
| Excel | openpyxl | 解析 .xlsx 文件 |
| 打包 | PyInstaller | 单文件 exe（onedir 模式） |

## ⚠️ 注意事项

- 首次运行需要联网下载浏览器（~150MB），后续不再需要
- 钉钉登录态有效期几小时，超时需重新扫码
- 填入时请不要关闭浏览器窗口
- 仅支持编辑已有单元格，不支持新增行

## License

MIT
