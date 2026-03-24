# 钉钉表格自动填入工具

<p align="center">
  <b>将本地 Excel 数据自动填入钉钉在线表格的 Windows 桌面工具</b>
</p>

---

## ✨ 功能特性

- 📱 **扫码登录** — 弹出浏览器扫码，无需手动复制 Cookie
- 📋 **智能映射** — 上传钉钉导出的 Excel 自动建立「ID → 行号」映射
- 🎯 **动态列选择** — 匹配列、填入列均由用户选择，不硬编码
- 👁️ **映射预览** — 填入前可查看匹配/未匹配数据，确认无误再执行
- 🖥️ **自动锁定** — 填入时浏览器窗口自动移到屏幕外，防止误操作
- 🔄 **断点恢复** — 定位偏移时自动回退重试，失败行跳过不中断
- 📦 **免安装发布** — 打包为 exe，用户双击即用，无需任何环境

## 🚀 快速开始

### 用户使用（exe 版本）

1. 解压分发的 zip 文件
2. 双击 `钉钉表格自动填入工具.exe`
3. 按照界面引导完成 4 步操作

### 开发者使用（源码版本）

**环境要求：** Python 3.9+

```bash
# 克隆项目
git clone https://github.com/aowind/ding_forms_input.git
cd ding_forms_input

# 安装依赖
pip install -r requirements.txt
playwright install chromium

# 运行
python main.py
```

## 📖 使用教程

### 第一步：登录钉钉文档

1. 输入钉钉文档 URL（如 `https://alidocs.dingtalk.com/i/nodes/...`）
2. 点击「🌐 打开浏览器」，弹出 Chrome 窗口
3. 在弹出的浏览器中**扫码登录**钉钉账号
4. 登录成功、看到表格内容后，点击「✅ 我已登录」

### 第二步：检测表格信息

工具自动检测文档中的 Sheet 标签页，确认后点击「下一步」。

### 第三步：文件配置

这一步需要**两个文件**：

#### 📌 上传钉钉表格导出文件（建立映射）

用于确定每行数据在钉钉表格中的位置。

**如何获取：**
1. 在已打开的钉钉表格页面，点击顶部工具栏的 **「表格」** 按钮
2. 在下拉菜单中悬停 **「下载」**
3. 选择 **「Excel (.xlsx, 整个表格)」**
4. 将下载的文件选择到工具中
5. 选择对应的 **Sheet** 和 **映射匹配列**（哪一列的值用于匹配定位行）

#### 📌 选择本地 Excel 数据文件

待填入的源数据。

1. 选择本地 `.xlsx` 文件
2. 选择对应的 **Sheet**
3. 选择 **数据匹配列**（与导出文件的匹配列对应）
4. **勾选要填入的列**（可多选）

#### 👁️ 查看映射预览

点击「确认并建立映射」后，工具会显示：
- 📊 统计信息：总映射数、源数据行数、匹配/未匹配数量
- ✅ **可匹配的数据**：`题包id → 钉钉表格第 N 行`
- ❌ **未匹配的数据**：在导出文件中找不到的 ID

确认无误后点击「🚀 开始执行填入」。

### 第四步：执行填入

1. 点击「▶️ 开始填入」
2. 浏览器窗口会**自动移到屏幕外**（防止误操作干扰）
3. 工具自动完成数据填入，实时显示进度和日志
4. 可随时点击「⏹️ 中断」停止执行
5. 填入完成后浏览器窗口自动恢复到原位

## 🏗️ 构建 exe

在 **Windows** 上执行：

```bash
# 一键构建（自动完成所有步骤）
python build.py
```

这会自动执行：
1. 安装 Python 依赖（customtkinter、playwright、openpyxl）
2. 安装 PyInstaller
3. 下载 Playwright Chromium 浏览器
4. 打包为 exe（浏览器一并打包）

**输出目录：** `dist/钉钉表格自动填入工具/`

**手动构建：**
```bash
pip install -r requirements.txt
pip install pyinstaller
playwright install chromium
pyinstaller build.spec --clean --noconfirm
```

## 📦 发布分发

1. 将 `dist/钉钉表格自动填入工具/` **整个文件夹**压缩为 zip
2. 发送给用户
3. 用户解压后双击 `.exe` 即可运行

> ⚠️ 必须**整个文件夹**一起分发，exe 依赖同目录下的 `browsers/` 等文件

## 📁 项目结构

```
ding_forms_input/
├── main.py              # 应用入口（打包时设置浏览器路径）
├── build.py             # 一键构建脚本
├── build.spec           # PyInstaller 打包配置
├── requirements.txt     # Python 依赖
├── README.md            # 说明文档
│
├── core/                # 核心逻辑层
│   ├── browser.py       # Playwright 浏览器管理（启动/关闭/窗口隐藏）
│   ├── cookie.py        # Cookie 字符串解析
│   ├── excel_reader.py  # Excel 读写（openpyxl）
│   ├── filler.py        # 核心填入逻辑（键盘模拟/相对导航/验证恢复）
│   └── table_downloader.py  # 钉钉表格自动下载
│
└── ui/                  # GUI 界面层
    ├── app.py           # 主窗口（4 步向导 + asyncio 后台线程）
    ├── step_login.py    # Step 1: 扫码登录
    ├── step_sheet.py    # Step 2: 检测 Sheet 标签
    ├── step_excel.py    # Step 3: 文件配置 + 映射预览
    └── step_execute.py  # Step 4: 执行填入（进度条 + 实时日志）
```

## 🔧 技术栈

| 组件 | 技术 | 用途 |
|------|------|------|
| GUI | customtkinter | 现代深色主题界面 |
| 浏览器自动化 | Playwright | 控制 Chromium 填入数据 |
| Excel 读写 | openpyxl | 解析本地 .xlsx 文件 |
| 打包 | PyInstaller | 打包为 Windows exe |

## ⚙️ 核心填入原理

钉钉在线表格使用 Canvas 渲染，没有 DOM 表格元素，也没有 REST API 写入接口。工具通过模拟键盘操作完成填入：

```
定位行：Ctrl+Home(A1) → ArrowDown(N次) → Home(A列) → ArrowRight(匹配列)
验证：  Ctrl+C → clipboard.readText() 比对值
填入：  Tab(到填入列) → F2(编辑) → Ctrl+A(全选) → type(输入) → Tab(确认+下一列)
```

**关键约束：**
- ✅ 用 Tab 确认编辑（不能用 Enter，会下移一行）
- ✅ 空值也要 Tab 跳过（否则列对齐错误）
- ✅ 按行号排序后相对移动（不回 A1）
- ❌ 不用 PageDown（虚拟滚动跳跃行数不确定）

## ⚠️ 注意事项

- 钉钉登录态有效期为几小时，超时需重新扫码
- 填入时请不要关闭浏览器窗口
- 每次键盘操作间隔 ~450ms，500 行数据约需 4-5 分钟
- 仅支持编辑已有单元格，不支持新增行
- 日期格式自动转为 `YYYY-MM-DD`

## License

MIT
