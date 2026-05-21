# 论文排版工具 - HPU Thesis Formatter

基于 PyQt6 的桌面应用，用于按模板自动排版 .docx 论文文档。

## 功能

- **可视化模板配置**：所有排版参数均可通过 UI 调整（字体、字号、行距、缩进、三线表、页眉页脚等）
- **一键排版**：打开 .docx 文件，选择模板，一键应用全部格式
- **模板管理**：保存、导入、导出排版模板（`.hputpl` JSON 格式）
- **内置 HPU 模板**：预置河南理工大学本科毕业设计（论文）撰写规范

## 快速开始

### 方式一：直接使用（推荐）

下载 `dist/HPU论文排版工具.exe`，双击运行。

### 方式二：源码运行

```bash
pip install -r requirements.txt
python main.py
```

## 项目结构

```
HPU/
├── main.py                     # 入口
├── app/
│   ├── main_window.py          # 主窗口
│   ├── models/config.py        # 数据模型
│   ├── engine/formatter.py     # 排版引擎
│   ├── templates/manager.py    # 模板管理器
│   ├── resources/              # 内置资源
│   └── ui/
│       ├── template_panel.py   # 模板列表面板
│       ├── settings_stack.py   # 设置导航框架
│       └── pages/              # 各设置页面
├── dist/
│   └── HPU论文排版工具.exe      # 打包的可执行文件
└── requirements.txt
```

## 技术栈

- Python 3.11+
- PyQt6（桌面 UI）
- python-docx（Word 文档处理）
- PyInstaller（打包）
