# SpectraMax Data Processor

> 将 SpectraMax 酶标仪导出的 `.xls` / `.txt` 数据文件批量转换为整理格式 `.xlsx` 的小工具。

[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## 📋 功能特性

- **自动解析**：直接读取仪器导出的原始 `.xls/.txt` 文件（UTF-16LE 编码）
- **批量转换**：支持 3 个、6 个或任意数量的数据 Block
- **独立输出**：每个 Block 自动生成一个独立的 `.xlsx` 文件
- **格式整理**：自动重排数据为标准格式（波长分组 × 子板布局）
- **灵活命名**：支持自定义输出文件名和子板标签
- **一键使用**：支持双击图形界面或把文件拖到 `easy_process.bat` 上直接处理

---

## 🚀 快速开始

### 方式 1：一键使用（推荐）

首次使用如果提示缺少依赖，先双击：

```text
install_dependencies.bat
```

日常处理数据有两种最简单方式：

```text
双击 easy_process.bat
```

然后在图形界面里选择 `.xls/.txt` 文件并点击 **开始转换**。

或者直接把一个或多个 `.xls/.txt` 文件拖到：

```text
easy_process.bat
```

默认输出目录为：

```text
E:\test
```

### 方式 2：命令行（高级/备用）

#### 环境要求

- Python 3.8+
- 依赖库：`numpy`、`pandas`、`openpyxl`

```bash
pip install -r requirements.txt
```

#### 基本用法

```bash
# Windows PowerShell（必须加 -X utf8 参数）
python -X utf8 process_xls.py "你的数据文件.xls"
```

#### 指定输出目录

```bash
python -X utf8 process_xls.py "你的数据文件.xls" --output-dir "E:\test"
```

#### 自定义输出文件名

```bash
# 6 个 Block → 提供 6 个名称
python -X utf8 process_xls.py "你的数据文件.xls" --names V1H1 V1S2 H2S1 0.5um 5nm 0.05nm --output-dir "E:\test"
```

#### 修改子板标签

```bash
python -X utf8 process_xls.py "你的数据文件.xls" --plate-names 样品A 样品B 样品C 样品D
```

---

## 📖 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input_file` | 输入 `.xls/.txt` 文件路径（必填） | — |
| `--output-dir DIR` | 输出目录 | `E:\test` |
| `--names N1 [N2 ...]` | 各 Block 的输出文件名（不含 `.xlsx` 后缀） | 使用原始板名 |
| `--plate-names P1 P2 P3 P4` | 四个子板的行标签 | `P1 P2 P3 P4` |

---

## 📊 输出格式说明

每个生成的 `.xlsx` 文件包含 **66 行 × 25 列**：

| 行号 | 内容 |
|------|------|
| 第 1 行 | 波长标签：`560/585`、`560/595`、`560/605`、`560/615` |
| 第 2 行 | 列头：`IA IB IC ID IE IF`（每组波长重复 4 次） |
| 第 3–18 行 | P1 数据 |
| 第 19–34 行 | P2 数据 |
| 第 35–50 行 | P3 数据 |
| 第 51–66 行 | P4 数据 |

> 💡 手动注释行（如 `(P4 only IL and QDS)`）需要自行添加

---

## 📁 项目结构

```
spectramax-processor/
├── easy_process.bat       # 一键入口：双击打开 GUI，或拖拽文件直接转换
├── easy_process.py        # 图形界面/拖拽友好包装器
├── install_dependencies.bat # 一键安装依赖
├── requirements.txt       # Python 依赖列表
├── process_xls.py          # 核心脚本
├── PROCESS_README.md       # 中文使用说明
├── README.md               # English README
├── .gitignore              # Git 忽略配置
└── LICENSE                 # MIT License
```

---

## ⚠️ 注意事项

1. **推荐入口**：日常使用优先双击或拖拽到 `easy_process.bat`
2. **编码问题**：命令行直接运行 `process_xls.py` 时，Windows PowerShell/CMD 下建议加 `-X utf8` 参数；使用 `easy_process.bat` 时不需要手动加
3. **数据安全**：输出文件不包含原始 `.xls/.txt`，请保留原始数据
4. **Block 数量**：脚本会自动检测，3 个、6 个或任意数量均可处理

---

## 📝 License

MIT License - 详见 [LICENSE](LICENSE) 文件

---

## 👤 作者

- GitHub: [bernis-web](https://github.com/bernis-web)
- Email: （可选，可自行添加）
