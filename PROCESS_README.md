# SpectraMax 数据转换使用说明

## 功能

将 SpectraMax 仪器导出的 `.xls` / `.txt` 文件（实为 UTF-16LE 编码的制表符分隔文本）
自动解析并重组为整理格式 `.xlsx`，**每个 Block 生成一个独立文件**。
主要为了方便导出之后不能直接用来做数据分析😋

## ✅ 最简单用法（推荐）

### 第一次使用

如果双击运行时提示缺少 `numpy` / `pandas` / `openpyxl`，先双击：

```text
install_dependencies.bat
```

它会自动执行：

```bash
python -m pip install -r requirements.txt
```

### 日常使用方式 A：双击图形界面

直接双击：

```text
easy_process.bat
```

然后按界面操作：

1. 点击 **添加文件...**，选择一个或多个 `.xls/.txt` 原始文件
2. 输出目录默认是 `E:\test`，一般不用改
3. 点击 **开始转换**
4. 转换完成后，到输出目录查看 `.xlsx` 文件

### 日常使用方式 B：拖拽文件

把一个或多个 `.xls/.txt` 原始文件直接拖到：

```text
easy_process.bat
```

脚本会自动转换并输出到 `E:\test`。这种方式不需要自己输入 `python -X utf8 ...`。

---

## 命令行用法（高级/备用）

```bash
python -X utf8 process_xls.py <输入文件.xls或txt> [选项]
```

### 选项

| 参数 | 说明 | 默认值 |
|---|---|---|
| `input_file` | 输入 `.xls/.txt` 文件路径（必填） | — |
| `--output-dir DIR` | 输出目录 | `E:\test` |
| `--names N1 [N2 ...]` | 各 Block 的输出文件名（不含 `.xlsx` 后缀），数量须与 Block 数量一致 | 使用 Block 原始板名 |
| `--plate-names P1 P2 P3 P4` | 四个子板的行标签 | `P1 P2 P3 P4` |

### 示例

```bash
# 基本用法（输出到 E:\test\，文件名用原始板名）
python -X utf8 "E:\workbuddy\process_xls.py" "E:\workbuddy\2026.4.25\混合比例测试与梯度稀释测试-560--585-615.xls" --output-dir "E:\test"

# txt 文件同样支持
python -X utf8 "E:\workbuddy\process_xls.py" "C:\Users\28417\Desktop\2026.5.9\cyt-c浓度梯度.txt" --output-dir "E:\test"

# 指定其他输出目录
python -X utf8 process_xls.py 实验数据.xls --output-dir "E:\test"

# 自定义输出文件名（6 个 Block 就提供 6 个名称）
python -X utf8 process_xls.py 实验数据.xls --names V1H1 V1S2 H2S1 0.5um 5nm 0.05nm --output-dir "E:\test"

# 修改行标签
python -X utf8 process_xls.py 实验数据.xls --plate-names 样品A 样品B 样品C 样品D --output-dir "E:\test"
```

## 输出格式

每个输出的 `.xlsx` 包含：

| 行号 | 内容 |
|---|---|
| 第 1 行 | 波长标签：`560/585`、`560/595`、`560/605`、`560/615` |
| 第 2 行 | 列头：`IA IB IC ID IE IF`（每组波长重复一次） |
| 第 3–18 行 | P1 数据（每行 24 个数值） |
| 第 19–34 行 | P2 数据 |
| 第 35–50 行 | P3 数据 |
| 第 51–66 行 | P4 数据 |

共 **66 行 × 25 列**。

## 注意事项

- 使用 `easy_process.bat` 时不需要手动输入 `-X utf8`，批处理文件已经自动加好
- 如果直接使用命令行运行 `process_xls.py`，Windows 环境建议加 `-X utf8`
- 文件编码通常为 UTF-16LE，脚本也会尝试 `utf-16` / `utf-8` / `gbk`
- Block 数量自动检测，**3 个、6 个或任意数量**均可处理
- 输出不含手动注释行（如 `(P4 only IL and QDS)`），整理后自行添加即可
- 需要依赖库：`numpy`、`pandas`、`openpyxl`；首次使用可双击 `install_dependencies.bat`
