# process_xls.py 使用说明

## 功能

将 SpectraMax 仪器导出的 `.xls` 文件（实为 UTF-16LE 编码的制表符分隔文本）
自动解析并重组为整理格式 `.xlsx`，**每个 Block 生成一个独立文件**。
主要为了方便导出之后不能直接用来做数据分析😋

## 用法

```bash
python -X utf8 process_xls.py <输入文件.xls> [选项]
```

### 选项

| 参数 | 说明 | 默认值 |
|---|---|---|
| `input_file` | 输入 `.xls` 文件路径（必填） | — |
| `--output-dir DIR` | 输出目录 | `E:\test` |
| `--names N1 [N2 ...]` | 各 Block 的输出文件名（不含 `.xlsx` 后缀），数量须与 Block 数量一致 | 使用 Block 原始板名 |
| `--plate-names P1 P2 P3 P4` | 四个子板的行标签 | `P1 P2 P3 P4` |

### 示例

```bash
# 基本用法（输出到 E:\test\，文件名用原始板名）
python -X utf8 "E:\workbuddy\process_xls.py" "E:\workbuddy\2026.4.25\混合比例测试与梯度稀释测试-560--585-615.xls" --output-dir "E:\test"

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

- 文件编码为 UTF-16LE，运行时需加 `-X utf8` 参数（Windows PowerShell 环境）
- Block 数量自动检测，**3 个、6 个或任意数量**均可处理
- 输出不含手动注释行（如 `(P4 only IL and QDS)`），整理后自行添加即可
- 需要 `openpyxl` 库：`pip install openpyxl`（如未安装）
