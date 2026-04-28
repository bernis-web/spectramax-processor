---
name: xls-data-processor
overview: 开发一个 Python 脚本，将 SpectraMax 仪器导出的 .xls 文件（UTF-16LE 文本格式）自动解析并重组为指定的 xlsx 格式，每个 Block 生成一个独立的 xlsx 文件。
todos:
  - id: write-parse-function
    content: 实现 parse_raw_block() 解析函数：从 Block 原始文本中提取 4 个波长 × 16 行 × 24 列的数值数组
    status: completed
  - id: write-build-dataframe
    content: 实现 build_dataframe() 重组函数：将解析结果转换为输出格式 DataFrame（67行×25列，含表头）
    status: completed
  - id: write-write-xlsx
    content: 实现 write_xlsx() 写入函数：将 DataFrame 按指定格式写入 xlsx（含两行表头：波长标签行 + IA-IF 列头行）
    status: completed
  - id: write-main-cli
    content: 实现 main() 主函数：命令行参数解析、文件读写流程、循环处理 6 个 Block 输出
    status: completed
  - id: test-with-sample-file
    content: 用原始 .xls 文件测试脚本，验证 6 个输出 xlsx 的数据正确性
    status: completed
---

## 用户需求

将 SpectraMax 荧光仪导出的 .xls 文件（UTF-16LE 编码制表符分隔文本）自动解析为整理格式的 xlsx 文件。

## 核心功能

1. **解析原始 .xls 文件**：读取 UTF-16LE 编码文本，按 `~End` 分隔提取 6 个数据 Block
2. **数据重组**：将 384 孔板格式（16行×24列/波长组）重组为整理格式（4个子板 P1-P4 各 16 行 × 6 列 IA-IF × 4 波长）
3. **输出 xlsx**：每个 Block 生成一个独立 xlsx 文件，表头含波长标签（560/585 等）和孔标识（IA-IF），数据行标签为通用 P1/P2/P3/P4
4. **命令行参数**：支持指定输出目录和输出文件名（可选，不指定则用 Block 原始名称）
5. **自动适配 Block 数量**：脚本自动检测文件中实际的 Block 数量（可能为 3 个、6 个或其他），有几个处理几个

## 技术栈

- **语言**：Python 3
- **核心库**：`openpyxl`（xlsx 写入）、标准库（文件读写、正则）
- **无需额外安装**：openpyxl 通常已预装

## 实现方案

### 核心解析逻辑

```
原始格式（每个 Block）：
  16 行数据，每行 103 列
  col0=空, col1=温度(仅首行), cols[2:26]=585nm(24值), col26=空,
  cols[27:51]=595nm, col51=空, cols[52:76]=605nm, col76=空, cols[77:101]=615nm

  每组 24 值 = 4个子板 × 6列(IA-IF)
    子板 P1: indices[0:6]
    子板 P2: indices[6:12]
    子板 P3: indices[12:18]
    子板 P4: indices[18:24]

输出格式（每个 xlsx）：
  行 0: wavelength header (NaN, 560/585×6, NaN, 560/595×6, 560/605×6, 560/615×6)
  行 1: column header (NaN, IA, IB, IC, ID, IE, IF × 4组)
  行 2-17:  P1 数据 (col0=P1, cols[1:25]=6列×4波长)
  行 18-33: P2 数据 (col0=P2, ...)
  行 34-49: P3 数据 (col0=P3, ...)
  行 50-65: P4 数据 (col0=P4, ...)
  总计 67 行 × 25 列
```

### Block 名称映射

原始 .xls 中的 Block 板名（plate_name）：

- Block 1: `1-4`
- Block 2: `p5-p8`
- Block 3: `p9-p12`
- Block 4: `梯度p1-p4`
- Block 5: `梯度p5-p8`
- Block 6: `梯度p9-p12`

用户可指定自定义文件名，也可使用默认板名作为文件名。

### 性能与边界

- 无 N+1 问题，单次顺序解析
- 空值列（col25/50/75 分隔符）跳过
- 温度行（第 1 行数据行的 col1）丢弃，不写入

## 目录结构

```
e:/workbuddy/
├── process_xls.py          # [NEW] 主脚本
└── output/                  # [NEW] 输出目录（命令行参数指定）
    ├── 1-4.xlsx              # Block 1 输出
    ├── p5-p8.xlsx            # Block 2 输出
    ├── p9-p12.xlsx           # Block 3 输出
    ├── 梯度p1-p4.xlsx        # Block 4 输出
    ├── 梯度p5-p8.xlsx        # Block 5 输出
    └── 梯度p9-p12.xlsx       # Block 6 输出
```

## 关键代码结构

```python
# process_xls.py

def parse_raw_block(raw_text: str) -> dict:
    """解析单个 Block，返回 {wavelengths: {name: [[24 vals per row] x 16 rows]}}"""
    ...

def build_dataframe(block_data: dict) -> pd.DataFrame:
    """将 block_data 重组为输出格式 DataFrame"""
    # 4 sub-plates × 16 rows × 4 wavelengths × 6 cols = 64 rows × 24 data cols
    ...

def write_xlsx(df: pd.DataFrame, output_path: str):
    """写入 xlsx，包含两行表头"""
    ...

def main():
    # 解析命令行参数（input_file, --output-dir, --names）
    # 读取文件（UTF-16LE decode）
    # 按 ~End 分割提取 6 个 Block
    # 对每个 Block 调用 parse_raw_block + build_dataframe + write_xlsx
```