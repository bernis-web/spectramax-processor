# WorkBuddy Memory

## 实验数据处理

### SpectraMax .xls 数据格式

- 仪器导出文件实为 **UTF-16LE 编码的制表符分隔文本**，非真正 .xls
- 文件用 `~End` 分隔多个 Block，数量不固定（3 个、6 个或更多）
- 每个数据 Block：
  - 16 行数据，每行 103 列（制表符分隔）
  - 列结构：col0=空, col1=温度, cols[2:26]=585nm(24值), col26=空, cols[27:51]=595nm, col51=空, cols[52:76]=605nm, col76=空, cols[77:101]=615nm
  - 每组 24 值 = 4 个子板 × 6 列(IA-IF)：P1→[0:6], P2→[6:12], P3→[12:18], P4→[18:24]

### 整理后 xlsx 目标格式

- 67 行 × 25 列：
  - 行 0：波长标签（NaN, 560/585×6, NaN, 560/595×6, 560/605×6, 560/615×6）
  - 行 1：IA-IF 列头 × 4 组
  - 行 2-65：4 个子板各 16 行数据，col0=板标签（P1/P2/P3/P4）
- 末尾的 `(P4 only IL and QDS)` 等注释为用户手动添加，脚本不处理

### 脚本 process_xls.py

- 位置：`e:\workbuddy\process_xls.py`
- 用法：`python -X utf8 process_xls.py 输入.xls --output-dir ./output --names 名1 名2 ...`
- 自动检测 Block 数量，有几个处理几个
- `--plate-names` 可自定义每行的板标签（默认 P1 P2 P3 P4）
