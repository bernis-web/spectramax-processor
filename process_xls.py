#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SpectraMax .xls 原始数据 → 整理格式 xlsx 转换脚本

用法:
    python process_xls.py <输入文件.xls> [选项]

选项:
    --output-dir DIR      输出目录（默认: ./output）
    --names N1 [N2 ...]   各 Block 的输出文件名（不含 .xlsx 后缀）
                          若不指定，则使用 Block 原始板名
    --plate-names P1 P2 P3 P4  四个子板的行标签（默认: P1 P2 P3 P4）

示例:
    python process_xls.py 原始数据.xls
    python process_xls.py 原始数据.xls --output-dir ./results
    python process_xls.py 原始数据.xls --names V1H1 V1S2 H2S1 0.5um 5nm 0.05nm
"""

import argparse
import math
import os
import re
import sys

import openpyxl
import numpy as np
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# 核心解析逻辑
# ─────────────────────────────────────────────────────────────────────────────

def parse_raw_blocks(raw_text: str):
    """
    将原始 .xls 文本解析为若干 Block。

    原始文件结构（已确认）:
      - 每个 Block 由 "~End" 分隔
      - 第 0 个 ~End 段是协议说明（跳过）
      - 每个数据 Block:
          行 0:  Plate:\t{name}\t...        ← 板名
          行 1:  \tTemperature(°C)\t1\t2...24\t\t1\t2...24  ← 列头
          行 2-17: 16 行数据
              col0=空, col1=温度(仅首行), cols[2:26]=585nm(24值),
              col26=空, cols[27:51]=595nm, col51=空,
              cols[52:76]=605nm, col76=空, cols[77:101]=615nm
              每组 24 值 = 4个子板×6列(IA-IF)

    返回:
        list[ dict(name=str, data={
            '585': [[float,...]×16行],
            '595': [[float,...]×16行],
            '605': [[float,...]×16行],
            '615': [[float,...]×16行],
        }) ]
    """
    segments = raw_text.split("~End")
    blocks_raw = segments[1:]  # 跳过协议说明

    blocks = []
    for seg in blocks_raw:
        lines = [ln for ln in seg.split("\r\n") if ln.strip()]
        if not lines:
            continue

        # 跳过"原始文件名"等元数据段（无 Plate: 前缀）
        if not lines[0].startswith("Plate:"):
            continue

        # 提取板名
        first_cols = lines[0].split("\t")
        plate_name = first_cols[1] if len(first_cols) > 1 else ""

        # 取数据行（跳过板名行 lines[0] 和列头行 lines[1]）
        data_lines = lines[2:]
        if len(data_lines) < 16:
            print(f"  [警告] Block '{plate_name}' 数据行数={len(data_lines)}（期望16），跳过。",
                  file=sys.stderr)
            continue

        # 只取前 16 行
        data_lines = data_lines[:16]

        wl_names = ["585", "595", "605", "615"]
        # 每个波长组在行中的列范围（Python slice 含头不含尾）
        wl_slices = [(2, 26), (27, 51), (52, 76), (77, 101)]

        block_data = {wl: [[] for _ in range(16)] for wl in wl_names}

        for row_idx, row in enumerate(data_lines):
            cols = row.split("\t")
            for wl_name, (start, end) in zip(wl_names, wl_slices):
                raw_vals = cols[start:end]  # 24 个字符串
                floats = []
                for v in raw_vals:
                    v = v.strip()
                    if v:
                        try:
                            floats.append(float(v))
                        except ValueError:
                            floats.append(math.nan)
                    else:
                        floats.append(math.nan)
                block_data[wl_name][row_idx] = floats

        blocks.append({"name": plate_name, "data": block_data})

    return blocks


def build_dataframe(block: dict, plate_labels=None):
    """
    将单个 Block 的数据重组为整理格式 DataFrame。

    输出格式（67 行 × 25 列）：
      行 0:   NaN, 560/585×6, NaN, 560/595×6, 560/605×6, 560/615×6
      行 1:   NaN, IA..IF×4组
      行 2-17:   P1 + 6列×4波长
      行 18-33:  P2 + ...
      行 34-49:  P3 + ...
      行 50-65:  P4 + ...
    """
    if plate_labels is None:
        plate_labels = ["P1", "P2", "P3", "P4"]

    wl_names = ["585", "595", "605", "615"]

    # 每个子板在 24 列组中的偏移（P1→0:6, P2→6:12, P3→12:18, P4→18:24）
    plate_slices = [(0, 6), (6, 12), (12, 18), (18, 24)]

    rows = []

    # ── 行 0：波长标签行 ──
    row0 = [np.nan]
    for wl in wl_names:
        row0.extend(["560/" + wl] + [np.nan] * 5)
    rows.append(row0)

    # ── 行 1：IA-IF 列头行 ──
    row1 = [np.nan]
    for _ in wl_names:
        row1.extend(["IA", "IB", "IC", "ID", "IE", "IF"])
    rows.append(row1)

    # ── 数据行：4 子板 × 16 行 ──
    for plate_idx, plate_lbl in enumerate(plate_labels):
        ps, pe = plate_slices[plate_idx]
        for row_idx in range(16):
            out_row = [plate_lbl]
            for wl_name in wl_names:
                wl_24 = block["data"][wl_name][row_idx]
                out_row.extend(wl_24[ps:pe])
            rows.append(out_row)

    df = pd.DataFrame(rows)
    return df


def write_xlsx(df: pd.DataFrame, output_path: str):
    """
    将 DataFrame 写入 xlsx（含两行表头格式）。
    df 预期 67 行 × 25 列。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx + 1, column=c_idx, value=val)

    wb.save(output_path)


# ─────────────────────────────────────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="将 SpectraMax .xls 原始数据转换为整理格式 xlsx",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("input_file", help="输入 .xls 文件路径")
    parser.add_argument(
        "--output-dir", default="E:\\test",
        help="输出目录（默认: E:\\test）"
    )
    parser.add_argument(
        "--names", nargs="+",
        help="各 Block 的输出文件名（不含 .xlsx 后缀），需与 Block 数量一致"
    )
    parser.add_argument(
        "--plate-names", nargs=4, default=["P1", "P2", "P3", "P4"],
        help="四个子板的行标签（默认: P1 P2 P3 P4）"
    )

    args = parser.parse_args()

    # ── 读取文件 ────────────────────────────────────────────────────────────
    input_path = args.input_file
    if not os.path.exists(input_path):
        print(f"错误: 文件不存在: {input_path}", file=sys.stderr)
        sys.exit(1)

    with open(input_path, "rb") as f:
        raw = f.read()

    # 尝试解码
    text = None
    for enc in ("utf-16-le", "utf-16", "utf-8", "gbk"):
        try:
            text = raw.decode(enc)
            break
        except Exception:
            continue
    if text is None:
        print("错误: 无法解码文件（尝试了 utf-16-le/utf-16/utf-8/gbk）。", file=sys.stderr)
        sys.exit(1)

    if text.startswith("\ufeff"):
        text = text[1:]

    # ── 解析所有 Block ─────────────────────────────────────────────────────
    blocks = parse_raw_blocks(text)
    if not blocks:
        print("错误: 未在文件中找到任何有效的数据 Block。", file=sys.stderr)
        sys.exit(1)

    print(f"发现 {len(blocks)} 个数据 Block:")
    for i, b in enumerate(blocks):
        print(f"  [{i+1}] {b['name']!r}")
    print()

    # ── 准备输出目录 ─────────────────────────────────────────────────────────
    output_dir = os.path.abspath(args.output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # ── 验证 --names 参数 ──────────────────────────────────────────────────
    if args.names and len(args.names) != len(blocks):
        print(f"警告: --names 提供了 {len(args.names)} 个名称，"
              f"但文件有 {len(blocks)} 个 Block。将使用 Block 原始名称。",
              file=sys.stderr)
        args.names = None

    # ── 逐 Block 处理并输出 ───────────────────────────────────────────────
    for i, block in enumerate(blocks):
        # 确定输出文件名
        if args.names and i < len(args.names):
            base_name = args.names[i]
        else:
            base_name = block["name"].strip()
            # 替换 Windows 非法字符
            for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                base_name = base_name.replace(ch, "_")

        output_path = os.path.join(output_dir, base_name + ".xlsx")

        # 构建 DataFrame 并写入
        df = build_dataframe(block, plate_labels=args.plate_names)
        write_xlsx(df, output_path)

        print(f"  [OK] [{i+1}/{len(blocks)}] {output_path}")

    print(f"\nDone. {len(blocks)} blocks -> {output_dir}")


if __name__ == "__main__":
    main()
