#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SpectraMax 原始数据一键转换入口。

推荐给日常实验使用：
  - 双击 easy_process.bat：打开图形界面，点选 .xls/.txt 文件即可转换
  - 拖拽 .xls/.txt 到 easy_process.bat：无需命令行，自动输出到默认目录

核心解析逻辑仍复用 process_xls.py，避免维护两套转换算法。
"""

import argparse
import os
import re
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Sequence


DEFAULT_PLATE_LABELS = ["P1", "P2", "P3", "P4"]
INVALID_FILENAME_CHARS = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
_CORE_FUNCS = None


@dataclass
class ProcessSummary:
    total_files: int = 0
    successful_files: int = 0
    blocks_written: int = 0
    failures: List[str] = field(default_factory=list)


def get_default_output_dir() -> Path:
    """优先使用实验电脑上的 E:\\test；没有 E 盘时退回到脚本目录下 output。"""
    if Path("E:/").exists():
        return Path("E:/test")
    return Path(__file__).resolve().parent / "output"


def import_core_functions():
    """延迟导入核心脚本，便于在缺依赖时给出友好提示。"""
    global _CORE_FUNCS
    if _CORE_FUNCS is not None:
        return _CORE_FUNCS

    try:
        from process_xls import build_dataframe, parse_raw_blocks, write_xlsx
    except ModuleNotFoundError as exc:
        missing = exc.name or "未知模块"
        if missing == "process_xls":
            message = "找不到核心脚本 process_xls.py，请确认 easy_process.py 与 process_xls.py 放在同一文件夹。"
        else:
            message = (
                f"缺少 Python 依赖库：{missing}\n"
                "请先双击 install_dependencies.bat，或在命令行运行：\n"
                "python -m pip install -r requirements.txt"
            )
        raise RuntimeError(message) from exc
    except ImportError as exc:
        raise RuntimeError(f"导入核心脚本失败：{exc}") from exc

    _CORE_FUNCS = (parse_raw_blocks, build_dataframe, write_xlsx)
    return _CORE_FUNCS


def read_raw_text(input_path: Path) -> str:
    """按 SpectraMax 常见编码读取原始文本。"""
    raw = input_path.read_bytes()
    for enc in ("utf-16-le", "utf-16", "utf-8-sig", "utf-8", "gbk"):
        try:
            text = raw.decode(enc)
            return text[1:] if text.startswith("\ufeff") else text
        except UnicodeDecodeError:
            continue
    raise ValueError("无法解码文件（已尝试 utf-16-le / utf-16 / utf-8 / gbk）。")


def sanitize_filename(name: str, fallback: str) -> str:
    """清理 Windows 文件名中的非法字符。"""
    cleaned = (name or "").strip()
    for ch in INVALID_FILENAME_CHARS:
        cleaned = cleaned.replace(ch, "_")
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or fallback


def unique_output_path(output_dir: Path, base_name: str) -> Path:
    """避免覆盖已有结果：若同名文件存在，自动追加 _2 / _3。"""
    candidate = output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    for index in range(2, 10000):
        candidate = output_dir / f"{base_name}_{index}.xlsx"
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法为 {base_name}.xlsx 生成不冲突的输出文件名。")


def normalize_input_paths(input_files: Iterable[str]) -> List[Path]:
    paths = []
    for item in input_files:
        raw = str(item).strip().strip('"')
        if not raw:
            continue
        path = Path(raw).expanduser()
        try:
            path = path.resolve()
        except OSError:
            pass
        paths.append(path)
    return paths


def process_files(
    input_files: Iterable[str],
    output_dir: str,
    plate_labels: Optional[Sequence[str]] = None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> ProcessSummary:
    """批量转换文件。供 GUI、拖拽入口和命令行入口共用。"""
    parse_raw_blocks, build_dataframe, write_xlsx = import_core_functions()

    paths = normalize_input_paths(input_files)
    labels = list(plate_labels or DEFAULT_PLATE_LABELS)
    if len(labels) != 4:
        raise ValueError("plate_labels 必须正好包含 4 个标签，例如：P1 P2 P3 P4")

    out_dir = Path(output_dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    summary = ProcessSummary(total_files=len(paths))
    use_source_prefix = len(paths) > 1

    def log(message: str) -> None:
        if log_callback:
            log_callback(message)
        else:
            print(message)

    for file_index, input_path in enumerate(paths, start=1):
        log(f"[{file_index}/{len(paths)}] 读取：{input_path}")

        if not input_path.exists():
            msg = f"文件不存在：{input_path}"
            summary.failures.append(msg)
            log(f"  [失败] {msg}")
            continue

        try:
            text = read_raw_text(input_path)
            blocks = parse_raw_blocks(text)
        except Exception as exc:  # noqa: BLE001 - GUI/批处理入口需要不中断后续文件
            msg = f"{input_path.name}: 读取或解析失败：{exc}"
            summary.failures.append(msg)
            log(f"  [失败] {msg}")
            continue

        if not blocks:
            msg = f"{input_path.name}: 未找到有效数据 Block"
            summary.failures.append(msg)
            log(f"  [失败] {msg}")
            continue

        log(f"  发现 {len(blocks)} 个 Block")
        source_prefix = sanitize_filename(input_path.stem, f"file_{file_index}")
        written_for_this_file = 0

        for block_index, block in enumerate(blocks, start=1):
            block_name = sanitize_filename(block.get("name", ""), f"block_{block_index}")
            if use_source_prefix:
                base_name = sanitize_filename(f"{source_prefix}__{block_name}", f"file_{file_index}_block_{block_index}")
            else:
                base_name = block_name

            try:
                output_path = unique_output_path(out_dir, base_name)
                df = build_dataframe(block, plate_labels=labels)
                write_xlsx(df, str(output_path))
            except Exception as exc:  # noqa: BLE001
                msg = f"{input_path.name} / Block {block_index} ({block_name}): 写出失败：{exc}"
                summary.failures.append(msg)
                log(f"    [失败] {msg}")
                continue

            written_for_this_file += 1
            summary.blocks_written += 1
            log(f"    [OK] [{block_index}/{len(blocks)}] {output_path}")

        if written_for_this_file:
            summary.successful_files += 1

    log("")
    log(f"完成：成功写出 {summary.blocks_written} 个 xlsx 文件。输出目录：{out_dir.resolve()}")
    if summary.failures:
        log(f"注意：有 {len(summary.failures)} 个问题，详情见上方日志。")
    return summary


def split_plate_labels(label_text: str) -> List[str]:
    """支持用空格、逗号或中文逗号分隔 4 个子板标签。"""
    return [part for part in re.split(r"[\s,，]+", label_text.strip()) if part]


def run_gui(default_output_dir: Optional[str] = None) -> int:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except Exception as exc:  # noqa: BLE001
        print(f"无法启动图形界面：{exc}", file=sys.stderr)
        print("你仍可把文件拖到 easy_process.bat 上，或使用 process_xls.py 命令行。", file=sys.stderr)
        return 1

    class EasyProcessApp:
        def __init__(self, root: "tk.Tk") -> None:
            self.root = root
            self.files: List[str] = []
            self.output_var = tk.StringVar(value=default_output_dir or str(get_default_output_dir()))
            self.plate_var = tk.StringVar(value=" ".join(DEFAULT_PLATE_LABELS))

            root.title("SpectraMax 一键转换")
            root.geometry("820x560")
            root.minsize(760, 500)

            main = ttk.Frame(root, padding=12)
            main.pack(fill="both", expand=True)
            main.columnconfigure(0, weight=1)
            main.rowconfigure(2, weight=1)
            main.rowconfigure(7, weight=1)

            title = ttk.Label(main, text="SpectraMax 原始数据一键转换", font=("Microsoft YaHei UI", 15, "bold"))
            title.grid(row=0, column=0, sticky="w")

            hint = ttk.Label(
                main,
                text="选择 .xls 或 .txt 原始文件 → 确认输出目录 → 点击开始转换。默认输出到 E:\\test。",
            )
            hint.grid(row=1, column=0, sticky="w", pady=(4, 10))

            file_frame = ttk.LabelFrame(main, text="1. 输入文件（可多选）", padding=8)
            file_frame.grid(row=2, column=0, sticky="nsew")
            file_frame.columnconfigure(0, weight=1)
            file_frame.rowconfigure(1, weight=1)

            file_buttons = ttk.Frame(file_frame)
            file_buttons.grid(row=0, column=0, sticky="ew", pady=(0, 6))
            ttk.Button(file_buttons, text="添加文件...", command=self.add_files).pack(side="left")
            ttk.Button(file_buttons, text="清空列表", command=self.clear_files).pack(side="left", padx=(8, 0))

            list_frame = ttk.Frame(file_frame)
            list_frame.grid(row=1, column=0, sticky="nsew")
            list_frame.columnconfigure(0, weight=1)
            list_frame.rowconfigure(0, weight=1)
            self.file_listbox = tk.Listbox(list_frame, height=6)
            self.file_listbox.grid(row=0, column=0, sticky="nsew")
            file_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
            file_scroll.grid(row=0, column=1, sticky="ns")
            self.file_listbox.configure(yscrollcommand=file_scroll.set)

            output_frame = ttk.LabelFrame(main, text="2. 输出设置", padding=8)
            output_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
            output_frame.columnconfigure(1, weight=1)

            ttk.Label(output_frame, text="输出目录：").grid(row=0, column=0, sticky="w")
            ttk.Entry(output_frame, textvariable=self.output_var).grid(row=0, column=1, sticky="ew", padx=(6, 6))
            ttk.Button(output_frame, text="选择...", command=self.choose_output_dir).grid(row=0, column=2)
            ttk.Button(output_frame, text="打开输出目录", command=self.open_output_dir).grid(row=0, column=3, padx=(6, 0))

            ttk.Label(output_frame, text="子板标签：").grid(row=1, column=0, sticky="w", pady=(8, 0))
            ttk.Entry(output_frame, textvariable=self.plate_var).grid(row=1, column=1, sticky="ew", padx=(6, 6), pady=(8, 0))
            ttk.Label(output_frame, text="一般不用改，默认 P1 P2 P3 P4").grid(row=1, column=2, columnspan=2, sticky="w", pady=(8, 0))

            action_frame = ttk.Frame(main)
            action_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
            self.start_button = ttk.Button(action_frame, text="开始转换", command=self.start_processing)
            self.start_button.pack(side="left")
            self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
            self.progress.pack(side="left", fill="x", expand=True, padx=(12, 0))

            log_frame = ttk.LabelFrame(main, text="运行日志", padding=8)
            log_frame.grid(row=7, column=0, sticky="nsew", pady=(10, 0))
            log_frame.columnconfigure(0, weight=1)
            log_frame.rowconfigure(0, weight=1)
            self.log_text = tk.Text(log_frame, height=9, wrap="word")
            self.log_text.grid(row=0, column=0, sticky="nsew")
            log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
            log_scroll.grid(row=0, column=1, sticky="ns")
            self.log_text.configure(yscrollcommand=log_scroll.set)

            self.append_log("准备就绪。首次使用如提示缺少依赖，请双击 install_dependencies.bat。")

        def add_files(self) -> None:
            selected = filedialog.askopenfilenames(
                parent=self.root,
                title="选择 SpectraMax 原始文件",
                filetypes=[("SpectraMax 原始文件", "*.xls *.txt"), ("所有文件", "*.*")],
            )
            for path in selected:
                if path not in self.files:
                    self.files.append(path)
                    self.file_listbox.insert("end", path)

        def clear_files(self) -> None:
            self.files.clear()
            self.file_listbox.delete(0, "end")

        def choose_output_dir(self) -> None:
            initial_dir = self.output_var.get() or str(get_default_output_dir())
            if not Path(initial_dir).exists():
                initial_dir = str(Path(__file__).resolve().parent)
            selected = filedialog.askdirectory(parent=self.root, title="选择输出目录", initialdir=initial_dir)
            if selected:
                self.output_var.set(selected)

        def open_output_dir(self) -> None:
            out_dir = Path(self.output_var.get()).expanduser()
            out_dir.mkdir(parents=True, exist_ok=True)
            try:
                if os.name == "nt":
                    os.startfile(str(out_dir))  # type: ignore[attr-defined]
                else:
                    import subprocess

                    subprocess.Popen(["xdg-open", str(out_dir)])
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("无法打开目录", str(exc), parent=self.root)

        def append_log(self, message: str) -> None:
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")

        def thread_safe_log(self, message: str) -> None:
            self.root.after(0, self.append_log, message)

        def start_processing(self) -> None:
            if not self.files:
                messagebox.showwarning("请先选择文件", "请先添加至少一个 .xls 或 .txt 原始文件。", parent=self.root)
                return

            labels = split_plate_labels(self.plate_var.get())
            if len(labels) != 4:
                messagebox.showwarning(
                    "子板标签数量不对",
                    "子板标签必须正好 4 个，例如：P1 P2 P3 P4",
                    parent=self.root,
                )
                return

            self.start_button.configure(state="disabled")
            self.progress.start(10)
            self.append_log("-" * 60)
            self.append_log("开始转换...")

            worker = threading.Thread(target=self._process_worker, args=(list(self.files), labels), daemon=True)
            worker.start()

        def _process_worker(self, files: List[str], labels: List[str]) -> None:
            try:
                summary = process_files(files, self.output_var.get(), labels, log_callback=self.thread_safe_log)
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, self._processing_failed, exc)
                return
            self.root.after(0, self._processing_finished, summary)

        def _processing_failed(self, exc: Exception) -> None:
            self.progress.stop()
            self.start_button.configure(state="normal")
            self.append_log(f"[失败] {exc}")
            messagebox.showerror("转换失败", str(exc), parent=self.root)

        def _processing_finished(self, summary: ProcessSummary) -> None:
            self.progress.stop()
            self.start_button.configure(state="normal")
            if summary.failures:
                messagebox.showwarning(
                    "转换完成（有问题）",
                    f"成功写出 {summary.blocks_written} 个 xlsx。\n"
                    f"有 {len(summary.failures)} 个问题，详情请查看运行日志。",
                    parent=self.root,
                )
            else:
                messagebox.showinfo(
                    "转换完成",
                    f"成功写出 {summary.blocks_written} 个 xlsx。\n输出目录：{Path(self.output_var.get()).resolve()}",
                    parent=self.root,
                )

    root = tk.Tk()
    EasyProcessApp(root)
    root.mainloop()
    return 0


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="SpectraMax 原始数据一键转换入口（双击/拖拽/GUI 友好版）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input_files", nargs="*", help="输入 .xls/.txt 文件；不提供时启动图形界面")
    parser.add_argument("--output-dir", default=str(get_default_output_dir()), help="输出目录（默认优先 E:\\test）")
    parser.add_argument("--plate-names", nargs=4, default=DEFAULT_PLATE_LABELS, help="四个子板标签，默认 P1 P2 P3 P4")
    parser.add_argument("--gui", action="store_true", help="强制启动图形界面")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    if args.gui or not args.input_files:
        return run_gui(default_output_dir=args.output_dir)

    try:
        summary = process_files(args.input_files, args.output_dir, args.plate_names)
    except Exception as exc:  # noqa: BLE001
        print(f"[失败] {exc}", file=sys.stderr)
        return 1

    return 0 if summary.blocks_written > 0 and not summary.failures else 1


if __name__ == "__main__":
    sys.exit(main())