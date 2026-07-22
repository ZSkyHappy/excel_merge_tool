from __future__ import annotations

import ctypes
import os
import queue
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from merge_engine import (
    MergeCancelled,
    MergeConfig,
    MergeError,
    MergeSummary,
    ProgressEvent,
    find_input_files,
    find_output_conflicts,
    merge_workbooks,
    normalize_output_name,
)


APP_TITLE = "Excel 批量合并工具"


def _enable_windows_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


class ExcelMergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("760x560")
        self.root.minsize(700, 520)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.input_dir_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.output_name_var = tk.StringVar(value=self._default_output_name())
        self.status_var = tk.StringVar(value="请选择包含 .xlsx 文件的文件夹。")
        self.progress_text_var = tk.StringVar(value="尚未开始")
        self.result_dir: Path | None = None

        self.cancel_event = threading.Event()
        self.event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.running = False
        self.close_requested = False

        self._configure_style()
        self._build_ui()
        self.root.after(100, self._poll_events)

    @staticmethod
    def _default_output_name() -> str:
        return f"合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 17, "bold"))
        style.configure("Hint.TLabel", foreground="#555555")
        style.configure("Status.TLabel", foreground="#1F4E78")
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=20)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(8, weight=1)

        ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").grid(
            row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 4)
        )
        ttk.Label(
            outer,
            text="合并每个文件第一个可见工作表的 A:Q，并在 R 列写入“文件名 - 主讲老师”。",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 18))

        ttk.Label(outer, text="输入文件夹").grid(row=2, column=0, sticky=tk.W, pady=6)
        self.input_entry = ttk.Entry(outer, textvariable=self.input_dir_var)
        self.input_entry.grid(row=2, column=1, sticky=tk.EW, padx=10, pady=6)
        self.input_button = ttk.Button(outer, text="选择…", command=self._choose_input_dir)
        self.input_button.grid(row=2, column=2, sticky=tk.EW, pady=6)

        ttk.Label(outer, text="输出文件夹").grid(row=3, column=0, sticky=tk.W, pady=6)
        self.output_entry = ttk.Entry(outer, textvariable=self.output_dir_var)
        self.output_entry.grid(row=3, column=1, sticky=tk.EW, padx=10, pady=6)
        self.output_button = ttk.Button(outer, text="选择…", command=self._choose_output_dir)
        self.output_button.grid(row=3, column=2, sticky=tk.EW, pady=6)

        ttk.Label(outer, text="输出文件名").grid(row=4, column=0, sticky=tk.W, pady=6)
        self.name_entry = ttk.Entry(outer, textvariable=self.output_name_var)
        self.name_entry.grid(row=4, column=1, columnspan=2, sticky=tk.EW, padx=(10, 0), pady=6)

        ttk.Separator(outer).grid(row=5, column=0, columnspan=3, sticky=tk.EW, pady=14)

        progress_frame = ttk.Frame(outer)
        progress_frame.grid(row=6, column=0, columnspan=3, sticky=tk.EW)
        progress_frame.columnconfigure(0, weight=1)
        ttk.Label(progress_frame, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Label(progress_frame, textvariable=self.progress_text_var).grid(
            row=0, column=1, sticky=tk.E
        )
        self.progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(6, 10))

        button_frame = ttk.Frame(outer)
        button_frame.grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=(0, 10))
        self.start_button = ttk.Button(
            button_frame,
            text="开始合并",
            style="Accent.TButton",
            command=self._start_merge,
        )
        self.start_button.pack(side=tk.LEFT)
        self.cancel_button = ttk.Button(
            button_frame,
            text="取消",
            command=self._cancel_merge,
            state=tk.DISABLED,
        )
        self.cancel_button.pack(side=tk.LEFT, padx=8)
        self.open_button = ttk.Button(
            button_frame,
            text="打开输出文件夹",
            command=self._open_output_dir,
            state=tk.DISABLED,
        )
        self.open_button.pack(side=tk.RIGHT)

        log_frame = ttk.LabelFrame(outer, text="处理记录", padding=8)
        log_frame.grid(row=8, column=0, columnspan=3, sticky=tk.NSEW)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_frame,
            height=10,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 9),
            state=tk.DISABLED,
        )
        self.log_text.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _choose_input_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
        if not selected:
            return
        self.input_dir_var.set(selected)
        if not self.output_dir_var.get().strip():
            self.output_dir_var.set(selected)
        self._append_log(f"已选择输入文件夹：{selected}")

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择结果保存文件夹")
        if selected:
            self.output_dir_var.set(selected)
            self._append_log(f"已选择输出文件夹：{selected}")

    def _start_merge(self) -> None:
        if self.running:
            return
        try:
            input_dir = Path(self.input_dir_var.get().strip()).expanduser().resolve()
            output_dir_text = self.output_dir_var.get().strip()
            output_dir = Path(output_dir_text).expanduser().resolve()
            output_name = normalize_output_name(self.output_name_var.get())
            if not input_dir.is_dir():
                raise MergeError("请选择有效的输入文件夹。")
            if not output_dir_text:
                raise MergeError("请选择输出文件夹。")
            output_dir.mkdir(parents=True, exist_ok=True)

            conflicts = find_output_conflicts(output_dir, output_name)
            excluded = conflicts if input_dir == output_dir else []
            files = find_input_files(input_dir, excluded_paths=excluded)
            if not files:
                raise MergeError("输入文件夹第一层没有可合并的 .xlsx 文件。")

            overwrite = False
            if conflicts:
                shown = "\n".join(f"• {path.name}" for path in conflicts[:8])
                if len(conflicts) > 8:
                    shown += f"\n• 另有 {len(conflicts) - 8} 个分卷"
                overwrite = messagebox.askyesno(
                    "确认覆盖旧结果",
                    "以下旧结果将在新结果完整生成后被替换：\n\n"
                    f"{shown}\n\n输入文件不会被修改。是否继续？",
                    icon=messagebox.WARNING,
                )
                if not overwrite:
                    return

            config = MergeConfig(
                input_dir=input_dir,
                output_dir=output_dir,
                output_name=output_name,
                overwrite=overwrite,
            )
        except (MergeError, OSError) as exc:
            messagebox.showerror("无法开始", str(exc))
            return

        self.output_name_var.set(output_name)
        self.cancel_event = threading.Event()
        self.result_dir = None
        self.running = True
        self.close_requested = False
        self._set_controls_running(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self.status_var.set(f"正在预检 {len(files)} 个文件……")
        self.progress_text_var.set("0 个文件")
        self._clear_log()
        self._append_log(f"发现 {len(files)} 个待处理文件。")

        self.worker = threading.Thread(
            target=self._run_merge,
            args=(config,),
            name="excel-merge-worker",
            daemon=True,
        )
        self.worker.start()

    def _run_merge(self, config: MergeConfig) -> None:
        try:
            summary = merge_workbooks(
                config,
                cancel_event=self.cancel_event,
                progress_callback=lambda event: self.event_queue.put(("progress", event)),
            )
            self.event_queue.put(("success", summary))
        except MergeCancelled as exc:
            self.event_queue.put(("cancelled", str(exc)))
        except Exception as exc:
            details = "".join(traceback.format_exception(exc))
            self.event_queue.put(("error", (str(exc), details)))

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.event_queue.get_nowait()
                if kind == "progress":
                    self._handle_progress(payload)
                elif kind == "success":
                    self._handle_success(payload)
                elif kind == "cancelled":
                    self._handle_cancelled(str(payload))
                elif kind == "error":
                    message, details = payload
                    self._handle_error(str(message), str(details))
        except queue.Empty:
            pass

        if self.close_requested and not self.running:
            self.root.destroy()
            return
        self.root.after(100, self._poll_events)

    def _handle_progress(self, event: object) -> None:
        if not isinstance(event, ProgressEvent):
            return
        self.status_var.set(event.message)
        if event.phase in {"scan_start", "scan_file", "scan_done"}:
            if self.progress.cget("mode") != "indeterminate":
                self.progress.configure(mode="indeterminate")
                self.progress.start(10)
            if event.files_total:
                self.progress_text_var.set(
                    f"预检 {event.files_completed}/{event.files_total} 个文件"
                )
        else:
            if self.progress.cget("mode") != "determinate":
                self.progress.stop()
                self.progress.configure(mode="determinate")
            maximum = max(event.total_rows, 1)
            self.progress.configure(maximum=maximum, value=min(event.rows_written, maximum))
            self.progress_text_var.set(
                f"{event.rows_written:,}/{event.total_rows:,} 行"
                if event.total_rows
                else f"{event.rows_written:,} 行"
            )
        if event.phase in {"scan_file", "merge_file", "saving", "done"}:
            self._append_log(event.message)

    def _handle_success(self, payload: object) -> None:
        if not isinstance(payload, MergeSummary):
            return
        self._finish_running_state()
        self.result_dir = payload.report_path.parent
        self.open_button.configure(state=tk.NORMAL)
        self.progress.configure(mode="determinate", maximum=1, value=1)
        self.progress_text_var.set(f"共 {payload.total_rows:,} 行")

        skipped = sum(1 for item in payload.file_results if item.status == "跳过")
        outputs = "\n".join(path.name for path in payload.output_files) or "无 Excel 输出"
        self._append_log(f"合并完成，耗时 {payload.elapsed_seconds:.2f} 秒。")
        self._append_log(f"报告：{payload.report_path.name}")
        message = (
            f"合并完成，共写入 {payload.total_rows:,} 行。\n\n"
            f"输出：\n{outputs}\n\n"
            f"跳过文件：{skipped} 个\n"
            f"详细信息见：{payload.report_path.name}"
        )
        if skipped:
            messagebox.showwarning("合并完成（有文件被跳过）", message)
        else:
            messagebox.showinfo("合并完成", message)

    def _handle_cancelled(self, message: str) -> None:
        self._finish_running_state()
        self.status_var.set("已取消，未发布半成品。")
        self.progress.configure(mode="determinate", value=0, maximum=100)
        self.progress_text_var.set("已取消")
        self._append_log(message)
        if not self.close_requested:
            messagebox.showinfo("已取消", "本次合并已取消，没有发布半成品。")

    def _handle_error(self, message: str, details: str) -> None:
        self._finish_running_state()
        self.status_var.set("合并失败，未发布半成品。")
        self.progress.configure(mode="determinate", value=0, maximum=100)
        self.progress_text_var.set("失败")
        self._append_log(f"错误：{message}")
        self._write_crash_log(details)
        if not self.close_requested:
            messagebox.showerror(
                "合并失败",
                f"{message}\n\n没有发布半成品。程序目录或输出目录中已保存错误日志。",
            )

    def _write_crash_log(self, details: str) -> None:
        candidates = [
            Path(self.output_dir_var.get().strip()),
            Path.cwd(),
        ]
        for directory in candidates:
            try:
                directory.mkdir(parents=True, exist_ok=True)
                path = directory / f"合并错误_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
                path.write_text(details, encoding="utf-8-sig")
                self._append_log(f"错误日志：{path}")
                return
            except OSError:
                continue

    def _cancel_merge(self) -> None:
        if not self.running or self.cancel_event.is_set():
            return
        if messagebox.askyesno("确认取消", "确定取消本次合并吗？临时结果将被清理。"):
            self.cancel_event.set()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status_var.set("正在取消，请等待当前读写操作结束……")
            self._append_log("已请求取消。")

    def _open_output_dir(self) -> None:
        if self.result_dir and self.result_dir.is_dir():
            try:
                os.startfile(self.result_dir)
            except OSError as exc:
                messagebox.showerror("无法打开文件夹", str(exc))

    def _on_close(self) -> None:
        if not self.running:
            self.root.destroy()
            return
        if messagebox.askyesno("退出程序", "合并仍在进行。是否取消任务并退出？"):
            self.close_requested = True
            self.cancel_event.set()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status_var.set("正在安全取消任务……")

    def _set_controls_running(self, running: bool) -> None:
        state = tk.DISABLED if running else tk.NORMAL
        for widget in (
            self.input_entry,
            self.output_entry,
            self.name_entry,
            self.input_button,
            self.output_button,
            self.start_button,
        ):
            widget.configure(state=state)
        self.cancel_button.configure(state=tk.NORMAL if running else tk.DISABLED)
        if running:
            self.open_button.configure(state=tk.DISABLED)

    def _finish_running_state(self) -> None:
        self.running = False
        self.progress.stop()
        self._set_controls_running(False)

    def _append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)


def main() -> None:
    _enable_windows_dpi_awareness()
    root = tk.Tk()
    ExcelMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

