from __future__ import annotations

import ctypes
import json
import os
import queue
import sys
import threading
import traceback
import uuid
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from merge_engine import (
    InspectionPlan,
    InspectedFile,
    MergeCancelled,
    MergeConfig,
    MergeError,
    MergeSummary,
    ProgressEvent,
    SourceOverride,
    create_profile_from_template,
    find_output_conflicts,
    inspect_sources,
    list_visible_sheets,
    merge_workbooks,
    normalize_output_name,
)
from merge_profiles import (
    APP_NAME,
    APP_VERSION,
    BUILTIN_CLASSIC_PROFILE_NAME,
    CsvOptions,
    LabelRule,
    MergeProfile,
    OutputColumnRule,
    ProfileError,
    ProfileStore,
    SpecialFieldRule,
    classic_profile,
    default_app_data_dir,
)


APP_TITLE = f"{APP_NAME} v{APP_VERSION}"
CSV_DELIMITER_DISPLAY = {
    "auto": "自动识别",
    ",": "逗号 ,",
    "\t": "制表符 Tab",
    ";": "分号 ;",
    "|": "竖线 |",
}
CSV_DELIMITER_VALUE = {value: key for key, value in CSV_DELIMITER_DISPLAY.items()}


def _enable_windows_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


def _resource_path(name: str) -> Path:
    bundle = getattr(sys, "_MEIPASS", None)
    if bundle:
        return Path(bundle) / name
    project = Path(__file__).resolve().parent
    direct = project / name
    return direct if direct.exists() else project / "build" / name


def _split_aliases(text: str) -> tuple[str, ...]:
    normalized = text.replace("；", ";").replace("，", ",")
    items: list[str] = []
    for part in normalized.replace(",", ";").split(";"):
        value = part.strip()
        if value:
            items.append(value)
    return tuple(items)


class SettingsStore:
    def __init__(self, path: Path | None = None):
        self.path = Path(path) if path else default_app_data_dir() / "settings.json"

    def load(self) -> dict[str, object]:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, UnicodeError, json.JSONDecodeError):
            return {}

    def save(self, data: dict[str, object]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_name(f".{self.path.name}.{uuid.uuid4().hex}.tmp")
        try:
            temp.write_text(
                json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            temp.replace(self.path)
        finally:
            if temp.exists():
                temp.unlink()


@dataclass(frozen=True)
class ProfileEntry:
    display_name: str
    profile: MergeProfile
    path: Path | None = None


class ChoiceDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, title: str, prompt: str, choices: tuple[str, ...]):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result: str | None = None
        self.choice_var = tk.StringVar(value=choices[0] if choices else "")
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=prompt).pack(anchor=tk.W, pady=(0, 8))
        combo = ttk.Combobox(
            frame,
            textvariable=self.choice_var,
            values=choices,
            state="readonly",
            width=42,
        )
        combo.pack(fill=tk.X)
        buttons = ttk.Frame(frame)
        buttons.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(buttons, text="确定", command=self._accept).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side=tk.RIGHT, padx=8)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_visibility()
        combo.focus_set()

    def _accept(self) -> None:
        value = self.choice_var.get()
        if value:
            self.result = value
            self.destroy()


class OutputRuleDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, rule: OutputColumnRule | None = None):
        super().__init__(parent)
        self.title("输出字段")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result: OutputColumnRule | None = None
        current = rule or OutputColumnRule("新字段")
        self.name_var = tk.StringVar(value=current.name)
        self.alias_var = tk.StringVar(value="；".join(current.aliases))
        self.required_var = tk.BooleanVar(value=current.required)
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="输出字段名").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_entry = ttk.Entry(frame, textvariable=self.name_var, width=46)
        name_entry.grid(row=0, column=1, sticky=tk.EW, padx=(10, 0), pady=5)
        ttk.Label(frame, text="可匹配别名").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.alias_var).grid(row=1, column=1, sticky=tk.EW, padx=(10, 0), pady=5)
        ttk.Label(frame, text="多个别名用分号分隔", foreground="#666666").grid(
            row=2, column=1, sticky=tk.W, padx=(10, 0)
        )
        ttk.Checkbutton(frame, text="必填字段（源文件缺少时跳过）", variable=self.required_var).grid(
            row=3, column=0, columnspan=2, sticky=tk.W, pady=(10, 4)
        )
        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=2, sticky=tk.E, pady=(14, 0))
        ttk.Button(buttons, text="确定", command=self._accept).pack(side=tk.LEFT)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side=tk.LEFT, padx=(8, 0))
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_visibility()
        name_entry.focus_set()
        name_entry.select_range(0, tk.END)

    def _accept(self) -> None:
        try:
            self.result = OutputColumnRule(
                name=self.name_var.get(),
                aliases=_split_aliases(self.alias_var.get()),
                required=self.required_var.get(),
            ).validated()
        except ProfileError as exc:
            messagebox.showerror("字段无效", str(exc), parent=self)
            return
        self.destroy()


class SpecialRuleDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, rule: SpecialFieldRule | None = None):
        super().__init__(parent)
        self.title("特殊字段")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result: SpecialFieldRule | None = None
        current = rule or SpecialFieldRule("主讲老师")
        self.name_var = tk.StringVar(value=current.name)
        self.alias_var = tk.StringVar(value="；".join(current.aliases))
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="特殊字段名").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_entry = ttk.Entry(frame, textvariable=self.name_var, width=46)
        name_entry.grid(row=0, column=1, sticky=tk.EW, padx=(10, 0), pady=5)
        ttk.Label(frame, text="可匹配别名").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.alias_var).grid(row=1, column=1, sticky=tk.EW, padx=(10, 0), pady=5)
        ttk.Label(frame, text="每个文件必须只有一个唯一非空值", foreground="#666666").grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=(6, 0)
        )
        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=0, columnspan=2, sticky=tk.E, pady=(14, 0))
        ttk.Button(buttons, text="确定", command=self._accept).pack(side=tk.LEFT)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side=tk.LEFT, padx=(8, 0))
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_visibility()
        name_entry.focus_set()
        name_entry.select_range(0, tk.END)

    def _accept(self) -> None:
        try:
            self.result = SpecialFieldRule(
                name=self.name_var.get(),
                aliases=_split_aliases(self.alias_var.get()),
            ).validated()
        except ProfileError as exc:
            messagebox.showerror("字段无效", str(exc), parent=self)
            return
        self.destroy()


class ProfileEditorDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, profile: MergeProfile):
        super().__init__(parent)
        self.title("编辑合并方案")
        self.geometry("850x650")
        self.minsize(760, 580)
        self.transient(parent)
        self.grab_set()
        self.result: MergeProfile | None = None
        self.output_rules = list(profile.output_columns)
        self.special_rules = list(profile.special_fields)
        self.name_var = tk.StringVar(value=profile.name)
        self.header_row_var = tk.IntVar(value=profile.default_header_row)
        self.recursive_var = tk.BooleanVar(value=profile.recursive)
        self.encoding_var = tk.StringVar(value=profile.csv_options.encoding)
        self.delimiter_var = tk.StringVar(value=CSV_DELIMITER_DISPLAY[profile.csv_options.delimiter])
        self.recognize_var = tk.BooleanVar(value=profile.recognize_special_fields)
        self.label_enabled_var = tk.BooleanVar(value=profile.label.enabled)
        self.label_header_var = tk.StringVar(value=profile.label.header)
        self.label_template_var = tk.StringVar(value=profile.label.template)
        self.variable_var = tk.StringVar()
        self._build_ui()
        self._refresh_output_tree()
        self._refresh_special_tree()
        self._refresh_variables()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill=tk.BOTH, expand=True)
        general = ttk.LabelFrame(outer, text="方案设置", padding=10)
        general.pack(fill=tk.X)
        general.columnconfigure(1, weight=1)
        ttk.Label(general, text="方案名称").grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(general, textvariable=self.name_var).grid(row=0, column=1, sticky=tk.EW, padx=8, pady=4)
        ttk.Label(general, text="默认表头行").grid(row=0, column=2, sticky=tk.W, pady=4)
        ttk.Spinbox(general, from_=1, to=1_048_576, textvariable=self.header_row_var, width=8).grid(
            row=0, column=3, sticky=tk.W, pady=4
        )
        ttk.Checkbutton(general, text="扫描子文件夹", variable=self.recursive_var).grid(
            row=1, column=0, sticky=tk.W, pady=4
        )
        ttk.Label(general, text="CSV编码").grid(row=1, column=1, sticky=tk.E, pady=4)
        ttk.Combobox(
            general,
            textvariable=self.encoding_var,
            values=("auto", "utf-8", "utf-8-sig", "gb18030"),
            state="readonly",
            width=12,
        ).grid(row=1, column=2, sticky=tk.W, padx=8, pady=4)
        ttk.Combobox(
            general,
            textvariable=self.delimiter_var,
            values=tuple(CSV_DELIMITER_DISPLAY.values()),
            state="readonly",
            width=14,
        ).grid(row=1, column=3, sticky=tk.W, pady=4)

        notebook = ttk.Notebook(outer)
        notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        fields_tab = ttk.Frame(notebook, padding=8)
        special_tab = ttk.Frame(notebook, padding=8)
        notebook.add(fields_tab, text="输出字段")
        notebook.add(special_tab, text="特殊字段与标签")
        fields_tab.columnconfigure(0, weight=1)
        fields_tab.rowconfigure(0, weight=1)
        self.output_tree = ttk.Treeview(
            fields_tab,
            columns=("order", "name", "required", "aliases"),
            show="headings",
            selectmode="browse",
        )
        for column, title, width in (
            ("order", "顺序", 60),
            ("name", "输出字段", 180),
            ("required", "必填", 70),
            ("aliases", "可匹配别名", 360),
        ):
            self.output_tree.heading(column, text=title)
            self.output_tree.column(column, width=width, anchor=tk.W)
        self.output_tree.grid(row=0, column=0, sticky=tk.NSEW)
        output_scroll = ttk.Scrollbar(fields_tab, orient=tk.VERTICAL, command=self.output_tree.yview)
        output_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.output_tree.configure(yscrollcommand=output_scroll.set)
        output_buttons = ttk.Frame(fields_tab)
        output_buttons.grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        for text, command in (
            ("添加", self._add_output),
            ("编辑", self._edit_output),
            ("删除", self._delete_output),
            ("上移", lambda: self._move_output(-1)),
            ("下移", lambda: self._move_output(1)),
        ):
            ttk.Button(output_buttons, text=text, command=command).pack(side=tk.LEFT, padx=(0, 6))
        self.output_tree.bind("<Double-1>", lambda _event: self._edit_output())

        special_tab.columnconfigure(0, weight=1)
        special_tab.rowconfigure(1, weight=1)
        ttk.Checkbutton(
            special_tab,
            text="识别文件级特殊字段（每个文件必须只有一个唯一非空值）",
            variable=self.recognize_var,
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))
        self.special_tree = ttk.Treeview(
            special_tab,
            columns=("name", "aliases"),
            show="headings",
            height=7,
        )
        self.special_tree.heading("name", text="特殊字段")
        self.special_tree.heading("aliases", text="可匹配别名")
        self.special_tree.column("name", width=190)
        self.special_tree.column("aliases", width=420)
        self.special_tree.grid(row=1, column=0, sticky=tk.NSEW)
        special_buttons = ttk.Frame(special_tab)
        special_buttons.grid(row=2, column=0, sticky=tk.W, pady=(6, 12))
        ttk.Button(special_buttons, text="添加", command=self._add_special).pack(side=tk.LEFT)
        ttk.Button(special_buttons, text="编辑", command=self._edit_special).pack(side=tk.LEFT, padx=6)
        ttk.Button(special_buttons, text="删除", command=self._delete_special).pack(side=tk.LEFT)
        self.special_tree.bind("<Double-1>", lambda _event: self._edit_special())

        label_frame = ttk.LabelFrame(special_tab, text="可选标签列", padding=8)
        label_frame.grid(row=3, column=0, columnspan=2, sticky=tk.EW)
        label_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(label_frame, text="在输出末尾追加标签列", variable=self.label_enabled_var).grid(
            row=0, column=0, columnspan=2, sticky=tk.W
        )
        ttk.Label(label_frame, text="列标题").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(label_frame, textvariable=self.label_header_var).grid(
            row=1, column=1, sticky=tk.EW, padx=(8, 0), pady=5
        )
        ttk.Label(label_frame, text="内容模板").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.template_entry = ttk.Entry(label_frame, textvariable=self.label_template_var)
        self.template_entry.grid(row=2, column=1, sticky=tk.EW, padx=(8, 0), pady=5)
        variable_frame = ttk.Frame(label_frame)
        variable_frame.grid(row=3, column=1, sticky=tk.W, padx=(8, 0))
        self.variable_combo = ttk.Combobox(
            variable_frame,
            textvariable=self.variable_var,
            state="readonly",
            width=27,
        )
        self.variable_combo.pack(side=tk.LEFT)
        ttk.Button(variable_frame, text="插入变量", command=self._insert_variable).pack(side=tk.LEFT, padx=6)

        buttons = ttk.Frame(outer)
        buttons.pack(fill=tk.X)
        ttk.Button(buttons, text="保存方案", command=self._accept).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side=tk.RIGHT, padx=8)

    def _selected_index(self, tree: ttk.Treeview) -> int | None:
        selected = tree.selection()
        return int(selected[0]) if selected else None

    def _refresh_output_tree(self, selected: int | None = None) -> None:
        self.output_tree.delete(*self.output_tree.get_children())
        for index, rule in enumerate(self.output_rules):
            self.output_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(index + 1, rule.name, "是" if rule.required else "否", "；".join(rule.aliases)),
            )
        if selected is not None and 0 <= selected < len(self.output_rules):
            self.output_tree.selection_set(str(selected))

    def _refresh_special_tree(self, selected: int | None = None) -> None:
        self.special_tree.delete(*self.special_tree.get_children())
        for index, rule in enumerate(self.special_rules):
            self.special_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(rule.name, "；".join(rule.aliases)),
            )
        if selected is not None and 0 <= selected < len(self.special_rules):
            self.special_tree.selection_set(str(selected))
        self._refresh_variables()

    def _add_output(self) -> None:
        dialog = OutputRuleDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            self.output_rules.append(dialog.result)
            self._refresh_output_tree(len(self.output_rules) - 1)

    def _edit_output(self) -> None:
        index = self._selected_index(self.output_tree)
        if index is None:
            return
        dialog = OutputRuleDialog(self, self.output_rules[index])
        self.wait_window(dialog)
        if dialog.result:
            self.output_rules[index] = dialog.result
            self._refresh_output_tree(index)

    def _delete_output(self) -> None:
        index = self._selected_index(self.output_tree)
        if index is None:
            return
        del self.output_rules[index]
        self._refresh_output_tree(min(index, len(self.output_rules) - 1))

    def _move_output(self, direction: int) -> None:
        index = self._selected_index(self.output_tree)
        if index is None:
            return
        target = index + direction
        if target < 0 or target >= len(self.output_rules):
            return
        self.output_rules[index], self.output_rules[target] = self.output_rules[target], self.output_rules[index]
        self._refresh_output_tree(target)

    def _add_special(self) -> None:
        dialog = SpecialRuleDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            self.special_rules.append(dialog.result)
            self._refresh_special_tree(len(self.special_rules) - 1)

    def _edit_special(self) -> None:
        index = self._selected_index(self.special_tree)
        if index is None:
            return
        dialog = SpecialRuleDialog(self, self.special_rules[index])
        self.wait_window(dialog)
        if dialog.result:
            self.special_rules[index] = dialog.result
            self._refresh_special_tree(index)

    def _delete_special(self) -> None:
        index = self._selected_index(self.special_tree)
        if index is None:
            return
        del self.special_rules[index]
        self._refresh_special_tree(min(index, len(self.special_rules) - 1))

    def _refresh_variables(self) -> None:
        if not hasattr(self, "variable_combo"):
            return
        values = ["{文件名}", "{文件名不含扩展名}", "{相对路径}", "{工作表}"]
        values.extend(f"{{特殊字段:{rule.name}}}" for rule in self.special_rules)
        self.variable_combo.configure(values=values)
        if values and self.variable_var.get() not in values:
            self.variable_var.set(values[0])

    def _insert_variable(self) -> None:
        token = self.variable_var.get()
        if not token:
            return
        self.template_entry.insert(self.template_entry.index(tk.INSERT), token)

    def _accept(self) -> None:
        try:
            delimiter = CSV_DELIMITER_VALUE[self.delimiter_var.get()]
            profile = MergeProfile(
                name=self.name_var.get(),
                output_columns=tuple(self.output_rules),
                special_fields=tuple(self.special_rules),
                recognize_special_fields=self.recognize_var.get(),
                label=LabelRule(
                    enabled=self.label_enabled_var.get(),
                    header=self.label_header_var.get(),
                    template=self.label_template_var.get(),
                ),
                recursive=self.recursive_var.get(),
                default_header_row=self.header_row_var.get(),
                csv_options=CsvOptions(
                    encoding=self.encoding_var.get(),
                    delimiter=delimiter,
                ),
            ).validated()
        except (ProfileError, KeyError, tk.TclError) as exc:
            messagebox.showerror("方案无法保存", str(exc), parent=self)
            return
        self.result = profile
        self.destroy()


class SourceOptionsDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        item: InspectedFile,
        current: SourceOverride,
        visible_sheets: tuple[str, ...],
    ):
        super().__init__(parent)
        self.title(f"文件设置 - {item.path.name}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.path = item.path
        self.include = current.include
        self.result: SourceOverride | None = None
        self.header_row_var = tk.IntVar(value=current.header_row or item.header_row or 1)
        self.sheet_var = tk.StringVar(value=current.worksheet or item.worksheet)
        self.encoding_var = tk.StringVar(value=current.csv_encoding or item.csv_encoding or "auto")
        delimiter = current.csv_delimiter or item.csv_delimiter or "auto"
        self.delimiter_var = tk.StringVar(value=CSV_DELIMITER_DISPLAY.get(delimiter, "自动识别"))
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="文件").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Label(frame, text=item.relative_path).grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(frame, text="表头行").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(frame, from_=1, to=1_048_576, textvariable=self.header_row_var, width=12).grid(
            row=1, column=1, sticky=tk.W, padx=10, pady=5
        )
        if item.file_type in {"XLSX", "XLSM"}:
            ttk.Label(frame, text="工作表").grid(row=2, column=0, sticky=tk.W, pady=5)
            ttk.Combobox(
                frame,
                textvariable=self.sheet_var,
                values=visible_sheets,
                state="readonly",
                width=34,
            ).grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        else:
            ttk.Label(frame, text="CSV编码").grid(row=2, column=0, sticky=tk.W, pady=5)
            ttk.Combobox(
                frame,
                textvariable=self.encoding_var,
                values=("auto", "utf-8", "utf-8-sig", "gb18030"),
                state="readonly",
                width=18,
            ).grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
            ttk.Label(frame, text="CSV分隔符").grid(row=3, column=0, sticky=tk.W, pady=5)
            ttk.Combobox(
                frame,
                textvariable=self.delimiter_var,
                values=tuple(CSV_DELIMITER_DISPLAY.values()),
                state="readonly",
                width=18,
            ).grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=2, sticky=tk.E, pady=(14, 0))
        ttk.Button(buttons, text="确定并重新扫描", command=self._accept).pack(side=tk.LEFT)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side=tk.LEFT, padx=8)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _accept(self) -> None:
        try:
            self.result = SourceOverride(
                path=self.path,
                include=self.include,
                worksheet=self.sheet_var.get(),
                header_row=self.header_row_var.get(),
                csv_encoding=self.encoding_var.get(),
                csv_delimiter=CSV_DELIMITER_VALUE[self.delimiter_var.get()],
            ).validated()
        except (MergeError, ProfileError, KeyError, tk.TclError) as exc:
            messagebox.showerror("文件设置无效", str(exc), parent=self)
            return
        self.destroy()


class ExcelMergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x760")
        self.root.minsize(980, 680)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        icon_path = _resource_path("app.ico")
        if icon_path.exists():
            try:
                self.root.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass

        self.settings_store = SettingsStore()
        self.settings = self.settings_store.load()
        geometry = self.settings.get("geometry")
        if isinstance(geometry, str) and "x" in geometry:
            try:
                self.root.geometry(geometry)
            except tk.TclError:
                pass
        self.profile_store = ProfileStore()
        self.profile_entries: dict[str, ProfileEntry] = {}
        self.current_plan: InspectionPlan | None = None
        self.source_overrides: dict[Path, SourceOverride] = {}
        self.preview_items: dict[str, InspectedFile] = {}
        self.last_summary: MergeSummary | None = None
        self.running = False
        self.operation = ""
        self.close_requested = False
        self.cancel_event = threading.Event()
        self.event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.input_dir_var = tk.StringVar(value=str(self.settings.get("input_dir", "")))
        self.output_dir_var = tk.StringVar(value=str(self.settings.get("output_dir", "")))
        self.output_name_var = tk.StringVar(value=self._default_output_name())
        self.profile_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="请选择方案和输入文件夹，然后执行精确扫描。")
        self.progress_text_var = tk.StringVar(value="尚未扫描")

        self._configure_style()
        self._build_menu()
        self._build_ui()
        self._load_profiles(preferred=str(self.settings.get("last_profile", "")))
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
        style.configure("Treeview", rowheight=25)

    def _build_menu(self) -> None:
        menu = tk.Menu(self.root)
        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label="使用帮助", command=self._show_help)
        help_menu.add_command(label="关于", command=self._show_about)
        menu.add_cascade(label="帮助", menu=help_menu)
        self.root.configure(menu=menu)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(5, weight=1)
        ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(
            outer,
            text="支持 XLSX、XLSM 和 CSV；先精确预检，再以流式方式生成结果与 Excel 报告。",
            style="Hint.TLabel",
        ).grid(row=1, column=0, sticky=tk.W, pady=(2, 10))

        config_frame = ttk.LabelFrame(outer, text="1. 方案与路径", padding=10)
        config_frame.grid(row=2, column=0, sticky=tk.EW)
        config_frame.columnconfigure(1, weight=1)
        ttk.Label(config_frame, text="合并方案").grid(row=0, column=0, sticky=tk.W, pady=5)
        profile_line = ttk.Frame(config_frame)
        profile_line.grid(row=0, column=1, columnspan=2, sticky=tk.EW, padx=(10, 0), pady=5)
        profile_line.columnconfigure(0, weight=1)
        self.profile_combo = ttk.Combobox(profile_line, textvariable=self.profile_var, state="readonly")
        self.profile_combo.grid(row=0, column=0, sticky=tk.EW)
        self.profile_combo.bind("<<ComboboxSelected>>", self._on_profile_changed)
        for column, (text, command) in enumerate(
            (
                ("新建", self._new_profile),
                ("编辑", self._edit_profile),
                ("删除", self._delete_profile),
                ("导入", self._import_profile),
                ("导出", self._export_profile),
            ),
            start=1,
        ):
            ttk.Button(profile_line, text=text, command=command).grid(row=0, column=column, padx=(6, 0))
        ttk.Checkbutton(
            config_frame,
            text="扫描子文件夹",
            variable=self.recursive_var,
            command=self._invalidate_plan,
        ).grid(row=0, column=3, sticky=tk.W, padx=(10, 0))

        ttk.Label(config_frame, text="输入文件夹").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.input_entry = ttk.Entry(config_frame, textvariable=self.input_dir_var)
        self.input_entry.grid(row=1, column=1, sticky=tk.EW, padx=10, pady=5)
        self.input_button = ttk.Button(config_frame, text="选择…", command=self._choose_input_dir)
        self.input_button.grid(row=1, column=2, sticky=tk.EW, pady=5)
        ttk.Label(config_frame, text="输出文件夹").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_entry = ttk.Entry(config_frame, textvariable=self.output_dir_var)
        self.output_entry.grid(row=2, column=1, sticky=tk.EW, padx=10, pady=5)
        self.output_button = ttk.Button(config_frame, text="选择…", command=self._choose_output_dir)
        self.output_button.grid(row=2, column=2, sticky=tk.EW, pady=5)
        ttk.Label(config_frame, text="输出文件名").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(config_frame, textvariable=self.output_name_var)
        self.name_entry.grid(row=3, column=1, columnspan=2, sticky=tk.EW, padx=10, pady=5)
        for variable in (self.input_dir_var, self.output_dir_var, self.output_name_var):
            variable.trace_add("write", lambda *_args: self._invalidate_plan())

        action_frame = ttk.Frame(outer)
        action_frame.grid(row=3, column=0, sticky=tk.EW, pady=10)
        self.scan_button = ttk.Button(
            action_frame,
            text="2. 精确扫描",
            style="Accent.TButton",
            command=self._start_scan,
        )
        self.scan_button.pack(side=tk.LEFT)
        self.toggle_button = ttk.Button(action_frame, text="包含/排除文件", command=self._toggle_selected)
        self.toggle_button.pack(side=tk.LEFT, padx=6)
        self.file_settings_button = ttk.Button(action_frame, text="逐文件设置", command=self._edit_source_options)
        self.file_settings_button.pack(side=tk.LEFT)
        self.start_button = ttk.Button(
            action_frame,
            text="3. 开始合并",
            style="Accent.TButton",
            command=self._start_merge,
            state=tk.DISABLED,
        )
        self.start_button.pack(side=tk.LEFT, padx=(18, 0))
        self.cancel_button = ttk.Button(
            action_frame,
            text="取消任务",
            command=self._cancel_task,
            state=tk.DISABLED,
        )
        self.cancel_button.pack(side=tk.LEFT, padx=6)
        self.open_report_button = ttk.Button(
            action_frame,
            text="打开报告",
            command=self._open_report,
            state=tk.DISABLED,
        )
        self.open_report_button.pack(side=tk.RIGHT)
        self.open_output_button = ttk.Button(
            action_frame,
            text="打开结果",
            command=self._open_output,
            state=tk.DISABLED,
        )
        self.open_output_button.pack(side=tk.RIGHT, padx=6)

        progress_frame = ttk.Frame(outer)
        progress_frame.grid(row=4, column=0, sticky=tk.EW, pady=(0, 8))
        progress_frame.columnconfigure(0, weight=1)
        ttk.Label(progress_frame, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Label(progress_frame, textvariable=self.progress_text_var).grid(row=0, column=1, sticky=tk.E)
        self.progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(5, 0))

        paned = ttk.Panedwindow(outer, orient=tk.VERTICAL)
        paned.grid(row=5, column=0, sticky=tk.NSEW)
        preview_frame = ttk.LabelFrame(paned, text="文件预览", padding=6)
        log_frame = ttk.LabelFrame(paned, text="处理记录", padding=6)
        paned.add(preview_frame, weight=4)
        paned.add(log_frame, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        columns = ("include", "status", "path", "type", "sheet", "header", "rows", "special", "reason")
        self.preview_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", selectmode="browse")
        settings = (
            ("include", "参与", 55),
            ("status", "状态", 70),
            ("path", "相对路径", 220),
            ("type", "格式", 60),
            ("sheet", "工作表", 120),
            ("header", "表头行", 65),
            ("rows", "数据行", 90),
            ("special", "特殊字段", 220),
            ("reason", "异常原因", 330),
        )
        for column, title, width in settings:
            self.preview_tree.heading(column, text=title)
            self.preview_tree.column(column, width=width, anchor=tk.W, stretch=column in {"path", "special", "reason"})
        self.preview_tree.grid(row=0, column=0, sticky=tk.NSEW)
        preview_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        preview_y.grid(row=0, column=1, sticky=tk.NS)
        preview_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        preview_x.grid(row=1, column=0, sticky=tk.EW)
        self.preview_tree.configure(yscrollcommand=preview_y.set, xscrollcommand=preview_x.set)
        self.preview_tree.bind("<Double-1>", lambda _event: self._edit_source_options())
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_frame,
            height=7,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 9),
            state=tk.DISABLED,
        )
        self.log_text.grid(row=0, column=0, sticky=tk.NSEW)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.log_text.configure(yscrollcommand=log_scroll.set)

    def _load_profiles(self, preferred: str = "") -> None:
        entries: dict[str, ProfileEntry] = {
            BUILTIN_CLASSIC_PROFILE_NAME: ProfileEntry(
                BUILTIN_CLASSIC_PROFILE_NAME,
                classic_profile(),
                None,
            )
        }
        for path, profile in self.profile_store.list_profiles():
            display = profile.name
            counter = 2
            while display in entries:
                display = f"{profile.name} ({counter})"
                counter += 1
            entries[display] = ProfileEntry(display, profile, path)
        self.profile_entries = entries
        values = tuple(entries)
        self.profile_combo.configure(values=values)
        selected = next(
            (name for name, entry in entries.items() if name == preferred or entry.profile.name == preferred),
            BUILTIN_CLASSIC_PROFILE_NAME,
        )
        self.profile_var.set(selected)
        self.recursive_var.set(entries[selected].profile.recursive)
        self._invalidate_plan()

    def _current_entry(self) -> ProfileEntry:
        selected = self.profile_var.get()
        if selected not in self.profile_entries:
            raise MergeError("请选择有效的合并方案。")
        return self.profile_entries[selected]

    def _current_profile(self) -> MergeProfile:
        return replace(self._current_entry().profile, recursive=self.recursive_var.get()).validated()

    def _on_profile_changed(self, _event: object = None) -> None:
        try:
            self.recursive_var.set(self._current_entry().profile.recursive)
        except MergeError:
            pass
        self.source_overrides.clear()
        self._invalidate_plan()

    def _new_profile(self) -> None:
        source = filedialog.askopenfilename(
            title="选择用作字段模板的文件",
            filetypes=[("支持的表格", "*.xlsx *.xlsm *.csv"), ("所有文件", "*.*")],
        )
        if not source:
            return
        name = simpledialog.askstring("新建方案", "方案名称：", parent=self.root)
        if not name:
            return
        header_row = simpledialog.askinteger(
            "新建方案",
            "模板表头所在行：",
            initialvalue=1,
            minvalue=1,
            maxvalue=1_048_576,
            parent=self.root,
        )
        if not header_row:
            return
        worksheet = ""
        path = Path(source)
        try:
            if path.suffix.casefold() in {".xlsx", ".xlsm"}:
                sheets = list_visible_sheets(path)
                dialog = ChoiceDialog(self.root, "选择模板工作表", "请选择表头所在工作表：", sheets)
                self.root.wait_window(dialog)
                if not dialog.result:
                    return
                worksheet = dialog.result
            profile, _header = create_profile_from_template(
                path,
                name,
                override=SourceOverride(path, worksheet=worksheet, header_row=header_row),
                default_header_row=header_row,
            )
        except (MergeError, ProfileError, OSError) as exc:
            messagebox.showerror("无法创建方案", str(exc), parent=self.root)
            return
        editor = ProfileEditorDialog(self.root, profile)
        self.root.wait_window(editor)
        if editor.result:
            try:
                self.profile_store.save(editor.result)
                self._load_profiles(preferred=editor.result.name)
            except (ProfileError, OSError) as exc:
                messagebox.showerror("无法保存方案", str(exc), parent=self.root)

    def _edit_profile(self) -> None:
        entry = self._current_entry()
        if entry.profile.classic:
            messagebox.showinfo("内置方案", "经典 A:Q 模式不可修改；请用“新建”从模板创建自定义方案。")
            return
        editor = ProfileEditorDialog(self.root, entry.profile)
        self.root.wait_window(editor)
        if not editor.result:
            return
        try:
            self.profile_store.save(editor.result, existing_path=entry.path)
            self._load_profiles(preferred=editor.result.name)
        except (ProfileError, OSError) as exc:
            messagebox.showerror("无法保存方案", str(exc), parent=self.root)

    def _delete_profile(self) -> None:
        entry = self._current_entry()
        if entry.path is None:
            messagebox.showinfo("内置方案", "经典 A:Q 模式不能删除。")
            return
        if not messagebox.askyesno("删除方案", f"确定删除方案“{entry.profile.name}”吗？"):
            return
        try:
            self.profile_store.delete(entry.path)
            self._load_profiles()
        except (ProfileError, OSError) as exc:
            messagebox.showerror("无法删除方案", str(exc))

    def _import_profile(self) -> None:
        source = filedialog.askopenfilename(title="导入方案", filetypes=[("JSON方案", "*.json")])
        if not source:
            return
        try:
            _path, profile = self.profile_store.import_profile(Path(source))
            self._load_profiles(preferred=profile.name)
        except (ProfileError, OSError) as exc:
            messagebox.showerror("无法导入方案", str(exc))

    def _export_profile(self) -> None:
        entry = self._current_entry()
        if entry.profile.classic:
            messagebox.showinfo("内置方案", "经典模式无需导出；请导出自定义方案。")
            return
        target = filedialog.asksaveasfilename(
            title="导出方案",
            defaultextension=".json",
            initialfile=f"{entry.profile.name}.json",
            filetypes=[("JSON方案", "*.json")],
        )
        if not target:
            return
        try:
            self.profile_store.export_profile(entry.profile, Path(target))
        except (ProfileError, OSError) as exc:
            messagebox.showerror("无法导出方案", str(exc))

    def _choose_input_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择包含 Excel/CSV 的文件夹")
        if not selected:
            return
        self.input_dir_var.set(selected)
        if not self.output_dir_var.get().strip():
            self.output_dir_var.set(selected)

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择结果保存文件夹")
        if selected:
            self.output_dir_var.set(selected)

    def _make_config(self, *, overwrite: bool = False) -> MergeConfig:
        input_text = self.input_dir_var.get().strip()
        output_text = self.output_dir_var.get().strip()
        if not input_text or not Path(input_text).expanduser().is_dir():
            raise MergeError("请选择有效的输入文件夹。")
        if not output_text:
            raise MergeError("请选择输出文件夹。")
        output_dir = Path(output_text).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        return MergeConfig(
            input_dir=Path(input_text),
            output_dir=output_dir,
            output_name=normalize_output_name(self.output_name_var.get()),
            overwrite=overwrite,
            profile=self._current_profile(),
            source_overrides=tuple(self.source_overrides.values()),
        ).validated()

    def _start_scan(self) -> None:
        if self.running:
            return
        try:
            config = self._make_config()
        except (MergeError, ProfileError, OSError) as exc:
            messagebox.showerror("无法扫描", str(exc))
            return
        self._clear_log()
        self._append_log(f"使用方案：{config.profile.name if config.profile else BUILTIN_CLASSIC_PROFILE_NAME}")
        self._append_log("开始精确预检；预检不会修改源文件。")
        self.current_plan = None
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_items.clear()
        self._set_running("scan")
        self.worker = threading.Thread(
            target=self._run_scan,
            args=(config,),
            name="excel-inspection-worker",
            daemon=True,
        )
        self.worker.start()

    def _run_scan(self, config: MergeConfig) -> None:
        try:
            plan = inspect_sources(
                config,
                cancel_event=self.cancel_event,
                progress_callback=lambda event: self.event_queue.put(("progress", event)),
            )
            self.event_queue.put(("scan_success", plan))
        except MergeCancelled as exc:
            self.event_queue.put(("cancelled", str(exc)))
        except Exception as exc:
            self.event_queue.put(("error", (str(exc), "".join(traceback.format_exception(exc)))))

    def _populate_preview(self, plan: InspectionPlan) -> None:
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_items.clear()
        for index, item in enumerate(plan.files):
            iid = f"source_{index}"
            self.preview_items[iid] = item
            specials = "；".join(f"{key}={value}" for key, value in item.special_values)
            included = self.source_overrides.get(item.path, SourceOverride(item.path)).include
            self.preview_tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    "是" if included else "否",
                    item.status,
                    item.relative_path,
                    item.file_type,
                    item.worksheet or "—",
                    item.header_row,
                    f"{item.data_rows:,}" if item.data_rows else "—",
                    specials,
                    item.reason,
                ),
                tags=(item.status,),
            )
        self.preview_tree.tag_configure("有效", background="#EFF8EA")
        self.preview_tree.tag_configure("跳过", background="#FFF1E8")
        self.preview_tree.tag_configure("排除", background="#F0F0F0")

    def _selected_item(self) -> tuple[str, InspectedFile] | None:
        selected = self.preview_tree.selection()
        if not selected:
            return None
        iid = selected[0]
        item = self.preview_items.get(iid)
        return (iid, item) if item else None

    def _toggle_selected(self) -> None:
        selected = self._selected_item()
        if not selected:
            messagebox.showinfo("请选择文件", "请先在预览表中选择一个文件。")
            return
        _iid, item = selected
        current = self.source_overrides.get(item.path, SourceOverride(item.path))
        self.source_overrides[item.path] = replace(current, include=not current.include).validated()
        self._invalidate_plan()
        self.status_var.set("文件参与状态已修改，请重新执行精确扫描。")

    def _edit_source_options(self) -> None:
        selected = self._selected_item()
        if not selected:
            messagebox.showinfo("请选择文件", "请先在预览表中选择一个文件。")
            return
        _iid, item = selected
        current = self.source_overrides.get(item.path, SourceOverride(item.path))
        sheets = item.visible_sheets
        if item.file_type in {"XLSX", "XLSM"} and not sheets:
            try:
                sheets = list_visible_sheets(item.path)
            except MergeError as exc:
                messagebox.showerror("无法读取工作表", str(exc))
                return
        dialog = SourceOptionsDialog(self.root, item, current, sheets)
        self.root.wait_window(dialog)
        if dialog.result:
            self.source_overrides[item.path] = dialog.result
            self._invalidate_plan()
            self._start_scan()

    def _start_merge(self) -> None:
        if self.running or self.current_plan is None:
            return
        try:
            initial = self._make_config()
            conflicts = find_output_conflicts(initial.output_dir, initial.output_name)
            overwrite = False
            if conflicts:
                shown = "\n".join(f"• {path.name}" for path in conflicts[:8])
                if len(conflicts) > 8:
                    shown += f"\n• 另有 {len(conflicts) - 8} 个分卷"
                overwrite = messagebox.askyesno(
                    "确认替换旧结果",
                    f"新结果完整生成后将替换：\n\n{shown}\n\n是否继续？",
                    icon=messagebox.WARNING,
                )
                if not overwrite:
                    return
            config = replace(initial, overwrite=overwrite).validated()
        except (MergeError, ProfileError, OSError) as exc:
            messagebox.showerror("无法开始合并", str(exc))
            return
        self._append_log("开始流式合并；源文件不会被修改。")
        self._set_running("merge")
        self.worker = threading.Thread(
            target=self._run_merge,
            args=(config, self.current_plan),
            name="excel-merge-worker",
            daemon=True,
        )
        self.worker.start()

    def _run_merge(self, config: MergeConfig, plan: InspectionPlan) -> None:
        try:
            summary = merge_workbooks(
                config,
                plan=plan,
                cancel_event=self.cancel_event,
                progress_callback=lambda event: self.event_queue.put(("progress", event)),
            )
            self.event_queue.put(("merge_success", summary))
        except MergeCancelled as exc:
            self.event_queue.put(("cancelled", str(exc)))
        except Exception as exc:
            self.event_queue.put(("error", (str(exc), "".join(traceback.format_exception(exc)))))

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.event_queue.get_nowait()
                if kind == "progress":
                    self._handle_progress(payload)
                elif kind == "scan_success":
                    self._handle_scan_success(payload)
                elif kind == "merge_success":
                    self._handle_merge_success(payload)
                elif kind == "cancelled":
                    self._handle_cancelled(str(payload))
                elif kind == "error":
                    message, details = payload
                    self._handle_error(str(message), str(details))
        except queue.Empty:
            pass
        if self.close_requested and not self.running:
            self._save_settings()
            self.root.destroy()
            return
        self.root.after(100, self._poll_events)

    def _handle_progress(self, payload: object) -> None:
        if not isinstance(payload, ProgressEvent):
            return
        event = payload
        self.status_var.set(event.message)
        if event.phase in {"scan_start", "scan_file", "scan_rows", "scan_done"}:
            self.progress.configure(mode="determinate", maximum=max(event.files_total, 1))
            self.progress.configure(value=min(event.files_completed, max(event.files_total, 1)))
            self.progress_text_var.set(f"预检 {event.files_completed}/{event.files_total} 个文件")
        else:
            maximum = max(event.total_rows, 1)
            self.progress.configure(mode="determinate", maximum=maximum, value=min(event.rows_written, maximum))
            self.progress_text_var.set(
                f"{event.rows_written:,}/{event.total_rows:,} 行"
                if event.total_rows
                else f"{event.rows_written:,} 行"
            )
        if event.phase in {"scan_file", "merge_file", "saving", "done"}:
            self._append_log(event.message)

    def _handle_scan_success(self, payload: object) -> None:
        if not isinstance(payload, InspectionPlan):
            return
        self._finish_running()
        self.current_plan = payload
        self._populate_preview(payload)
        valid = len(payload.valid_files)
        skipped = sum(1 for item in payload.files if item.status == "跳过")
        self.status_var.set(f"扫描完成：{valid} 个有效，{skipped} 个跳过，共 {payload.total_rows:,} 行。")
        self.progress_text_var.set(f"{len(payload.files)} 个文件")
        self.progress.configure(maximum=1, value=1)
        self.start_button.configure(state=tk.NORMAL)
        self._append_log(f"精确扫描完成：预计合并 {payload.total_rows:,} 行。")
        if skipped:
            messagebox.showwarning("扫描完成（存在异常文件）", "部分文件无法参与合并，请查看预览表中的异常原因。")

    def _handle_merge_success(self, payload: object) -> None:
        if not isinstance(payload, MergeSummary):
            return
        self._finish_running()
        self.last_summary = payload
        self.open_report_button.configure(state=tk.NORMAL)
        self.open_output_button.configure(state=tk.NORMAL if payload.output_files else tk.DISABLED)
        self.progress.configure(maximum=1, value=1)
        self.progress_text_var.set(f"共 {payload.total_rows:,} 行")
        skipped = sum(1 for item in payload.file_results if item.status == "跳过")
        self.status_var.set(f"合并完成：{payload.total_rows:,} 行，报告 {payload.report_path.name}")
        self._append_log(f"合并完成，耗时 {payload.elapsed_seconds:.2f} 秒。")
        self._append_log(f"报告：{payload.report_path.name}")
        outputs = "\n".join(path.name for path in payload.output_files) or "无 Excel 数据输出"
        message = (
            f"共写入 {payload.total_rows:,} 行。\n\n输出：\n{outputs}\n\n"
            f"跳过文件：{skipped} 个\n报告：{payload.report_path.name}"
        )
        if skipped:
            messagebox.showwarning("合并完成（有文件被跳过）", message)
        else:
            messagebox.showinfo("合并完成", message)

    def _handle_cancelled(self, message: str) -> None:
        self._finish_running()
        self.status_var.set("任务已取消，未发布半成品。")
        self.progress.configure(maximum=100, value=0)
        self.progress_text_var.set("已取消")
        self._append_log(message)
        if not self.close_requested:
            messagebox.showinfo("任务已取消", "临时结果已清理，没有发布半成品。")

    def _handle_error(self, message: str, details: str) -> None:
        self._finish_running()
        self.status_var.set("任务失败，未发布半成品。")
        self.progress.configure(maximum=100, value=0)
        self.progress_text_var.set("失败")
        self._append_log(f"错误：{message}")
        self._write_crash_log(details)
        if not self.close_requested:
            messagebox.showerror("任务失败", f"{message}\n\n没有发布半成品；技术日志已保存。")

    def _write_crash_log(self, details: str) -> None:
        candidates = [Path(self.output_dir_var.get().strip()), Path.cwd()]
        for directory in candidates:
            try:
                directory.mkdir(parents=True, exist_ok=True)
                path = directory / f"合并错误_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
                path.write_text(details, encoding="utf-8-sig")
                self._append_log(f"技术日志：{path}")
                return
            except OSError:
                continue

    def _set_running(self, operation: str) -> None:
        self.running = True
        self.operation = operation
        self.close_requested = False
        self.cancel_event = threading.Event()
        self._set_controls_state(tk.DISABLED)
        self.cancel_button.configure(state=tk.NORMAL)
        self.start_button.configure(state=tk.DISABLED)
        self.progress.configure(mode="determinate", maximum=100, value=0)

    def _finish_running(self) -> None:
        self.running = False
        self.operation = ""
        self._set_controls_state(tk.NORMAL)
        self.profile_combo.configure(state="readonly")
        self.cancel_button.configure(state=tk.DISABLED)
        if self.current_plan is not None:
            self.start_button.configure(state=tk.NORMAL)

    def _set_controls_state(self, state: str) -> None:
        for widget in (
            self.profile_combo,
            self.input_entry,
            self.output_entry,
            self.name_entry,
            self.input_button,
            self.output_button,
            self.scan_button,
            self.toggle_button,
            self.file_settings_button,
        ):
            widget.configure(state=state)

    def _cancel_task(self) -> None:
        if not self.running or self.cancel_event.is_set():
            return
        if messagebox.askyesno("确认取消", "确定取消当前任务吗？临时结果将被清理。"):
            self.cancel_event.set()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status_var.set("正在安全取消，请等待当前读写操作结束……")

    def _invalidate_plan(self, *_args: object) -> None:
        if self.running:
            return
        self.current_plan = None
        if hasattr(self, "start_button"):
            self.start_button.configure(state=tk.DISABLED)

    def _open_output(self) -> None:
        if not self.last_summary or not self.last_summary.output_files:
            return
        target = self.last_summary.output_files[0]
        try:
            os.startfile(target if len(self.last_summary.output_files) == 1 else target.parent)
        except OSError as exc:
            messagebox.showerror("无法打开结果", str(exc))

    def _open_report(self) -> None:
        if not self.last_summary:
            return
        try:
            os.startfile(self.last_summary.report_path)
        except OSError as exc:
            messagebox.showerror("无法打开报告", str(exc))

    def _show_help(self) -> None:
        messagebox.showinfo(
            "使用帮助",
            "1. 选择经典方案或自定义方案。\n"
            "2. 选择输入和输出文件夹。\n"
            "3. 点击“精确扫描”，检查每个文件的状态。\n"
            "4. 如需调整工作表、表头行或CSV参数，使用“逐文件设置”。\n"
            "5. 确认后点击“开始合并”。\n\n"
            "XLSM 宏不会执行或复制；公式仅读取上次保存的缓存值。",
        )

    def _show_about(self) -> None:
        messagebox.showinfo(
            "关于",
            f"{APP_NAME}\n版本 {APP_VERSION}\n\n"
            "面向非技术用户的流式表格合并工具。\n"
            "支持 XLSX、XLSM、CSV 和可复用字段映射方案。",
        )

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

    def _save_settings(self) -> None:
        try:
            self.settings_store.save(
                {
                    "input_dir": self.input_dir_var.get().strip(),
                    "output_dir": self.output_dir_var.get().strip(),
                    "last_profile": self._current_entry().profile.name,
                    "geometry": self.root.geometry(),
                }
            )
        except (OSError, MergeError):
            pass

    def _on_close(self) -> None:
        if not self.running:
            self._save_settings()
            self.root.destroy()
            return
        if messagebox.askyesno("退出程序", "任务仍在进行。是否取消任务并退出？"):
            self.close_requested = True
            self.cancel_event.set()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status_var.set("正在安全取消任务……")


def main() -> None:
    _enable_windows_dpi_awareness()
    root = tk.Tk()
    ExcelMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
