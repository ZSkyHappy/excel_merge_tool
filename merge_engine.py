from __future__ import annotations

import os
import re
import shutil
import tempfile
import threading
import time
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Iterable, Literal, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


INPUT_COLUMN_COUNT = 17
OUTPUT_COLUMN_COUNT = 18
MAX_DATA_ROWS_PER_FILE = 1_048_575
TEACHER_HEADER = "主讲老师"
SOURCE_HEADER = "所属表格+主讲老师"
OUTPUT_SHEET_NAME = "合并结果"
PROGRESS_INTERVAL_ROWS = 1_000

_INVALID_OUTPUT_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]")
_PART_NAME_TEMPLATE = "{stem}_第{number}部分.xlsx"
_WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
}


class MergeError(RuntimeError):
    """Base error for merge operations."""


class MergeCancelled(MergeError):
    """Raised after a cooperative cancellation request."""


class OutputExistsError(MergeError):
    """Raised when output files already exist and overwrite is disabled."""

    def __init__(self, paths: Sequence[Path]):
        self.paths = tuple(paths)
        names = "、".join(path.name for path in self.paths)
        super().__init__(f"输出目录中已存在同名结果：{names}")


Status = Literal["成功", "跳过"]
Phase = Literal[
    "scan_start",
    "scan_file",
    "scan_done",
    "merge_start",
    "merge_file",
    "merge_rows",
    "saving",
    "done",
]


@dataclass(frozen=True)
class MergeConfig:
    input_dir: Path
    output_dir: Path
    output_name: str
    overwrite: bool = False
    max_data_rows_per_file: int = MAX_DATA_ROWS_PER_FILE
    progress_interval_rows: int = PROGRESS_INTERVAL_ROWS

    def validated(self) -> "MergeConfig":
        input_dir = Path(self.input_dir).expanduser().resolve()
        output_dir = Path(self.output_dir).expanduser().resolve()
        output_name = normalize_output_name(self.output_name)

        if not input_dir.is_dir():
            raise MergeError(f"输入文件夹不存在：{input_dir}")
        if self.max_data_rows_per_file < 1 or self.max_data_rows_per_file > MAX_DATA_ROWS_PER_FILE:
            raise MergeError(
                f"每个文件的数据行上限必须在 1 到 {MAX_DATA_ROWS_PER_FILE:,} 之间。"
            )
        if self.progress_interval_rows < 1:
            raise MergeError("进度更新间隔必须大于 0。")

        return MergeConfig(
            input_dir=input_dir,
            output_dir=output_dir,
            output_name=output_name,
            overwrite=self.overwrite,
            max_data_rows_per_file=self.max_data_rows_per_file,
            progress_interval_rows=self.progress_interval_rows,
        )


@dataclass(frozen=True)
class ProgressEvent:
    phase: Phase
    message: str
    current_file: str = ""
    files_completed: int = 0
    files_total: int = 0
    rows_written: int = 0
    total_rows: int = 0


@dataclass
class FileResult:
    path: Path
    status: Status
    rows_written: int = 0
    teacher: str = ""
    worksheet: str = ""
    reason: str = ""
    adjusted_cells: int = 0


@dataclass(frozen=True)
class MergeSummary:
    output_files: tuple[Path, ...]
    report_path: Path
    file_results: tuple[FileResult, ...]
    total_rows: int
    elapsed_seconds: float
    overwritten_files: tuple[Path, ...] = field(default_factory=tuple)


@dataclass(frozen=True)
class _InputMetadata:
    path: Path
    worksheet: str
    teacher: str
    header: tuple[object, ...]
    data_rows: int
    size: int
    modified_ns: int


ProgressCallback = Callable[[ProgressEvent], None]


def normalize_output_name(output_name: str) -> str:
    name = str(output_name).strip()
    if not name:
        raise MergeError("请输入输出文件名。")
    if name in {".", ".."} or Path(name).name != name:
        raise MergeError("输出文件名不能包含路径。")
    if _INVALID_OUTPUT_CHARS.search(name):
        raise MergeError('输出文件名不能包含 < > : " / \\ | ? * 等字符。')
    if not name.casefold().endswith(".xlsx"):
        name += ".xlsx"
    if not Path(name).stem.strip(" ."):
        raise MergeError("输出文件名无效。")
    if Path(name).stem.rstrip(" .").upper() in _WINDOWS_RESERVED_NAMES:
        raise MergeError("输出文件名是 Windows 保留名称，请更换名称。")
    return name


def find_input_files(
    input_dir: Path,
    *,
    excluded_paths: Iterable[Path] = (),
) -> list[Path]:
    root = Path(input_dir).resolve()
    excluded = {Path(path).resolve() for path in excluded_paths}
    files = [
        path.resolve()
        for path in root.iterdir()
        if path.is_file()
        and path.suffix.casefold() == ".xlsx"
        and not path.name.startswith("~$")
        and path.resolve() not in excluded
    ]
    return sorted(files, key=lambda path: (path.name.casefold(), path.name))


def find_output_conflicts(output_dir: Path, output_name: str) -> list[Path]:
    directory = Path(output_dir).expanduser().resolve()
    if not directory.is_dir():
        return []

    normalized_name = normalize_output_name(output_name)
    stem = Path(normalized_name).stem
    exact_name = normalized_name.casefold()
    part_pattern = re.compile(
        rf"^{re.escape(stem)}_第\d+部分\.xlsx$",
        flags=re.IGNORECASE,
    )
    matches = [
        path.resolve()
        for path in directory.iterdir()
        if path.is_file()
        and (path.name.casefold() == exact_name or part_pattern.fullmatch(path.name))
    ]
    return sorted(matches, key=lambda path: (path.name.casefold(), path.name))


def merge_workbooks(
    config: MergeConfig,
    *,
    cancel_event: threading.Event | None = None,
    progress_callback: ProgressCallback | None = None,
) -> MergeSummary:
    started = time.perf_counter()
    cfg = config.validated()
    cfg.output_dir.mkdir(parents=True, exist_ok=True)
    cancellation = cancel_event or threading.Event()

    conflicts = find_output_conflicts(cfg.output_dir, cfg.output_name)
    if conflicts and not cfg.overwrite:
        raise OutputExistsError(conflicts)

    excluded = conflicts if cfg.input_dir == cfg.output_dir else []
    input_files = find_input_files(cfg.input_dir, excluded_paths=excluded)
    if not input_files:
        raise MergeError("所选文件夹第一层没有可合并的 .xlsx 文件。")

    temp_dir = Path(
        tempfile.mkdtemp(prefix=".excel_merge_", dir=str(cfg.output_dir))
    ).resolve()
    writer: _OutputWriter | None = None
    published = False
    try:
        metadata, results, canonical_header = _preflight(
            input_files,
            cancellation,
            progress_callback,
        )
        _check_cancelled(cancellation)

        total_expected_rows = sum(item.data_rows for item in metadata)
        _emit(
            progress_callback,
            ProgressEvent(
                phase="merge_start",
                message=f"预检完成，准备合并 {len(metadata)} 个有效文件。",
                files_total=len(metadata),
                total_rows=total_expected_rows,
            ),
        )

        writer = _OutputWriter(
            temp_dir=temp_dir,
            header=canonical_header,
            max_data_rows=cfg.max_data_rows_per_file,
        )
        successful_results: dict[Path, FileResult] = {
            result.path: result for result in results if result.status == "成功"
        }
        total_written = 0
        valid_completed = 0

        for item in metadata:
            _check_cancelled(cancellation)
            current_stat = item.path.stat()
            if current_stat.st_size != item.size or current_stat.st_mtime_ns != item.modified_ns:
                result = successful_results[item.path]
                result.status = "跳过"
                result.reason = "文件在预检后发生变化，请重新运行。"
                continue

            _emit(
                progress_callback,
                ProgressEvent(
                    phase="merge_file",
                    message=f"正在合并：{item.path.name}",
                    current_file=item.path.name,
                    files_completed=valid_completed,
                    files_total=len(metadata),
                    rows_written=total_written,
                    total_rows=total_expected_rows,
                ),
            )

            file_written = 0
            adjusted_cells = 0
            workbook = None
            try:
                workbook = load_workbook(
                    item.path,
                    read_only=True,
                    data_only=True,
                    keep_links=False,
                )
                sheet = workbook[item.worksheet]
                _reset_read_only_dimensions(sheet)
                source_label = f"{item.path.stem} - {item.teacher}"
                for row_number, row in enumerate(
                    sheet.iter_rows(
                        min_row=2,
                        min_col=1,
                        max_col=INPUT_COLUMN_COUNT,
                        values_only=True,
                    ),
                    start=2,
                ):
                    if row_number % cfg.progress_interval_rows == 0:
                        _check_cancelled(cancellation)
                    values = _pad_row(row, INPUT_COLUMN_COUNT)
                    if _is_blank_row(values):
                        continue
                    adjusted_row, changed = _prepare_output_values(values)
                    adjusted_cells += changed
                    writer.append(adjusted_row, source_label)
                    file_written += 1
                    total_written += 1
                    if total_written % cfg.progress_interval_rows == 0:
                        _emit(
                            progress_callback,
                            ProgressEvent(
                                phase="merge_rows",
                                message=f"已写入 {total_written:,} 行",
                                current_file=item.path.name,
                                files_completed=valid_completed,
                                files_total=len(metadata),
                                rows_written=total_written,
                                total_rows=total_expected_rows,
                            ),
                        )
            except MergeCancelled:
                raise
            except Exception as exc:
                raise MergeError(
                    f"合并 {item.path.name} 时发生错误。为避免输出部分数据，本批次已取消：{exc}"
                ) from exc
            finally:
                if workbook is not None:
                    workbook.close()

            result = successful_results[item.path]
            result.rows_written = file_written
            result.adjusted_cells = adjusted_cells
            valid_completed += 1
            _emit(
                progress_callback,
                ProgressEvent(
                    phase="merge_rows",
                    message=f"已完成：{item.path.name}（{file_written:,} 行）",
                    current_file=item.path.name,
                    files_completed=valid_completed,
                    files_total=len(metadata),
                    rows_written=total_written,
                    total_rows=total_expected_rows,
                ),
            )

        _check_cancelled(cancellation)
        _emit(
            progress_callback,
            ProgressEvent(
                phase="saving",
                message="正在保存并发布结果，请稍候……",
                files_completed=valid_completed,
                files_total=len(metadata),
                rows_written=total_written,
                total_rows=total_expected_rows,
            ),
        )
        part_files = writer.finish()
        writer = None

        elapsed = time.perf_counter() - started
        target_paths = _target_output_paths(cfg.output_dir, cfg.output_name, len(part_files))
        effective_conflicts = conflicts if target_paths else []
        report_name = _unique_report_name(cfg.output_dir)
        report_temp_path = temp_dir / report_name
        report_target_path = cfg.output_dir / report_name
        report_text = _build_report_text(
            config=cfg,
            input_count=len(input_files),
            results=results,
            output_paths=target_paths,
            total_rows=total_written,
            elapsed_seconds=elapsed,
            overwritten_files=effective_conflicts,
        )
        report_temp_path.write_text(report_text, encoding="utf-8-sig")

        published_outputs, overwritten = _publish_results(
            part_files=part_files,
            target_paths=target_paths,
            report_temp_path=report_temp_path,
            report_target_path=report_target_path,
            conflicts=effective_conflicts,
            overwrite=cfg.overwrite,
            temp_dir=temp_dir,
        )
        published = True

        summary = MergeSummary(
            output_files=tuple(published_outputs),
            report_path=report_target_path,
            file_results=tuple(results),
            total_rows=total_written,
            elapsed_seconds=elapsed,
            overwritten_files=tuple(overwritten),
        )
        _emit(
            progress_callback,
            ProgressEvent(
                phase="done",
                message=f"合并完成，共写入 {total_written:,} 行。",
                files_completed=valid_completed,
                files_total=len(metadata),
                rows_written=total_written,
                total_rows=total_expected_rows,
            ),
        )
        return summary
    finally:
        if writer is not None:
            writer.abort()
        if not published or temp_dir.exists():
            _safe_remove_temp_dir(temp_dir, cfg.output_dir)


def _preflight(
    input_files: Sequence[Path],
    cancel_event: threading.Event,
    progress_callback: ProgressCallback | None,
) -> tuple[list[_InputMetadata], list[FileResult], tuple[object, ...]]:
    _emit(
        progress_callback,
        ProgressEvent(
            phase="scan_start",
            message=f"开始预检 {len(input_files)} 个文件……",
            files_total=len(input_files),
        ),
    )
    metadata: list[_InputMetadata] = []
    results: list[FileResult] = []
    canonical_header: tuple[object, ...] | None = None
    canonical_normalized: tuple[str, ...] | None = None

    for index, path in enumerate(input_files, start=1):
        _check_cancelled(cancel_event)
        _emit(
            progress_callback,
            ProgressEvent(
                phase="scan_file",
                message=f"正在预检：{path.name}",
                current_file=path.name,
                files_completed=index - 1,
                files_total=len(input_files),
            ),
        )
        workbook = None
        try:
            stat = path.stat()
            workbook = load_workbook(
                path,
                read_only=True,
                data_only=True,
                keep_links=False,
            )
            sheet = _first_visible_sheet(workbook)
            if sheet is None:
                raise MergeError("没有可见工作表。")
            _reset_read_only_dimensions(sheet)

            header_row = next(
                sheet.iter_rows(min_row=1, max_row=1, values_only=True),
                (),
            )
            if not header_row or _is_blank_row(header_row):
                raise MergeError("第一行表头为空。")
            output_header = _pad_row(header_row, INPUT_COLUMN_COUNT)
            normalized_header = tuple(_normalize_header(value) for value in output_header)

            teacher_indexes = [
                position
                for position, value in enumerate(header_row)
                if _normalize_header(value) == TEACHER_HEADER
            ]
            if not teacher_indexes:
                raise MergeError('首行缺少“主讲老师”字段。')
            if len(teacher_indexes) > 1:
                raise MergeError('首行存在多个“主讲老师”字段。')
            teacher_index = teacher_indexes[0]

            teacher = ""
            data_rows = 0
            max_col = max(INPUT_COLUMN_COUNT, teacher_index + 1)
            for row_number, row in enumerate(
                sheet.iter_rows(
                    min_row=2,
                    min_col=1,
                    max_col=max_col,
                    values_only=True,
                ),
                start=2,
            ):
                if row_number % PROGRESS_INTERVAL_ROWS == 0:
                    _check_cancelled(cancel_event)
                data_values = _pad_row(row, INPUT_COLUMN_COUNT)
                if not _is_blank_row(data_values):
                    data_rows += 1
                if not teacher and teacher_index < len(row):
                    teacher = _value_as_text(row[teacher_index])

            if not teacher:
                raise MergeError('“主讲老师”列的数据区没有非空值。')
            if data_rows == 0:
                raise MergeError("A:Q 数据区没有有效数据行。")

            if canonical_normalized is None:
                canonical_header = tuple(output_header)
                canonical_normalized = normalized_header
            elif normalized_header != canonical_normalized:
                raise MergeError("A:Q 表头与首个有效文件不一致。")

            metadata.append(
                _InputMetadata(
                    path=path,
                    worksheet=sheet.title,
                    teacher=teacher,
                    header=tuple(output_header),
                    data_rows=data_rows,
                    size=stat.st_size,
                    modified_ns=stat.st_mtime_ns,
                )
            )
            results.append(
                FileResult(
                    path=path,
                    status="成功",
                    teacher=teacher,
                    worksheet=sheet.title,
                )
            )
        except MergeCancelled:
            raise
        except Exception as exc:
            reason = str(exc).strip() or exc.__class__.__name__
            results.append(FileResult(path=path, status="跳过", reason=reason))
        finally:
            if workbook is not None:
                workbook.close()

        _emit(
            progress_callback,
            ProgressEvent(
                phase="scan_file",
                message=f"已预检 {index}/{len(input_files)} 个文件",
                current_file=path.name,
                files_completed=index,
                files_total=len(input_files),
            ),
        )

    if canonical_header is None:
        canonical_header = tuple(f"列{get_column_letter(index)}" for index in range(1, 18))

    _emit(
        progress_callback,
        ProgressEvent(
            phase="scan_done",
            message=f"预检完成：{len(metadata)} 个有效，{len(input_files) - len(metadata)} 个跳过。",
            files_completed=len(input_files),
            files_total=len(input_files),
            total_rows=sum(item.data_rows for item in metadata),
        ),
    )
    return metadata, results, canonical_header


class _OutputWriter:
    def __init__(self, *, temp_dir: Path, header: Sequence[object], max_data_rows: int):
        self.temp_dir = temp_dir
        self.header = tuple(_pad_row(header, INPUT_COLUMN_COUNT)) + (SOURCE_HEADER,)
        self.max_data_rows = max_data_rows
        self.part_files: list[Path] = []
        self.workbook: Workbook | None = None
        self.worksheet = None
        self.current_data_rows = 0
        self.total_rows = 0

    def append(self, values: Sequence[object], source_label: str) -> None:
        if self.workbook is None or self.current_data_rows >= self.max_data_rows:
            self._close_part()
            self._open_part()
        assert self.worksheet is not None
        output_values = list(values[:INPUT_COLUMN_COUNT]) + [source_label]
        self.worksheet.append(self._write_only_row(output_values))
        self.current_data_rows += 1
        self.total_rows += 1

    def finish(self) -> list[Path]:
        self._close_part()
        return list(self.part_files)

    def abort(self) -> None:
        if self.worksheet is not None:
            try:
                self.worksheet.close()
            except Exception:
                pass
        if self.workbook is not None:
            try:
                self.workbook.close()
            except Exception:
                pass
        self.workbook = None
        self.worksheet = None

    def _open_part(self) -> None:
        self.workbook = Workbook(write_only=True)
        self.worksheet = self.workbook.create_sheet(OUTPUT_SHEET_NAME)
        self.current_data_rows = 0

        self.worksheet.freeze_panes = "A2"
        self.worksheet.sheet_view.showGridLines = False
        self.worksheet.row_dimensions[1].height = 24
        for index in range(1, OUTPUT_COLUMN_COUNT + 1):
            letter = get_column_letter(index)
            self.worksheet.column_dimensions[letter].width = 34 if index == 18 else 15

        header_cells = []
        for value in self.header:
            cell = WriteOnlyCell(self.worksheet, value=value)
            cell.font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            header_cells.append(cell)
        self.worksheet.append(header_cells)

    def _close_part(self) -> None:
        if self.workbook is None or self.worksheet is None:
            return
        self.worksheet.auto_filter.ref = f"A1:R{self.current_data_rows + 1}"
        part_path = self.temp_dir / f"part_{len(self.part_files) + 1:04d}.xlsx"
        self.workbook.save(part_path)
        self.workbook.close()
        self.part_files.append(part_path)
        self.workbook = None
        self.worksheet = None

    def _write_only_row(self, values: Sequence[object]) -> list[object]:
        assert self.worksheet is not None
        row: list[object] = []
        for value in values:
            if isinstance(value, str) and value.startswith("="):
                cell = WriteOnlyCell(self.worksheet, value=value)
                cell.data_type = "s"
                row.append(cell)
            else:
                row.append(value)
        return row


def _first_visible_sheet(workbook):
    for sheet in workbook.worksheets:
        if sheet.sheet_state == "visible":
            return sheet
    return None


def _reset_read_only_dimensions(sheet) -> None:
    reset = getattr(sheet, "reset_dimensions", None)
    if callable(reset):
        reset()


def _pad_row(row: Sequence[object], width: int) -> tuple[object, ...]:
    values = tuple(row[:width])
    if len(values) < width:
        values += (None,) * (width - len(values))
    return values


def _normalize_header(value: object) -> str:
    return _value_as_text(value)


def _value_as_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return str(value).strip()


def _is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_blank_row(row: Sequence[object]) -> bool:
    return all(_is_blank(value) for value in row)


def _prepare_output_values(values: Sequence[object]) -> tuple[tuple[object, ...], int]:
    adjusted = 0
    output: list[object] = []
    for value in values[:INPUT_COLUMN_COUNT]:
        if isinstance(value, str):
            cleaned = _ILLEGAL_XML_CHARS.sub("", value)
            if len(cleaned) > 32_767:
                cleaned = cleaned[:32_767]
            if cleaned != value:
                adjusted += 1
            output.append(cleaned)
        else:
            output.append(value)
    return _pad_row(output, INPUT_COLUMN_COUNT), adjusted


def _target_output_paths(output_dir: Path, output_name: str, part_count: int) -> list[Path]:
    if part_count == 0:
        return []
    stem = Path(output_name).stem
    if part_count == 1:
        return [output_dir / output_name]
    return [
        output_dir / _PART_NAME_TEMPLATE.format(stem=stem, number=index)
        for index in range(1, part_count + 1)
    ]


def _unique_report_name(output_dir: Path) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = f"合并报告_{timestamp}"
    candidate = f"{base}.txt"
    counter = 2
    while (output_dir / candidate).exists():
        candidate = f"{base}_{counter}.txt"
        counter += 1
    return candidate


def _build_report_text(
    *,
    config: MergeConfig,
    input_count: int,
    results: Sequence[FileResult],
    output_paths: Sequence[Path],
    total_rows: int,
    elapsed_seconds: float,
    overwritten_files: Sequence[Path],
) -> str:
    succeeded = sum(1 for result in results if result.status == "成功")
    skipped = sum(1 for result in results if result.status == "跳过")
    lines = [
        "Excel 批量合并报告",
        "=" * 72,
        f"完成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"输入文件夹：{config.input_dir}",
        f"输出文件夹：{config.output_dir}",
        f"扫描文件数：{input_count}",
        f"成功文件数：{succeeded}",
        f"跳过文件数：{skipped}",
        f"合并数据行：{total_rows:,}",
        f"耗时：{elapsed_seconds:.2f} 秒",
        "",
        "输出文件：",
    ]
    if output_paths:
        lines.extend(f"- {path.name}" for path in output_paths)
    else:
        lines.append("- 无（没有有效数据）")

    if overwritten_files:
        lines.extend(["", "已确认替换的旧结果："])
        lines.extend(f"- {path.name}" for path in overwritten_files)

    lines.extend(["", "文件明细：", "-" * 72])
    for result in results:
        detail = f"[{result.status}] {result.path.name}"
        if result.worksheet:
            detail += f" | 工作表：{result.worksheet}"
        if result.teacher:
            detail += f" | 主讲老师：{result.teacher}"
        if result.status == "成功":
            detail += f" | 写入：{result.rows_written:,} 行"
            if result.adjusted_cells:
                detail += f" | 清理/截断单元格：{result.adjusted_cells}"
        if result.reason:
            detail += f" | 原因：{result.reason}"
        lines.append(detail)
    lines.append("")
    return "\n".join(lines)


def _publish_results(
    *,
    part_files: Sequence[Path],
    target_paths: Sequence[Path],
    report_temp_path: Path,
    report_target_path: Path,
    conflicts: Sequence[Path],
    overwrite: bool,
    temp_dir: Path,
) -> tuple[list[Path], list[Path]]:
    current_conflicts = [path for path in conflicts if path.exists()]
    if current_conflicts and not overwrite:
        raise OutputExistsError(current_conflicts)

    backup_dir = temp_dir / "backup"
    backup_dir.mkdir(exist_ok=True)
    moved_backups: list[tuple[Path, Path]] = []
    published_paths: list[Path] = []
    try:
        for original in current_conflicts:
            backup = backup_dir / f"{uuid.uuid4().hex}_{original.name}"
            original.replace(backup)
            moved_backups.append((original, backup))

        for source, target in zip(part_files, target_paths, strict=True):
            if target.exists():
                if not overwrite:
                    raise OutputExistsError([target])
                backup = backup_dir / f"{uuid.uuid4().hex}_{target.name}"
                target.replace(backup)
                moved_backups.append((target, backup))
            source.replace(target)
            published_paths.append(target)

        report_temp_path.replace(report_target_path)
        published_paths.append(report_target_path)
    except Exception:
        for path in reversed(published_paths):
            if path.exists():
                path.unlink()
        for original, backup in reversed(moved_backups):
            if backup.exists():
                backup.replace(original)
        raise

    return list(target_paths), [original for original, _ in moved_backups]


def _safe_remove_temp_dir(temp_dir: Path, output_dir: Path) -> None:
    try:
        resolved_temp = temp_dir.resolve()
        resolved_output = output_dir.resolve()
        if (
            resolved_temp.parent == resolved_output
            and resolved_temp.name.startswith(".excel_merge_")
            and resolved_temp.is_dir()
        ):
            shutil.rmtree(resolved_temp, ignore_errors=True)
    except OSError:
        pass


def _check_cancelled(cancel_event: threading.Event) -> None:
    if cancel_event.is_set():
        raise MergeCancelled("用户已取消本次合并。")


def _emit(callback: ProgressCallback | None, event: ProgressEvent) -> None:
    if callback is None:
        return
    try:
        callback(event)
    except Exception:
        # Progress reporting must never corrupt the merge itself.
        pass

