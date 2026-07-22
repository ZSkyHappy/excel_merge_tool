from __future__ import annotations

import json
import os
import re
import shutil
import unicodedata
import uuid
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Iterable, Mapping, Sequence


APP_NAME = "Excel 批量合并工具"
APP_VERSION = "2.0.0"
PROFILE_SCHEMA_VERSION = 1
MAX_OUTPUT_COLUMNS = 16_384

_WHITESPACE_RE = re.compile(r"\s+")
_INVALID_PROFILE_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_TOKEN_RE = re.compile(r"\{([^{}]+)\}")

TOKEN_FILE_NAME = "文件名"
TOKEN_FILE_STEM = "文件名不含扩展名"
TOKEN_RELATIVE_PATH = "相对路径"
TOKEN_WORKSHEET = "工作表"
SPECIAL_TOKEN_PREFIX = "特殊字段:"
BUILTIN_CLASSIC_PROFILE_NAME = "经典 A:Q 模式"


class ProfileError(ValueError):
    """Raised when a merge profile is invalid or cannot be loaded."""


def normalize_header(value: object) -> str:
    """Normalize a header for deterministic cross-file matching."""

    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return _WHITESPACE_RE.sub(" ", text.strip()).casefold()


def _clean_display_name(value: object, *, label: str) -> str:
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = _WHITESPACE_RE.sub(" ", text.strip())
    if not text:
        raise ProfileError(f"{label}不能为空。")
    return text


def _clean_aliases(values: Iterable[object], *, primary_name: str) -> tuple[str, ...]:
    aliases: list[str] = []
    seen = {normalize_header(primary_name)}
    for value in values:
        text = unicodedata.normalize("NFKC", str(value or ""))
        text = _WHITESPACE_RE.sub(" ", text.strip())
        normalized = normalize_header(text)
        if not normalized or normalized in seen:
            continue
        aliases.append(text)
        seen.add(normalized)
    return tuple(aliases)


@dataclass(frozen=True)
class OutputColumnRule:
    name: str
    aliases: tuple[str, ...] = field(default_factory=tuple)
    required: bool = True

    def validated(self) -> "OutputColumnRule":
        name = _clean_display_name(self.name, label="输出字段名")
        aliases = _clean_aliases(self.aliases, primary_name=name)
        return OutputColumnRule(name=name, aliases=aliases, required=bool(self.required))

    @property
    def candidates(self) -> tuple[str, ...]:
        return (self.name, *self.aliases)

    def to_dict(self) -> dict[str, object]:
        return {
            "name": self.name,
            "aliases": list(self.aliases),
            "required": self.required,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, object]) -> "OutputColumnRule":
        aliases = data.get("aliases", [])
        if not isinstance(aliases, list):
            raise ProfileError("输出字段 aliases 必须是数组。")
        return cls(
            name=str(data.get("name", "")),
            aliases=tuple(str(item) for item in aliases),
            required=bool(data.get("required", True)),
        ).validated()


@dataclass(frozen=True)
class SpecialFieldRule:
    name: str
    aliases: tuple[str, ...] = field(default_factory=tuple)

    def validated(self) -> "SpecialFieldRule":
        name = _clean_display_name(self.name, label="特殊字段名")
        aliases = _clean_aliases(self.aliases, primary_name=name)
        return SpecialFieldRule(name=name, aliases=aliases)

    @property
    def candidates(self) -> tuple[str, ...]:
        return (self.name, *self.aliases)

    def to_dict(self) -> dict[str, object]:
        return {"name": self.name, "aliases": list(self.aliases)}

    @classmethod
    def from_dict(cls, data: Mapping[str, object]) -> "SpecialFieldRule":
        aliases = data.get("aliases", [])
        if not isinstance(aliases, list):
            raise ProfileError("特殊字段 aliases 必须是数组。")
        return cls(
            name=str(data.get("name", "")),
            aliases=tuple(str(item) for item in aliases),
        ).validated()


@dataclass(frozen=True)
class LabelRule:
    enabled: bool = False
    header: str = "来源标签"
    template: str = "{文件名不含扩展名}"

    def validated(self) -> "LabelRule":
        if not self.enabled:
            return LabelRule(enabled=False, header=str(self.header), template=str(self.template))
        header = _clean_display_name(self.header, label="标签列标题")
        template = str(self.template).strip()
        if not template:
            raise ProfileError("标签模板不能为空。")
        if "{" in _TOKEN_RE.sub("", template) or "}" in _TOKEN_RE.sub("", template):
            raise ProfileError("标签模板的大括号不完整。")
        return LabelRule(enabled=True, header=header, template=template)

    def to_dict(self) -> dict[str, object]:
        return {
            "enabled": self.enabled,
            "header": self.header,
            "template": self.template,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, object]) -> "LabelRule":
        return cls(
            enabled=bool(data.get("enabled", False)),
            header=str(data.get("header", "来源标签")),
            template=str(data.get("template", "{文件名不含扩展名}")),
        ).validated()


@dataclass(frozen=True)
class CsvOptions:
    encoding: str = "auto"
    delimiter: str = "auto"

    def validated(self) -> "CsvOptions":
        encoding = str(self.encoding or "auto").strip().lower()
        if encoding not in {"auto", "utf-8", "utf-8-sig", "gb18030"}:
            raise ProfileError("CSV 编码仅支持自动、UTF-8、UTF-8 BOM 或 GB18030。")
        delimiter = str(self.delimiter or "auto")
        allowed = {"auto", ",", "\t", ";", "|"}
        if delimiter not in allowed:
            raise ProfileError("CSV 分隔符仅支持自动、逗号、制表符、分号或竖线。")
        return CsvOptions(encoding=encoding, delimiter=delimiter)

    def to_dict(self) -> dict[str, object]:
        return {"encoding": self.encoding, "delimiter": self.delimiter}

    @classmethod
    def from_dict(cls, data: Mapping[str, object]) -> "CsvOptions":
        return cls(
            encoding=str(data.get("encoding", "auto")),
            delimiter=str(data.get("delimiter", "auto")),
        ).validated()


@dataclass(frozen=True)
class MergeProfile:
    name: str
    output_columns: tuple[OutputColumnRule, ...] = field(default_factory=tuple)
    special_fields: tuple[SpecialFieldRule, ...] = field(default_factory=tuple)
    recognize_special_fields: bool = False
    label: LabelRule = field(default_factory=LabelRule)
    recursive: bool = False
    default_header_row: int = 1
    csv_options: CsvOptions = field(default_factory=CsvOptions)
    classic: bool = False
    schema_version: int = PROFILE_SCHEMA_VERSION

    def validated(self) -> "MergeProfile":
        if int(self.schema_version) != PROFILE_SCHEMA_VERSION:
            raise ProfileError(
                f"不支持方案版本 {self.schema_version}，当前仅支持 {PROFILE_SCHEMA_VERSION}。"
            )
        name = _clean_display_name(self.name, label="方案名称")
        header_row = int(self.default_header_row)
        if header_row < 1 or header_row > 1_048_576:
            raise ProfileError("默认表头行必须在 1 到 1,048,576 之间。")

        output_columns = tuple(item.validated() for item in self.output_columns)
        special_fields = tuple(item.validated() for item in self.special_fields)
        label = self.label.validated()
        csv_options = self.csv_options.validated()

        if self.classic:
            output_columns = ()
        elif not output_columns:
            raise ProfileError("自定义方案至少需要一个输出字段。")

        output_count = (17 if self.classic else len(output_columns)) + int(label.enabled)
        if output_count > MAX_OUTPUT_COLUMNS:
            raise ProfileError(f"输出列数不能超过 {MAX_OUTPUT_COLUMNS:,}。")

        _validate_unique_rule_candidates(output_columns, "输出字段")
        _validate_unique_rule_candidates(special_fields, "特殊字段")

        output_names = {normalize_header(item.name) for item in output_columns}
        if label.enabled and normalize_header(label.header) in output_names:
            raise ProfileError("标签列标题不能与输出字段重名。")

        configured_specials = {normalize_header(item.name) for item in special_fields}
        for token in label_tokens(label.template) if label.enabled else ():
            if token in {TOKEN_FILE_NAME, TOKEN_FILE_STEM, TOKEN_RELATIVE_PATH, TOKEN_WORKSHEET}:
                continue
            if token.startswith(SPECIAL_TOKEN_PREFIX):
                special_name = token[len(SPECIAL_TOKEN_PREFIX) :]
                if not self.recognize_special_fields:
                    raise ProfileError("标签模板引用了特殊字段，但特殊字段识别未启用。")
                if normalize_header(special_name) not in configured_specials:
                    raise ProfileError(f"标签模板引用了未配置的特殊字段：{special_name}")
                continue
            raise ProfileError(f"标签模板包含未知变量：{{{token}}}")

        if self.recognize_special_fields and not special_fields:
            raise ProfileError("启用特殊字段识别后，至少需要配置一个特殊字段。")

        return MergeProfile(
            name=name,
            output_columns=output_columns,
            special_fields=special_fields,
            recognize_special_fields=bool(self.recognize_special_fields),
            label=label,
            recursive=bool(self.recursive),
            default_header_row=header_row,
            csv_options=csv_options,
            classic=bool(self.classic),
            schema_version=PROFILE_SCHEMA_VERSION,
        )

    @property
    def output_headers(self) -> tuple[str, ...]:
        headers = tuple(item.name for item in self.output_columns)
        return headers + ((self.label.header,) if self.label.enabled else ())

    def with_name(self, name: str) -> "MergeProfile":
        return replace(self, name=name).validated()

    def to_dict(self) -> dict[str, object]:
        validated = self.validated()
        return {
            "schema_version": validated.schema_version,
            "name": validated.name,
            "classic": validated.classic,
            "recursive": validated.recursive,
            "default_header_row": validated.default_header_row,
            "recognize_special_fields": validated.recognize_special_fields,
            "output_columns": [item.to_dict() for item in validated.output_columns],
            "special_fields": [item.to_dict() for item in validated.special_fields],
            "label": validated.label.to_dict(),
            "csv_options": validated.csv_options.to_dict(),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, object]) -> "MergeProfile":
        output_data = data.get("output_columns", [])
        special_data = data.get("special_fields", [])
        label_data = data.get("label", {})
        csv_data = data.get("csv_options", {})
        if not isinstance(output_data, list) or not all(isinstance(item, dict) for item in output_data):
            raise ProfileError("output_columns 必须是对象数组。")
        if not isinstance(special_data, list) or not all(isinstance(item, dict) for item in special_data):
            raise ProfileError("special_fields 必须是对象数组。")
        if not isinstance(label_data, dict) or not isinstance(csv_data, dict):
            raise ProfileError("label 和 csv_options 必须是对象。")
        return cls(
            name=str(data.get("name", "")),
            output_columns=tuple(OutputColumnRule.from_dict(item) for item in output_data),
            special_fields=tuple(SpecialFieldRule.from_dict(item) for item in special_data),
            recognize_special_fields=bool(data.get("recognize_special_fields", False)),
            label=LabelRule.from_dict(label_data),
            recursive=bool(data.get("recursive", False)),
            default_header_row=int(data.get("default_header_row", 1)),
            csv_options=CsvOptions.from_dict(csv_data),
            classic=bool(data.get("classic", False)),
            schema_version=int(data.get("schema_version", 0)),
        ).validated()


def _validate_unique_rule_candidates(
    rules: Sequence[OutputColumnRule | SpecialFieldRule],
    label: str,
) -> None:
    owners: dict[str, int] = {}
    for rule_index, rule in enumerate(rules):
        for candidate in rule.candidates:
            normalized = normalize_header(candidate)
            previous = owners.get(normalized)
            if previous is not None and previous != rule_index:
                previous_name = rules[previous].name
                raise ProfileError(
                    f'{label}“{previous_name}”和“{rule.name}”使用了相同名称或别名“{candidate}”。'
                )
            owners[normalized] = rule_index


def classic_profile() -> MergeProfile:
    return MergeProfile(
        name=BUILTIN_CLASSIC_PROFILE_NAME,
        special_fields=(SpecialFieldRule("主讲老师"),),
        recognize_special_fields=True,
        label=LabelRule(
            enabled=True,
            header="所属表格+主讲老师",
            template="{文件名不含扩展名} - {特殊字段:主讲老师}",
        ),
        classic=True,
    ).validated()


def create_profile_from_headers(
    name: str,
    headers: Sequence[object],
    *,
    header_row: int = 1,
    recursive: bool = False,
) -> MergeProfile:
    columns: list[OutputColumnRule] = []
    seen: set[str] = set()
    for index, value in enumerate(headers, start=1):
        display = _WHITESPACE_RE.sub(" ", str(value or "").strip())
        if not display:
            continue
        normalized = normalize_header(display)
        if normalized in seen:
            raise ProfileError(f"模板表头存在重复字段：{display}")
        seen.add(normalized)
        columns.append(OutputColumnRule(display, required=True))
    if not columns:
        raise ProfileError("模板表头没有可用字段。")
    return MergeProfile(
        name=name,
        output_columns=tuple(columns),
        recursive=recursive,
        default_header_row=header_row,
    ).validated()


def label_tokens(template: str) -> tuple[str, ...]:
    return tuple(match.group(1).strip() for match in _TOKEN_RE.finditer(str(template)))


def render_label(
    rule: LabelRule,
    *,
    file_name: str,
    relative_path: str,
    worksheet: str,
    special_values: Mapping[str, str],
) -> str:
    if not rule.enabled:
        return ""
    normalized_specials = {normalize_header(key): value for key, value in special_values.items()}
    file_stem = Path(file_name).stem

    def replace_token(match: re.Match[str]) -> str:
        token = match.group(1).strip()
        if token == TOKEN_FILE_NAME:
            return file_name
        if token == TOKEN_FILE_STEM:
            return file_stem
        if token == TOKEN_RELATIVE_PATH:
            return relative_path
        if token == TOKEN_WORKSHEET:
            return worksheet
        if token.startswith(SPECIAL_TOKEN_PREFIX):
            name = token[len(SPECIAL_TOKEN_PREFIX) :]
            return normalized_specials.get(normalize_header(name), "")
        raise ProfileError(f"标签模板包含未知变量：{{{token}}}")

    return _TOKEN_RE.sub(replace_token, rule.template)


def load_profile(path: Path) -> MergeProfile:
    profile_path = Path(path)
    try:
        data = json.loads(profile_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ProfileError(f"无法读取方案 {profile_path.name}：{exc}") from exc
    if not isinstance(data, dict):
        raise ProfileError("方案文件的根内容必须是对象。")
    return MergeProfile.from_dict(data)


def save_profile(profile: MergeProfile, path: Path) -> Path:
    validated = profile.validated()
    profile_path = Path(path)
    profile_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = profile_path.with_name(f".{profile_path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temp_path.write_text(
            json.dumps(validated.to_dict(), ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        temp_path.replace(profile_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return profile_path


def default_app_data_dir() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / "ExcelMergeTool"
    return Path.home() / ".excel_merge_tool"


class ProfileStore:
    def __init__(self, directory: Path | None = None):
        self.directory = Path(directory) if directory is not None else default_app_data_dir() / "profiles"

    def list_profiles(self) -> list[tuple[Path, MergeProfile]]:
        if not self.directory.is_dir():
            return []
        profiles: list[tuple[Path, MergeProfile]] = []
        for path in sorted(self.directory.glob("*.json"), key=lambda item: item.name.casefold()):
            try:
                profile = load_profile(path)
            except ProfileError:
                continue
            if not profile.classic:
                profiles.append((path, profile))
        return profiles

    def save(self, profile: MergeProfile, *, existing_path: Path | None = None) -> Path:
        validated = profile.validated()
        if validated.classic:
            raise ProfileError("内置经典方案不能保存或覆盖。")
        if existing_path is None:
            self.directory.mkdir(parents=True, exist_ok=True)
            base = _profile_file_stem(validated.name)
            path = self.directory / f"{base}.json"
            counter = 2
            while path.exists():
                path = self.directory / f"{base}_{counter}.json"
                counter += 1
        else:
            path = Path(existing_path)
            if path.resolve().parent != self.directory.resolve():
                raise ProfileError("只能覆盖本机方案目录中的文件。")
        return save_profile(validated, path)

    def import_profile(self, source: Path) -> tuple[Path, MergeProfile]:
        profile = load_profile(source)
        if profile.classic:
            profile = replace(profile, classic=False).validated()
        path = self.save(profile)
        return path, profile

    def export_profile(self, profile: MergeProfile, target: Path) -> Path:
        target_path = Path(target)
        if target_path.suffix.casefold() != ".json":
            target_path = target_path.with_suffix(".json")
        return save_profile(replace(profile, classic=False), target_path)

    def delete(self, path: Path) -> None:
        candidate = Path(path).resolve()
        if candidate.parent != self.directory.resolve():
            raise ProfileError("只能删除本机方案目录中的文件。")
        candidate.unlink()


def _profile_file_stem(name: str) -> str:
    stem = _INVALID_PROFILE_FILENAME.sub("_", name).strip(" ._")
    return stem[:80] or "merge_profile"


def copy_profile_file(source: Path, target: Path) -> Path:
    """Copy a validated profile without preserving unrelated file metadata."""

    profile = load_profile(source)
    target_path = save_profile(profile, target)
    try:
        shutil.copymode(source, target_path)
    except OSError:
        pass
    return target_path
