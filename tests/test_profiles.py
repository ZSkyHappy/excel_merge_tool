from __future__ import annotations

import json
import unittest
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory

from merge_profiles import (
    CsvOptions,
    LabelRule,
    MergeProfile,
    OutputColumnRule,
    ProfileError,
    ProfileStore,
    SpecialFieldRule,
    classic_profile,
    create_profile_from_headers,
    load_profile,
    normalize_header,
    render_label,
    save_profile,
)


def sample_profile() -> MergeProfile:
    return MergeProfile(
        name="课程合并",
        output_columns=(
            OutputColumnRule("编号", aliases=("序号",), required=True),
            OutputColumnRule("姓名", aliases=("学员姓名",), required=False),
        ),
        special_fields=(
            SpecialFieldRule("主讲老师", aliases=("讲师",)),
            SpecialFieldRule("部门"),
        ),
        recognize_special_fields=True,
        label=LabelRule(
            enabled=True,
            header="来源标签",
            template="{文件名不含扩展名}-{特殊字段:主讲老师}-{特殊字段:部门}",
        ),
        recursive=True,
        default_header_row=2,
        csv_options=CsvOptions(encoding="gb18030", delimiter=";"),
    ).validated()


class ProfileTests(unittest.TestCase):
    def test_normalize_header_nfkc_case_and_whitespace(self) -> None:
        self.assertEqual(normalize_header("  ＩＤ\tName  "), "id name")

    def test_output_rule_cleans_duplicate_aliases(self) -> None:
        rule = OutputColumnRule("编号", aliases=(" 编号 ", "序号", "序号")).validated()
        self.assertEqual(rule.aliases, ("序号",))

    def test_output_rule_requires_name(self) -> None:
        with self.assertRaisesRegex(ProfileError, "不能为空"):
            OutputColumnRule(" ").validated()

    def test_special_rule_requires_name(self) -> None:
        with self.assertRaises(ProfileError):
            SpecialFieldRule("").validated()

    def test_csv_options_accept_supported_values(self) -> None:
        self.assertEqual(CsvOptions("UTF-8", "\t").validated(), CsvOptions("utf-8", "\t"))

    def test_csv_options_reject_unknown_encoding(self) -> None:
        with self.assertRaisesRegex(ProfileError, "编码"):
            CsvOptions("big5", ",").validated()

    def test_csv_options_reject_unknown_delimiter(self) -> None:
        with self.assertRaisesRegex(ProfileError, "分隔符"):
            CsvOptions("auto", ":").validated()

    def test_profile_rejects_duplicate_output_names(self) -> None:
        profile = MergeProfile(
            name="重复",
            output_columns=(OutputColumnRule("编号"), OutputColumnRule(" 编号 ")),
        )
        with self.assertRaisesRegex(ProfileError, "相同名称"):
            profile.validated()

    def test_profile_rejects_alias_collision(self) -> None:
        profile = MergeProfile(
            name="重复别名",
            output_columns=(
                OutputColumnRule("编号", aliases=("ID",)),
                OutputColumnRule("姓名", aliases=("id",)),
            ),
        )
        with self.assertRaisesRegex(ProfileError, "相同名称"):
            profile.validated()

    def test_profile_rejects_label_header_collision(self) -> None:
        profile = MergeProfile(
            name="标签重名",
            output_columns=(OutputColumnRule("来源"),),
            label=LabelRule(True, "来源", "{文件名}"),
        )
        with self.assertRaisesRegex(ProfileError, "重名"):
            profile.validated()

    def test_profile_rejects_unknown_label_token(self) -> None:
        profile = MergeProfile(
            name="未知变量",
            output_columns=(OutputColumnRule("编号"),),
            label=LabelRule(True, "来源", "{未知变量}"),
        )
        with self.assertRaisesRegex(ProfileError, "未知变量"):
            profile.validated()

    def test_profile_requires_special_recognition_for_special_token(self) -> None:
        profile = MergeProfile(
            name="特殊变量",
            output_columns=(OutputColumnRule("编号"),),
            special_fields=(SpecialFieldRule("老师"),),
            label=LabelRule(True, "来源", "{特殊字段:老师}"),
        )
        with self.assertRaisesRegex(ProfileError, "未启用"):
            profile.validated()

    def test_profile_requires_special_fields_when_enabled(self) -> None:
        profile = MergeProfile(
            name="缺特殊字段",
            output_columns=(OutputColumnRule("编号"),),
            recognize_special_fields=True,
        )
        with self.assertRaisesRegex(ProfileError, "至少"):
            profile.validated()

    def test_label_disabled_does_not_validate_template(self) -> None:
        profile = MergeProfile(
            name="关闭标签",
            output_columns=(OutputColumnRule("编号"),),
            label=LabelRule(False, "", "{不完整"),
        ).validated()
        self.assertFalse(profile.label.enabled)

    def test_create_profile_from_headers_skips_blank_headers(self) -> None:
        profile = create_profile_from_headers("模板", ["编号", None, "姓名"])
        self.assertEqual([item.name for item in profile.output_columns], ["编号", "姓名"])

    def test_create_profile_from_headers_rejects_duplicate_headers(self) -> None:
        with self.assertRaisesRegex(ProfileError, "重复"):
            create_profile_from_headers("模板", ["编号", " 编号 "])

    def test_classic_profile_keeps_expected_contract(self) -> None:
        profile = classic_profile()
        self.assertTrue(profile.classic)
        self.assertTrue(profile.recognize_special_fields)
        self.assertEqual(profile.special_fields[0].name, "主讲老师")
        self.assertEqual(profile.label.header, "所属表格+主讲老师")

    def test_render_label_supports_all_variables(self) -> None:
        rule = LabelRule(
            True,
            "来源",
            "{文件名}|{文件名不含扩展名}|{相对路径}|{工作表}|{特殊字段:老师}",
        )
        value = render_label(
            rule,
            file_name="课程.xlsx",
            relative_path="一部/课程.xlsx",
            worksheet="数据",
            special_values={"老师": "张老师"},
        )
        self.assertEqual(value, "课程.xlsx|课程|一部/课程.xlsx|数据|张老师")

    def test_profile_json_round_trip(self) -> None:
        with TemporaryDirectory() as temp:
            path = Path(temp) / "profile.json"
            saved = save_profile(sample_profile(), path)
            self.assertEqual(saved, path)
            self.assertEqual(load_profile(path).to_dict(), sample_profile().to_dict())

    def test_profile_rejects_future_schema_version(self) -> None:
        data = sample_profile().to_dict()
        data["schema_version"] = 99
        with self.assertRaisesRegex(ProfileError, "不支持"):
            MergeProfile.from_dict(data)

    def test_invalid_json_is_reported(self) -> None:
        with TemporaryDirectory() as temp:
            path = Path(temp) / "broken.json"
            path.write_text("{", encoding="utf-8")
            with self.assertRaisesRegex(ProfileError, "无法读取"):
                load_profile(path)

    def test_profile_store_save_list_export_import_and_delete(self) -> None:
        with TemporaryDirectory() as temp:
            root = Path(temp)
            store = ProfileStore(root / "profiles")
            path = store.save(sample_profile())
            listed = store.list_profiles()
            self.assertEqual(len(listed), 1)
            self.assertEqual(listed[0][1].name, "课程合并")
            exported = store.export_profile(sample_profile(), root / "shared")
            self.assertEqual(exported.suffix, ".json")
            imported_path, imported = store.import_profile(exported)
            self.assertTrue(imported_path.exists())
            self.assertEqual(imported.name, "课程合并")
            store.delete(path)
            self.assertFalse(path.exists())

    def test_profile_dict_contains_no_local_paths(self) -> None:
        payload = json.dumps(sample_profile().to_dict(), ensure_ascii=False)
        self.assertNotIn("C:\\", payload)
        self.assertNotIn("input_dir", payload)

    def test_with_name_returns_validated_copy(self) -> None:
        renamed = sample_profile().with_name("新方案")
        self.assertEqual(renamed.name, "新方案")
        self.assertEqual(renamed.output_columns, sample_profile().output_columns)

    def test_classic_cannot_be_saved_in_profile_store(self) -> None:
        with TemporaryDirectory() as temp:
            with self.assertRaisesRegex(ProfileError, "内置"):
                ProfileStore(Path(temp)).save(classic_profile())


if __name__ == "__main__":
    unittest.main()
