"""Tests for DocSpec parsing and validation."""
from __future__ import annotations

import pytest

from docx_cli.errors import SpecError
from docx_cli.spec import DocSpec


class TestSpecParsing:
    def test_package_and_project_versions_match(self):
        import re
        from pathlib import Path
        import docx_cli

        root = Path(__file__).parent.parent
        match = re.search(r'(?m)^version = "([^"]+)"', (root / "pyproject.toml").read_text())
        assert match is not None
        assert docx_cli.__version__ == match.group(1)
        from importlib.metadata import version
        assert docx_cli.__version__ == match.group(1)

    def test_minimal_spec_parses(self, minimal_spec_dict):
        spec = DocSpec.from_dict(minimal_spec_dict)
        assert spec.profile == "gost"
        assert spec.output.file_name == "test.docx"
        assert spec.title.main == "ТЕСТОВЫЙ ДОКУМЕНТ"

    def test_compact_profile_parses(self, minimal_spec_dict):
        minimal_spec_dict["profile"] = "compact"
        spec = DocSpec.from_dict(minimal_spec_dict)
        assert spec.profile == "compact"

    def test_unknown_profile_raises(self, minimal_spec_dict):
        minimal_spec_dict["profile"] = "fancy"
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_unknown_top_level_field_raises(self, minimal_spec_dict):
        minimal_spec_dict["extra_field"] = 123
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_misspelled_output_field_raises(self, minimal_spec_dict):
        minimal_spec_dict["output"] = {"file_name": "x.docx", "autor": "A"}
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_missing_output_raises(self, minimal_spec_dict):
        del minimal_spec_dict["output"]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_missing_header_raises(self, minimal_spec_dict):
        del minimal_spec_dict["header"]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_missing_title_raises(self, minimal_spec_dict):
        del minimal_spec_dict["title"]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_missing_body_raises(self, minimal_spec_dict):
        del minimal_spec_dict["body"]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_empty_body_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = []
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_empty_header_lines_raises(self, minimal_spec_dict):
        minimal_spec_dict["header"]["lines"] = []
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_wrong_type_for_profile_raises(self, minimal_spec_dict):
        minimal_spec_dict["profile"] = 123
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)


class TestBlockSpecs:
    def test_table_row_width_mismatch_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [
            {
                "type": "table",
                "headers": ["A", "B"],
                "rows": [["1", "2"], ["3"]],
            }
        ]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_table_widths_mismatch_headers_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [
            {
                "type": "table",
                "headers": ["A", "B"],
                "rows": [["1", "2"]],
                "widths_cm": [5.0, 5.0, 5.0],
            }
        ]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_table_empty_row_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [
            {"type": "table", "headers": [], "rows": [[]]}
        ]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_table_empty_cell_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [
            {"type": "table", "headers": ["A"], "rows": [[""]]}
        ]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_unknown_block_type_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [{"type": "bullet_list", "items": ["a"]}]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_paragraph_unknown_field_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [
            {"type": "paragraph", "text": "x", "font": "Arial"}
        ]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_numbered_list_empty_items_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [{"type": "numbered_list", "items": []}]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_attachments_list_empty_items_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [{"type": "attachments_list", "items": []}]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)

    def test_signature_block_missing_field_raises(self, minimal_spec_dict):
        minimal_spec_dict["body"] = [{"type": "signature_block", "left": "x"}]
        with pytest.raises(SpecError):
            DocSpec.from_dict(minimal_spec_dict)
