"""Tests for MCP tool functions and path validation in the server layer."""

import pytest
from mcp.types import ImageContent, TextContent

from excel_vision_mcp.server import (
    _validate_file_path,
    extract_images,
    get_workbook_overview,
    list_sheets,
    read_excel_data,
    read_full_content,
    search_excel,
)


class TestValidateFilePath:
    def test_valid_xlsx(self, sample_workbook):
        path = _validate_file_path(str(sample_workbook))
        assert path.is_absolute()
        assert path.exists()

    def test_missing_file(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            _validate_file_path(str(tmp_path / "missing.xlsx"))

    def test_unsupported_extension(self, tmp_path):
        csv = tmp_path / "data.csv"
        csv.write_text("a,b,c")
        with pytest.raises(ValueError, match="Unsupported format"):
            _validate_file_path(str(csv))


class TestListSheets:
    def test_output_contains_structure(self, sample_workbook):
        text = list_sheets(str(sample_workbook))

        assert "sample.xlsx" in text
        assert "Total sheets: 2" in text
        assert "Data" in text
        assert "Notes" in text


class TestReadExcelData:
    def test_formats_rows_and_pagination_hint(self, sample_workbook):
        text = read_excel_data(str(sample_workbook), sheet_name="Data", max_rows=5)

        assert "## Sheet: Data" in text
        assert "Quarterly Report" in text
        assert "[M]" in text  # merged-cell marker
        assert "More data available" in text
        assert "start_row=6" in text


class TestExtractImages:
    def test_returns_text_and_image_content(self, image_workbook):
        contents = extract_images(str(image_workbook))

        images = [c for c in contents if isinstance(c, ImageContent)]
        texts = [c for c in contents if isinstance(c, TextContent)]
        assert len(images) == 2
        assert any("Total images found: 2" in t.text for t in texts)

    def test_no_images_message(self, sample_workbook):
        contents = extract_images(str(sample_workbook))

        assert all(isinstance(c, TextContent) for c in contents)
        assert any("No images found" in c.text for c in contents)


class TestReadFullContent:
    def test_includes_all_sheets_and_images(self, image_workbook):
        contents = read_full_content(str(image_workbook))

        full_text = "\n".join(c.text for c in contents if isinstance(c, TextContent))
        images = [c for c in contents if isinstance(c, ImageContent)]

        assert "images.xlsx" in full_text
        assert "## Sheet: Diagrams" in full_text
        assert "Diagram sheet" in full_text
        assert len(images) == 2


class TestGetWorkbookOverview:
    def test_summary_lines(self, sample_workbook):
        text = get_workbook_overview(str(sample_workbook))

        assert "sample.xlsx" in text
        assert "**Sheets**: 2" in text


class TestSearchExcel:
    def test_match_formatting(self, sample_workbook):
        text = search_excel(str(sample_workbook), "Quarterly")

        assert "Found 1 match(es)" in text
        assert "Data!A1" in text

    def test_no_match_message(self, sample_workbook):
        text = search_excel(str(sample_workbook), "zzz-not-present")
        assert "No results found" in text
