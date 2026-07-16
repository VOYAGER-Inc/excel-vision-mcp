"""Tests for workbook summary, sheet reading, pagination, and search."""

import pytest

from excel_vision_mcp.excel_reader import (
    get_workbook_summary,
    read_all_sheets_data,
    read_sheet_data,
    search_in_workbook,
)


class TestGetWorkbookSummary:
    def test_summary_structure(self, sample_workbook):
        summary = get_workbook_summary(str(sample_workbook))

        assert summary.file_name == "sample.xlsx"
        assert summary.file_size_bytes > 0
        assert summary.sheet_count == 2
        assert [s.name for s in summary.sheets] == ["Data", "Notes"]
        assert summary.total_image_count == 0

    def test_sheet_dimensions_and_merges(self, sample_workbook):
        summary = get_workbook_summary(str(sample_workbook))
        data_sheet = summary.sheets[0]

        # 1 title row + 1 header row + 30 data rows
        assert data_sheet.row_count == 32
        assert data_sheet.col_count == 3
        assert data_sheet.data_range == "A1:C32"
        assert data_sheet.merged_cell_count == 1
        assert data_sheet.has_images is False

    def test_has_images_flag(self, image_workbook):
        summary = get_workbook_summary(str(image_workbook))

        assert summary.total_image_count == 2
        assert summary.sheets[0].has_images is True

    def test_file_not_found(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            get_workbook_summary(str(tmp_path / "missing.xlsx"))

    def test_unsupported_extension(self, tmp_path):
        txt = tmp_path / "data.txt"
        txt.write_text("hello")
        with pytest.raises(ValueError):
            get_workbook_summary(str(txt))


class TestReadSheetData:
    def test_reads_cell_values(self, sample_workbook):
        data = read_sheet_data(str(sample_workbook), sheet_name="Data", max_rows=5)

        assert data.sheet_name == "Data"
        assert data.total_rows == 32
        assert data.total_cols == 3

        by_coord = {c.coordinate: c for c in data.cells}
        assert by_coord["A1"].value == "Quarterly Report"
        assert by_coord["A2"].value == "ID"
        assert by_coord["C3"].value == "10"

    def test_merged_cell_flags(self, sample_workbook):
        data = read_sheet_data(str(sample_workbook), sheet_name="Data", max_rows=5)

        by_coord = {c.coordinate: c for c in data.cells}
        assert by_coord["A1"].is_merged is True
        assert by_coord["A2"].is_merged is False
        assert "A1:C1" in data.merged_cells

    def test_pagination(self, sample_workbook):
        first = read_sheet_data(str(sample_workbook), sheet_name="Data", max_rows=10)
        assert first.start_row == 1
        assert first.end_row == 10
        assert first.has_more is True

        second = read_sheet_data(
            str(sample_workbook), sheet_name="Data", start_row=11, max_rows=100
        )
        assert second.start_row == 11
        assert second.end_row == 32
        assert second.has_more is False

        first_rows = {c.row for c in first.cells}
        second_rows = {c.row for c in second.cells}
        assert first_rows.isdisjoint(second_rows)

    def test_defaults_to_first_sheet(self, sample_workbook):
        data = read_sheet_data(str(sample_workbook))
        assert data.sheet_name == "Data"

    def test_missing_sheet_raises(self, sample_workbook):
        with pytest.raises(ValueError, match="not found"):
            read_sheet_data(str(sample_workbook), sheet_name="Nonexistent")

    def test_file_not_found(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            read_sheet_data(str(tmp_path / "missing.xlsx"))


class TestReadAllSheetsData:
    def test_returns_one_entry_per_sheet(self, sample_workbook):
        all_data = read_all_sheets_data(str(sample_workbook))

        assert [d.sheet_name for d in all_data] == ["Data", "Notes"]
        notes = all_data[1]
        by_coord = {c.coordinate: c.value for c in notes.cells}
        assert by_coord == {"B2": "Reviewed by QA", "B3": "Approved"}

    def test_row_limit_applies_per_sheet(self, sample_workbook):
        all_data = read_all_sheets_data(str(sample_workbook), max_rows_per_sheet=5)

        data_sheet = all_data[0]
        assert data_sheet.end_row == 5
        assert data_sheet.has_more is True


class TestSearchInWorkbook:
    def test_case_insensitive_match(self, sample_workbook):
        matches = search_in_workbook(str(sample_workbook), "quarterly")

        assert len(matches) == 1
        assert matches[0]["sheet"] == "Data"
        assert matches[0]["cell"] == "A1"
        assert matches[0]["value"] == "Quarterly Report"

    def test_case_sensitive_match(self, sample_workbook):
        assert search_in_workbook(str(sample_workbook), "quarterly", case_sensitive=True) == []
        assert len(search_in_workbook(str(sample_workbook), "Quarterly", case_sensitive=True)) == 1

    def test_sheet_filter(self, sample_workbook):
        matches = search_in_workbook(str(sample_workbook), "Approved", sheet_name="Data")
        assert matches == []

        matches = search_in_workbook(str(sample_workbook), "Approved", sheet_name="Notes")
        assert len(matches) == 1
        assert matches[0]["cell"] == "B3"

    def test_result_cap(self, sample_workbook):
        # "Item" appears in 30 cells; "1" appears in far more — verify the cap holds
        matches = search_in_workbook(str(sample_workbook), "1")
        assert len(matches) <= 100

    def test_no_match(self, sample_workbook):
        assert search_in_workbook(str(sample_workbook), "zzz-not-present") == []
