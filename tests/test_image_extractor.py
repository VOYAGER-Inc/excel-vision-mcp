"""Tests for dual-strategy image extraction and archive counting."""

from excel_mcp.image_extractor import count_images_in_archive, extract_all_images


class TestCountImagesInArchive:
    def test_counts_embedded_images(self, image_workbook):
        assert count_images_in_archive(str(image_workbook)) == 2

    def test_zero_for_workbook_without_images(self, sample_workbook):
        assert count_images_in_archive(str(sample_workbook)) == 0

    def test_zero_for_invalid_archive(self, not_an_xlsx):
        assert count_images_in_archive(str(not_an_xlsx)) == 0


class TestExtractAllImages:
    def test_extracts_cell_mapped_images(self, image_workbook):
        result = extract_all_images(str(image_workbook))

        assert result.total_count == 2
        assert len(result.cell_mapped_images) == 2
        assert result.errors == []

        refs = {img.cell_reference for img in result.cell_mapped_images}
        assert refs == {"B2", "D5"}
        for img in result.cell_mapped_images:
            assert img.sheet_name == "Diagrams"
            assert img.data_base64
            assert img.mime_type.startswith("image/")
            assert img.original_width > 0
            assert img.original_height > 0

    def test_dedup_leaves_no_orphans(self, image_workbook):
        # Both archive images are already cell-mapped, so the archive
        # fallback must dedup them all — including images whose encoded
        # headers are near-identical (same encoder settings).
        result = extract_all_images(str(image_workbook))
        assert result.orphan_images == []

    def test_no_images(self, sample_workbook):
        result = extract_all_images(str(sample_workbook))

        assert result.total_count == 0
        assert result.cell_mapped_images == []
        assert result.orphan_images == []

    def test_sheet_filter(self, image_workbook):
        result = extract_all_images(str(image_workbook), sheet_name="Diagrams")
        assert len(result.cell_mapped_images) == 2

        # Nonexistent sheet: no cell-mapped hits, archive images become orphans
        result = extract_all_images(str(image_workbook), sheet_name="Nope")
        assert result.cell_mapped_images == []
        assert len(result.orphan_images) == 2

    def test_resizes_to_max_dimensions(self, large_image_workbook):
        result = extract_all_images(
            str(large_image_workbook), max_width=512, max_height=512
        )

        assert result.total_count == 1
        img = result.cell_mapped_images[0]
        assert img.original_width == 1600
        assert img.original_height == 800
        assert img.processed_width <= 512
        assert img.processed_height <= 512
        # Aspect ratio preserved (2:1)
        assert abs(img.processed_width / img.processed_height - 2.0) < 0.05

    def test_invalid_archive_reports_errors_gracefully(self, not_an_xlsx):
        result = extract_all_images(str(not_an_xlsx))

        assert result.total_count == 0
        # Cell-mapping strategy fails on a broken zip and must be reported,
        # not raised.
        assert result.errors
