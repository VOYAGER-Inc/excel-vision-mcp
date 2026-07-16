"""
Excel MCP Server — FastMCP server with tools for reading Excel files and extracting embedded images.

Provides AI agents with comprehensive Excel reading capabilities including:
- Workbook structure discovery (sheets, dimensions, image counts)
- Paginated cell data reading for large files
- Embedded image extraction with cell-position mapping
- Full-content reading (text + images) for analysis workflows
- Text search across workbooks

All images are returned as base64-encoded ImageContent for multimodal AI consumption.

TODO: Add write capabilities (create/update cells, insert images) in future version.
"""

import logging
import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from mcp.types import ImageContent, TextContent

from .excel_reader import (
    CellData,
    SheetData,
    WorkbookSummary,
    get_workbook_summary as _get_workbook_summary,
    read_all_sheets_data,
    read_sheet_data,
    search_in_workbook,
)
from .image_extractor import ExtractionResult, extract_all_images
from .utils import format_file_size

# Route logs to stderr so they don't corrupt the stdio JSON-RPC stream
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)

mcp = FastMCP("ExcelReader")


def _validate_file_path(file_path: str) -> Path:
    """
    Validate and resolve the file path.

    @param file_path: User-provided file path.
    @return: Resolved Path object.
    @throws FileNotFoundError: If file doesn't exist.
    @throws ValueError: If file format is unsupported.
    """
    path = Path(file_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError(
            f"Unsupported format: {path.suffix}. Only .xlsx and .xlsm files are supported."
        )
    return path


def _format_sheet_data_as_text(sheet_data: SheetData) -> str:
    """
    Format sheet data as a readable text table for AI consumption.

    Builds a structured text representation with row/column headers,
    merged cell indicators, and pagination info.

    @param sheet_data: Parsed sheet data.
    @return: Formatted text string.
    """
    lines: list[str] = []
    lines.append(f"## Sheet: {sheet_data.sheet_name}")
    lines.append(f"Dimensions: {sheet_data.total_rows} rows × {sheet_data.total_cols} columns")
    lines.append(f"Showing rows: {sheet_data.start_row} to {sheet_data.end_row}")

    if sheet_data.merged_cells:
        lines.append(f"Merged cells: {', '.join(sheet_data.merged_cells[:20])}")
        if len(sheet_data.merged_cells) > 20:
            lines.append(f"  ... and {len(sheet_data.merged_cells) - 20} more merged ranges")

    lines.append("")

    if not sheet_data.cells:
        lines.append("(No data in this range)")
        return "\n".join(lines)

    # Group cells by row for table format
    rows_data: dict[int, list[CellData]] = {}
    for cell in sheet_data.cells:
        rows_data.setdefault(cell.row, []).append(cell)

    for row_num in sorted(rows_data.keys()):
        row_cells = sorted(rows_data[row_num], key=lambda c: c.col)
        cell_texts = []
        for c in row_cells:
            merged_marker = " [M]" if c.is_merged else ""
            val = c.value if c.value else ""
            # Truncate very long values to avoid context overflow
            if len(val) > 200:
                val = val[:200] + "..."
            cell_texts.append(f"{c.coordinate}: {val}{merged_marker}")
        lines.append(f"Row {row_num}: {' | '.join(cell_texts)}")

    if sheet_data.has_more:
        lines.append("")
        lines.append(
            f"⚠️ More data available. Showing {sheet_data.end_row - sheet_data.start_row + 1} "
            f"of {sheet_data.total_rows} rows. Use start_row={sheet_data.end_row + 1} for next page."
        )

    return "\n".join(lines)


@mcp.tool()
def list_sheets(file_path: str) -> str:
    """
    List all sheets in an Excel workbook with their dimensions and metadata.

    Returns sheet names, row/column counts, data ranges, merged cell counts,
    and total embedded image count. Use this to understand the structure of
    an Excel file before reading its contents.

    @param file_path: Absolute path to the .xlsx file.
    @return: Formatted text with workbook structure overview.
    """
    path = _validate_file_path(file_path)
    summary = _get_workbook_summary(str(path))

    lines = [
        f"# Workbook: {summary.file_name}",
        f"File size: {format_file_size(summary.file_size_bytes)}",
        f"Total sheets: {summary.sheet_count}",
        f"Total embedded images: {summary.total_image_count}",
        "",
        "## Sheets:",
    ]

    for i, sheet in enumerate(summary.sheets, 1):
        lines.append(f"\n### {i}. {sheet.name}")
        lines.append(f"  - Dimensions: {sheet.row_count} rows × {sheet.col_count} columns")
        lines.append(f"  - Data range: {sheet.data_range or 'Empty'}")
        lines.append(f"  - Merged cells: {sheet.merged_cell_count}")

    return "\n".join(lines)


@mcp.tool()
def read_excel_data(
    file_path: str,
    sheet_name: str | None = None,
    start_row: int = 1,
    max_rows: int = 200,
) -> str:
    """
    Read cell data from an Excel sheet with pagination for large files.

    Returns cell values organized by row with coordinate labels.
    Supports pagination via start_row and max_rows for efficient
    handling of large spreadsheets.

    @param file_path: Absolute path to the .xlsx file.
    @param sheet_name: Sheet to read. None = first/active sheet.
    @param start_row: Starting row number (1-indexed).
    @param max_rows: Maximum rows to return (default 200).
    @return: Formatted text table with cell values.
    """
    path = _validate_file_path(file_path)
    sheet_data = read_sheet_data(
        str(path),
        sheet_name=sheet_name,
        start_row=start_row,
        max_rows=max_rows,
    )
    return _format_sheet_data_as_text(sheet_data)


@mcp.tool()
def extract_images(
    file_path: str,
    sheet_name: str | None = None,
    max_width: int = 1024,
    max_height: int = 1024,
) -> list[TextContent | ImageContent]:
    """
    Extract all embedded images from an Excel file.

    Uses dual extraction strategy: cell-position mapping (primary) and
    ZIP archive scanning (fallback) to ensure no images are missed.
    Returns images as base64-encoded ImageContent that AI can visually analyze.

    @param file_path: Absolute path to the .xlsx file.
    @param sheet_name: Specific sheet, or None for all sheets.
    @param max_width: Max width in pixels for image optimization (default 1024).
    @param max_height: Max height in pixels for image optimization (default 1024).
    @return: Mixed list of TextContent (metadata) and ImageContent (images).
    """
    path = _validate_file_path(file_path)
    result = extract_all_images(
        str(path),
        sheet_name=sheet_name,
        max_width=max_width,
        max_height=max_height,
    )

    contents: list[TextContent | ImageContent] = []

    # Summary header
    summary_lines = [
        f"# Image Extraction Results",
        f"File: {path.name}",
        f"Total images found: {result.total_count}",
        f"Cell-mapped: {len(result.cell_mapped_images)}",
        f"Archive-only (orphan): {len(result.orphan_images)}",
    ]
    if result.errors:
        summary_lines.append(f"Errors: {'; '.join(result.errors)}")

    contents.append(TextContent(type="text", text="\n".join(summary_lines)))

    # Cell-mapped images with position context
    for img in result.cell_mapped_images:
        label = (
            f"[Image {img.index + 1}] Sheet: {img.sheet_name}, "
            f"Cell: {img.cell_reference}, "
            f"Size: {img.original_width}×{img.original_height}"
        )
        contents.append(TextContent(type="text", text=label))
        contents.append(ImageContent(
            type="image",
            data=img.data_base64,
            mimeType=img.mime_type,
        ))

    # Orphan images (from archive)
    for img in result.orphan_images:
        label = (
            f"[Orphan Image] File: {img.original_filename}, "
            f"Size: {img.original_width}×{img.original_height}"
        )
        contents.append(TextContent(type="text", text=label))
        contents.append(ImageContent(
            type="image",
            data=img.data_base64,
            mimeType=img.mime_type,
        ))

    if result.total_count == 0:
        contents.append(TextContent(type="text", text="No images found in this workbook."))

    return contents


@mcp.tool()
def read_full_content(
    file_path: str,
    max_rows_per_sheet: int = 500,
    max_image_width: int = 1024,
    max_image_height: int = 1024,
) -> list[TextContent | ImageContent]:
    """
    Read the FULL content of an Excel file including all text data AND embedded images.

    This is the primary tool for comprehensive document analysis. Returns all sheet
    data as structured text followed by all extracted images with their cell positions.
    Ideal for analyzing documents where both text and diagrams/screenshots are
    essential, such as requirement definitions, reports, or design specs.

    For very large files, data is paginated per sheet. Image extraction uses
    dual strategy (cell-mapping + archive) for maximum coverage.

    @param file_path: Absolute path to the .xlsx file.
    @param max_rows_per_sheet: Max rows to read per sheet (default 500).
    @param max_image_width: Max width for image optimization (default 1024).
    @param max_image_height: Max height for image optimization (default 1024).
    @return: Mixed list of TextContent and ImageContent covering entire workbook.
    """
    path = _validate_file_path(file_path)
    file_path_str = str(path)

    contents: list[TextContent | ImageContent] = []

    # Part 1: Workbook overview
    summary = _get_workbook_summary(file_path_str)
    overview = [
        f"# Excel Full Content: {summary.file_name}",
        f"File size: {format_file_size(summary.file_size_bytes)}",
        f"Sheets: {summary.sheet_count}",
        f"Total images: {summary.total_image_count}",
        "",
        "---",
        "",
    ]
    contents.append(TextContent(type="text", text="\n".join(overview)))

    # Part 2: All sheet data
    all_sheets = read_all_sheets_data(file_path_str, max_rows_per_sheet=max_rows_per_sheet)
    for sheet_data in all_sheets:
        text = _format_sheet_data_as_text(sheet_data)
        contents.append(TextContent(type="text", text=text + "\n\n---\n"))

    # Part 3: All images
    image_result = extract_all_images(
        file_path_str,
        max_width=max_image_width,
        max_height=max_image_height,
    )

    if image_result.total_count > 0:
        contents.append(TextContent(
            type="text",
            text=f"\n# Embedded Images ({image_result.total_count} total)\n",
        ))

        for img in image_result.cell_mapped_images:
            label = (
                f"**[Image {img.index + 1}]** Sheet: `{img.sheet_name}` | "
                f"Cell: `{img.cell_reference}` | "
                f"Original: {img.original_width}×{img.original_height}px"
            )
            contents.append(TextContent(type="text", text=label))
            contents.append(ImageContent(
                type="image",
                data=img.data_base64,
                mimeType=img.mime_type,
            ))

        for img in image_result.orphan_images:
            label = (
                f"**[Archive Image]** File: `{img.original_filename}` | "
                f"Original: {img.original_width}×{img.original_height}px"
            )
            contents.append(TextContent(type="text", text=label))
            contents.append(ImageContent(
                type="image",
                data=img.data_base64,
                mimeType=img.mime_type,
            ))

    return contents


@mcp.tool()
def get_workbook_overview(file_path: str) -> str:
    """
    Get a quick summary overview of an Excel workbook.

    Returns file metadata, sheet list with dimensions, image count,
    and merged cell information. Use this for a fast assessment
    before deeper analysis.

    @param file_path: Absolute path to the .xlsx file.
    @return: Formatted text summary.
    """
    path = _validate_file_path(file_path)
    summary = _get_workbook_summary(str(path))

    lines = [
        f"# Workbook Summary: {summary.file_name}",
        f"- **Path**: {summary.file_path}",
        f"- **Size**: {format_file_size(summary.file_size_bytes)}",
        f"- **Sheets**: {summary.sheet_count}",
        f"- **Total images**: {summary.total_image_count}",
        "",
    ]

    for sheet in summary.sheets:
        lines.append(
            f"| {sheet.name} | {sheet.row_count}×{sheet.col_count} | "
            f"Merged: {sheet.merged_cell_count} | Range: {sheet.data_range or 'Empty'} |"
        )

    return "\n".join(lines)


@mcp.tool()
def search_excel(
    file_path: str,
    query: str,
    sheet_name: str | None = None,
) -> str:
    """
    Search for text content across all cells in an Excel workbook.

    Performs case-insensitive substring search and returns matching cells
    with their coordinates and values. Limited to 100 results.

    @param file_path: Absolute path to the .xlsx file.
    @param query: Text to search for.
    @param sheet_name: Limit to specific sheet, or None for all sheets.
    @return: Formatted text with search results.
    """
    path = _validate_file_path(file_path)
    matches = search_in_workbook(str(path), query, sheet_name=sheet_name)

    if not matches:
        return f"No results found for '{query}' in {path.name}"

    lines = [
        f"# Search Results: '{query}' in {path.name}",
        f"Found {len(matches)} match(es):",
        "",
    ]

    for m in matches:
        value_preview = m["value"][:150] + "..." if len(m["value"]) > 150 else m["value"]
        lines.append(f"- **{m['sheet']}!{m['cell']}**: {value_preview}")

    return "\n".join(lines)


def main():
    """Entry point for the Excel MCP Server."""
    logger.info("Starting Excel MCP Server...")
    mcp.run()


if __name__ == "__main__":
    main()
