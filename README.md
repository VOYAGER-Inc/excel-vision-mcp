<div align="center">

# 📊 Excel Vision MCP

**The first MCP server that lets AI agents _see_ images inside your spreadsheets.**

Read Excel files with full content extraction — cell data, formulas, merged cells, **and embedded images** — all returned as multimodal content your AI can actually understand.

[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![MCP](https://img.shields.io/badge/MCP-Compatible-purple.svg)](https://modelcontextprotocol.io)
[![PyPI](https://img.shields.io/pypi/v/excel-vision-mcp.svg)](https://pypi.org/project/excel-vision-mcp/)

[Installation](#-quick-start) · [Tools](#-available-tools) · [Configuration](#-configuration) · [How It Works](#-how-it-works) · [FAQ](#-faq)

</div>

---

## 🤔 The Problem

You ask your AI assistant to analyze an Excel document. It reads the text just fine — but **completely misses the diagrams, screenshots, and charts** embedded in the file. That's because every existing Excel MCP server ignores images.

**Excel MCP Server fixes this.** It extracts embedded images, optimizes them, and returns them as native `ImageContent` that vision-capable AI models can see and analyze — alongside all the text data.

## ✨ Key Features

| Feature | Description |
|---------|-------------|
| 🖼️ **Image Extraction** | Extracts all embedded images with cell-position mapping |
| 📄 **Full Content Reading** | Text + images in a single call — nothing is missed |
| 📊 **Smart Pagination** | Handles massive spreadsheets without blowing up context |
| 🔍 **Text Search** | Find content across all sheets instantly |
| 🔒 **100% Local** | Your files never leave your machine |
| ⚡ **Fast** | 16MB file with 40 images processed in ~4 seconds |
| 🖥️ **Cross-Platform** | macOS, Linux, Windows |

### Image Extraction — What Makes This Different

Most Excel MCP servers only read cell values. This server uses a **dual extraction strategy**:

1. **Cell-Position Mapping** (primary) — Maps each image to its exact cell location using `openpyxl-image-loader`
2. **Archive Scanning** (fallback) — Scans the xlsx ZIP archive's `xl/media/` directory to catch any images missed by method 1

The result: **zero images left behind**, with position metadata when available.

---

## 🚀 Quick Start

### Install via `uvx` (Recommended)

No installation needed — runs directly:

```bash
uvx excel-vision-mcp
```

### Install via `pip`

```bash
pip install excel-vision-mcp
```

Then run:

```bash
excel-vision-mcp
```

### Install from source

```bash
git clone https://github.com/VOYAGER-Inc/excel-mcp-server.git
cd excel-mcp-server
uv sync
uv run excel-vision-mcp
```

---

## 🔧 Configuration

Add the server to your MCP client's configuration file.

### Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows):

```json
{
  "mcpServers": {
    "excel-reader": {
      "command": "uvx",
      "args": ["excel-vision-mcp"]
    }
  }
}
```

### Cursor

Edit `.cursor/mcp.json` in your project root:

```json
{
  "mcpServers": {
    "excel-reader": {
      "command": "uvx",
      "args": ["excel-vision-mcp"]
    }
  }
}
```

### Windsurf / VS Code (Copilot)

Edit your MCP settings file:

```json
{
  "mcpServers": {
    "excel-reader": {
      "command": "uvx",
      "args": ["excel-vision-mcp"]
    }
  }
}
```

### Antigravity IDE

Edit `~/.gemini/config/mcp_config.json`:

```json
{
  "mcpServers": {
    "excel-reader": {
      "command": "uvx",
      "args": ["excel-vision-mcp"]
    }
  }
}
```

> **Note:** After editing the config, restart your IDE/client to load the new server.

---

## 🛠️ Available Tools

### `list_sheets`

List all sheets with dimensions, merged cell counts, and image totals. Use this first to understand a workbook's structure.

```
list_sheets(file_path="/path/to/file.xlsx")
```

**Returns:** Sheet names, row×column dimensions, data ranges, merged cell counts, total image count.

---

### `read_excel_data`

Read cell data from a specific sheet with pagination support.

```
read_excel_data(
    file_path="/path/to/file.xlsx",
    sheet_name="Sheet1",      # optional, defaults to first sheet
    start_row=1,              # optional, 1-indexed
    max_rows=200              # optional, default 200
)
```

**Returns:** Cell values organized by row with coordinate labels and merged cell indicators.

---

### `extract_images`

Extract all embedded images from the workbook as base64 `ImageContent`.

```
extract_images(
    file_path="/path/to/file.xlsx",
    sheet_name="Overview",    # optional, None = all sheets
    max_width=1024,           # optional, resize limit
    max_height=1024           # optional, resize limit
)
```

**Returns:** List of `ImageContent` (base64) with metadata — cell position, sheet name, original dimensions.

---

### `read_full_content` ⭐

**The star tool.** Reads ALL text data AND all embedded images in a single call. Ideal for comprehensive document analysis.

```
read_full_content(
    file_path="/path/to/file.xlsx",
    max_rows_per_sheet=500,   # optional
    max_image_width=1024,     # optional
    max_image_height=1024     # optional
)
```

**Returns:** Complete workbook contents — every sheet's data as structured text, followed by every embedded image with cell-position mapping.

**Example use case:** _"Analyze this requirements document and summarize all use cases, including the workflow diagrams."_

---

### `get_workbook_overview`

Quick structural summary of a workbook — file size, sheet list, dimensions, image count.

```
get_workbook_overview(file_path="/path/to/file.xlsx")
```

---

### `search_excel`

Case-insensitive text search across all cells in the workbook.

```
search_excel(
    file_path="/path/to/file.xlsx",
    query="revenue",
    sheet_name="Q4 Report"    # optional, None = all sheets
)
```

**Returns:** Matching cells with sheet name, coordinate, and value. Limited to 100 results.

---

## ⚙️ How It Works

### Architecture

```
Your AI Client (Claude, Cursor, etc.)
       │
       │ stdio (JSON-RPC)
       ▼
┌─────────────────────────────┐
│     Excel MCP Server        │
│                             │
│  ┌───────────────────────┐  │
│  │   openpyxl            │  │──→ Cell data, formulas, merged cells
│  │   (Excel parser)      │  │
│  └───────────────────────┘  │
│                             │
│  ┌───────────────────────┐  │
│  │ openpyxl-image-loader │  │──→ Images with cell positions
│  │ + zipfile (fallback)  │  │
│  └───────────────────────┘  │
│                             │
│  ┌───────────────────────┐  │
│  │   Pillow              │  │──→ Resize, optimize, base64 encode
│  │   (image processing)  │  │
│  └───────────────────────┘  │
└─────────────────────────────┘
       │
       │ TextContent + ImageContent
       ▼
  AI sees text AND images
```

### Data Flow & Privacy

1. **Your file stays on your machine.** The server runs locally via `stdio` — no network requests, no uploads, no cloud.
2. **Nothing is written to disk.** All image processing happens in-memory (`BytesIO` buffers). The original `.xlsx` file is never modified.
3. **Memory is freed automatically.** After each request, Python's garbage collector reclaims all buffers.

### Image Processing Pipeline

```
Original image in .xlsx (e.g., 2048×1536px PNG)
  ↓ Extract from ZIP archive / drawing layer
  ↓ Resize to fit max dimensions (default 1024px)
  ↓ Compress (JPEG 80% / PNG optimized)
  ↓ Base64 encode
  → ImageContent returned to AI client (~100-300KB per image)
```

---

## 📋 Supported Formats

| Format | Status | Notes |
|--------|--------|-------|
| `.xlsx` | ✅ Fully supported | Excel 2007+ Open XML |
| `.xlsm` | ✅ Fully supported | Macro-enabled workbooks |
| `.xls` | ❌ Not supported | Legacy Excel 97-2003 format |
| `.csv` | ❌ Not supported | Use a CSV-specific tool |

### Image Types

| Image Type | Cell-Mapped | Archive Extraction |
|------------|:-----------:|:------------------:|
| PNG | ✅ | ✅ |
| JPEG | ✅ | ✅ |
| GIF | ✅ | ✅ |
| BMP | ✅ | ✅ |
| TIFF | ⚠️ Partial | ✅ |
| EMF/WMF | ❌ | ✅ |
| `=IMAGE()` formula | ❌ | ❌ |
| Images in comments | ❌ | ❌ |

---

## 📊 Performance

Tested on real-world enterprise Excel files (macOS, Apple Silicon):

| File | Size | Sheets | Images Extracted | Time |
|------|------|--------|:----------------:|-----:|
| Requirements Doc A | 4.5 MB | 12 | 24 | 2.4s |
| Requirements Doc B | 5.0 MB | 6 | 18 | 2.4s |
| Requirements Doc C | 10.7 MB | 6 | 13 | 1.5s |
| Master Spec | 16.0 MB | 12 | 40 | 4.4s |

---

## ❓ FAQ

<details>
<summary><b>Why can't it extract images from .xls files?</b></summary>

`.xls` is the legacy binary format (Excel 97-2003). It uses a completely different internal structure (BIFF) compared to `.xlsx` (ZIP-based Open XML). The libraries used (`openpyxl`, `openpyxl-image-loader`) only support the modern Open XML format. If you have `.xls` files, convert them to `.xlsx` using Excel or LibreOffice first.
</details>

<details>
<summary><b>Why are some images marked as "orphan"?</b></summary>

The primary extraction method (`openpyxl-image-loader`) maps images to specific cells but may miss images that aren't anchored to the standard drawing layer. The fallback archive scanner catches these "orphan" images from the `xl/media/` directory — you get every image, just without cell-position metadata for orphans.
</details>

<details>
<summary><b>Can I use this with models that don't support vision?</b></summary>

Yes! Text data extraction works perfectly with any model. Image extraction will still return `ImageContent`, but text-only models will simply ignore the image data. You won't get errors.
</details>

<details>
<summary><b>Is my data safe?</b></summary>

Yes. The server runs **entirely on your local machine** via `stdio` transport. No data is sent over the network, no files are uploaded anywhere, and no temporary files are created on disk. Your Excel files are read in-place and never modified.
</details>

<details>
<summary><b>How do I handle very large files (100MB+)?</b></summary>

The server uses `read_only` mode for data iteration and processes images in-memory one at a time. For extremely large files, use `read_excel_data` with pagination (`start_row` + `max_rows`) instead of `read_full_content` to control memory usage.
</details>

---

## 🗺️ Roadmap

- [ ] **Write support** — Create and update cells, insert images
- [ ] **Chart extraction** — Render charts as images
- [ ] **Formula evaluation** — Show calculated values alongside formulas
- [ ] **Conditional formatting** — Extract formatting rules
- [ ] **CSV/TSV support** — Extend to other tabular formats

---

## 🤝 Contributing

Contributions are welcome! Please open an issue first to discuss what you'd like to change.

```bash
git clone https://github.com/VOYAGER-Inc/excel-mcp-server.git
cd excel-mcp-server
uv sync
uv run pytest  # Run the test suite
```

## 📄 License

[MIT](LICENSE) — use it however you want.

---

<div align="center">

**Built for AI agents that need to see the whole picture, not just the text.**

⭐ Star this repo if it helped you!

</div>
