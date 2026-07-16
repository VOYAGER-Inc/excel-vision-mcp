"""Tests for image optimization and formatting utilities."""

import base64
import io

from PIL import Image as PILImage

from excel_vision_mcp.utils import (
    bytes_to_base64,
    format_file_size,
    image_to_base64,
    optimize_image,
)


class TestFormatFileSize:
    def test_units(self):
        assert format_file_size(500) == "500.0 B"
        assert format_file_size(2048) == "2.0 KB"
        assert format_file_size(5 * 1024 * 1024) == "5.0 MB"
        assert format_file_size(3 * 1024**3) == "3.0 GB"


class TestOptimizeImage:
    def test_rgb_image_becomes_jpeg(self):
        img = PILImage.new("RGB", (100, 100), (200, 50, 50))
        data, mime = optimize_image(img)

        assert mime == "image/jpeg"
        assert PILImage.open(io.BytesIO(data)).format == "JPEG"

    def test_rgba_image_stays_png(self):
        img = PILImage.new("RGBA", (100, 100), (200, 50, 50, 128))
        data, mime = optimize_image(img)

        assert mime == "image/png"
        assert PILImage.open(io.BytesIO(data)).format == "PNG"

    def test_downscales_oversized_image(self):
        img = PILImage.new("RGB", (2000, 1000), (10, 20, 30))
        data, _ = optimize_image(img, max_width=500, max_height=500)

        result = PILImage.open(io.BytesIO(data))
        assert result.width <= 500
        assert result.height <= 500

    def test_small_image_not_upscaled(self):
        img = PILImage.new("RGB", (80, 60), (10, 20, 30))
        data, _ = optimize_image(img, max_width=1024, max_height=1024)

        result = PILImage.open(io.BytesIO(data))
        assert (result.width, result.height) == (80, 60)


class TestBase64Conversion:
    def test_image_to_base64_roundtrip(self):
        img = PILImage.new("RGB", (64, 32), (0, 100, 200))
        encoded, mime, width, height = image_to_base64(img)

        decoded = base64.standard_b64decode(encoded)
        reloaded = PILImage.open(io.BytesIO(decoded))
        assert (reloaded.width, reloaded.height) == (width, height) == (64, 32)
        assert mime == "image/jpeg"

    def test_bytes_to_base64_reports_both_dimensions(self):
        buffer = io.BytesIO()
        PILImage.new("RGB", (1200, 600), (5, 5, 5)).save(buffer, format="PNG")

        encoded, mime, orig_w, orig_h, proc_w, proc_h = bytes_to_base64(
            buffer.getvalue(), max_width=300, max_height=300
        )

        assert (orig_w, orig_h) == (1200, 600)
        assert proc_w <= 300 and proc_h <= 300
        assert base64.standard_b64decode(encoded)
