"""
Image processing utilities for Excel MCP Server.

Handles resizing, format conversion, and base64 encoding of extracted images
to optimize context window usage when returning images to AI agents.
"""

import base64
import io
import logging
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image as PILImage

logger = logging.getLogger(__name__)

# Aggressive defaults to keep context window small
DEFAULT_MAX_WIDTH = 1024
DEFAULT_MAX_HEIGHT = 1024
DEFAULT_JPEG_QUALITY = 80


@dataclass
class ImageInfo:
    """
    Metadata and processed data for an extracted Excel image.

    Bundles the base64-encoded image with positional and dimensional metadata
    so the AI agent can correlate images with their Excel cell locations.
    """

    data_base64: str
    mime_type: str
    original_filename: str = ""
    cell_reference: str = ""
    sheet_name: str = ""
    original_width: int = 0
    original_height: int = 0
    processed_width: int = 0
    processed_height: int = 0
    file_size_bytes: int = 0
    index: int = 0


def optimize_image(
    image: PILImage.Image,
    max_width: int = DEFAULT_MAX_WIDTH,
    max_height: int = DEFAULT_MAX_HEIGHT,
    quality: int = DEFAULT_JPEG_QUALITY,
    force_format: str | None = None,
) -> tuple[bytes, str]:
    """
    Resize and compress an image for efficient transmission via MCP.

    Applies proportional scaling to fit within max dimensions, converts RGBA to RGB
    for JPEG compatibility, and returns compressed bytes with the appropriate MIME type.

    @param image: PIL Image object to process.
    @param max_width: Maximum width in pixels after resizing.
    @param max_height: Maximum height in pixels after resizing.
    @param quality: JPEG compression quality (1-100).
    @param force_format: Force output format ('PNG', 'JPEG'). Auto-detects if None.
    @return: Tuple of (compressed image bytes, MIME type string).
    """
    original_size = image.size

    # Proportional resize if exceeds max dimensions
    ratio = min(max_width / image.width, max_height / image.height)
    if ratio < 1.0:
        new_size = (int(image.width * ratio), int(image.height * ratio))
        image = image.resize(new_size, PILImage.Resampling.LANCZOS)
        logger.debug(
            "Resized image from %s to %s", original_size, new_size
        )

    # Determine output format
    if force_format:
        fmt = force_format.upper()
    elif image.mode == "RGBA" or (hasattr(image, "info") and image.info.get("transparency")):
        fmt = "PNG"
    else:
        fmt = "JPEG"

    # RGBA -> RGB conversion for JPEG
    if fmt == "JPEG" and image.mode in ("RGBA", "P", "LA"):
        background = PILImage.new("RGB", image.size, (255, 255, 255))
        if image.mode == "P":
            image = image.convert("RGBA")
        background.paste(image, mask=image.split()[-1] if "A" in image.mode else None)
        image = background

    # Compress
    buffer = io.BytesIO()
    save_kwargs = {"format": fmt}
    if fmt == "JPEG":
        save_kwargs["quality"] = quality
        save_kwargs["optimize"] = True
    elif fmt == "PNG":
        save_kwargs["optimize"] = True

    image.save(buffer, **save_kwargs)
    image_bytes = buffer.getvalue()

    mime_type = f"image/{fmt.lower()}"
    return image_bytes, mime_type


def image_to_base64(
    image: PILImage.Image,
    max_width: int = DEFAULT_MAX_WIDTH,
    max_height: int = DEFAULT_MAX_HEIGHT,
    quality: int = DEFAULT_JPEG_QUALITY,
) -> tuple[str, str, int, int]:
    """
    Convert a PIL Image to optimized base64 string.

    @param image: PIL Image object.
    @param max_width: Maximum width for resizing.
    @param max_height: Maximum height for resizing.
    @param quality: JPEG compression quality.
    @return: Tuple of (base64 string, MIME type, processed width, processed height).
    """
    image_bytes, mime_type = optimize_image(image, max_width, max_height, quality)

    # Reload to get processed dimensions
    processed = PILImage.open(io.BytesIO(image_bytes))
    processed_w, processed_h = processed.size

    encoded = base64.standard_b64encode(image_bytes).decode("utf-8")
    return encoded, mime_type, processed_w, processed_h


def bytes_to_base64(
    raw_bytes: bytes,
    max_width: int = DEFAULT_MAX_WIDTH,
    max_height: int = DEFAULT_MAX_HEIGHT,
    quality: int = DEFAULT_JPEG_QUALITY,
) -> tuple[str, str, int, int, int, int]:
    """
    Convert raw image bytes to optimized base64 string.

    @param raw_bytes: Raw image file bytes.
    @param max_width: Maximum width for resizing.
    @param max_height: Maximum height for resizing.
    @param quality: JPEG compression quality.
    @return: Tuple of (base64 string, MIME type, original width, original height,
             processed width, processed height).
    """
    image = PILImage.open(io.BytesIO(raw_bytes))
    original_w, original_h = image.size

    encoded, mime_type, processed_w, processed_h = image_to_base64(
        image, max_width, max_height, quality
    )
    return encoded, mime_type, original_w, original_h, processed_w, processed_h


def format_file_size(size_bytes: int) -> str:
    """
    Format byte count to human-readable string.

    @param size_bytes: Number of bytes.
    @return: Formatted string like '1.5 MB'.
    """
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"
