"""Screenshot generation for Excel sheets.

Uses LibreOffice in headless mode for cross-platform support.
Exports all sheets by converting to PDF (multi-page), then to PNG per page.
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl
from PIL import Image

logger = logging.getLogger(__name__)


class ScreenshotError(RuntimeError):
    """Raised when screenshot generation fails and is required."""


def is_screenshot_available() -> bool:
    """Check if screenshot generation is available (LibreOffice installed)."""
    return _find_libreoffice() is not None


def generate_screenshots(
    excel_path: str | Path,
    required: bool = True,
) -> dict[str, bytes]:
    """Generate PNG screenshots for each sheet in the Excel file.

    Args:
        excel_path: Path to the .xlsx file.
        required: If True, raise ScreenshotError when LibreOffice is missing.

    Returns:
        Dict mapping sheet_name → PNG bytes.
    """
    excel_path = Path(excel_path)
    lo_path = _find_libreoffice()

    if lo_path is None:
        msg = (
            "LibreOffice is required for screenshot generation but was not found. "
            "Install it from https://www.libreoffice.org/download/ "
            "or via: winget install TheDocumentFoundation.LibreOffice"
        )
        if required:
            raise ScreenshotError(msg)
        logger.warning(msg)
        return {}

    # Get sheet names from the workbook
    try:
        wb = openpyxl.load_workbook(str(excel_path), read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception:
        sheet_names = []

    try:
        return _generate_via_pdf(excel_path, lo_path, sheet_names)
    except Exception as e:
        logger.warning(f"Screenshot generation failed: {e}")
        return {}


def _find_libreoffice() -> str | None:
    """Find LibreOffice executable on the system."""
    for name in ["libreoffice", "soffice", "soffice.exe"]:
        path = shutil.which(name)
        if path:
            return path

    common_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in common_paths:
        if Path(p).exists():
            return p

    return None


def _find_poppler() -> str | None:
    """Find poppler bin directory (needed by pdf2image on Windows)."""
    if shutil.which("pdftoppm"):
        return None  # Already in PATH, no need to specify

    common_paths = [
        r"C:\tools\poppler\poppler-24.08.0\Library\bin",
        r"C:\tools\poppler\Library\bin",
    ]
    # Also scan C:\tools\poppler for any version
    poppler_root = Path(r"C:\tools\poppler")
    if poppler_root.exists():
        for bin_dir in poppler_root.rglob("bin"):
            if (bin_dir / "pdftoppm.exe").exists():
                return str(bin_dir)

    for p in common_paths:
        if Path(p).exists() and (Path(p) / "pdftoppm.exe").exists():
            return p

    return None


def _generate_via_pdf(
    excel_path: Path,
    lo_path: str,
    sheet_names: list[str],
) -> dict[str, bytes]:
    """Convert Excel → PDF (all sheets) → PNG per page."""
    screenshots: dict[str, bytes] = {}

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Step 1: Convert Excel to PDF (all sheets become pages)
        cmd = [
            lo_path,
            "--headless",
            "--calc",
            "--convert-to", "pdf",
            "--outdir", str(tmpdir_path),
            str(excel_path),
        ]

        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=120,
        )

        pdf_files = sorted(tmpdir_path.glob("*.pdf"))
        if not pdf_files:
            logger.warning(f"LibreOffice PDF conversion produced no output: {result.stderr}")
            # Fallback to single PNG export
            return _generate_single_png(excel_path, lo_path, sheet_names)

        pdf_path = pdf_files[0]

        # Step 2: Convert PDF pages to PNGs using Pillow (reads PDF if pillow has PDF support)
        try:
            screenshots = _pdf_to_pngs(pdf_path, sheet_names)
        except Exception as e:
            logger.warning(f"PDF to PNG conversion failed ({e}), falling back to single PNG")
            return _generate_single_png(excel_path, lo_path, sheet_names)

    return screenshots


def _pdf_to_pngs(pdf_path: Path, sheet_names: list[str]) -> dict[str, bytes]:
    """Convert PDF pages to PNG images.

    Maps pages to sheet names (best-effort: 1 page per sheet assumption).
    """

    try:
        from pdf2image import convert_from_path
        poppler_path = _find_poppler()
        kwargs = {"dpi": 150}
        if poppler_path:
            kwargs["poppler_path"] = poppler_path
        images = convert_from_path(str(pdf_path), **kwargs)
        result: dict[str, bytes] = {}
        for i, img in enumerate(images):
            name = sheet_names[i] if i < len(sheet_names) else f"Page_{i + 1}"
            import io
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            result[name] = buf.getvalue()
        return result
    except ImportError:
        pass

    # Fallback: use Pillow directly (only works if Pillow has PDF plugin)
    try:
        import io
        result = {}
        img = Image.open(str(pdf_path))
        page = 0
        while True:
            try:
                img.seek(page)
                name = sheet_names[page] if page < len(sheet_names) else f"Sheet_{page + 1}"
                buf = io.BytesIO()
                img.convert("RGB").save(buf, format="PNG")
                result[name] = buf.getvalue()
                page += 1
            except EOFError:
                break
        if result:
            return result
    except Exception:
        pass

    # Last resort: read PDF as single image
    logger.info("Multi-page PDF extraction not available, using single-page fallback")
    return {}


def _generate_single_png(
    excel_path: Path,
    lo_path: str,
    sheet_names: list[str],
) -> dict[str, bytes]:
    """Fallback: export as single PNG (first sheet only)."""
    screenshots: dict[str, bytes] = {}

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        cmd = [
            lo_path, "--headless", "--calc",
            "--convert-to", "png",
            "--outdir", str(tmpdir_path),
            str(excel_path),
        ]
        subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        png_files = sorted(tmpdir_path.glob("*.png"))
        for i, png_file in enumerate(png_files):
            name = sheet_names[i] if i < len(sheet_names) else f"Sheet_{i + 1}"
            screenshots[name] = png_file.read_bytes()

    return screenshots
