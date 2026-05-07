"""Recalculate Excel formulas via Excel COM automation or LibreOffice.

When Excel files are generated programmatically (e.g. by openpyxl), formula
cells may not have cached computed values.  openpyxl with ``data_only=True``
reads the cached value, so missing caches appear as ``None``.

This module re-opens the workbook in Excel (via COM automation on Windows)
or LibreOffice (headless), triggering a full recalculation and saving the
result — populating every formula cache.

Strategy: Excel COM first (full formula compatibility), LibreOffice fallback.
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)


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


def _excel_com_available() -> bool:
    """Check whether Excel COM automation is available (Windows + pywin32 + Excel installed)."""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401

        return True
    except ImportError:
        return False


def _sanitize_for_excel(src: Path) -> Path:
    """Re-save *src* via openpyxl to fix XML issues that block Excel COM.

    Some AI-generated xlsx files have structural problems (e.g. cells with
    ``t="n"`` but no ``<v>`` tag, inline strings, missing sharedStrings.xml)
    that openpyxl tolerates but Excel rejects.  Opening and re-saving through
    openpyxl normalises the XML.

    Additionally, formulas using newer Excel functions (SORT, UNIQUE,
    FILTER, XLOOKUP, etc.) need ``_xlfn.`` / ``_xlfn._xlws.`` prefixes
    in the raw XML for Excel to recognise them.  openpyxl does not add
    these prefixes, so we patch them here.

    Returns the path to the sanitized temp file (caller must clean up).
    """
    import re
    from zipfile import ZipFile

    # --- Step 1: re-save via openpyxl to normalise structure ---
    import openpyxl

    tmp = tempfile.NamedTemporaryFile(
        suffix=".xlsx",
        prefix="sanitized_",
        delete=False,
    )
    tmp.close()
    try:
        wb = openpyxl.load_workbook(src)
        wb.save(tmp.name)
    except Exception:
        Path(tmp.name).unlink(missing_ok=True)
        raise

    # --- Step 2: patch formulas in the raw XML to add _xlfn. prefixes ---
    # Functions that require the _xlfn. prefix (Excel 2013+ / 365 functions).
    _XLFN_FUNCTIONS = [
        "SORT",
        "SORTBY",
        "UNIQUE",
        "FILTER",
        "SEQUENCE",
        "RANDARRAY",
        "XLOOKUP",
        "XMATCH",
        "LET",
        "LAMBDA",
        "MAP",
        "REDUCE",
        "SCAN",
        "MAKEARRAY",
        "BYROW",
        "BYCOL",
        "ISOMITTED",
        "CHOOSECOLS",
        "CHOOSEROWS",
        "DROP",
        "TAKE",
        "EXPAND",
        "WRAPCOLS",
        "WRAPROWS",
        "TOCOL",
        "TOROW",
        "VSTACK",
        "HSTACK",
        "TEXTSPLIT",
        "TEXTBEFORE",
        "TEXTAFTER",
        "VALUETOTEXT",
        "ARRAYTOTEXT",
        "CONCAT",
        "IFS",
        "SWITCH",
        "MINIFS",
        "MAXIFS",
        "TEXTJOIN",
    ]

    # Build a regex that matches these function names when NOT already prefixed
    # with _xlfn. or _xlfn._xlws.  We look for them inside <f>...</f> tags.
    _fn_pattern = re.compile(
        r"(?<![\w.])(" + "|".join(_XLFN_FUNCTIONS) + r")\(",
        re.IGNORECASE,
    )

    try:
        patched = False
        tmp_path = Path(tmp.name)
        out_tmp = tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            prefix="patched_",
            delete=False,
        )
        out_tmp.close()

        with ZipFile(tmp_path, "r") as zin, ZipFile(out_tmp.name, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith(
                    "xl/worksheets/"
                ) and item.filename.endswith(".xml"):
                    xml_text = data.decode("utf-8")

                    def _add_prefix(m: re.Match) -> str:  # noqa: E301
                        return f"_xlfn.{m.group(1)}("

                    # Only patch inside <f>...</f> formula tags
                    def _patch_formulas(fm: re.Match) -> str:
                        original = fm.group(0)
                        replaced = _fn_pattern.sub(_add_prefix, original)
                        return replaced

                    new_xml = re.sub(
                        r"<f[^>]*>.*?</f>", _patch_formulas, xml_text, flags=re.DOTALL
                    )
                    if new_xml != xml_text:
                        patched = True
                    data = new_xml.encode("utf-8")
                zout.writestr(item, data)

        if patched:
            # Replace the original temp with the patched version
            Path(tmp.name).unlink()
            shutil.move(out_tmp.name, tmp.name)
            logger.debug("Patched _xlfn. prefixes in formulas for %s", src.name)
        else:
            Path(out_tmp.name).unlink(missing_ok=True)

    except Exception as exc:
        logger.debug("Formula prefix patching failed (non-fatal): %s", exc)
        # If patching fails, we still have the openpyxl-normalised file

    return Path(tmp.name)


def _recalc_via_excel_com(src: Path, dest: Path) -> bool:
    """Open *src* in Excel via COM, recalculate, save to *dest*.

    Returns True on success, False on failure.
    The original file is never modified — we work on a copy.
    """
    import pythoncom
    import win32com.client

    # COM requires absolute paths
    dest_abs = str(dest.resolve())

    # Try to open the file directly first; if Excel rejects it (corrupted
    # XML etc.), sanitize via openpyxl and retry once.
    sanitized: Path | None = None

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        src_abs = str(src.resolve())
        try:
            wb = excel.Workbooks.Open(src_abs, ReadOnly=False, UpdateLinks=0)
        except Exception:
            # Sanitize and retry
            logger.debug("Direct open failed, sanitizing %s via openpyxl", src.name)
            sanitized = _sanitize_for_excel(src)
            wb = excel.Workbooks.Open(
                str(sanitized.resolve()),
                ReadOnly=False,
                UpdateLinks=0,
            )

        # Re-stamp every formula so Excel re-parses it from scratch.
        # Files written by openpyxl may have formula cells with no cached
        # value; Excel sometimes honours the stale cache and skips recalc.
        #
        # We try FormulaArray first — formulas that use array operations
        # inside non-array-native functions (e.g. TEXTJOIN+IF, AGGREGATE
        # with array args) only evaluate correctly when entered as CSE
        # (Ctrl+Shift+Enter).  If FormulaArray fails (e.g. the formula is
        # too long for CSE or doesn't support it), we fall back to plain
        # Formula re-stamp.
        for ws in wb.Worksheets:
            used = ws.UsedRange
            if used is None:
                continue
            for cell in used:
                try:
                    if cell.HasFormula:
                        f = cell.Formula
                        try:
                            cell.FormulaArray = f  # CSE entry
                        except Exception:
                            cell.Formula = f  # plain re-stamp fallback
                except Exception:
                    pass

        # Force full recalculation after re-stamping
        excel.CalculateFullRebuild()

        # xlOpenXMLWorkbook = 51
        wb.SaveAs(dest_abs, FileFormat=51)

        logger.info(
            "Recalculated formulas via Excel COM: %s -> %s", src.name, dest.name
        )
        return True

    except Exception as exc:
        logger.warning("Excel COM recalc failed for %s: %s", src, exc)
        return False
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        if sanitized is not None:
            sanitized.unlink(missing_ok=True)
        pythoncom.CoUninitialize()


def _recalc_via_libreoffice(src: Path, dest: Path) -> bool:
    """Open *src* in LibreOffice headless, recalculate, save to *dest*.

    Returns True on success, False on failure.
    """
    lo_path = _find_libreoffice()
    if lo_path is None:
        logger.debug("LibreOffice not found — skipping formula recalculation")
        return False

    tmpdir = tempfile.mkdtemp(prefix="recalc_")
    tmp = Path(tmpdir)

    try:
        tmp_src = tmp / src.name
        shutil.copy2(src, tmp_src)

        cmd = [
            lo_path,
            "--headless",
            "--calc",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(tmp),
            str(tmp_src),
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
        )

        out_file = tmp / src.with_suffix(".xlsx").name
        if out_file.exists() and out_file.stat().st_size > 0:
            shutil.copy2(out_file, dest)
            logger.info(
                "Recalculated formulas via LibreOffice: %s -> %s",
                src.name,
                dest.name,
            )
            return True
        else:
            logger.warning(
                "LibreOffice recalc produced no output (stderr: %s)",
                (result.stderr or "").strip()[:200],
            )
            return False

    except subprocess.TimeoutExpired:
        logger.warning("LibreOffice recalc timed out for %s", src)
        return False
    except Exception as exc:
        logger.warning("LibreOffice recalc failed for %s: %s", src, exc)
        return False
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


RECALC_SUFFIX = ".recalc.xlsx"


def _recalc_sibling_path(src: Path) -> Path:
    """Return the sibling path used to store the recalculated copy.

    For ``foo.xlsx`` this is ``foo.recalc.xlsx`` next to the original.
    """
    return src.with_name(src.stem + RECALC_SUFFIX)


def recalculate_excel(excel_path: Path) -> Path:
    """Recalculate all formulas in *excel_path*.

    Tries Excel COM automation first (full formula compatibility on Windows),
    then falls back to LibreOffice headless.

    The original file is never modified.  A sibling file
    ``<stem>.recalc.xlsx`` is produced next to it and its path is returned.

    If a fresh recalc sibling already exists (mtime newer than the source),
    it is reused without re-running any engine.

    If all engines are unavailable or fail, returns the original
    *excel_path* unchanged (best-effort, so downstream code keeps working).
    """
    src = Path(excel_path)

    # Don't try to recalc an already-recalculated file
    if src.name.endswith(RECALC_SUFFIX):
        return src

    sibling = _recalc_sibling_path(src)

    # Reuse cached recalc if it's newer than the source
    try:
        if sibling.exists() and sibling.stat().st_mtime >= src.stat().st_mtime:
            logger.debug("Reusing existing recalc sibling: %s", sibling.name)
            return sibling
    except OSError:
        pass

    # Strategy: Excel COM first, LibreOffice fallback
    if _excel_com_available():
        if _recalc_via_excel_com(src, sibling):
            return sibling
        logger.info("Excel COM failed, falling back to LibreOffice for %s", src.name)

    if _recalc_via_libreoffice(src, sibling):
        return sibling

    logger.debug("No recalc engine succeeded for %s", src.name)
    return src
