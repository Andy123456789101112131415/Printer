# -*- coding: utf-8 -*-
"""Printing and file-conversion backend.

- Converting and printing Office files (Word/Excel/PowerPoint) requires Windows,
  Microsoft Office, and pywin32.
- PDF printing prefers SumatraPDF (place SumatraPDF.exe in the tools/ directory),
  then Adobe Reader, and finally Edge or the system print handler.
"""

import contextlib
import os
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

OFFICE_DOC_EXTS = {".docx", ".doc", ".txt"}
SHEET_EXTS = {".xlsx", ".xls", ".csv"}
SLIDE_EXTS = {".pptx", ".ppt"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
PDF_EXTS = {".pdf"}

BASE_DIR = Path(__file__).resolve().parent


def ext_of(path):
    return Path(path).suffix.lower()


def _dispatch(app_name):
    try:
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("pywin32 is missing. Install it with: pip install pywin32") from exc
    import win32com.client
    return win32com.client.Dispatch(app_name)


@contextlib.contextmanager
def _com():
    """Initialize COM on the calling thread.

    Flask/waitress uses multiple threads, and each thread must call CoInitialize
    separately to avoid intermittent COM initialization errors.
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        yield
        return
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _find_sumatra():
    candidates = [
        BASE_DIR / "tools" / "SumatraPDF.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "SumatraPDF" / "SumatraPDF.exe",
        Path(os.environ.get("ProgramFiles", "")) / "SumatraPDF" / "SumatraPDF.exe",
        Path(os.environ.get("ProgramFiles(x86)", "")) / "SumatraPDF" / "SumatraPDF.exe",
    ]
    for c in candidates:
        try:
            if c.is_file():
                return str(c)
        except OSError:
            continue
    return None


def _find_adobe():
    candidates = [
        r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Acrobat Reader\Reader\AcroRd32.exe",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def _find_edge():
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


# ---------------------------------------------------------------- Printing ----

WD_PAPER_A4 = 7     # Word WdPaperSize.wdPaperA4
XL_PAPER_A4 = 9     # Excel XlPaperSize.xlPaperA4


def _print_word(path, printer=None):
    word = _dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        if printer:
            word.ActivePrinter = printer
        doc = word.Documents.Open(str(path), ReadOnly=True, ConfirmConversions=False)
        try:
            doc.PageSetup.PaperSize = WD_PAPER_A4
            doc.PrintOut(Background=False)
        finally:
            doc.Close(False)
        return True, "Word printed successfully (A4)."
    except Exception as exc:
        return False, f"Word printing failed: {exc}"
    finally:
        try:
            word.Quit()
        except Exception:
            pass


def _print_excel(path, printer=None):
    xl = _dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path), ReadOnly=True)
        try:
            for ws in wb.Worksheets:
                try:
                    ws.PageSetup.PaperSize = XL_PAPER_A4
                except Exception:
                    pass
            if printer:
                wb.PrintOut(ActivePrinter=printer)
            else:
                wb.PrintOut()
        finally:
            wb.Close(False)
        return True, "Excel printed successfully (A4)."
    except Exception as exc:
        return False, f"Excel printing failed: {exc}"
    finally:
        try:
            xl.Quit()
        except Exception:
            pass


def _print_powerpoint(path, printer=None):
    """Export PowerPoint to PDF first, then print it on A4 paper."""
    ppt = _dispatch("PowerPoint.Application")
    tmp_dir = tempfile.mkdtemp(prefix="ppt2pdf_")
    try:
        pres = ppt.Presentations.Open(str(path), WithWindow=False)
        try:
            pdf = str(Path(tmp_dir) / (Path(path).stem + ".pdf"))
            pres.SaveAs(pdf, 32)  # ppSaveAsPDF
        finally:
            pres.Close()
        if not os.path.isfile(pdf):
            return False, "Failed to convert PowerPoint to PDF."
        return _print_pdf(pdf, printer)
    except Exception as exc:
        return False, f"PowerPoint printing failed: {exc}"
    finally:
        try:
            ppt.Quit()
        except Exception:
            pass
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _shell_print(path, printer=None):
    """Print images or use a fallback through the Windows ShellExecute print verb."""
    import ctypes
    try:
        SW_HIDE = 0
        if printer:
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "printto", str(path), f'"{printer}"', None, SW_HIDE)
        else:
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "print", str(path), None, None, SW_HIDE)
        if ret <= 32:
            return False, f"System printing failed (error code {ret})."
        return True, "Sent to the system print handler."
    except Exception as exc:
        return False, f"System printing failed: {exc}"


def _print_pdf_edge(path):
    edge = _find_edge()
    if not edge:
        return False, "Edge was not found, and neither SumatraPDF nor Adobe Reader is available."
    udd = tempfile.mkdtemp(prefix="edge_print_")
    proc = subprocess.Popen(
        [edge, "--kiosk", "--kiosk-printing", "--no-first-run",
         f"--user-data-dir={udd}", str(path)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(10)
    try:
        proc.terminate()
    except Exception:
        pass
    shutil.rmtree(udd, ignore_errors=True)
    return True, "Sent to Edge for silent printing."


def _print_pdf(path, printer=None):
    sumatra = _find_sumatra()
    if sumatra:
        args = [sumatra, "-silent"]
        if printer:
            args += ["-print-to", printer]
        else:
            args += ["-print-to-default"]
        args += ["-print-settings", "paper=A4,fit"]
        args.append(str(path))
        try:
            subprocess.Popen(args, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL,
                             creationflags=subprocess.CREATE_NO_WINDOW)
            return True, "Printed on A4 paper."
        except Exception as exc:
            return False, f"SumatraPDF failed: {exc}"

    adobe = _find_adobe()
    if adobe:
        args = [adobe, "/t", str(path)]
        if printer:
            args.append(printer)
        try:
            subprocess.Popen(args, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL,
                             creationflags=subprocess.CREATE_NO_WINDOW)
            return True, "Adobe Reader printed successfully. Paper size depends on the printer's A4 default."
        except Exception as exc:
            return False, f"Adobe Reader failed: {exc}"

    return _print_pdf_edge(path)


def printer_exists(name):
    """Check whether a printer is currently available in Windows."""
    if not name:
        return False
    try:
        import win32print
        printers = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        return name in printers
    except Exception:
        # If printers cannot be enumerated, let the actual print attempt report errors.
        return True


def _print_file_inner(path, printer=None):
    ext = ext_of(path)
    if ext in OFFICE_DOC_EXTS:
        return _print_word(path, printer)
    if ext in SHEET_EXTS:
        return _print_excel(path, printer)
    if ext in SLIDE_EXTS:
        return _print_powerpoint(path, printer)
    if ext in PDF_EXTS:
        return _print_pdf(path, printer)
    if ext in IMAGE_EXTS:
        return _shell_print(path, printer)
    return False, f"Unsupported file type: {ext or '(no extension)'}"


def print_file(path, printer=None):
    """Print one file on A4 paper. Returns (success, message)."""
    with _com():
        return _print_file_inner(path, printer)


def _safe_job_name(name):
    """Remove characters that are invalid in Windows file names from a print job name."""
    cleaned = re.sub(r'[<>:"/\\|?*\r\n\t]', "", str(name)).strip()
    return cleaned or "print"


def print_file_as(src, job_name, printer=None):
    """Print a file on A4 paper using a specified printer-side job name.

    A temporary copy gives Office, PDF, and image jobs the requested name in the queue.
    Returns (success, message).
    """
    src = Path(src)
    ext = src.suffix.lower()
    tmp_dir = tempfile.mkdtemp(prefix="printjob_")
    try:
        tmp = Path(tmp_dir) / (_safe_job_name(job_name) + ext)
        shutil.copy2(src, tmp)
        return print_file(tmp, printer)
    except Exception as exc:
        return False, f"Printing failed: {exc}"
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ------------------------------------------------------------ PDF Conversion ----

def convert_to_pdf(src: Path, out_dir: Path):
    """Convert a file to PDF for online preview.

    Returns the PDF path. Existing PDFs and images are returned as-is; images
    are displayed directly by the browser. Returns None if conversion fails.
    """
    with _com():
        return _convert_to_pdf_inner(src, out_dir)


def _convert_to_pdf_inner(src: Path, out_dir: Path):
    ext = ext_of(src)
    if ext in PDF_EXTS or ext in IMAGE_EXTS:
        return src

    out = out_dir / f"{Path(src).stem}.pdf"
    if out.is_file():
        return out

    try:
        if ext in OFFICE_DOC_EXTS:
            word = _dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(str(src), ReadOnly=True)
                try:
                    doc.SaveAs(str(out), FileFormat=17)  # wdFormatPDF
                finally:
                    doc.Close(False)
            finally:
                word.Quit()
        elif ext in SHEET_EXTS:
            xl = _dispatch("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            try:
                wb = xl.Workbooks.Open(str(src), ReadOnly=True)
                try:
                    wb.ExportAsFixedFormat(0, str(out))  # 0 = xlTypePDF
                finally:
                    wb.Close(False)
            finally:
                xl.Quit()
        elif ext in SLIDE_EXTS:
            ppt = _dispatch("PowerPoint.Application")
            try:
                pres = ppt.Presentations.Open(str(src), WithWindow=False)
                try:
                    pres.SaveAs(str(out), 32)  # ppSaveAsPDF
                finally:
                    pres.Close()
            finally:
                ppt.Quit()
        else:
            return None
    except Exception:
        return None

    return out if out.is_file() else None
