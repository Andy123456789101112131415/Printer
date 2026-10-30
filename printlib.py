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
import socket
import subprocess
import tempfile
import threading
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

# GDI+ printing keeps output independent of file associations and lets us force
# A4, fit each page to the printable area, and set the queue job name.
IMAGE_PRINT_PS = r"""
param(
  [Parameter(Mandatory=$true)][string]$List,
  [string]$Printer = "",
  [string]$JobName = ""
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$paths = @([System.IO.File]::ReadAllLines($List) | Where-Object { $_.Trim() -ne "" })
if ($paths.Count -eq 0) {
  [Console]::Error.WriteLine("No pages to print.")
  exit 3
}
$doc = New-Object System.Drawing.Printing.PrintDocument
if ($JobName) { $doc.DocumentName = $JobName }
if ($Printer) {
  $doc.PrinterSettings.PrinterName = $Printer
  if (-not $doc.PrinterSettings.IsValid) {
    [Console]::Error.WriteLine("Printer not found: " + $Printer)
    exit 2
  }
}
$a4 = $null
foreach ($ps in $doc.PrinterSettings.PaperSizes) {
  if ($ps.Kind -eq [System.Drawing.Printing.PaperKind]::A4) { $a4 = $ps; break }
}
if ($a4) { $doc.DefaultPageSettings.PaperSize = $a4 }
$doc.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(40, 40, 40, 40)
$index = 0
$state = @{ i = 0; total = $paths.Count }
$doc.add_PrintPage({
  param($sender, $e)
  $i = $state.i
  if ($i -ge $state.total) {
    $e.HasMorePages = $false
    return
  }
  $img = [System.Drawing.Image]::FromFile($paths[$i])
  try {
    $e.PageSettings.Landscape = ($img.Width -gt $img.Height)
    $b = $e.MarginBounds
    $scale = [Math]::Min($b.Width / $img.Width, $b.Height / $img.Height)
    $w = [int]($img.Width * $scale)
    $h = [int]($img.Height * $scale)
    $x = $b.X + [int](($b.Width - $w) / 2)
    $y = $b.Y + [int](($b.Height - $h) / 2)
    $e.Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $e.Graphics.DrawImage($img, $x, $y, $w, $h)
  } finally {
    $img.Dispose()
  }
  $state.i = $i + 1
  $e.HasMorePages = (($i + 1) -lt $state.total)
})
$doc.Print()
$doc.Dispose()
Write-Output ("OK " + $state.i + "/" + $state.total)
"""

PRINT_TIMEOUT = 150.0
OFFICE_TIMEOUT = 150.0
PRINT_PORTS = (445, 515, 631, 9100)


def printer_reachable(name, timeout=1.5):
    """Check that a network print server answers before submitting a job.

    Without this, Office printing can block for minutes when the computer lost
    the school network and the job never starts.
    """
    if not name or not name.startswith("\\\\"):
        return True
    host = name[2:].split("\\", 1)[0]
    if not host:
        return True
    for port in PRINT_PORTS:
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except OSError:
            continue
    return False


def _call_with_timeout(func, timeout, label):
    """Run a print backend in its own COM-initialized thread with a deadline."""
    box = {}

    def runner():
        try:
            with _com():
                box["value"] = func()
        except Exception as exc:
            box["error"] = exc

    thread = threading.Thread(target=runner, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        return False, (f"Printing {label} did not finish within {int(timeout)} seconds. "
                       "The printer may be unreachable or the print queue is stuck.")
    if "error" in box:
        return False, f"Printing failed: {box['error']}"
    return box.get("value") or (False, "Printing failed.")



def _error_detail(text):
    """Pick an English error line from tool output, skipping benign repair notes."""
    skip = ("startxref", "repair", "LoadSettings", "ParseTranslations",
            "Untranslated strings", "Starting SumatraPDF", "GetCommandLineW")
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or not line.isascii() or any(s in line for s in skip):
            continue
        low = line.lower()
        if any(word in low for word in ("error", "failed", "cannot", "could not",
                                        "not found", "denied", "invalid")):
            return line
    return ""


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
    """Print slides by exporting each one to an image and drawing it on A4.

    PowerPoint cannot set a paper size through automation, so printing slides
    directly would use the printer default. Exported images keep the layout and
    let us force A4 with a readable queue job name.
    """
    ppt = _dispatch("PowerPoint.Application")
    work_dir = Path(tempfile.mkdtemp(prefix="psppt_"))
    try:
        pres = ppt.Presentations.Open(str(path), WithWindow=False)
        try:
            count = pres.Slides.Count
            if count < 1:
                return False, "The presentation has no slides."
            width = int(pres.PageSetup.SlideWidth)
            height = int(pres.PageSetup.SlideHeight)
            export_w = 1600
            export_h = max(1, int(round(1600 * height / width))) if width else 1200
            pages = []
            for index in range(1, count + 1):
                out = work_dir / f"slide{index:04d}.png"
                pres.Slides(index).Export(str(out), "PNG", export_w, export_h)
                pages.append(out)
            if len(pages) != count or not all(p.is_file() for p in pages):
                return False, "PowerPoint slides could not be exported."
        finally:
            pres.Close()
    except Exception as exc:
        return False, f"PowerPoint printing failed: {exc}"
    finally:
        try:
            ppt.Quit()
        except Exception:
            pass

    try:
        return _print_image_pages(pages, Path(path).stem, printer)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _print_image_pages(pages, job_name, printer=None):
    """Print one image per A4 page through GDI+, in a single print job."""
    pages = [Path(p) for p in pages]
    if not pages:
        return False, "Nothing to print."
    work_dir = Path(tempfile.mkdtemp(prefix="psimg_"))
    try:
        script = work_dir / "print_pages.ps1"
        script.write_text(IMAGE_PRINT_PS, encoding="utf-8-sig")
        listing = work_dir / "pages.txt"
        listing.write_text("\n".join(str(p) for p in pages), encoding="utf-8")
        args = [_powershell(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                str(script), "-List", str(listing),
                "-JobName", _safe_job_name(job_name or pages[0].stem)]
        if printer:
            args += ["-Printer", printer]
        try:
            proc = subprocess.run(args, capture_output=True, text=True, encoding="utf-8",
                                  errors="replace", timeout=PRINT_TIMEOUT,
                                  creationflags=subprocess.CREATE_NO_WINDOW)
        except FileNotFoundError:
            if len(pages) == 1:
                return _shell_print(pages[0], printer)
            return False, "PowerShell was not found, so the pages could not be printed."
        except subprocess.TimeoutExpired:
            return False, (f"Printing did not finish within {int(PRINT_TIMEOUT)} seconds. "
                           "The printer may be unreachable.")
        except Exception as exc:
            return False, f"Printing failed: {exc}"
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    if proc.returncode != 0:
        detail = _error_detail(proc.stderr) or _error_detail(proc.stdout)
        return False, f"Printing failed. {detail}" if detail else "Printing failed."
    return True, "Printed on A4 paper."


def _print_image(path, printer=None):
    """Print one image on A4 with GDI+, so no print association is required."""
    return _print_image_pages([path], Path(path).stem, printer)


def _render_pdf_pages(path, work_dir, dpi=150, max_pages=200):
    """Render PDF pages to PNG so they can be printed through GDI+.

    Returns None when PyMuPDF is unavailable, the PDF needs a password, or the
    document is too long to render.
    """
    try:
        import pymupdf as fitz
    except ImportError:
        return None
    try:
        doc = fitz.open(str(path))
    except Exception:
        return None
    try:
        if doc.needs_pass or doc.page_count == 0 or doc.page_count > max_pages:
            return None
        matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pages = []
        for index, page in enumerate(doc):
            out = Path(work_dir) / f"page{index + 1:04d}.png"
            page.get_pixmap(matrix=matrix).save(str(out))
            pages.append(out)
        return pages or None
    except Exception:
        return None
    finally:
        doc.close()



def _powershell():
    return shutil.which("powershell") or shutil.which("pwsh") or "powershell"


def _shell_print(path, printer=None):
    """Fallback printing through the Windows ShellExecute print verb."""
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


def _print_pdf_edge(path, printer=None):
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
    work_dir = Path(tempfile.mkdtemp(prefix="pspdf_"))
    try:
        pages = _render_pdf_pages(path, work_dir)
        if pages:
            return _print_image_pages(pages, Path(path).stem, printer)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    return _print_pdf_viewer(path, printer)


def _print_pdf_viewer(path, printer=None):
    """Fallback for PDFs that cannot be rendered locally."""
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
            # Wait for SumatraPDF to finish so the temporary copy is not removed
            # while it is still opening the file.
            proc = subprocess.run(args, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace", timeout=180,
                                  creationflags=subprocess.CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired:
            return False, "SumatraPDF did not finish printing within 180 seconds."
        except Exception as exc:
            return False, f"SumatraPDF failed: {exc}"
        if proc.returncode != 0:
            detail = _error_detail(proc.stdout) or _error_detail(proc.stderr)
            message = f"PDF printing failed (SumatraPDF exit code {proc.returncode})."
            return False, f"{message} {detail}" if detail else (
                f"{message} Check that the printer is online and available.")
        return True, "Printed on A4 paper."

    adobe = _find_adobe()
    if adobe:
        args = [adobe, "/t", str(path)]
        if printer:
            args.append(printer)
        try:
            proc = subprocess.Popen(args, stdout=subprocess.DEVNULL,
                                    stderr=subprocess.DEVNULL,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            proc.wait(timeout=20)
        except subprocess.TimeoutExpired:
            return True, "Adobe Reader printed successfully. Paper size depends on the printer's A4 default."
        except Exception as exc:
            return False, f"Adobe Reader failed: {exc}"
        if proc.returncode not in (0, None):
            return False, f"Adobe Reader failed (exit code {proc.returncode})."
        return True, "Adobe Reader printed successfully. Paper size depends on the printer's A4 default."

    return _print_pdf_edge(path, printer)


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
        return _print_image(path, printer)
    return False, f"Unsupported file type: {ext or '(no extension)'}"


def _target_printer(printer=None):
    """Return the printer that will actually receive the job."""
    if printer:
        return printer
    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def print_file(path, printer=None):
    """Print one file on A4 paper. Returns (success, message)."""
    path = Path(path)
    target = _target_printer(printer)
    if target and not printer_reachable(target):
        return False, (f"The print server for {target} does not answer. Check that this "
                       "computer is on the school network and that the printer is on.")
    ext = ext_of(path)
    timeout = OFFICE_TIMEOUT if ext in (OFFICE_DOC_EXTS | SHEET_EXTS | SLIDE_EXTS) else PRINT_TIMEOUT
    return _call_with_timeout(lambda: _print_file_inner(path, printer), timeout, path.name)


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
    tmp_dir = tempfile.mkdtemp(prefix="ps_")
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
