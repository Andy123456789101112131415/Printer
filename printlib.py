# -*- coding: utf-8 -*-
"""打印与格式转换后端。

- 所有 Office 文件（Word/Excel/PowerPoint）的转换与打印依赖 Windows + 已安装的
  Microsoft Office + pywin32。
- PDF 打印优先使用 SumatraPDF（把 SumatraPDF.exe 放到 tools/ 目录即可自动识别），
  其次 Adobe Reader，最后回退到 Edge / 系统打印。
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
        raise RuntimeError("缺少 pywin32，请先执行: pip install pywin32") from exc
    import win32com.client
    return win32com.client.Dispatch(app_name)


@contextlib.contextmanager
def _com():
    """在调用线程初始化 COM。

    Flask/waitress 是多线程的，每个线程都必须单独 CoInitialize，
    否则会随机报 "尚未调用 CoInitialize"。
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


# ---------------------------------------------------------------- 打印 ----

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
        return True, "Word 打印成功（A4）"
    except Exception as exc:
        return False, f"Word 打印失败: {exc}"
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
        return True, "Excel 打印成功（A4）"
    except Exception as exc:
        return False, f"Excel 打印失败: {exc}"
    finally:
        try:
            xl.Quit()
        except Exception:
            pass


def _print_powerpoint(path, printer=None):
    """PowerPoint 无法直接设置纸张，先导出 PDF 再按 A4 打印。"""
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
            return False, "PowerPoint 转 PDF 失败"
        return _print_pdf(pdf, printer)
    except Exception as exc:
        return False, f"PowerPoint 打印失败: {exc}"
    finally:
        try:
            ppt.Quit()
        except Exception:
            pass
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _shell_print(path, printer=None):
    """通过 Windows ShellExecute 的打印动词打印（图片/兜底）。"""
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
            return False, f"系统打印失败 (错误码 {ret})"
        return True, "已发送到系统打印"
    except Exception as exc:
        return False, f"系统打印失败: {exc}"


def _print_pdf_edge(path):
    edge = _find_edge()
    if not edge:
        return False, "未找到 Edge，且没有 SumatraPDF / Adobe Reader"
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
    return True, "Edge 静默打印"


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
            return True, "已按 A4 打印"
        except Exception as exc:
            return False, f"SumatraPDF 调用失败: {exc}"

    adobe = _find_adobe()
    if adobe:
        args = [adobe, "/t", str(path)]
        if printer:
            args.append(printer)
        try:
            subprocess.Popen(args, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL,
                             creationflags=subprocess.CREATE_NO_WINDOW)
            return True, "Adobe Reader 打印成功（纸张取决于打印机默认 A4）"
        except Exception as exc:
            return False, f"Adobe 调用失败: {exc}"

    return _print_pdf_edge(path)


def printer_exists(name):
    """检查打印机当前是否在系统中可用（网络打印机未连线时会返回 False）。"""
    if not name:
        return False
    try:
        import win32print
        printers = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        return name in printers
    except Exception:
        # 无法枚举时不做阻断，交由实际打印流程报错
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
    return False, f"不支持的文件类型: {ext or '(无扩展名)'}"


def print_file(path, printer=None):
    """打印单个文件（A4）。返回 (成功与否, 描述)。"""
    with _com():
        return _print_file_inner(path, printer)


def _safe_job_name(name):
    """打印任务名去掉 Windows 文件名非法字符。"""
    cleaned = re.sub(r'[<>:"/\\|?*\r\n\t]', "", str(name)).strip()
    return cleaned or "print"


def print_file_as(src, job_name, printer=None):
    """以指定名字（用于打印机端识别）打印文件，统一 A4。

    通过复制成同名临时副本实现：Office/PDF/图片进队列时任务名即为该名字。
    返回 (成功与否, 描述)。
    """
    src = Path(src)
    ext = src.suffix.lower()
    tmp_dir = tempfile.mkdtemp(prefix="printjob_")
    try:
        tmp = Path(tmp_dir) / (_safe_job_name(job_name) + ext)
        shutil.copy2(src, tmp)
        return print_file(tmp, printer)
    except Exception as exc:
        return False, f"打印失败: {exc}"
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ------------------------------------------------------------ 转 PDF ----

def convert_to_pdf(src: Path, out_dir: Path):
    """把文件转成 PDF 用于在线预览。

    返回 PDF 的 Path；已是 PDF / 图片时直接返回原路径（图片由浏览器直接显示）；
    转换失败返回 None。
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
