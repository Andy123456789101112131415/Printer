# -*- coding: utf-8 -*-
"""打印与格式转换后端。

- 所有 Office 文件（Word/Excel/PowerPoint）的转换与打印依赖 Windows + 已安装的
  Microsoft Office + pywin32。
- PDF 打印优先使用 SumatraPDF（把 SumatraPDF.exe 放到 tools/ 目录即可自动识别），
  其次 Adobe Reader，最后回退到 Edge / 系统打印。
"""

import os
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

def _print_word(path, printer=None):
    word = _dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        if printer:
            word.ActivePrinter = printer
        doc = word.Documents.Open(str(path), ReadOnly=True, ConfirmConversions=False)
        try:
            doc.PrintOut(Background=False)
        finally:
            doc.Close(False)
        return True, "Word 打印成功"
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
            if printer:
                wb.PrintOut(ActivePrinter=printer)
            else:
                wb.PrintOut()
        finally:
            wb.Close(False)
        return True, "Excel 打印成功"
    except Exception as exc:
        return False, f"Excel 打印失败: {exc}"
    finally:
        try:
            xl.Quit()
        except Exception:
            pass


def _print_powerpoint(path, printer=None):
    ppt = _dispatch("PowerPoint.Application")
    try:
        if printer:
            ppt.ActivePrinter = printer
        pres = ppt.Presentations.Open(str(path), WithWindow=False)
        try:
            pres.PrintOut()
        finally:
            pres.Close()
        return True, "PowerPoint 打印成功"
    except Exception as exc:
        return False, f"PowerPoint 打印失败: {exc}"
    finally:
        try:
            ppt.Quit()
        except Exception:
            pass


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
        args.append(str(path))
        try:
            subprocess.Popen(args, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL,
                             creationflags=subprocess.CREATE_NO_WINDOW)
            return True, "SumatraPDF 打印成功"
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
            return True, "Adobe Reader 打印成功"
        except Exception as exc:
            return False, f"Adobe 调用失败: {exc}"

    return _print_pdf_edge(path)


def print_file(path, printer=None):
    """打印单个文件。返回 (成功与否, 描述)。"""
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


def print_cover(text_lines, printer=None):
    """打印一页分隔封面（含中文）。

    通过 Word 打印一份带 UTF-8 BOM 的纯文本，避免编码对话框。
    返回 (成功与否, 描述)。
    """
    fd, tmp = tempfile.mkstemp(suffix=".txt")
    os.close(fd)
    try:
        with open(tmp, "w", encoding="utf-8-sig") as f:
            f.write("\n\n" + "\n".join(text_lines) + "\n")
        return _print_word(tmp, printer)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


# ------------------------------------------------------------ 转 PDF ----

def convert_to_pdf(src: Path, out_dir: Path):
    """把文件转成 PDF 用于在线预览。

    返回 PDF 的 Path；已是 PDF / 图片时直接返回原路径（图片由浏览器直接显示）；
    转换失败返回 None。
    """
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
