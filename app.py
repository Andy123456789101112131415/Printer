# -*- coding: utf-8 -*-
"""打印站 Flask 应用。

功能：
- 用户选择房间号、姓名并上传文件（支持多选）
- 管理员登录后按房间查看、下载、删除文件
- 按房间批量打印：打印视图（浏览器打印）或服务器直接打印
"""

import mimetypes
import os
import re
import secrets
import sqlite3
from datetime import datetime
from functools import wraps
from pathlib import Path

from flask import (Flask, abort, flash, jsonify, redirect, render_template,
                   request, send_file, session, url_for)

import printlib

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
PREVIEW_DIR = UPLOAD_DIR / ".preview"
DB_PATH = BASE_DIR / "data.db"
SECRET_KEY_FILE = BASE_DIR / ".secret_key"
ADMIN_PASSWORD_FILE = BASE_DIR / ".admin_password"

ALLOWED_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".txt", ".csv", ".md",
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp",
}

MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 单次请求最大 100MB

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


def _load_secret_key():
    key = os.environ.get("SECRET_KEY")
    if key:
        return key
    if SECRET_KEY_FILE.exists():
        key = SECRET_KEY_FILE.read_text(encoding="utf-8").strip()
        if key:
            return key
    key = secrets.token_hex(32)
    SECRET_KEY_FILE.write_text(key, encoding="utf-8")
    return key


app.config["SECRET_KEY"] = _load_secret_key()


def admin_password():
    pw = os.environ.get("ADMIN_PASSWORD")
    if pw:
        return pw
    if ADMIN_PASSWORD_FILE.exists():
        pw = ADMIN_PASSWORD_FILE.read_text(encoding="utf-8").strip()
        if pw:
            return pw
    return "admin123"


def ensure_dirs():
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    ensure_dirs()
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            room TEXT NOT NULL,
            name TEXT NOT NULL,
            original_name TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            size INTEGER NOT NULL,
            created_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS rooms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def get_setting(key, default=None):
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key, value):
    conn = get_db()
    conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                 (key, value))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------- 工具 ----

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login"))
        return view(*args, **kwargs)
    return wrapped


def csrf_token():
    token = session.get("_csrf")
    if not token:
        token = secrets.token_hex(16)
        session["_csrf"] = token
    return token


def check_csrf():
    token = session.get("_csrf")
    if not token or not secrets.compare_digest(str(token), request.form.get("_csrf", "")):
        abort(400, "CSRF 校验失败，请刷新页面重试")


def clean_room(raw):
    room = (raw or "").strip()
    room = re.sub(r'[<>:"/\\|?*\r\n\t]', "", room)
    return room[:50]


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_room_names():
    conn = get_db()
    rows = conn.execute("SELECT name FROM rooms ORDER BY name").fetchall()
    conn.close()
    return [r["name"] for r in rows]


def _fetch_file(file_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM files WHERE id = ?", (file_id,)).fetchone()
    conn.close()
    return row


def _serve(file_id, as_attachment):
    row = _fetch_file(file_id)
    if row is None:
        abort(404)
    path = UPLOAD_DIR / row["stored_name"]
    if not path.is_file():
        abort(404)
    mime = mimetypes.guess_type(row["stored_name"])[0] or "application/octet-stream"
    return send_file(path, mimetype=mime, as_attachment=as_attachment,
                     download_name=row["original_name"])


app.jinja_env.globals["csrf_token"] = csrf_token


# ------------------------------------------------------------ 上传 ----

@app.route("/")
def index():
    return redirect(url_for("upload"))


@app.route("/upload", methods=["GET", "POST"])
def upload():
    rooms = get_room_names()
    if request.method == "POST":
        room = clean_room(request.form.get("room", ""))
        name = (request.form.get("name") or "").strip()[:100]
        files = request.files.getlist("files")

        if not room:
            flash("请选择房间号", "error")
            return render_template("upload.html", rooms=rooms), 400
        if room not in rooms:
            flash("房间不存在，请选择列表中的房间", "error")
            return render_template("upload.html", rooms=rooms), 400
        if not name:
            flash("请填写姓名", "error")
            return render_template("upload.html", rooms=rooms), 400
        if not files or all(not f.filename for f in files):
            flash("请选择要上传的文件", "error")
            return render_template("upload.html", rooms=rooms), 400

        saved, errors = [], []
        conn = get_db()
        try:
            for f in files:
                if not f or not f.filename:
                    continue
                original = Path(f.filename).name
                ext = Path(original).suffix.lower()
                if ext not in ALLOWED_EXTENSIONS:
                    errors.append(f"{original}：不支持的文件类型")
                    continue
                stored = f"{secrets.token_hex(16)}{ext}"
                f.save(UPLOAD_DIR / stored)
                size = (UPLOAD_DIR / stored).stat().st_size
                conn.execute(
                    "INSERT INTO files (room, name, original_name, stored_name, size, created_at)"
                    " VALUES (?,?,?,?,?,?)",
                    (room, name, original, stored, size, now_str()),
                )
                saved.append(original)
        finally:
            conn.commit()
            conn.close()

        if saved:
            flash(f"已上传 {len(saved)} 个文件到房间 {room}", "ok")
        for e in errors:
            flash(e, "error")
        return redirect(url_for("upload"))

    return render_template("upload.html", rooms=rooms)


# ------------------------------------------------------------ 管理 ----

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        check_csrf()
        if request.form.get("password") == admin_password():
            session["is_admin"] = True
            return redirect(url_for("admin_rooms"))
        flash("密码错误", "error")
    return render_template("admin_login.html")


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    check_csrf()
    session.pop("is_admin", None)
    return redirect(url_for("upload"))


@app.route("/admin")
@login_required
def admin_rooms():
    conn = get_db()
    rows = conn.execute(
        "SELECT r.name AS room, COUNT(f.id) AS cnt,"
        " COALESCE(SUM(f.size), 0) AS total, MAX(f.created_at) AS last"
        " FROM rooms r LEFT JOIN files f ON f.room = r.name"
        " GROUP BY r.name ORDER BY r.name"
    ).fetchall()
    conn.close()
    return render_template("admin_rooms.html", rooms=rows)


@app.route("/admin/room/<path:room>")
@login_required
def admin_room(room):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM files WHERE room = ? ORDER BY created_at DESC, id DESC", (room,)
    ).fetchall()
    conn.close()
    return render_template("admin_room.html", room=room, files=rows,
                           default_printer=get_setting("default_printer", ""))


@app.route("/admin/room/add", methods=["POST"])
@login_required
def room_add():
    check_csrf()
    name = clean_room(request.form.get("room", ""))
    if not name:
        flash("房间号不能为空", "error")
        return redirect(url_for("admin_rooms"))
    conn = get_db()
    try:
        exists = conn.execute("SELECT 1 FROM rooms WHERE name = ?", (name,)).fetchone()
        if exists:
            flash(f"房间 {name} 已存在", "error")
        else:
            conn.execute("INSERT INTO rooms (name, created_at) VALUES (?, ?)",
                         (name, now_str()))
            conn.commit()
            flash(f"已添加房间 {name}", "ok")
    finally:
        conn.close()
    return redirect(url_for("admin_rooms"))


@app.route("/admin/room/<path:room>/delete", methods=["POST"])
@login_required
def room_delete(room):
    check_csrf()
    conn = get_db()
    rows = conn.execute("SELECT stored_name FROM files WHERE room = ?", (room,)).fetchall()
    conn.execute("DELETE FROM files WHERE room = ?", (room,))
    conn.execute("DELETE FROM rooms WHERE name = ?", (room,))
    conn.commit()
    conn.close()
    for row in rows:
        try:
            (UPLOAD_DIR / row["stored_name"]).unlink(missing_ok=True)
            (PREVIEW_DIR / f"{Path(row['stored_name']).stem}.pdf").unlink(missing_ok=True)
        except OSError:
            pass
    flash(f"已删除房间 {room} 及其 {len(rows)} 个文件", "ok")
    return redirect(url_for("admin_rooms"))


@app.route("/admin/printers")
@login_required
def admin_printers():
    printers = []
    try:
        import win32print
        printers = [
            p[2] for p in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        ]
    except Exception:
        pass
    return jsonify(printers=printers, default=get_setting("default_printer", ""))


@app.route("/admin/printer/set_default", methods=["POST"])
@login_required
def printer_set_default():
    check_csrf()
    printer = (request.form.get("printer") or "").strip()
    if printer:
        set_setting("default_printer", printer)
        flash(f"已将默认打印打印机设为：{printer}", "ok")
    else:
        set_setting("default_printer", "")
        flash("已清除默认打印打印机，将使用系统默认打印机", "ok")
    return redirect(request.referrer or url_for("admin_rooms"))


@app.route("/admin/room/<path:room>/print", methods=["POST"])
@login_required
def admin_room_print(room):
    check_csrf()
    printer = (request.form.get("printer") or "").strip() or None
    separator = request.form.get("separator") == "1"
    conn = get_db()
    # 按姓名分组排序，同一个人文件相邻，便于取件后分拣
    rows = conn.execute(
        "SELECT * FROM files WHERE room = ? ORDER BY name, created_at, id", (room,)
    ).fetchall()
    conn.close()

    results = []
    for row in rows:
        path = UPLOAD_DIR / row["stored_name"]
        if not path.is_file():
            results.append((row["original_name"], False, "文件丢失"))
            continue
        if separator:
            ok, msg = printlib.print_cover(
                [f"房间: {row['room']}",
                 f"姓名: {row['name']}",
                 f"文件: {row['original_name']}",
                 f"上传: {row['created_at']}"],
                printer)
            if not ok:
                results.append((f"[分隔页] {row['original_name']}", False, msg))
        ok, msg = printlib.print_file(path, printer)
        results.append((row["original_name"], ok, msg))

    ok_count = sum(1 for r in results if r[1])
    if ok_count == len(results):
        flash(f"打印完成：成功 {ok_count}/{len(results)} 个。已提交到打印队列，请到打印机刷教师卡取件。", "ok")
    else:
        flash(f"打印完成：成功 {ok_count}/{len(results)} 个（部分失败）。成功的已提交，请到打印机刷教师卡取件。", "error")
    return render_template("admin_room.html", room=room, files=rows,
                           default_printer=get_setting("default_printer", ""),
                           print_results=results)


@app.route("/admin/room/<path:room>/printview")
@login_required
def admin_room_printview(room):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM files WHERE room = ? ORDER BY created_at, id", (room,)
    ).fetchall()
    conn.close()

    items = []
    for row in rows:
        path = UPLOAD_DIR / row["stored_name"]
        if not path.is_file():
            items.append({"row": row, "kind": "missing"})
            continue
        ext = Path(row["stored_name"]).suffix.lower()
        if ext in printlib.IMAGE_EXTS:
            items.append({"row": row, "kind": "image",
                          "url": url_for("file_inline", file_id=row["id"])})
        elif ext in printlib.PDF_EXTS:
            items.append({"row": row, "kind": "pdf",
                          "url": url_for("file_inline", file_id=row["id"])})
        else:
            pdf = printlib.convert_to_pdf(path, PREVIEW_DIR)
            if pdf and pdf.is_file() and Path(pdf).suffix.lower() == ".pdf":
                items.append({"row": row, "kind": "pdf",
                              "url": url_for("file_preview", file_id=row["id"])})
            else:
                items.append({"row": row, "kind": "unsupported"})

    return render_template("print_view.html", room=room, items=items)


# ------------------------------------------------------------ 文件 ----

@app.route("/file/<int:file_id>")
@login_required
def file_inline(file_id):
    return _serve(file_id, as_attachment=False)


@app.route("/file/<int:file_id>/download")
@login_required
def file_download(file_id):
    return _serve(file_id, as_attachment=True)


@app.route("/preview/<int:file_id>")
@login_required
def file_preview(file_id):
    row = _fetch_file(file_id)
    if row is None:
        abort(404)
    src = UPLOAD_DIR / row["stored_name"]
    pdf = PREVIEW_DIR / f"{Path(row['stored_name']).stem}.pdf"
    if not pdf.is_file():
        converted = printlib.convert_to_pdf(src, PREVIEW_DIR)
        if not converted or not converted.is_file() or Path(converted).suffix.lower() != ".pdf":
            abort(415, "无法生成预览")
        pdf = converted
    return send_file(pdf, mimetype="application/pdf", as_attachment=False)


@app.route("/admin/file/<int:file_id>/delete", methods=["POST"])
@login_required
def file_delete(file_id):
    check_csrf()
    conn = get_db()
    row = conn.execute("SELECT * FROM files WHERE id = ?", (file_id,)).fetchone()
    if row is None:
        conn.close()
        abort(404)
    room = row["room"]
    conn.execute("DELETE FROM files WHERE id = ?", (file_id,))
    conn.commit()
    conn.close()
    try:
        (UPLOAD_DIR / row["stored_name"]).unlink(missing_ok=True)
        (PREVIEW_DIR / f"{Path(row['stored_name']).stem}.pdf").unlink(missing_ok=True)
    except OSError:
        pass
    flash(f"已删除 {row['original_name']}", "ok")
    return redirect(url_for("admin_room", room=room))


@app.errorhandler(413)
def too_large(_e):
    return render_template("upload.html", rooms=get_room_names(),
                           error="文件太大，单次请求最大 100MB"), 413


init_db()

if __name__ == "__main__":
    pw = admin_password()
    print("=" * 62)
    print("  打印站已启动")
    print(f"  管理入口: http://127.0.0.1:5000/admin")
    if pw == "admin123":
        print("  [警告] 正在使用默认管理密码 admin123，请设置环境变量 ADMIN_PASSWORD 修改")
    print("=" * 62)
    try:
        from waitress import serve
        print("using waitress production server (32 threads)")
        serve(app, host="0.0.0.0", port=5000, threads=32)
    except ImportError:
        print("waitress not found, using dev server")
        app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
