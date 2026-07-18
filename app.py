# -*- coding: utf-8 -*-
"""Mebel360 Boshqaruv Pro V9.

Asosiy imkoniyatlar:
- Zamonaviy Rahbar, Konstruktor, Menejer, Ishchi va Shafyor bo'limlari.
- Har bir xodim uchun alohida shifrlangan login-parol va rol bo'yicha kirish.
- Konstruktor: DBF/CSV/TXT, PRO100 STO, kroy, kromka x1.1 va A4 PDF.
- Ishchi “Kroy kesildi” yoki “Kromka tayyor”ni belgilasa, mijoz kuzatuvi avtomatik yangilanadi.
- Mijoz uchun maxfiy havola, jonli tayyorlik foizi va buyurtma yangiliklari.
- Shafyor: “Yetkazishga tayyor”, “Yo'lda”, “Yetkazildi” holatlari va mijozga avtomatik xabar.
- SQLite baza, kunlik zaxira nusxa, CSRF himoyasi va login bloklash.

Ishga tushirish:
    pip install -r requirements.txt
    python app.py
Brauzer: http://127.0.0.1:5000
"""
from __future__ import annotations

import base64
import csv
import io
import json
import os
import re
import secrets
import sqlite3
import random
import struct
from dataclasses import dataclass, asdict, field
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path
from zoneinfo import ZoneInfo
from typing import Any, Iterable

from flask import (
    Flask,
    Response,
    abort,
    flash,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

try:
    import qrcode
except ImportError:  # QR bo'lmasa ham dastur ishlaydi
    qrcode = None

APP_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("MEBEL360_KROY_DB", APP_DIR / "kroy_demo.db"))
BACKUP_DIR = Path(os.environ.get("MEBEL360_BACKUP_DIR", APP_DIR / "backups"))
UPLOAD_DIR = Path(os.environ.get("MEBEL360_UPLOAD_DIR", APP_DIR / "uploads"))
MAX_REQUEST_BYTES = 25 * 1024 * 1024
MAX_IMPORT_BYTES = 5 * 1024 * 1024
MAX_STO_BYTES = 20 * 1024 * 1024
MAX_PART_TYPES = 500
MAX_TOTAL_PARTS = 2500
MIN_PART_MM = 10
MAX_PART_MM = 10000
BACKUP_KEEP_DAYS = 14
EDGE_MULTIPLIER = 1.10
OPTIMIZATION_LABELS = {
    "fast": "Tezkor - 60 variantgacha",
    "full": "Standart - 120 variantgacha",
    "large": "Chuqur - 240 variantgacha",
}

ROLE_LABELS = {
    "admin": "Rahbar",
    "constructor": "Konstruktor",
    "manager": "Menejer",
    "worker": "Ishchi",
    "driver": "Shafyor",
}
ROLE_HOME_ENDPOINTS = {
    "admin": "dashboard",
    "constructor": "constructor",
    "manager": "manager_dashboard",
    "worker": "worker_center",
    "driver": "driver_dashboard",
}
DELIVERY_STATUSES = ("Kutilmoqda", "Yetkazishga tayyor", "Yo'lda", "Yetkazildi")


def optimization_label(mode: str) -> str:
    return OPTIMIZATION_LABELS.get(mode or "large", OPTIMIZATION_LABELS["large"])


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_REQUEST_BYTES
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=(
        os.environ.get(
            "MEBEL360_COOKIE_SECURE",
            "1" if os.environ.get("RENDER") else "0",
        )
        == "1"
    ),
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
)


def _secret_key() -> str:
    env = os.environ.get("MEBEL360_KROY_SECRET")
    if env:
        return env
    p = APP_DIR / ".kroy_secret"
    if p.exists():
        value = p.read_text("utf-8").strip()
        if value:
            return value
    value = secrets.token_hex(32)
    try:
        p.write_text(value, encoding="utf-8")
    except OSError:
        pass
    return value


app.secret_key = _secret_key()

TASHKENT_TZ = ZoneInfo("Asia/Tashkent")
MAX_LOGIN_ATTEMPTS = 5
LOGIN_BLOCK_SECONDS = 15 * 60


def tashkent_now() -> datetime:
    return datetime.now(TASHKENT_TZ)


def now_iso() -> str:
    return tashkent_now().isoformat(timespec="seconds")


def _as_bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value).strip().lower() in {"1", "true", "yes", "ha", "on", "y", "да"}


def csrf_token() -> str:
    token = session.get("_csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_csrf_token"] = token
    return token


@app.context_processor
def inject_template_helpers() -> dict[str, Any]:
    return {"csrf_token": csrf_token}


@app.before_request
def verify_csrf() -> None:
    if request.method not in {"POST", "PUT", "PATCH", "DELETE"}:
        return
    expected = session.get("_csrf_token", "")
    supplied = request.form.get("_csrf", "") or request.headers.get("X-CSRF-Token", "")
    if not expected or not supplied or not secrets.compare_digest(str(expected), str(supplied)):
        abort(400, description="Xavfsizlik tokeni eskirgan. Sahifani yangilab qayta urinib ko'ring.")


def create_database_backup() -> Path | None:
    """Bazani har muhim o'zgarishdan keyin kunlik nusxaga xavfsiz yangilaydi."""
    try:
        if str(DB_PATH) == ":memory:" or not DB_PATH.exists():
            return None
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        target = BACKUP_DIR / f"Mebel360_kroy_{tashkent_now():%Y-%m-%d}.db"
        temp = target.with_suffix(".tmp")
        source_conn = sqlite3.connect(DB_PATH)
        backup_conn = sqlite3.connect(temp)
        try:
            source_conn.backup(backup_conn)
        finally:
            backup_conn.close()
            source_conn.close()
        os.replace(temp, target)
        backups = sorted(BACKUP_DIR.glob("Mebel360_kroy_*.db"), key=lambda x: x.stat().st_mtime, reverse=True)
        for old_backup in backups[BACKUP_KEEP_DAYS:]:
            try:
                old_backup.unlink()
            except OSError:
                pass
        return target
    except (OSError, sqlite3.Error):
        return None


def _decode_text_file(data: bytes) -> str:
    for encoding in ("utf-8-sig", "cp1251", "cp866", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


def _header_key(value: str) -> str:
    return re.sub(r"[^a-zа-яё0-9]+", "", value.strip().lower())


def parse_table_parts(data: bytes) -> list[Part]:
    """CSV/TXT jadvalidan detallarni o'qiydi. Ustun nomlari o'zbek/rus/inglizcha bo'lishi mumkin."""
    text = _decode_text_file(data).strip()
    if not text:
        raise ValueError("CSV/TXT fayl bo'sh")
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ";" if ";" in sample else ("\t" if "\t" in sample else ",")
    rows = [row for row in csv.reader(io.StringIO(text), delimiter=delimiter) if any(cell.strip() for cell in row)]
    if not rows:
        raise ValueError("CSV/TXT ichida ma'lumot topilmadi")

    aliases = {
        "name": {"name", "nom", "nomi", "detal", "detail", "наименование", "наим", "деталь"},
        "length": {"length", "uzunlik", "dlina", "длина", "l"},
        "width": {"width", "eni", "en", "shirina", "ширина", "w"},
        "qty": {"qty", "quantity", "many", "soni", "dona", "количество", "колво", "кол"},
        "rotate": {"rotate", "aylantirish", "povorot", "поворот"},
        "edge_left": {"edgeleft", "left", "lborder", "chap", "лево", "левый"},
        "edge_right": {"edgeright", "right", "rborder", "ong", "o'ng", "право", "правый"},
        "edge_top": {"edgetop", "top", "tborder", "tepa", "верх"},
        "edge_bottom": {"edgebottom", "bottom", "bborder", "past", "низ"},
        "edges": {"edges", "edge", "kromka", "кромка"},
    }
    normalized_aliases = {key: {_header_key(v) for v in values} for key, values in aliases.items()}
    first = [_header_key(x) for x in rows[0]]
    has_header = any(cell in normalized_aliases["length"] | normalized_aliases["width"] for cell in first)
    mapping: dict[str, int] = {}
    start = 0
    if has_header:
        for idx, cell in enumerate(first):
            for key, values in normalized_aliases.items():
                if cell in values and key not in mapping:
                    mapping[key] = idx
        start = 1
    else:
        mapping = {"name": 0, "length": 1, "width": 2, "qty": 3, "rotate": 4, "edges": 5}

    if "length" not in mapping or "width" not in mapping:
        raise ValueError("Jadvalda uzunlik va en ustunlari topilmadi")

    def cell(row: list[str], key: str, default: str = "") -> str:
        idx = mapping.get(key)
        return row[idx].strip() if idx is not None and idx < len(row) else default

    parts: list[Part] = []
    for row_no, row in enumerate(rows[start:], start + 1):
        try:
            length = int(float(cell(row, "length", "0").replace(" ", "").replace(",", ".")))
            width = int(float(cell(row, "width", "0").replace(" ", "").replace(",", ".")))
            qty = int(float(cell(row, "qty", "1").replace(" ", "").replace(",", ".")))
        except ValueError:
            raise ValueError(f"{row_no}-qatorda o'lcham yoki son noto'g'ri")
        name = cell(row, "name", f"Detal {row_no}") or f"Detal {row_no}"
        raw_edges = cell(row, "edges", "")
        edge_tokens = {_header_key(token) for token in re.split(r"[,;/|+\s]+", raw_edges) if token.strip()}
        parts.append(
            Part(
                uid=secrets.token_hex(4),
                name=name[:80],
                length=length,
                width=width,
                qty=qty,
                rotate=_as_bool(cell(row, "rotate", "1"), True),
                edge_left=_as_bool(cell(row, "edge_left")) or bool(edge_tokens & {"ch", "e1", "left", "chap", "l"}),
                edge_right=_as_bool(cell(row, "edge_right")) or bool(edge_tokens & {"o", "e2", "right", "ong", "r"}),
                edge_top=_as_bool(cell(row, "edge_top")) or bool(edge_tokens & {"t", "u1", "top", "tepa"}),
                edge_bottom=_as_bool(cell(row, "edge_bottom")) or bool(edge_tokens & {"p", "u2", "bottom", "past"}),
            )
        )
    if not parts:
        raise ValueError("CSV/TXT ichidan yaroqli detallar topilmadi")
    return validate_parts(parts)


def validate_parts(parts: list[Part]) -> list[Part]:
    if not parts:
        raise ValueError("Kamida bitta detal kiriting")
    if len(parts) > MAX_PART_TYPES:
        raise ValueError(f"Detal turlari {MAX_PART_TYPES} tadan oshmasin")
    total = 0
    seen_uids: set[str] = set()
    for index, part in enumerate(parts, 1):
        part.name = str(part.name).strip()[:80]
        if not part.name:
            raise ValueError(f"{index}-detal nomi yozilmagan")
        if not (MIN_PART_MM <= int(part.length) <= MAX_PART_MM):
            raise ValueError(f"{part.name}: uzunlik {MIN_PART_MM}-{MAX_PART_MM} mm oralig'ida bo'lsin")
        if not (MIN_PART_MM <= int(part.width) <= MAX_PART_MM):
            raise ValueError(f"{part.name}: en {MIN_PART_MM}-{MAX_PART_MM} mm oralig'ida bo'lsin")
        if not (1 <= int(part.qty) <= 999):
            raise ValueError(f"{part.name}: soni 1-999 oralig'ida bo'lsin")
        total += int(part.qty)
        if total > MAX_TOTAL_PARTS:
            raise ValueError(f"Jami detallar {MAX_TOTAL_PARTS} tadan oshmasin")
        uid = re.sub(r"[^A-Za-z0-9_-]", "", str(part.uid))[:64]
        if not uid or uid in seen_uids:
            uid = secrets.token_hex(8)
        part.uid = uid
        seen_uids.add(uid)
    return parts


def save_sto_attachment(job_id: int, original_name: str, data: bytes) -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    safe = secure_filename(original_name) or "pro100.sto"
    stored = f"job_{job_id}_{secrets.token_hex(8)}.sto"
    path = UPLOAD_DIR / stored
    path.write_bytes(data)
    conn = get_db()
    conn.execute(
        """INSERT INTO kroy_attachments(job_id,kind,original_name,stored_name,size_bytes,created_at)
           VALUES(?, 'PRO100_STO', ?, ?, ?, ?)""",
        (job_id, safe[:200], stored, len(data), now_iso()),
    )
    # MEBEL360_KORXONA_MIGRATION_V1
    # Eski ma'lumotlar bazasidagi Seh yozuvlarini ma'lumot yo'qotmasdan Korxonaga o'tkazadi.
    _korxona_pairs = [
        ('Seh uchun', 'Korxona uchun'), ('seh uchun', 'korxona uchun'), ('SEH UCHUN', 'KORXONA UCHUN'),
        ('Sehda', 'Korxonada'), ('sehda', 'korxonada'), ('SEHDA', 'KORXONADA'),
        ('Sehga', 'Korxonaga'), ('sehga', 'korxonaga'), ('SEHGA', 'KORXONAGA'),
        ('Sehdan', 'Korxonadan'), ('sehdan', 'korxonadan'), ('SEHDAN', 'KORXONADAN'),
        ('Sehning', 'Korxonaning'), ('sehning', 'korxonaning'), ('SEHNING', 'KORXONANING'),
        ('Sehni', 'Korxonani'), ('sehni', 'korxonani'), ('SEHNI', 'KORXONANI'),
        ('Sehlar', 'Korxonalar'), ('sehlar', 'korxonalar'), ('SEHLAR', 'KORXONALAR'),
        ('Seh', 'Korxona'), ('seh', 'korxona'), ('SEH', 'KORXONA'),
    ]
    try:
        _tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").fetchall()
        for _table_row in _tables:
            _table = str(_table_row[0] if not hasattr(_table_row, 'keys') else _table_row['name'])
            _safe_table = _table.replace('\"', '\"\"')
            _columns = conn.execute(f'PRAGMA table_info("{_safe_table}")').fetchall()
            for _col in _columns:
                _col_name = str(_col[1])
                _col_type = str(_col[2] or '').upper()
                if _col_type and 'TEXT' not in _col_type and 'CHAR' not in _col_type and 'CLOB' not in _col_type:
                    continue
                _safe_col = _col_name.replace('\"', '\"\"')
                for _old, _new in _korxona_pairs:
                    conn.execute(
                        f'UPDATE "{_safe_table}" SET "{_safe_col}"=REPLACE("{_safe_col}", ?, ?) '
                        f'WHERE "{_safe_col}" LIKE ?',
                        (_old, _new, f'%{_old}%'),
                    )
    except Exception as _korxona_migration_error:
        print('Korxona migratsiyasi ogohlantirishi:', _korxona_migration_error)

    conn.commit()
    conn.close()


def _register_pdf_fonts() -> tuple[str, str]:
    """Windows/Linuxda topilgan Unicode shriftni ishlatadi; faylni paketga qo'shmaydi."""
    candidates = [
        (r"C:\\Windows\\Fonts\\arial.ttf", r"C:\\Windows\\Fonts\\arialbd.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"),
    ]
    for regular, bold in candidates:
        if os.path.exists(regular) and os.path.exists(bold):
            try:
                pdfmetrics.registerFont(TTFont("M360Unicode", regular))
                pdfmetrics.registerFont(TTFont("M360UnicodeBold", bold))
                return "M360Unicode", "M360UnicodeBold"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"


PDF_FONT, PDF_FONT_BOLD = _register_pdf_fonts()


def pdf_text(value: Any) -> str:
    text = str(value)
    if PDF_FONT == "Helvetica":
        replacements = {"‘": "'", "’": "'", "ʻ": "'", "ʼ": "'", "×": "x", "–": "-", "—": "-"}
        for old, new in replacements.items():
            text = text.replace(old, new)
        text = text.encode("latin-1", "replace").decode("latin-1")
    return text


BASE_CSS = r"""
:root{--nav:#0f1b33;--blue:#2563eb;--line:#dbe2ea;--bg:#f4f7fb;--ok:#059669;--warn:#d97706;--bad:#dc2626;--text:#172033}
*{box-sizing:border-box}body{margin:0;background:var(--bg);font-family:Arial,Helvetica,sans-serif;color:var(--text)}
header{background:linear-gradient(135deg,#0f1b33,#1d4ed8);color:#fff;padding:18px 16px}.top{max-width:1450px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:12px}.brand{font-size:25px;font-weight:800}.sub{font-size:12px;opacity:.85;margin-top:4px}
.wrap{max-width:1450px;margin:auto;padding:15px}.grid{display:grid;grid-template-columns:410px 1fr;gap:14px}.panel{background:#fff;border-radius:15px;padding:14px;box-shadow:0 8px 26px #17203312;margin-bottom:14px}h2,h3{margin:0 0 11px}label{font-size:12px;font-weight:700;display:block;margin-top:8px}input,select,textarea{width:100%;padding:9px;border:1px solid #cbd5e1;border-radius:8px;margin-top:4px;background:#fff}input[type=checkbox]{width:auto;margin:0 5px 0 0}.row{display:grid;grid-template-columns:1fr 1fr;gap:8px}.row4{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}.checks{display:flex;flex-wrap:wrap;gap:8px;margin-top:7px}.check{background:#f1f5f9;border-radius:7px;padding:7px;font-size:12px}
button,.btn{display:inline-block;border:0;border-radius:9px;padding:10px 13px;background:var(--blue);color:#fff;font-weight:700;text-decoration:none;cursor:pointer}.btn2{background:#334155}.ok{background:var(--ok)}.warn{background:var(--warn)}.danger{background:var(--bad)}.light{background:#e2e8f0;color:#1e293b}.actions{display:flex;gap:7px;flex-wrap:wrap;margin-top:10px}.actions form{margin:0}
table{width:100%;border-collapse:collapse;font-size:12px}th,td{border-bottom:1px solid #e5e7eb;padding:8px;text-align:left;vertical-align:top}th{background:#f8fafc;position:sticky;top:0}.tablewrap{overflow:auto;max-height:500px}.muted{color:#64748b;font-size:12px}.badge{display:inline-block;background:#e0e7ff;color:#3730a3;border-radius:20px;padding:3px 8px;font-size:11px;font-weight:700}.flash{padding:10px 12px;border-radius:9px;background:#dcfce7;color:#166534;margin-bottom:10px}.flash.bad{background:#fee2e2;color:#991b1b}
.sheet-grid{display:grid;grid-template-columns:repeat(2,minmax(350px,1fr));gap:12px}.sheet-card{background:#fff;border-radius:14px;padding:10px;box-shadow:0 5px 18px #17203310}.sheet-head{display:flex;justify-content:space-between;align-items:start;gap:8px;margin-bottom:7px}.sheet-title{font-weight:800}.sheet-svg{width:100%;height:auto;background:#fff;border:1px solid #94a3b8;touch-action:pan-x pan-y}.legend{font-size:11px;color:#475569;margin-top:6px}.edge-list{font-weight:700;color:#b91c1c}.small{font-size:11px}.empty{padding:25px;text-align:center;color:#64748b;border:2px dashed #cbd5e1;border-radius:12px}.opt-box{margin-top:9px;padding:10px;background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px}.offcut-stat{color:#9a3412;font-weight:800}.variant-note{font-size:11px;color:#047857;margin-top:5px}
.worker-header{display:flex;justify-content:space-between;gap:10px;align-items:center}.top-actions{display:flex;align-items:center;gap:8px}.top-actions form{margin:0}.auth-wrap{max-width:460px;margin:35px auto}.auth-note{padding:10px 12px;border-radius:9px;background:#eff6ff;border:1px solid #bfdbfe;color:#1e3a8a;font-size:12px;margin-bottom:12px}.statusline{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.stat{padding:8px 11px;border-radius:10px;background:#eef2ff;font-size:12px;font-weight:700}.qr{width:150px;height:150px;background:#fff;padding:6px;border-radius:10px}.part-cards{display:none}.part-card{border:1px solid #dbe2ea;border-radius:10px;padding:9px;margin:7px 0;background:#fff}
.edge-help{margin:8px 0 2px;padding:9px 10px;background:#fff7ed;border:1px solid #fed7aa;border-radius:9px;color:#9a3412;font-size:12px}
.edge-pickers{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:7px}.edge-picker{border:1px solid #dbe2ea;border-radius:10px;padding:8px;background:#f8fafc}.edge-picker-title{display:flex;justify-content:space-between;align-items:center;gap:6px;margin-bottom:5px}.edge-dim{font-weight:800;font-size:15px}.edge-side-name{font-size:10px;color:#64748b}.edge-lines{display:grid;grid-template-columns:1fr 1fr;gap:6px}
button.edge-toggle{display:block;width:100%;height:25px;padding:0 6px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:#475569;box-shadow:none}.edge-toggle .edge-stroke{display:block;height:3px;border-radius:5px;background:#94a3b8;transition:.15s}.edge-toggle.active{border-color:#ef4444;background:#fff1f2}.edge-toggle.active .edge-stroke{height:6px;background:#dc2626;box-shadow:0 0 0 1px #b91c1c}.edge-toggle:hover{border-color:#64748b}.edge-toggle.active:hover{border-color:#b91c1c}.edge-no{font-size:9px;position:absolute;opacity:0}.dim-cell{min-width:112px}.dim-value{font-weight:800;font-size:13px;text-align:center;margin-bottom:4px}.table-edge-lines{display:grid;grid-template-columns:1fr 1fr;gap:4px}.table-edge-lines button.edge-toggle{height:20px}.table-edge-lines .edge-toggle .edge-stroke{height:2px}.table-edge-lines .edge-toggle.active .edge-stroke{height:5px}.part-card .edge-pickers{margin:8px 0}.part-card .edge-picker{padding:6px}.part-card .edge-dim{font-size:13px}
.grain-box{margin-top:10px;padding:10px;border:1px solid #bfdbfe;background:#eff6ff;border-radius:10px;display:flex;align-items:center;justify-content:space-between;gap:10px}.grain-title{font-weight:800}.grain-note{font-size:11px;color:#475569;margin-top:3px}.grain-toggle{display:flex;align-items:center;gap:7px;padding:8px 10px;border-radius:9px;background:#fff;border:1px solid #cbd5e1;cursor:pointer;white-space:nowrap}.grain-toggle input{margin:0;width:auto}.grain-badge{display:inline-block;padding:4px 7px;border-radius:16px;background:#dcfce7;color:#166534;font-size:11px;font-weight:800}.grain-badge.locked{background:#fee2e2;color:#991b1b}.rotate-action{padding:6px 8px;font-size:11px;background:#0f766e}.rotate-action.locked{background:#b91c1c}.meter-box{margin-top:10px;padding:10px 12px;border-radius:10px;background:#f0fdf4;border:1px solid #bbf7d0;display:flex;align-items:center;justify-content:space-between;gap:10px}.meter-box strong{font-size:20px;color:#166534}.meter-box small{color:#64748b}.meter-total{margin-top:10px;padding:10px 12px;border-radius:10px;background:#fff7ed;border:1px solid #fed7aa;font-weight:800;color:#9a3412}.meter-cell{font-weight:800;color:#166534;white-space:nowrap}.grain-cell{min-width:135px}
@media(max-width:1000px){.grid{grid-template-columns:1fr}.sheet-grid{grid-template-columns:1fr}}
@media(max-width:600px){header{padding:13px 10px}.wrap{padding:8px}.brand{font-size:21px}.panel{padding:10px;border-radius:11px}.row,.row4{grid-template-columns:1fr 1fr}.tablewrap.desktop{display:none}.part-cards{display:block}.sheet-card{padding:7px}.sheet-head{display:block}.actions .btn,.actions button{flex:1;text-align:center}.qr{width:125px;height:125px}}

/* Mebel360 Pro V8 zamonaviy ko'rinish */
body{background:radial-gradient(circle at top left,#dbeafe 0,transparent 28%),radial-gradient(circle at top right,#dcfce7 0,transparent 26%),#f5f7fb;min-height:100vh}
header{position:sticky;top:0;z-index:30;padding:13px 16px;background:rgba(10,20,38,.94);backdrop-filter:blur(14px);box-shadow:0 10px 30px #0f172a22}
.brand{letter-spacing:-.5px}.brand-mark{display:inline-flex;align-items:center;justify-content:center;width:36px;height:36px;margin-right:9px;border-radius:11px;background:linear-gradient(135deg,#60a5fa,#22c55e);box-shadow:0 8px 20px #22c55e33}.module-pill{display:inline-block;margin-left:8px;padding:4px 8px;border:1px solid #ffffff30;border-radius:999px;font-size:10px;vertical-align:middle;color:#dbeafe}
.wrap{padding-top:20px}.panel{border:1px solid #e2e8f0;box-shadow:0 14px 38px #0f172a0d;transition:transform .18s ease,box-shadow .18s ease}.panel:hover{box-shadow:0 18px 45px #0f172a14}.panel h2,.panel h3{letter-spacing:-.3px}
.hero{position:relative;overflow:hidden;background:linear-gradient(135deg,#0f172a 0%,#1e3a8a 55%,#0f766e 100%);color:#fff;border:0;padding:22px}.hero:after{content:"";position:absolute;width:260px;height:260px;border-radius:50%;right:-85px;top:-125px;background:#ffffff12}.hero-grid{position:relative;z-index:1;display:flex;align-items:center;justify-content:space-between;gap:18px}.hero h1{margin:0 0 6px;font-size:27px}.hero p{margin:0;color:#dbeafe;max-width:760px}.hero-actions{display:flex;gap:8px;flex-wrap:wrap}.hero .btn{background:#fff;color:#0f172a}.hero .btn2{background:#ffffff18;border:1px solid #ffffff35}
.step-title{display:flex;align-items:center;gap:9px}.step-no{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:9px;background:linear-gradient(135deg,#2563eb,#0f766e);color:#fff;font-size:13px}.section-note{margin:-4px 0 12px;color:#64748b;font-size:12px}
.import-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px}.file-box{border:1px dashed #93c5fd;background:#eff6ff;border-radius:12px;padding:11px}.file-box.green{border-color:#86efac;background:#f0fdf4}.file-box h4{margin:0 0 6px;font-size:13px}.file-box input{background:#fff}
input,select,textarea{border:1px solid #d6deea;background:#fbfdff;transition:border .16s,box-shadow .16s,background .16s}input:focus,select:focus,textarea:focus{outline:none;border-color:#3b82f6;background:#fff;box-shadow:0 0 0 4px #3b82f61a}
button,.btn{box-shadow:0 5px 14px #2563eb22;transition:transform .13s ease,filter .13s ease,box-shadow .13s ease}button:hover,.btn:hover{transform:translateY(-1px);filter:brightness(1.03);box-shadow:0 8px 20px #0f172a1c}button:active,.btn:active{transform:translateY(0)}
.stat{border:1px solid #c7d2fe;background:linear-gradient(180deg,#f8faff,#eef2ff)}.badge{border:1px solid #c7d2fe}.badge.ready{background:#dcfce7;color:#166534;border-color:#86efac}.badge.progress{background:#fef3c7;color:#92400e;border-color:#fcd34d}.badge.revoked{background:#fee2e2;color:#991b1b;border-color:#fecaca}
.sheet-card{border:1px solid #dbe3ef;box-shadow:0 12px 28px #0f172a0d}.sheet-svg{border-radius:8px}.security-note{padding:10px 12px;border-radius:11px;background:#f0fdf4;border:1px solid #bbf7d0;color:#166534;font-size:12px;margin-top:10px}.warning-note{padding:10px 12px;border-radius:11px;background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;font-size:12px;margin-top:10px}
.attachment{display:flex;align-items:center;justify-content:space-between;gap:10px;border:1px solid #dbe2ea;border-radius:10px;padding:9px 10px;background:#f8fafc;margin-top:8px}.attachment-name{font-weight:700}.attachment-meta{font-size:11px;color:#64748b}
.empty{background:#f8fafc}.tablewrap{border:1px solid #eef2f7;border-radius:10px}th{background:#f8fafc;color:#334155;text-transform:none}
@media(max-width:760px){.hero-grid{display:block}.hero-actions{margin-top:14px}.import-grid{grid-template-columns:1fr}.hero h1{font-size:22px}.grid{gap:8px}.top-actions .small{display:none}}

/* Mebel360 Pro V9 — rollar va mijoz kuzatuvi */
.navbar{display:flex;gap:7px;align-items:center;flex-wrap:wrap}.navlink{display:inline-flex;align-items:center;gap:6px;padding:8px 10px;border-radius:10px;color:#e2e8f0;text-decoration:none;font-size:12px;font-weight:800;border:1px solid transparent}.navlink:hover{background:#ffffff12;border-color:#ffffff1f}.navlink.active{background:#ffffff18;color:#fff}.role-chip{display:inline-flex;align-items:center;gap:6px;padding:6px 9px;border-radius:999px;background:#0ea5e922;border:1px solid #38bdf84d;color:#bae6fd;font-size:11px;font-weight:800}
.dashboard-hero{background:linear-gradient(135deg,#0f172a,#1d4ed8 52%,#0f766e);color:#fff;border:0;padding:26px}.dashboard-hero h1{font-size:30px;margin:0 0 7px}.dashboard-hero p{margin:0;color:#dbeafe;max-width:850px}.role-grid{display:grid;grid-template-columns:repeat(4,minmax(210px,1fr));gap:14px;margin:14px 0}.role-card{position:relative;overflow:hidden;display:block;background:#fff;border:1px solid #e2e8f0;border-radius:18px;padding:18px;text-decoration:none;color:#172033;box-shadow:0 14px 35px #0f172a0d;transition:.18s}.role-card:hover{transform:translateY(-3px);box-shadow:0 18px 42px #0f172a18}.role-card:after{content:"";position:absolute;width:105px;height:105px;border-radius:50%;right:-42px;top:-44px;background:linear-gradient(135deg,#60a5fa33,#22c55e33)}.role-icon{display:flex;align-items:center;justify-content:center;width:48px;height:48px;border-radius:14px;background:linear-gradient(135deg,#dbeafe,#dcfce7);font-size:24px;margin-bottom:12px}.role-card h3{margin:0 0 5px;font-size:18px}.role-card p{margin:0;color:#64748b;font-size:12px;line-height:1.45}.role-arrow{margin-top:13px;font-weight:800;color:#2563eb;font-size:12px}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.kpi{padding:15px;border-radius:15px;background:linear-gradient(180deg,#fff,#f8fafc);border:1px solid #e2e8f0}.kpi b{display:block;font-size:25px;letter-spacing:-1px}.kpi span{font-size:11px;color:#64748b;font-weight:700}
.progress-track{height:10px;background:#e2e8f0;border-radius:999px;overflow:hidden;min-width:100px}.progress-fill{height:100%;border-radius:999px;background:linear-gradient(90deg,#2563eb,#10b981);transition:width .3s}.progress-label{font-size:11px;font-weight:800;color:#334155;margin-top:4px}.order-card{border:1px solid #e2e8f0;border-radius:15px;padding:13px;background:#fff;margin-bottom:10px}.order-head{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}.order-code{font-size:17px;font-weight:900}.split-grid{display:grid;grid-template-columns:1.15fr .85fr;gap:14px}.copy-box{display:flex;gap:6px;align-items:center}.copy-box input{margin:0;font-size:11px}.copy-box button{white-space:nowrap}
.customer-shell{max-width:760px;margin:12px auto}.customer-top{position:relative;overflow:hidden;background:linear-gradient(135deg,#0f172a,#1e40af 55%,#0f766e);color:#fff;border-radius:22px;padding:24px;box-shadow:0 20px 50px #0f172a22}.customer-top:after{content:"360°";position:absolute;right:-12px;top:-22px;font-size:110px;font-weight:900;color:#ffffff0b}.customer-code{font-size:13px;color:#bfdbfe;font-weight:800}.customer-title{font-size:28px;font-weight:900;margin:4px 0}.customer-stage{font-size:14px;color:#dbeafe}.customer-progress{margin-top:18px}.customer-progress .progress-track{height:14px;background:#ffffff2b}.customer-progress .progress-fill{background:linear-gradient(90deg,#38bdf8,#4ade80)}.customer-progress-row{display:flex;justify-content:space-between;margin-bottom:6px;font-weight:800}.timeline{position:relative;margin:6px 0}.timeline-item{position:relative;padding:4px 0 16px 34px}.timeline-item:before{content:"";position:absolute;left:10px;top:7px;width:12px;height:12px;border-radius:50%;background:#10b981;box-shadow:0 0 0 5px #d1fae5}.timeline-item:after{content:"";position:absolute;left:15px;top:23px;bottom:-2px;width:2px;background:#dbe2ea}.timeline-item:last-child:after{display:none}.timeline-title{font-weight:900}.timeline-time{font-size:11px;color:#64748b;margin-top:2px}.timeline-message{font-size:12px;color:#475569;margin-top:3px;line-height:1.45}.stage-list{display:grid;grid-template-columns:repeat(5,1fr);gap:7px;margin-top:12px}.stage-step{padding:9px 6px;text-align:center;border-radius:11px;background:#f1f5f9;color:#64748b;font-size:10px;font-weight:800;border:1px solid #e2e8f0}.stage-step.done{background:#ecfdf5;color:#047857;border-color:#a7f3d0}.stage-step.current{background:#eff6ff;color:#1d4ed8;border-color:#bfdbfe}.auto-note{font-size:11px;color:#64748b;text-align:center;margin-top:10px}
.notice-preview{padding:14px;border-radius:14px;background:linear-gradient(135deg,#eff6ff,#ecfdf5);border:1px solid #bfdbfe}.notice-preview b{color:#1e40af}.delivery-buttons{display:flex;gap:6px;flex-wrap:wrap}.delivery-buttons form{margin:0}.delivery-status{font-weight:900}.delivery-status.road{color:#d97706}.delivery-status.done{color:#059669}.mini-form{padding:12px;border-radius:13px;background:#f8fafc;border:1px solid #e2e8f0}.user-grid{display:grid;grid-template-columns:360px 1fr;gap:14px}
@media(max-width:1050px){.role-grid{grid-template-columns:repeat(2,1fr)}.split-grid,.user-grid{grid-template-columns:1fr}.kpi-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:600px){.navbar{display:none}.role-grid{grid-template-columns:1fr}.kpi-grid{grid-template-columns:1fr 1fr}.dashboard-hero{padding:19px}.dashboard-hero h1{font-size:23px}.customer-title{font-size:23px}.stage-list{grid-template-columns:1fr}.stage-step{text-align:left}.copy-box{display:block}.copy-box button{margin-top:6px;width:100%}}

@media print{
@page{size:A4 portrait;margin:8mm}
html,body{background:#fff}
header,.no-print{display:none!important}
.wrap{max-width:none;padding:0}
.panel,.sheet-card{box-shadow:none}
.sheet-grid{display:grid;grid-template-columns:1fr;gap:4mm}
.sheet-card{border:1px solid #94a3b8;border-radius:0;padding:3mm;height:132mm;overflow:hidden;break-inside:avoid;page-break-inside:avoid;margin:0}
.sheet-card:nth-child(2n){page-break-after:always}
.sheet-svg{width:100%;max-height:101mm}
.sheet-head{margin-bottom:2mm}
.legend{font-size:8px;line-height:1.15;margin-top:1mm}
}
"""


@dataclass
class Part:
    uid: str
    name: str
    length: int
    width: int
    qty: int = 1
    rotate: bool = True
    edge_left: bool = False
    edge_right: bool = False
    edge_top: bool = False
    edge_bottom: bool = False


@dataclass
class Placement:
    uid: str
    name: str
    x: int
    y: int
    w: int
    h: int
    original_length: int
    original_width: int
    rotated: bool
    edge_left: bool
    edge_right: bool
    edge_top: bool
    edge_bottom: bool
    rotate_allowed: bool = True


@dataclass
class FreeRect:
    x: int
    y: int
    w: int
    h: int


@dataclass
class SheetPlan:
    number: int
    width: int
    height: int
    placements: list[Placement]
    leftovers: list[FreeRect] = field(default_factory=list)
    variant: str = ""


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db() -> None:
    conn = get_db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS kroy_jobs(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_code TEXT NOT NULL,
            customer TEXT DEFAULT '',
            material TEXT NOT NULL,
            sheet_length INTEGER NOT NULL,
            sheet_width INTEGER NOT NULL,
            kerf INTEGER DEFAULT 4,
            trim INTEGER DEFAULT 10,
            worker_name TEXT DEFAULT '',
            token TEXT UNIQUE NOT NULL,
            status TEXT DEFAULT 'Yangi',
            created_at TEXT NOT NULL,
            sent_at TEXT DEFAULT '',
            finished_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS kroy_parts(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id INTEGER NOT NULL,
            uid TEXT NOT NULL,
            name TEXT NOT NULL,
            length INTEGER NOT NULL,
            width INTEGER NOT NULL,
            qty INTEGER DEFAULT 1,
            rotate INTEGER DEFAULT 1,
            edge_left INTEGER DEFAULT 0,
            edge_right INTEGER DEFAULT 0,
            edge_top INTEGER DEFAULT 0,
            edge_bottom INTEGER DEFAULT 0,
            FOREIGN KEY(job_id) REFERENCES kroy_jobs(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS kroy_sheets(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id INTEGER NOT NULL,
            sheet_no INTEGER NOT NULL,
            plan_json TEXT NOT NULL,
            cut_done INTEGER DEFAULT 0,
            edge_done INTEGER DEFAULT 0,
            note TEXT DEFAULT '',
            FOREIGN KEY(job_id) REFERENCES kroy_jobs(id) ON DELETE CASCADE,
            UNIQUE(job_id,sheet_no)
        );
        CREATE TABLE IF NOT EXISTS kroy_import_drafts(
            token TEXT PRIMARY KEY,
            source_name TEXT DEFAULT '',
            parts_json TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS app_users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'admin',
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS login_attempts(
            ip TEXT PRIMARY KEY,
            attempts INTEGER NOT NULL DEFAULT 0,
            blocked_until INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS audit_log(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT DEFAULT '',
            action TEXT NOT NULL,
            ip TEXT DEFAULT '',
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS kroy_attachments(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id INTEGER NOT NULL,
            kind TEXT NOT NULL DEFAULT 'FILE',
            original_name TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            size_bytes INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(job_id) REFERENCES kroy_jobs(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS customer_updates(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id INTEGER NOT NULL,
            event_key TEXT NOT NULL,
            title TEXT NOT NULL,
            message TEXT DEFAULT '',
            stage TEXT DEFAULT '',
            progress INTEGER DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(job_id) REFERENCES kroy_jobs(id) ON DELETE CASCADE,
            UNIQUE(job_id,event_key)
        );
        """
    )
    columns = {row[1] for row in conn.execute("PRAGMA table_info(kroy_jobs)").fetchall()}
    if "optimization_mode" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN optimization_mode TEXT DEFAULT 'large'")
    if "tested_variants" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN tested_variants INTEGER DEFAULT 0")
    if "best_variant" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN best_variant TEXT DEFAULT ''")
    if "worker_link_active" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN worker_link_active INTEGER DEFAULT 1")
    if "worker_token_created_at" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN worker_token_created_at TEXT DEFAULT ''")
        conn.execute("UPDATE kroy_jobs SET worker_token_created_at=COALESCE(NULLIF(sent_at,''),created_at)")
    if "customer_token" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN customer_token TEXT DEFAULT ''")
    if "customer_link_active" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN customer_link_active INTEGER DEFAULT 1")
    if "delivery_status" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN delivery_status TEXT DEFAULT 'Kutilmoqda'")
    if "delivery_note" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN delivery_note TEXT DEFAULT ''")
    if "delivered_at" not in columns:
        conn.execute("ALTER TABLE kroy_jobs ADD COLUMN delivered_at TEXT DEFAULT ''")
    for row in conn.execute("SELECT id FROM kroy_jobs WHERE COALESCE(customer_token,'')='' ").fetchall():
        conn.execute("UPDATE kroy_jobs SET customer_token=? WHERE id=?", (secrets.token_urlsafe(24), row["id"]))
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_kroy_jobs_customer_token ON kroy_jobs(customer_token)")
    conn.commit()
    conn.close()


def _truthy_border(value: str) -> bool:
    value = (value or "").strip().upper()
    return value not in {"", "0", "F", "FALSE", "N", "NO"}


def parse_dbf(data: bytes) -> list[Part]:
    """2D-PLACE xBase/DBF faylidan detallarni o'qiydi.

    Kutilyotgan ustunlar: NAME, LENGTH, WIDTH, MANY, ROTATE,
    LBORDER, RBORDER, TBORDER, BBORDER.
    """
    if len(data) < 65:
        raise ValueError("DBF fayl juda kichik yoki buzilgan")
    records = struct.unpack("<I", data[4:8])[0]
    header_len = struct.unpack("<H", data[8:10])[0]
    record_len = struct.unpack("<H", data[10:12])[0]
    if header_len <= 32 or record_len <= 1 or header_len > len(data):
        raise ValueError("DBF sarlavhasi noto'g'ri")

    fields: list[tuple[str, str, int, int]] = []
    offset = 32
    pos = 1
    while offset + 32 <= header_len and data[offset] != 0x0D:
        desc = data[offset : offset + 32]
        name = desc[:11].split(b"\0", 1)[0].decode("ascii", "ignore").upper()
        ftype = chr(desc[11])
        length = desc[16]
        if not name or length <= 0:
            raise ValueError("DBF ustun tavsifi buzilgan")
        fields.append((name, ftype, length, pos))
        pos += length
        offset += 32

    available = {f[0] for f in fields}
    required = {"LENGTH", "WIDTH", "MANY"}
    if not required.issubset(available):
        raise ValueError("Bu DBFda LENGTH, WIDTH yoki MANY ustuni topilmadi")

    def text(raw: bytes, ftype: str) -> str:
        raw = raw.rstrip(b" \x00")
        if ftype in {"N", "F", "L"}:
            return raw.decode("ascii", "ignore").strip()
        for enc in ("cp866", "cp1251", "utf-8", "latin1"):
            try:
                return raw.decode(enc).strip()
            except UnicodeDecodeError:
                continue
        return raw.decode("latin1", "replace").strip()

    parts: list[Part] = []
    for idx in range(min(records, 10000)):
        start = header_len + idx * record_len
        rec = data[start : start + record_len]
        if len(rec) < record_len:
            break
        if rec[:1] == b"*":
            continue
        row: dict[str, str] = {}
        for name, ftype, length, field_pos in fields:
            row[name] = text(rec[field_pos : field_pos + length], ftype)
        try:
            length = int(float(row.get("LENGTH", "0") or 0))
            width = int(float(row.get("WIDTH", "0") or 0))
            qty = max(1, int(float(row.get("MANY", "1") or 1)))
        except ValueError:
            continue
        if length <= 0 or width <= 0:
            continue
        parts.append(
            Part(
                uid=secrets.token_hex(4),
                name=row.get("NAME") or f"Detal {idx + 1}",
                length=length,
                width=width,
                qty=qty,
                rotate=(row.get("ROTATE", "0").strip() not in {"1", "N", "F"}),
                edge_left=_truthy_border(row.get("LBORDER", "")),
                edge_right=_truthy_border(row.get("RBORDER", "")),
                edge_top=_truthy_border(row.get("TBORDER", "")),
                edge_bottom=_truthy_border(row.get("BBORDER", "")),
            )
        )
    if not parts:
        raise ValueError("DBF ichidan yaroqli detallar topilmadi")
    return parts


def rotated_edges(part: Part, rotated: bool) -> tuple[bool, bool, bool, bool]:
    if not rotated:
        return part.edge_left, part.edge_right, part.edge_top, part.edge_bottom
    # 90 daraja soat yo'nalishi bo'yicha: yangi L=eski B, R=eski T, T=eski L, B=eski R
    return part.edge_bottom, part.edge_top, part.edge_left, part.edge_right


def _intersects(a: FreeRect, b: FreeRect) -> bool:
    return not (b.x >= a.x + a.w or b.x + b.w <= a.x or b.y >= a.y + a.h or b.y + b.h <= a.y)


def _contains(a: FreeRect, b: FreeRect) -> bool:
    return b.x >= a.x and b.y >= a.y and b.x + b.w <= a.x + a.w and b.y + b.h <= a.y + a.h


def _merge_free_rects(rects: list[FreeRect]) -> list[FreeRect]:
    """Yonma-yon qoldiqlarni birlashtirib, katta foydalaniladigan qoldiq saqlaydi."""
    result = [FreeRect(r.x, r.y, r.w, r.h) for r in rects if r.w > 0 and r.h > 0]
    changed = True
    while changed:
        changed = False
        for i in range(len(result)):
            if changed:
                break
            for j in range(i + 1, len(result)):
                a, b = result[i], result[j]
                merged: FreeRect | None = None
                # gorizontal yonma-yon
                if a.y == b.y and a.h == b.h and (a.x + a.w == b.x or b.x + b.w == a.x):
                    merged = FreeRect(min(a.x, b.x), a.y, a.w + b.w, a.h)
                # vertikal yonma-yon
                elif a.x == b.x and a.w == b.w and (a.y + a.h == b.y or b.y + b.h == a.y):
                    merged = FreeRect(a.x, min(a.y, b.y), a.w, a.h + b.h)
                if merged:
                    result.pop(j)
                    result.pop(i)
                    result.append(merged)
                    changed = True
                    break
    return result


def _prune_free_rects(rects: list[FreeRect]) -> list[FreeRect]:
    cleaned: list[FreeRect] = []
    for i, rect in enumerate(rects):
        if rect.w <= 0 or rect.h <= 0:
            continue
        if any(i != j and _contains(other, rect) for j, other in enumerate(rects)):
            continue
        if not any((r.x, r.y, r.w, r.h) == (rect.x, rect.y, rect.w, rect.h) for r in cleaned):
            cleaned.append(rect)
    return _merge_free_rects(cleaned)


def _guillotine_split(fr: FreeRect, used_w: int, used_h: int, split_mode: str) -> list[FreeRect]:
    """Detalni bo'sh to'rtburchakning chap-tepasiga qo'yib, qoldiqni 2 katta bo'lakka ajratadi."""
    dw = fr.w - used_w
    dh = fr.h - used_h
    out: list[FreeRect] = []
    if split_mode == "vertical":
        # o'ngdagi qoldiq to'liq balandlikda qoladi
        if dw > 0:
            out.append(FreeRect(fr.x + used_w, fr.y, dw, fr.h))
        if dh > 0:
            out.append(FreeRect(fr.x, fr.y + used_h, used_w, dh))
    else:
        # pastdagi qoldiq to'liq kenglikda qoladi
        if dh > 0:
            out.append(FreeRect(fr.x, fr.y + used_h, fr.w, dh))
        if dw > 0:
            out.append(FreeRect(fr.x + used_w, fr.y, dw, used_h))
    return out


def _expanded_parts(parts: Iterable[Part]) -> list[tuple[Part, int]]:
    expanded: list[tuple[Part, int]] = []
    for part in parts:
        for copy_no in range(1, part.qty + 1):
            expanded.append((part, copy_no))
    return expanded


def _ordered_parts(expanded: list[tuple[Part, int]], mode: str, seed: int = 0) -> list[tuple[Part, int]]:
    data = list(expanded)
    if mode == "area":
        key = lambda item: (item[0].length * item[0].width, max(item[0].length, item[0].width), min(item[0].length, item[0].width))
    elif mode == "max_side":
        key = lambda item: (max(item[0].length, item[0].width), item[0].length * item[0].width, min(item[0].length, item[0].width))
    elif mode == "length":
        key = lambda item: (item[0].length, item[0].width, item[0].length * item[0].width)
    elif mode == "width":
        key = lambda item: (item[0].width, item[0].length, item[0].length * item[0].width)
    elif mode == "perimeter":
        key = lambda item: (item[0].length + item[0].width, item[0].length * item[0].width)
    elif mode == "thin_first":
        key = lambda item: (max(item[0].length, item[0].width) / max(1, min(item[0].length, item[0].width)), item[0].length * item[0].width)
    elif mode == "square_first":
        key = lambda item: (-abs(item[0].length - item[0].width), item[0].length * item[0].width)
    else:
        key = lambda item: (item[0].length * item[0].width, max(item[0].length, item[0].width))
    data.sort(key=key, reverse=True)
    if seed:
        # Bir xil yoki yaqin o'lchamli detallar tartibini xavfsiz ravishda o'zgartirib ko'radi.
        rng = random.Random(seed)
        buckets: dict[int, list[tuple[Part, int]]] = {}
        for item in data:
            area = item[0].length * item[0].width
            bucket = max(1, area // 50000)
            buckets.setdefault(bucket, []).append(item)
        data = []
        for bucket in sorted(buckets, reverse=True):
            group = buckets[bucket]
            rng.shuffle(group)
            data.extend(group)
    return data


def _candidate_score(
    heuristic: str,
    sheet_index: int,
    fr: FreeRect,
    used_w: int,
    used_h: int,
    new_free: list[FreeRect],
    split_mode: str,
) -> tuple[Any, ...]:
    area_fit = fr.w * fr.h - used_w * used_h
    short_fit = min(fr.w - used_w, fr.h - used_h)
    long_fit = max(fr.w - used_w, fr.h - used_h)
    largest = max((r.w * r.h for r in new_free), default=0)
    fragments = len(new_free)
    # Kichik y/x = detallar bir burchakka yig'iladi va qoldiq bir joyda qoladi.
    if heuristic == "large_offcut":
        return (-largest, fragments, fr.y, fr.x, area_fit, short_fit, sheet_index, split_mode)
    if heuristic == "compact":
        return (fr.y, fr.x, fragments, -largest, area_fit, short_fit, sheet_index, split_mode)
    if heuristic == "long_side":
        return (long_fit, short_fit, fragments, -largest, fr.y, fr.x, sheet_index, split_mode)
    return (area_fit, short_fit, fragments, -largest, fr.y, fr.x, sheet_index, split_mode)


def _pack_once(
    ordered: list[tuple[Part, int]],
    sheet_length: int,
    sheet_width: int,
    kerf: int,
    trim: int,
    heuristic: str,
    split_preference: str,
    variant_name: str,
) -> list[SheetPlan]:
    usable_w = sheet_length - 2 * trim
    usable_h = sheet_width - 2 * trim
    if usable_w <= 0 or usable_h <= 0:
        raise ValueError("List o'lchami yoki chet kesimi noto'g'ri")

    sheets: list[dict[str, Any]] = []
    initial = FreeRect(trim, trim, usable_w + kerf, usable_h + kerf)

    for part, copy_no in ordered:
        if part.length <= 0 or part.width <= 0:
            continue
        orientations = [(part.length, part.width, False)]
        if part.rotate and part.length != part.width:
            orientations.append((part.width, part.length, True))

        best: tuple[tuple[Any, ...], int, int, int, bool, str, list[FreeRect]] | None = None
        for si, sheet in enumerate(sheets):
            for fi, fr in enumerate(sheet["free"]):
                for w, h, rotated in orientations:
                    reserve_w, reserve_h = w + kerf, h + kerf
                    if reserve_w > fr.w or reserve_h > fr.h:
                        continue
                    modes = [split_preference] if split_preference in {"vertical", "horizontal"} else ["vertical", "horizontal"]
                    for split_mode in modes:
                        replacement = _guillotine_split(fr, reserve_w, reserve_h, split_mode)
                        candidate_free = sheet["free"][:fi] + sheet["free"][fi + 1:] + replacement
                        candidate_free = _prune_free_rects(candidate_free)
                        score = _candidate_score(heuristic, si, fr, reserve_w, reserve_h, candidate_free, split_mode)
                        candidate = (score, si, fi, w, h, rotated, split_mode, candidate_free)
                        if best is None or candidate[0] < best[0]:
                            best = candidate

        if best is None:
            fits: list[tuple[tuple[Any, ...], int, int, bool, str, list[FreeRect]]] = []
            for w, h, rotated in orientations:
                reserve_w, reserve_h = w + kerf, h + kerf
                if reserve_w > initial.w or reserve_h > initial.h:
                    continue
                modes = [split_preference] if split_preference in {"vertical", "horizontal"} else ["vertical", "horizontal"]
                for split_mode in modes:
                    candidate_free = _prune_free_rects(_guillotine_split(initial, reserve_w, reserve_h, split_mode))
                    score = _candidate_score(heuristic, len(sheets), initial, reserve_w, reserve_h, candidate_free, split_mode)
                    fits.append((score, w, h, rotated, split_mode, candidate_free))
            if not fits:
                raise ValueError(f"'{part.name}' ({part.length}x{part.width}) listga sig'maydi")
            _, w, h, rotated, split_mode, candidate_free = min(fits, key=lambda x: x[0])
            sheets.append({"free": candidate_free, "placements": []})
            si = len(sheets) - 1
            fr = initial
        else:
            _, si, fi, w, h, rotated, split_mode, candidate_free = best
            fr = sheets[si]["free"][fi]
            sheets[si]["free"] = candidate_free

        left, right, top, bottom = rotated_edges(part, rotated)
        sheets[si]["placements"].append(
            Placement(
                uid=f"{part.uid}-{copy_no}", name=part.name, x=fr.x, y=fr.y, w=w, h=h,
                original_length=part.length, original_width=part.width, rotated=rotated,
                edge_left=left, edge_right=right, edge_top=top, edge_bottom=bottom,
                rotate_allowed=part.rotate,
            )
        )

    plans: list[SheetPlan] = []
    for index, sheet in enumerate(sheets, 1):
        leftovers: list[FreeRect] = []
        for fr in _prune_free_rects(sheet["free"]):
            # Ishlatilgan detallar orasidagi arra izi qoldiq sifatida hisoblanmaydi.
            w = max(0, fr.w - kerf)
            h = max(0, fr.h - kerf)
            if w >= 80 and h >= 80:
                leftovers.append(FreeRect(fr.x, fr.y, w, h))
        leftovers.sort(key=lambda r: r.w * r.h, reverse=True)
        plans.append(SheetPlan(index, sheet_length, sheet_width, sheet["placements"], leftovers, variant_name))
    return plans


def _plan_score(plans: list[SheetPlan]) -> tuple[Any, ...]:
    # 1) listlar soni kam; 2) yirik qoldiq ko'p; 3) qoldiq bo'laklari kam; 4) detallar bir burchakka zich.
    largest_areas = [max((r.w * r.h for r in p.leftovers), default=0) for p in plans]
    total_largest = sum(largest_areas)
    fragment_count = sum(len(p.leftovers) for p in plans)
    bounding = 0
    for plan in plans:
        if plan.placements:
            min_x = min(p.x for p in plan.placements)
            min_y = min(p.y for p in plan.placements)
            max_x = max(p.x + p.w for p in plan.placements)
            max_y = max(p.y + p.h for p in plan.placements)
            bounding += (max_x - min_x) * (max_y - min_y)
    last_largest = largest_areas[-1] if largest_areas else 0
    return (len(plans), -total_largest, fragment_count, bounding, -last_largest)


def pack_parts(
    parts: Iterable[Part], sheet_length: int, sheet_width: int, kerf: int, trim: int,
    optimization_mode: str = "large",
) -> tuple[list[SheetPlan], int, str]:
    """Ko'p variantni tekshiradi va list soni/qoldiq bo'yicha eng yaxshisini qaytaradi."""
    expanded = _expanded_parts(parts)
    if not expanded:
        return [], 0, ""

    if optimization_mode == "fast":
        limit = 60
    elif optimization_mode == "full":
        limit = 120
    else:
        limit = 240

    # Juda ko'p detal bo'lsa eski kompyuterda uzoq kutmasligi uchun kamayadi,
    # ammo baribir 48 tadan ko'p variant tekshiriladi.
    if len(expanded) > 1200:
        limit = min(limit, 60)
    elif len(expanded) > 500:
        limit = min(limit, 80)
    elif len(expanded) > 250:
        limit = min(limit, 120)

    sort_modes = ["area", "max_side", "length", "width", "perimeter", "thin_first", "square_first"]
    heuristics = ["large_offcut", "compact", "best_area", "long_side"]
    splits = ["auto", "vertical", "horizontal"]
    variants: list[tuple[str, str, str, int]] = []
    # Avval turli detal tartiblari va avtomatik bo'lish usuli sinovdan o'tadi.
    # Shuning uchun Tezkor rejim ham faqat bitta tartibga qamalib qolmaydi.
    for heuristic in heuristics:
        for sort_mode in sort_modes:
            variants.append((sort_mode, heuristic, "auto", 0))
    # Keyin vertikal va gorizontal kesim ustuvorligi alohida tekshiriladi.
    for split in ("vertical", "horizontal"):
        for heuristic in ("large_offcut", "best_area"):
            for sort_mode in sort_modes:
                variants.append((sort_mode, heuristic, split, 0))
    # Bir xil/yaqin o'lchamli detallar uchun qo'shimcha tartib variantlari.
    for seed in range(1, 101):
        variants.append(("area", "large_offcut", "auto", seed))
        variants.append(("max_side", "compact", "auto", seed))

    best_plans: list[SheetPlan] | None = None
    best_score: tuple[Any, ...] | None = None
    best_name = ""
    tested = 0
    for sort_mode, heuristic, split, seed in variants[:limit]:
        ordered = _ordered_parts(expanded, sort_mode, seed)
        name = f"{sort_mode}/{heuristic}/{split}" + (f"/s{seed}" if seed else "")
        plans = _pack_once(ordered, sheet_length, sheet_width, kerf, trim, heuristic, split, name)
        score = _plan_score(plans)
        tested += 1
        if best_score is None or score < best_score:
            best_score = score
            best_plans = plans
            best_name = name

    assert best_plans is not None
    return best_plans, tested, best_name

def sheet_usage(plan: SheetPlan) -> float:
    total = plan.width * plan.height
    used = sum(p.w * p.h for p in plan.placements)
    return round((used / total * 100) if total else 0, 1)


def largest_offcut(plan: SheetPlan) -> FreeRect | None:
    return max(plan.leftovers, key=lambda r: r.w * r.h, default=None)


def offcut_text(plan: SheetPlan) -> str:
    r = largest_offcut(plan)
    if not r:
        return "Qoldiq yo'q"
    return f"Eng katta qoldiq {r.w}×{r.h} mm"


def edge_letters(p: Placement | Part) -> str:
    values = []
    if p.edge_top:
        values.append("T")
    if p.edge_bottom:
        values.append("P")
    if p.edge_left:
        values.append("Ch")
    if p.edge_right:
        values.append("O")
    return ", ".join(values) if values else "yo'q"



def edge_length_mm_values(
    length: int,
    width: int,
    qty: int,
    edge_left: bool,
    edge_right: bool,
    edge_top: bool,
    edge_bottom: bool,
) -> int:
    return max(0, qty) * (
        (length if edge_top else 0)
        + (length if edge_bottom else 0)
        + (width if edge_left else 0)
        + (width if edge_right else 0)
    )


def parts_edge_m(parts: Iterable[Any]) -> float:
    total_mm = 0
    for p in parts:
        get = (lambda key, default=0: p[key] if key in p.keys() else default) if isinstance(p, sqlite3.Row) else (lambda key, default=0: getattr(p, key, default))
        total_mm += edge_length_mm_values(
            int(get("length", 0)),
            int(get("width", 0)),
            int(get("qty", 1)),
            bool(get("edge_left", False)),
            bool(get("edge_right", False)),
            bool(get("edge_top", False)),
            bool(get("edge_bottom", False)),
        )
    return round((total_mm / 1000.0) * EDGE_MULTIPLIER, 3)


def plan_edge_m(plan: SheetPlan) -> float:
    total_mm = 0
    for p in plan.placements:
        total_mm += (
            (p.w if p.edge_top else 0)
            + (p.w if p.edge_bottom else 0)
            + (p.h if p.edge_left else 0)
            + (p.h if p.edge_right else 0)
        )
    return round((total_mm / 1000.0) * EDGE_MULTIPLIER, 3)


def plan_to_dict(plan: SheetPlan) -> dict[str, Any]:
    return {
        "number": plan.number, "width": plan.width, "height": plan.height,
        "placements": [asdict(p) for p in plan.placements],
        "leftovers": [asdict(r) for r in plan.leftovers],
        "variant": plan.variant,
    }


def dict_to_plan(data: dict[str, Any]) -> SheetPlan:
    return SheetPlan(
        number=int(data["number"]),
        width=int(data["width"]),
        height=int(data["height"]),
        placements=[Placement(**({"rotate_allowed": True, **p})) for p in data["placements"]],
        leftovers=[FreeRect(**r) for r in data.get("leftovers", [])],
        variant=str(data.get("variant", "")),
    )



def _add_customer_update(
    conn: sqlite3.Connection,
    job_id: int,
    event_key: str,
    title: str,
    message: str = "",
    stage: str = "",
    progress: int = 0,
) -> bool:
    cur = conn.execute(
        """INSERT OR IGNORE INTO customer_updates(job_id,event_key,title,message,stage,progress,created_at)
           VALUES(?,?,?,?,?,?,?)""",
        (job_id, event_key[:120], title[:140], message[:500], stage[:80], max(0, min(100, int(progress))), now_iso()),
    )
    return cur.rowcount == 1


def _job_progress_from_counts(total: int, cuts: int, edges: int, delivery_status: str = "Kutilmoqda") -> int:
    if total <= 0:
        base = 0
    else:
        base = round(((cuts + edges) / (total * 2)) * 80)
    if delivery_status == "Yetkazishga tayyor":
        return max(base, 85)
    if delivery_status == "Yo'lda":
        return max(base, 95)
    if delivery_status == "Yetkazildi":
        return 100
    return min(80, base)


def _progress_snapshot(conn: sqlite3.Connection, job_id: int) -> dict[str, Any]:
    stats = conn.execute(
        """SELECT COUNT(*) total,COALESCE(SUM(cut_done),0) cuts,COALESCE(SUM(edge_done),0) edges
           FROM kroy_sheets WHERE job_id=?""",
        (job_id,),
    ).fetchone()
    job = conn.execute("SELECT status,delivery_status FROM kroy_jobs WHERE id=?", (job_id,)).fetchone()
    total = int(stats["total"] or 0)
    cuts = int(stats["cuts"] or 0)
    edges = int(stats["edges"] or 0)
    delivery = (job["delivery_status"] if job else "Kutilmoqda") or "Kutilmoqda"
    progress = _job_progress_from_counts(total, cuts, edges, delivery)
    if delivery == "Yetkazildi":
        stage = "Buyurtma yetkazildi"
    elif delivery == "Yo'lda":
        stage = "Buyurtma yo'lda"
    elif delivery == "Yetkazishga tayyor":
        stage = "Yetkazishga tayyor"
    elif total and edges == total:
        stage = "Kroy va kromka tayyor"
    elif cuts:
        stage = "Kromka ishlari davom etmoqda" if total and cuts == total else "Kroy ishlari davom etmoqda"
    else:
        stage = "Buyurtma qabul qilindi"
    return {"total": total, "cuts": cuts, "edges": edges, "progress": progress, "stage": stage, "delivery_status": delivery}


def _customer_stages(snapshot: dict[str, Any]) -> list[dict[str, Any]]:
    total = int(snapshot["total"] or 0)
    cuts = int(snapshot["cuts"] or 0)
    edges = int(snapshot["edges"] or 0)
    delivery = snapshot["delivery_status"]
    flags = [
        ("Buyurtma", True),
        ("Kroy", bool(total and cuts == total)),
        ("Kromka", bool(total and edges == total)),
        ("Yo'lda", delivery in {"Yo'lda", "Yetkazildi"}),
        ("Yetkazildi", delivery == "Yetkazildi"),
    ]
    result = []
    current_used = False
    for name, done in flags:
        state = "done" if done else ""
        if not done and not current_used:
            state = "current"
            current_used = True
        result.append({"name": name, "state": state})
    return result

def save_job(meta: dict[str, Any], parts: list[Part], plans: list[SheetPlan]) -> tuple[int, str]:
    token = secrets.token_urlsafe(24)
    now = now_iso()
    conn = get_db()
    cur = conn.execute(
        """INSERT INTO kroy_jobs(
               order_code,customer,material,sheet_length,sheet_width,kerf,trim,worker_name,token,status,created_at,sent_at,
               optimization_mode,tested_variants,best_variant,worker_link_active,worker_token_created_at
           ) VALUES(?,?,?,?,?,?,?,?,?,'Ishchiga yuborildi',?,?,?,?,?,1,?)""",
        (
            meta["order_code"], meta.get("customer", ""), meta["material"], meta["sheet_length"], meta["sheet_width"],
            meta["kerf"], meta["trim"], meta.get("worker_name", ""), token, now, now,
            meta.get("optimization_mode", "large"), int(meta.get("tested_variants", 0)), meta.get("best_variant", ""), now,
        ),
    )
    job_id = int(cur.lastrowid)
    customer_token = secrets.token_urlsafe(24)
    conn.execute(
        "UPDATE kroy_jobs SET customer_token=?,customer_link_active=1,delivery_status='Kutilmoqda' WHERE id=?",
        (customer_token, job_id),
    )
    _add_customer_update(
        conn, job_id, "order-created", "Buyurtma qabul qilindi",
        f"{meta['order_code']} buyurtmasi tizimga kiritildi va konstruktor ishiga yuborildi.",
        "Buyurtma", 0,
    )
    conn.executemany(
        """INSERT INTO kroy_parts(job_id,uid,name,length,width,qty,rotate,edge_left,edge_right,edge_top,edge_bottom)
           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
        [
            (
                job_id, p.uid, p.name, p.length, p.width, p.qty, int(p.rotate), int(p.edge_left), int(p.edge_right),
                int(p.edge_top), int(p.edge_bottom),
            )
            for p in parts
        ],
    )
    conn.executemany(
        "INSERT INTO kroy_sheets(job_id,sheet_no,plan_json) VALUES(?,?,?)",
        [(job_id, plan.number, json.dumps(plan_to_dict(plan), ensure_ascii=False)) for plan in plans],
    )
    conn.commit()
    conn.close()
    return job_id, token


def load_job(job_id: int | None = None, token: str | None = None) -> tuple[sqlite3.Row, list[sqlite3.Row], list[tuple[sqlite3.Row, SheetPlan]]]:
    conn = get_db()
    if job_id is not None:
        job = conn.execute("SELECT * FROM kroy_jobs WHERE id=?", (job_id,)).fetchone()
    else:
        job = conn.execute("SELECT * FROM kroy_jobs WHERE token=?", (token,)).fetchone()
        if job and not int(job["worker_link_active"] or 0):
            conn.close()
            abort(410)
    if not job:
        conn.close()
        abort(404)
    parts = conn.execute("SELECT * FROM kroy_parts WHERE job_id=? ORDER BY id", (job["id"],)).fetchall()
    sheet_rows = conn.execute("SELECT * FROM kroy_sheets WHERE job_id=? ORDER BY sheet_no", (job["id"],)).fetchall()
    sheets = [(row, dict_to_plan(json.loads(row["plan_json"]))) for row in sheet_rows]
    conn.close()
    return job, parts, sheets


def svg_for_plan(plan: SheetPlan, compact: bool = False) -> str:
    view_w, view_h = plan.width, plan.height
    font = 34 if compact else 42
    elements = [f'<svg class="sheet-svg" viewBox="0 0 {view_w} {view_h}" xmlns="http://www.w3.org/2000/svg" aria-label="Kroy list {plan.number}">']
    elements.append(f'<rect x="0" y="0" width="{view_w}" height="{view_h}" fill="#ffffff" stroke="#0f172a" stroke-width="10"/>')
    largest = largest_offcut(plan)
    for r in plan.leftovers:
        is_largest = largest is not None and (r.x, r.y, r.w, r.h) == (largest.x, largest.y, largest.w, largest.h)
        fill = "#fb7185" if is_largest else "#fecdd3"
        opacity = "0.72" if is_largest else "0.48"
        elements.append(f'<rect x="{r.x}" y="{r.y}" width="{r.w}" height="{r.h}" fill="{fill}" fill-opacity="{opacity}" stroke="#be123c" stroke-width="5"/>')
        if r.w >= 240 and r.h >= 120:
            fs = max(20, min(38, int(min(r.w, r.h) * .13)))
            elements.append(f'<text x="{r.x+r.w/2}" y="{r.y+r.h/2}" text-anchor="middle" font-size="{fs}" font-weight="700" fill="#881337">Qoldiq {r.w}×{r.h}</text>')
    for idx, p in enumerate(plan.placements, 1):
        fill = "#eff6ff" if idx % 2 else "#f8fafc"
        elements.append(f'<rect x="{p.x}" y="{p.y}" width="{p.w}" height="{p.h}" fill="{fill}" stroke="#475569" stroke-width="5"/>')
        min_dim = min(p.w, p.h)
        label_font = max(22, min(font, int(min_dim * .18)))
        cx, cy = p.x + p.w / 2, p.y + p.h / 2

        # Kromka belgisi detalning eng tashqi chegarasida emas,
        # markazdagi nom/o'lcham yozuvi atrofida qisqa chiziq bilan ko'rsatiladi.
        # Shu sabab yonma-yon detallar va list tashqi chegarasi bilan aralashmaydi.
        edge_stroke = "#dc2626"
        sw = max(9, min(17, int(min_dim * .055)))
        # Belgilar markazdagi o'lcham yozuvining to'rt tomonida turadi.
        dims_y = cy + label_font * .35
        marker_cy = dims_y - label_font * .28
        marker_half_w = max(18, min(p.w * .34, label_font * 3.8))
        marker_half_h = max(12, min(p.h * .24, label_font * .62))
        h_len = max(16, min(p.w * .26, label_font * 2.5))
        v_len = max(14, min(p.h * .25, label_font * 1.35))
        if p.edge_top:
            y_mark = marker_cy - marker_half_h
            elements.append(f'<line x1="{cx-h_len/2}" y1="{y_mark}" x2="{cx+h_len/2}" y2="{y_mark}" stroke="{edge_stroke}" stroke-width="{sw}" stroke-linecap="round"/>')
        if p.edge_bottom:
            y_mark = marker_cy + marker_half_h
            elements.append(f'<line x1="{cx-h_len/2}" y1="{y_mark}" x2="{cx+h_len/2}" y2="{y_mark}" stroke="{edge_stroke}" stroke-width="{sw}" stroke-linecap="round"/>')
        if p.edge_left:
            x_mark = cx - marker_half_w
            elements.append(f'<line x1="{x_mark}" y1="{marker_cy-v_len/2}" x2="{x_mark}" y2="{marker_cy+v_len/2}" stroke="{edge_stroke}" stroke-width="{sw}" stroke-linecap="round"/>')
        if p.edge_right:
            x_mark = cx + marker_half_w
            elements.append(f'<line x1="{x_mark}" y1="{marker_cy-v_len/2}" x2="{x_mark}" y2="{marker_cy+v_len/2}" stroke="{edge_stroke}" stroke-width="{sw}" stroke-linecap="round"/>')
        name = p.name[:16].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        elements.append(f'<text x="{cx}" y="{cy-label_font*.95}" text-anchor="middle" font-size="{label_font}" font-weight="700" fill="#0f172a">#{idx} {name}</text>')
        elements.append(f'<text x="{cx}" y="{dims_y}" text-anchor="middle" font-size="{max(20,label_font-7)}" fill="#334155">{p.original_length}×{p.original_width}{" ↻" if p.rotated else ""}</text>')
        if not p.rotate_allowed and min_dim > 130:
            elements.append(f'<text x="{cx}" y="{cy+label_font*1.35}" text-anchor="middle" font-size="{max(18,label_font-10)}" font-weight="700" fill="#7c3aed">GUL ↕</text>')
    elements.append("</svg>")
    return "".join(elements)


def qr_data_uri(url: str) -> str:
    if qrcode is None:
        return ""
    img = qrcode.make(url)
    out = io.BytesIO()
    img.save(out, format="PNG")
    return "data:image/png;base64," + base64.b64encode(out.getvalue()).decode("ascii")


def _fit_text(c: canvas.Canvas, text: str, max_width: float, start_size: float, min_size: float = 5) -> float:
    size = start_size
    while size > min_size and stringWidth(pdf_text(text), PDF_FONT, size) > max_width:
        size -= 0.5
    return size


def draw_plan_pdf(c: canvas.Canvas, plan: SheetPlan, job: sqlite3.Row, x: float, y: float, w: float, h: float) -> None:
    c.saveState()
    c.setStrokeColor(colors.HexColor("#0f172a"))
    c.setLineWidth(0.8)
    c.rect(x, y, w, h)
    header_h = 43
    footer_h = 24
    c.setFont(PDF_FONT_BOLD, 9)
    c.drawString(x + 7, y + h - 13, pdf_text(f"Mebel360 | {job['order_code']} | {job['material']}"))
    c.setFont(PDF_FONT, 7)
    c.drawString(x + 7, y + h - 24, pdf_text(f"List {plan.number}: {plan.width}x{plan.height} mm | Ishchi: {job['worker_name'] or '-'} | Foydalanish: {sheet_usage(plan)}%"))
    c.drawRightString(x + w - 7, y + h - 24, pdf_text(f"Kromka x1.1: {plan_edge_m(plan):.2f} m"))
    c.drawString(
        x + 7, y + h - 35,
        pdf_text(f"Hisoblash: {optimization_label(job['optimization_mode'])} | Tekshirildi: {job['tested_variants']} variant")
    )

    draw_x = x + 8
    draw_y = y + footer_h + 5
    draw_w = w - 16
    draw_h = h - header_h - footer_h - 10
    scale = min(draw_w / plan.width, draw_h / plan.height)
    actual_w = plan.width * scale
    actual_h = plan.height * scale
    ox = draw_x + (draw_w - actual_w) / 2
    oy = draw_y + (draw_h - actual_h) / 2

    c.setFillColor(colors.white)
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.7)
    c.rect(ox, oy, actual_w, actual_h, fill=1, stroke=1)

    largest = largest_offcut(plan)
    for r in plan.leftovers:
        rx = ox + r.x * scale
        ry = oy + actual_h - (r.y + r.h) * scale
        rw, rh = r.w * scale, r.h * scale
        is_largest = largest is not None and (r.x, r.y, r.w, r.h) == (largest.x, largest.y, largest.w, largest.h)
        c.setFillColor(colors.HexColor("#fb7185" if is_largest else "#fecdd3"))
        c.setStrokeColor(colors.HexColor("#be123c"))
        c.setLineWidth(0.35)
        c.rect(rx, ry, rw, rh, fill=1, stroke=1)
        label = f"Qoldiq {r.w}x{r.h}"
        if rw > 55 and rh > 14:
            c.setFillColor(colors.HexColor("#881337"))
            label_size = _fit_text(c, label, rw - 6, min(6.5, max(4, rh / 5)), 3.5)
            if stringWidth(pdf_text(label), PDF_FONT_BOLD, label_size) <= rw - 6:
                c.setFont(PDF_FONT_BOLD, label_size)
                c.drawCentredString(rx + rw / 2, ry + rh / 2, pdf_text(label))
            elif rh > 70 and rw > 14:
                rotated_size = _fit_text(c, label, rh - 6, min(6.5, max(4, rw / 4)), 3.5)
                c.saveState()
                c.translate(rx + rw / 2, ry + rh / 2)
                c.rotate(90)
                c.setFont(PDF_FONT_BOLD, rotated_size)
                c.drawCentredString(0, -rotated_size / 3, pdf_text(label))
                c.restoreState()

    for idx, p in enumerate(plan.placements, 1):
        px = ox + p.x * scale
        # SVGda y pastga, PDFda y yuqoriga; joylashni ag'daramiz
        py = oy + actual_h - (p.y + p.h) * scale
        pw, ph = p.w * scale, p.h * scale
        c.setFillColor(colors.HexColor("#f8fafc") if idx % 2 else colors.HexColor("#eef2ff"))
        c.setStrokeColor(colors.HexColor("#64748b"))
        c.setLineWidth(0.35)
        c.rect(px, py, pw, ph, fill=1, stroke=1)

        label = f"#{idx} {p.name[:12]}"
        dims = f"{p.original_length}x{p.original_width}{' R' if p.rotated else ''}{' GUL' if not p.rotate_allowed else ''}"
        edge = f"K:{edge_letters(p)}" if any((p.edge_left,p.edge_right,p.edge_top,p.edge_bottom)) else ""
        font_size = _fit_text(c, label, max(8, pw - 3), min(7.5, max(4.5, ph / 4)), 3.8)
        c.setFillColor(colors.black)
        c.setFont(PDF_FONT_BOLD, font_size)
        label_y = py + ph / 2 + font_size * .80
        c.drawCentredString(px + pw / 2, label_y, pdf_text(label))
        dim_font = max(3.5, font_size - 0.8)
        dims_y = py + ph / 2 - font_size * .15
        c.setFont(PDF_FONT, dim_font)
        c.drawCentredString(px + pw / 2, dims_y, pdf_text(dims))

        # A4da ham kromka detal tashqi chetida emas, o'lcham yozuvi atrofida turadi.
        if any((p.edge_left, p.edge_right, p.edge_top, p.edge_bottom)):
            cx_pdf = px + pw / 2
            marker_cy = dims_y + dim_font * .25
            dims_width = stringWidth(pdf_text(dims), PDF_FONT, dim_font)
            marker_half_w = max(5, min(pw * .34, dims_width / 2 + 6))
            marker_half_h = max(3.5, min(ph * .22, dim_font * .75))
            h_len = max(5, min(pw * .25, max(8, dims_width * .48)))
            v_len = max(4, min(ph * .22, font_size * 1.25))
            c.setStrokeColor(colors.HexColor("#dc2626"))
            c.setLineWidth(1.8)
            if p.edge_top:
                y_mark = marker_cy + marker_half_h
                c.line(cx_pdf - h_len / 2, y_mark, cx_pdf + h_len / 2, y_mark)
            if p.edge_bottom:
                y_mark = marker_cy - marker_half_h
                c.line(cx_pdf - h_len / 2, y_mark, cx_pdf + h_len / 2, y_mark)
            if p.edge_left:
                x_mark = cx_pdf - marker_half_w
                c.line(x_mark, marker_cy - v_len / 2, x_mark, marker_cy + v_len / 2)
            if p.edge_right:
                x_mark = cx_pdf + marker_half_w
                c.line(x_mark, marker_cy - v_len / 2, x_mark, marker_cy + v_len / 2)
            c.setStrokeColor(colors.black)

    c.setFillColor(colors.black)
    footer = "Kromka x1.1 | T=tepa, P=past, Ch=chap, O=o'ng | GUL=aylantirilmaydi | Qoldiq=pushti"
    footer_size = _fit_text(c, footer, w - 14, 6.5, 4.2)
    c.setFont(PDF_FONT, footer_size)
    c.drawString(x + 7, y + 8, pdf_text(footer))
    c.restoreState()


def build_pdf(job: sqlite3.Row, sheets: list[tuple[sqlite3.Row, SheetPlan]]) -> io.BytesIO:
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    page_w, page_h = A4
    margin = 18
    gap = 10
    slot_h = (page_h - 2 * margin - gap) / 2
    slot_w = page_w - 2 * margin
    for index, (_, plan) in enumerate(sheets):
        slot = index % 2
        if index and slot == 0:
            c.showPage()
        y = page_h - margin - slot_h if slot == 0 else margin
        draw_plan_pdf(c, plan, job, margin, y, slot_w, slot_h)
    c.save()
    out.seek(0)
    return out


BASE_TEMPLATE = r"""
<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{ title }}</title><style>""" + BASE_CSS + r"""</style></head>
<body><header><div class="top"><div><div class="brand"><span class="brand-mark">360°</span>Mebel360° <span class="module-pill">BOSHQARUV PRO V9</span></div><div class="sub">Konstruktor · menejer · ishchi · shafyor · mijoz kuzatuvi</div></div>
{% if session.get('admin_user_id') %}<div class="navbar no-print"><a class="navlink" href="{{ url_for('dashboard') }}">Boshqaruv</a>{% if session.get('user_role') in ['admin','constructor'] %}<a class="navlink" href="{{ url_for('constructor') }}">Konstruktor</a>{% endif %}{% if session.get('user_role') in ['admin','manager'] %}<a class="navlink" href="{{ url_for('manager_dashboard') }}">Menejer</a>{% endif %}{% if session.get('user_role') in ['admin','manager','constructor','worker'] %}<a class="navlink" href="{{ url_for('worker_center') }}">Ishchi</a>{% endif %}{% if session.get('user_role') in ['admin','manager','driver'] %}<a class="navlink" href="{{ url_for('driver_dashboard') }}">Shafyor</a>{% endif %}{% if session.get('user_role')=='admin' %}<a class="navlink" href="{{ url_for('users') }}">Xodimlar</a>{% endif %}<span class="role-chip">{{ role_labels.get(session.get('user_role','admin'),'Rahbar') }} · {{ session.get('admin_username','') }}</span><form method="post" action="{{ url_for('logout') }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><button class="btn2" type="submit">Chiqish</button></form></div>{% else %}<div class="top-actions no-print"><a class="btn light" href="{{ url_for('login') }}">Kirish</a></div>{% endif %}</div></header>
<div class="wrap">{% with msgs=get_flashed_messages(with_categories=true) %}{% for cat,msg in msgs %}<div class="flash {{ 'bad' if cat=='bad' else '' }}">{{ msg }}</div>{% endfor %}{% endwith %}{{ body|safe }}</div></body></html>
"""


AUTH_BODY = r"""
<div class="auth-wrap">
  <div class="panel">
    {% if setup %}
      <h2>Birinchi xavfsiz sozlash</h2>
      <div class="auth-note">Rahbar paroli TXT faylga yozilmaydi va CMD oynasida ko'rsatilmaydi. U bazada faqat shifrlangan holatda saqlanadi.</div>
      <form method="post" action="{{ url_for('setup') }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}">
        <label>Rahbar login<input name="username" value="admin" minlength="3" maxlength="40" autocomplete="username" required></label>
        <label>Yangi parol<input name="password" type="password" minlength="8" autocomplete="new-password" required></label>
        <label>Parolni qayta yozing<input name="password_confirm" type="password" minlength="8" autocomplete="new-password" required></label>
        <button class="ok" type="submit">Rahbarni yaratish</button>
      </form>
    {% else %}
      <h2>Mebel360° tizimiga kirish</h2>
      <div class="auth-note">Rahbar, konstruktor, menejer, ishchi va shafyor o'z login-paroli bilan kiradi. 5 marta noto'g'ri urinishdan keyin kirish 15 daqiqaga bloklanadi.</div>
      <form method="post" action="{{ url_for('login') }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}">
        <label>Login<input name="username" autocomplete="username" required autofocus></label>
        <label>Parol<input name="password" type="password" autocomplete="current-password" required></label>
        <button class="ok" type="submit">Tizimga kirish</button>
      </form>
    {% endif %}
  </div>
</div>
"""


CONSTRUCTOR_BODY = r"""
<div class="panel hero">
  <div class="hero-grid"><div><h1>Konstruktor boshqaruv markazi</h1><p>Detalni kiriting yoki fayldan yuklang. Dastur eng yaxshi kroy variantini tanlaydi, kromkani hisoblaydi va ishchiga xavfsiz topshiriq yuboradi.</p></div><div class="hero-actions no-print">{% if session.get('user_role')=='admin' %}<a class="btn" href="{{ url_for('download_backup') }}">Baza nusxasini yuklash</a>{% endif %}<span class="btn2 btn">Avto-zaxira: faol</span></div></div>
</div>
<div class="grid">
  <div>
    <div class="panel">
      <h3 class="step-title"><span class="step-no">1</span>Buyurtma va list</h3><div class="section-note">Asosiy ma’lumotlar va PRO100 chizmasini biriktiring.</div>
      <form method="post" action="{{ url_for('job_create') }}" id="jobForm" enctype="multipart/form-data"><input type="hidden" name="_csrf" value="{{ csrf_token() }}">
        <label>Buyurtma kodi<input name="order_code" required value="{{ draft.order_code }}" placeholder="AD-001"></label>
        <label>Mijoz<input name="customer" value="{{ draft.customer }}" placeholder="Akmal aka"></label>
        <label>Material / rang<input name="material" required value="{{ draft.material }}" placeholder="LMDEF oq 16 mm"></label>
        <div class="row"><label>List uzunligi, mm<input type="number" name="sheet_length" required value="{{ draft.sheet_length }}"></label><label>List eni, mm<input type="number" name="sheet_width" required value="{{ draft.sheet_width }}"></label></div>
        <div class="row"><label>Arraning izi, mm<input type="number" name="kerf" required value="{{ draft.kerf }}"></label><label>Chet kesimi, mm<input type="number" name="trim" required value="{{ draft.trim }}"></label></div>
        <div class="opt-box"><label>Kroy hisoblash usuli<select name="optimization_mode"><option value="large" {{ 'selected' if draft.optimization_mode=='large' else '' }}>Chuqur hisob - 240 variantgacha</option><option value="full" {{ 'selected' if draft.optimization_mode=='full' else '' }}>Standart - 120 variantgacha</option><option value="fast" {{ 'selected' if draft.optimization_mode=='fast' else '' }}>Tezkor - 60 variantgacha</option></select></label><div class="variant-note">Tanlangan usul topshiriqda, ishchi oynasida va A4 PDFda yoziladi. Dastur tekshirilgan variantlar ichidan listi kam va katta qoldiq qoladigan taxlashni tanlaydi.</div></div>
        <label>Ishchi<input name="worker_name" value="{{ draft.worker_name }}" placeholder="Kesuvchi ishchi"></label>
        <label>PRO100 chizmasi — ixtiyoriy `.STO`<input type="file" name="sto_file" accept=".sto"></label><div class="muted">STO fayl kroy topshirig‘iga biriktiriladi va keyin yuklab olish mumkin.</div>
        <input type="hidden" name="parts_json" id="partsJson">
        <button class="ok" type="submit" onclick="return prepareSubmit()">Kroy qil va ishchiga yubor</button>
      </form>
    </div>
    <div class="panel">
      <h3 class="step-title"><span class="step-no">2</span>Fayldan detal yuklash</h3><div class="section-note">DBF, CSV yoki TXT faylidan detal ro‘yxatini bir bosishda oling.</div>
      <div class="import-grid">
        <form class="file-box" method="post" action="{{ url_for('dbf_import') }}" enctype="multipart/form-data"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><h4>2D-PLACE DBF</h4><input type="file" name="dbf" accept=".dbf" required><button type="submit" class="btn2">DBF yuklash</button></form>
        <form class="file-box green" method="post" action="{{ url_for('table_import') }}" enctype="multipart/form-data"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><h4>CSV / TXT jadval</h4><input type="file" name="table_file" accept=".csv,.txt" required><button type="submit" class="ok">CSV/TXT yuklash</button></form>
      </div><p class="muted">Jadval ustunlari: detal nomi, uzunlik, en, soni, aylantirish va kromka. O‘zbek, rus yoki inglizcha sarlavhalar qabul qilinadi.</p>
    </div>
  </div>
  <div>
    <div class="panel">
      <h3 class="step-title"><span class="step-no">3</span>Detallar va kromka</h3><div class="section-note">O‘lcham 10–10000 mm, jami 2500 tagacha detal. Noto‘g‘ri qiymat serverda ham to‘xtatiladi.</div>
      <div class="row"><label>Detal nomi<input id="pName" placeholder="Bokovina"></label><label>Soni<input id="pQty" type="number" min="1" max="999" value="1" oninput="renderEntryEdges()"></label></div>
      <div class="row"><label>Uzunligi, mm<input id="pLength" type="number" min="10" max="10000" placeholder="2100" oninput="renderEntryEdges()"></label><label>Eni, mm<input id="pWidth" type="number" min="10" max="10000" placeholder="300" oninput="renderEntryEdges()"></label></div>
      <div class="edge-help"><b>Kromkani chiziqdan belgilang:</b> uzunlik ostidagi 2 chiziq — uzun ikki tomon, en ostidagi 2 chiziq — kalta ikki tomon. Bosilgan qizil chiziq kromka bo'ladi.</div>
      <div class="edge-pickers">
        <div class="edge-picker"><div class="edge-picker-title"><span class="edge-dim" id="entryLengthDim">2100 mm</span><span class="edge-side-name">uzun tomonlar</span></div><div class="edge-lines"><button id="entryTop" class="edge-toggle" type="button" onclick="toggleEntryEdge('top')" title="Uzun tomon 1"><span class="edge-stroke"></span></button><button id="entryBottom" class="edge-toggle" type="button" onclick="toggleEntryEdge('bottom')" title="Uzun tomon 2"><span class="edge-stroke"></span></button></div></div>
        <div class="edge-picker"><div class="edge-picker-title"><span class="edge-dim" id="entryWidthDim">300 mm</span><span class="edge-side-name">kalta tomonlar</span></div><div class="edge-lines"><button id="entryLeft" class="edge-toggle" type="button" onclick="toggleEntryEdge('left')" title="Kalta tomon 1"><span class="edge-stroke"></span></button><button id="entryRight" class="edge-toggle" type="button" onclick="toggleEntryEdge('right')" title="Kalta tomon 2"><span class="edge-stroke"></span></button></div></div>
      </div>
      <div class="meter-box"><div><b>Bu detal kromkasi ×1.1</b><br><small>Tanlangan chiziqlar, soni va 10% zaxira bilan</small></div><strong id="entryEdgeMeters">0.00 m</strong></div>
      <div class="grain-box"><div><div class="grain-title">Gul / tekstura yo'nalishi</div><div class="grain-note" id="rotateExplain">Detalni 90° aylantirib joylash mumkin.</div></div><label class="grain-toggle"><input id="pRotate" type="checkbox" checked onchange="renderRotateStatus()"><span id="rotateStatus" class="grain-badge">Aylantirish mumkin</span></label></div>
      <div class="actions"><button type="button" onclick="addPart()">Detal qo'shish</button><button type="button" class="light" onclick="addSample()">Namuna qo'shish</button><button type="button" class="danger" onclick="clearParts()">Tozalash</button></div>
      <div class="tablewrap desktop"><table><thead><tr><th>#</th><th>Detal</th><th>Uzunlik</th><th>En</th><th>Soni</th><th>Gul yo'nalishi</th><th>Kromka</th><th>Kromka metri</th><th></th></tr></thead><tbody id="partsBody"></tbody></table></div>
      <div class="part-cards" id="partCards"></div>
      <p class="muted" id="partSummary"></p>
      <div class="meter-total">Jami kromka ×1.1 (10% zaxira bilan): <span id="totalEdgeMeters">0.00 m</span></div>
    </div>
    <div class="panel">
      <h3 class="step-title"><span class="step-no">4</span>Oldingi topshiriqlar</h3>
      {% if jobs %}<div class="tablewrap"><table><thead><tr><th>Kod</th><th>Material</th><th>Ishchi</th><th>Holat</th><th></th></tr></thead><tbody>{% for j in jobs %}<tr><td>{{ j.order_code }}</td><td>{{ j.material }}</td><td>{{ j.worker_name }}</td><td><span class="badge">{{ j.status }}</span></td><td><a class="btn light" href="{{ url_for('job_view',job_id=j.id) }}">Ochish</a></td></tr>{% endfor %}</tbody></table></div>{% else %}<div class="empty">Hozircha topshiriq yo'q</div>{% endif %}
    </div>
  </div>
</div>
<script>
let parts={{ parts_json|safe }};
let entryEdges={top:false,bottom:false,left:false,right:false};
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));}
function edges(p){let a=[];if(p.edge_top)a.push('U1');if(p.edge_bottom)a.push('U2');if(p.edge_left)a.push('E1');if(p.edge_right)a.push('E2');return a.length?a.join(', '):"yo'q";}
const EDGE_MULTIPLIER=1.1;
function edgeMeters(p){const q=Math.max(1,Number(p.qty||1)),l=Math.max(0,Number(p.length||0)),w=Math.max(0,Number(p.width||0));const mm=(p.edge_top?l:0)+(p.edge_bottom?l:0)+(p.edge_left?w:0)+(p.edge_right?w:0);return mm*q/1000*EDGE_MULTIPLIER;}
function totalEdgeMeters(){return parts.reduce((s,p)=>s+edgeMeters(p),0);}
function meterText(v){return `${Number(v||0).toFixed(2)} m`;}
function grainText(p){return p.rotate?'Aylantirish mumkin':"Aylantirib bo'lmaydi";}
function rotateButton(index,p){return `<button type="button" class="rotate-action ${p.rotate?'':'locked'}" onclick="togglePartRotate(${index})">${grainText(p)}</button>`;}
function edgeButton(index,key,active,title){return `<button class="edge-toggle ${active?'active':''}" type="button" onclick="togglePartEdge(${index},'${key}')" title="${title}" aria-pressed="${active?'true':'false'}"><span class="edge-stroke"></span></button>`;}
function dimensionPicker(index,value,kind,p){
 const isLength=kind==='length';
 const key1=isLength?'top':'left',key2=isLength?'bottom':'right';
 const active1=isLength?p.edge_top:p.edge_left,active2=isLength?p.edge_bottom:p.edge_right;
 const title1=isLength?'Uzun tomon 1':'Kalta tomon 1',title2=isLength?'Uzun tomon 2':'Kalta tomon 2';
 return `<div class="dim-cell"><div class="dim-value">${value} mm</div><div class="table-edge-lines">${edgeButton(index,key1,active1,title1)}${edgeButton(index,key2,active2,title2)}</div></div>`;
}
function cardPicker(index,p){return `<div class="edge-pickers"><div class="edge-picker"><div class="edge-picker-title"><span class="edge-dim">${p.length} mm</span><span class="edge-side-name">uzun tomonlar</span></div><div class="edge-lines">${edgeButton(index,'top',p.edge_top,'Uzun tomon 1')}${edgeButton(index,'bottom',p.edge_bottom,'Uzun tomon 2')}</div></div><div class="edge-picker"><div class="edge-picker-title"><span class="edge-dim">${p.width} mm</span><span class="edge-side-name">kalta tomonlar</span></div><div class="edge-lines">${edgeButton(index,'left',p.edge_left,'Kalta tomon 1')}${edgeButton(index,'right',p.edge_right,'Kalta tomon 2')}</div></div></div>`;}
function renderParts(){
 const body=document.getElementById('partsBody'),cards=document.getElementById('partCards');body.innerHTML='';cards.innerHTML='';let total=0;
 parts.forEach((p,i)=>{total+=Number(p.qty||1);body.innerHTML+=`<tr><td>${i+1}</td><td>${esc(p.name)}</td><td>${dimensionPicker(i,p.length,'length',p)}</td><td>${dimensionPicker(i,p.width,'width',p)}</td><td>${p.qty}</td><td class="grain-cell">${rotateButton(i,p)}</td><td class="edge-list">${edges(p)}</td><td class="meter-cell">${meterText(edgeMeters(p))}</td><td><button class="danger" type="button" onclick="removePart(${i})">×</button></td></tr>`;cards.innerHTML+=`<div class="part-card"><b>${i+1}. ${esc(p.name)}</b><div>${p.qty} dona · <span class="meter-cell">${meterText(edgeMeters(p))} kromka</span></div><div style="margin-top:7px">${rotateButton(i,p)}</div>${cardPicker(i,p)}<div class="edge-list">Kromka: ${edges(p)}</div><button class="danger" type="button" onclick="removePart(${i})">O'chirish</button></div>`;});
 document.getElementById('partSummary').textContent=`${parts.length} tur detal, jami ${total} dona`;
 document.getElementById('totalEdgeMeters').textContent=meterText(totalEdgeMeters());
}
function toggleEntryEdge(key){entryEdges[key]=!entryEdges[key];renderEntryEdges();}
function renderRotateStatus(){
 const rot=document.getElementById('pRotate'),status=document.getElementById('rotateStatus'),note=document.getElementById('rotateExplain');if(!rot||!status||!note)return;
 status.textContent=rot.checked?'Aylantirish mumkin':"Aylantirib bo'lmaydi";status.classList.toggle('locked',!rot.checked);
 note.textContent=rot.checked?"Detalni 90° aylantirib joylash mumkin.":"Gul yo'nalishi saqlanadi, detal 90° aylantirilmaydi.";
}
function renderEntryEdges(){
 const lenValue=Number(document.getElementById('pLength').value||0),widValue=Number(document.getElementById('pWidth').value||0),qtyValue=Math.max(1,Number(document.getElementById('pQty').value||1));
 const len=document.getElementById('pLength').value||document.getElementById('pLength').placeholder||'Uzunlik';
 const wid=document.getElementById('pWidth').value||document.getElementById('pWidth').placeholder||'En';
 document.getElementById('entryLengthDim').textContent=`${len} mm`;document.getElementById('entryWidthDim').textContent=`${wid} mm`;
 const ids={top:'entryTop',bottom:'entryBottom',left:'entryLeft',right:'entryRight'};
 Object.keys(ids).forEach(k=>{const b=document.getElementById(ids[k]);b.classList.toggle('active',entryEdges[k]);b.setAttribute('aria-pressed',entryEdges[k]?'true':'false');});
 document.getElementById('entryEdgeMeters').textContent=meterText(edgeMeters({length:lenValue,width:widValue,qty:qtyValue,edge_top:entryEdges.top,edge_bottom:entryEdges.bottom,edge_left:entryEdges.left,edge_right:entryEdges.right}));
}
function togglePartEdge(i,key){const p=parts[i];if(!p)return;const map={top:'edge_top',bottom:'edge_bottom',left:'edge_left',right:'edge_right'};p[map[key]]=!p[map[key]];renderParts();}
function togglePartRotate(i){const p=parts[i];if(!p)return;p.rotate=!p.rotate;renderParts();}
function addPart(){
 const nameEl=document.getElementById('pName'),lenEl=document.getElementById('pLength'),widEl=document.getElementById('pWidth'),qtyEl=document.getElementById('pQty');
 const rotEl=document.getElementById('pRotate');
 let name=nameEl.value.trim(),length=Number(lenEl.value),width=Number(widEl.value),qty=Math.floor(Number(qtyEl.value||1));if(!name||length<10||length>10000||width<10||width>10000||qty<1||qty>999){alert("Detal nomini yozing. O'lcham 10–10000 mm, soni 1–999 oralig'ida bo'lsin.");return;}if(parts.length>=500||parts.reduce((s,p)=>s+Number(p.qty||1),0)+qty>2500){alert('Limit: 500 tur va jami 2500 ta detal.');return;}parts.push({uid:Math.random().toString(36).slice(2),name,length,width,qty,rotate:rotEl.checked,edge_left:entryEdges.left,edge_right:entryEdges.right,edge_top:entryEdges.top,edge_bottom:entryEdges.bottom});nameEl.value='';lenEl.value='';widEl.value='';qtyEl.value=1;entryEdges={top:false,bottom:false,left:false,right:false};renderEntryEdges();renderParts();}
function addSample(){parts=[{uid:'a1',name:'800 bakavoy',length:2100,width:300,qty:2,rotate:false,edge_left:true,edge_right:false,edge_top:true,edge_bottom:true},{uid:'a2',name:'Polka',length:700,width:420,qty:5,rotate:true,edge_left:false,edge_right:false,edge_top:true,edge_bottom:false},{uid:'a3',name:'Niz',length:700,width:450,qty:2,rotate:true,edge_left:true,edge_right:true,edge_top:true,edge_bottom:false},{uid:'a4',name:'Planka',length:700,width:100,qty:4,rotate:true,edge_left:false,edge_right:false,edge_top:true,edge_bottom:false},{uid:'a5',name:'Fasad',length:716,width:397,qty:4,rotate:false,edge_left:true,edge_right:true,edge_top:true,edge_bottom:true}];renderParts();}
function removePart(i){parts.splice(i,1);renderParts();}function clearParts(){if(confirm('Hamma detallar o\'chirilsinmi?')){parts=[];renderParts();}}
function prepareSubmit(){const total=parts.reduce((s,p)=>s+Number(p.qty||1),0);if(!parts.length){alert('Kamida bitta detal kiriting');return false;}if(parts.length>500||total>2500){alert('Detal limiti oshgan: 500 tur / 2500 dona.');return false;}document.getElementById('partsJson').value=JSON.stringify(parts);return true;}renderRotateStatus();renderEntryEdges();renderParts();
</script>
"""


JOB_BODY = r"""
<div class="panel">
  <div class="worker-header"><div><h2>{{ job.order_code }} — {{ job.material }}</h2><div class="muted">Mijoz: {{ job.customer or '-' }} · Ishchi: {{ job.worker_name or '-' }} · {{ job.created_at }}</div></div><span class="badge {{ 'ready' if snapshot.progress==100 else 'progress' }}">{{ snapshot.progress }}% · {{ snapshot.stage }}</span></div>
  <div class="statusline"><span class="stat">{{ sheets|length }} ta list</span><span class="stat">{{ total_parts }} ta detal</span><span class="stat">Kromka ×1.1: {{ '%.2f'|format(edge_m) }} m</span><span class="stat">{{ opt_label }}</span><span class="stat">Holat: {{ job.status }}</span><span class="stat">Yetkazish: {{ job.delivery_status or 'Kutilmoqda' }}</span></div>
  <div class="actions no-print">
    {% if job.worker_link_active %}<a class="btn ok" href="{{ worker_url }}" target="_blank">Ishchi oynasi</a><button class="btn2" type="button" onclick="navigator.clipboard.writeText('{{ worker_url }}').then(()=>alert('Ishchi havolasi nusxalandi'))">Ishchi havolasini olish</button>{% if session.get('user_role') in ['admin','constructor','manager'] %}<form method="post" action="{{ url_for('worker_link_revoke',job_id=job.id) }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><button class="danger" type="submit" onclick="return confirm('Ishchi havolasi bekor qilinsinmi?')">Bekor qilish</button></form>{% endif %}{% else %}<span class="badge revoked">Ishchi havolasi bekor qilingan</span>{% if session.get('user_role') in ['admin','constructor','manager'] %}<form method="post" action="{{ url_for('worker_link_regenerate',job_id=job.id) }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><button class="ok" type="submit">Yangi havola</button></form>{% endif %}{% endif %}
    <a class="btn" href="{{ customer_url }}" target="_blank">Mijoz kuzatuvi</a><button class="btn2" type="button" onclick="navigator.clipboard.writeText('{{ customer_url }}').then(()=>alert('Mijoz havolasi nusxalandi'))">Mijozga yuborish</button>
    <a class="btn warn" href="{{ url_for('job_pdf',job_id=job.id) }}" target="_blank">A4 PDF</a>
  </div>
  <div class="split-grid" style="margin-top:12px">
    <div class="notice-preview"><b>Avtomatik bildirishnoma faol:</b> ishchi “Kroy kesildi” yoki “Kromka tayyor”ni bossa, mijozning kuzatuv sahifasida shu zahoti yangi xabar va foiz paydo bo'ladi.</div>
    <div class="copy-box"><input readonly value="{{ customer_url }}"><button type="button" onclick="navigator.clipboard.writeText('{{ customer_url }}').then(()=>alert('Nusxalandi'))">Nusxalash</button></div>
  </div>
  {% if attachments %}<div class="security-note"><b>Biriktirilgan chizma:</b>{% for a in attachments %}<div class="attachment"><div><div class="attachment-name">{{ a.original_name }}</div><div class="attachment-meta">{{ (a.size_bytes/1024)|round(1) }} KB · {{ a.created_at }}</div></div><a class="btn light" href="{{ url_for('attachment_download',attachment_id=a.id) }}">STO yuklash</a></div>{% endfor %}</div>{% endif %}
</div>
<div class="split-grid">
  <div><div class="sheet-grid">{% for row,plan,svg in sheets %}<div class="sheet-card"><div class="sheet-head"><div><div class="sheet-title">List {{ plan.number }} · {{ plan.width }}×{{ plan.height }} mm</div><div class="muted">{{ plan.placements|length }} detal · Foydalanish {{ usage(plan) }}% · <span class="offcut-stat">{{ offcut(plan) }}</span></div></div><span class="badge {{ 'ready' if row.cut_done and row.edge_done else ('progress' if row.cut_done or row.edge_done else '') }}">{{ 'Tayyor' if row.cut_done and row.edge_done else ('Jarayonda' if row.cut_done or row.edge_done else 'Kutilmoqda') }}</span></div>{{ svg|safe }}<div class="legend"><b>O'lcham atrofidagi qizil chiziq = kromka.</b> Pushti maydon = qoldiq. GUL ↕ = aylantirilmaydi.</div></div>{% endfor %}</div></div>
  <div>
    <div class="panel"><h3>Mijozga ko'ringan yangiliklar</h3>{% if updates %}<div class="timeline">{% for u in updates %}<div class="timeline-item"><div class="timeline-title">{{ u.title }}</div><div class="timeline-time">{{ u.created_at }} · {{ u.progress }}%</div><div class="timeline-message">{{ u.message }}</div></div>{% endfor %}</div>{% else %}<div class="empty">Yangilik yo'q</div>{% endif %}</div>
    {% if session.get('user_role') in ['admin','manager','constructor'] %}<div class="panel"><h3>Menejer xabari</h3><form method="post" action="{{ url_for('manager_customer_update',job_id=job.id) }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><label>Sarlavha<input name="title" placeholder="Masalan: Yig'ish boshlandi" required></label><label>Mijozga xabar<textarea name="message" placeholder="Qisqa va tushunarli xabar" required></textarea></label><label>Tayyorlik foizi<input name="progress" type="number" min="0" max="100" value="{{ snapshot.progress }}"></label><button class="ok" type="submit">Mijozga ko'rsatish</button></form></div>{% endif %}
  </div>
</div>
"""


WORKER_BODY = r"""
<div class="panel no-print">
  <div class="worker-header"><div><h2>{{ job.order_code }}</h2><div class="muted">{{ job.material }} · {{ job.worker_name or 'Ishchi' }} · {{ job.customer or '' }}</div></div><span class="badge {{ 'ready' if snapshot.progress>=80 else 'progress' }}">{{ snapshot.progress }}% · {{ snapshot.stage }}</span></div>
  <div class="notice-preview" style="margin-top:10px"><b>Mijozga avtomatik xabar:</b> “Kroy kesildi” yoki “Kromka tayyor”ni belgilab saqlashingiz bilan buyurtmachining telefonidagi kuzatuv sahifasi yangilanadi.</div>
  <div class="statusline"><span class="stat">List: {{ sheets|length }}</span><span class="stat">Detal: {{ total_parts }}</span><span class="stat">Jami kromka ×1.1: {{ '%.2f'|format(edge_m) }} m</span><span class="stat">Kroy usuli: {{ opt_label }}</span></div>
  <div class="actions"><button class="btn ok" type="button" onclick="window.print()">Printerdan chiqarish</button><a class="btn warn" href="{{ url_for('worker_pdf',token=job.token) }}" target="_blank">A4 PDF — 2 list/bet</a></div>
</div>
<div class="sheet-grid">{% for row,plan,svg in sheets %}<div class="sheet-card"><div class="sheet-head"><div><div class="sheet-title">{{ job.order_code }} · List {{ plan.number }} · {{ plan.width }}×{{ plan.height }}</div><div class="muted">{{ plan.placements|length }} detal · {{ usage(plan) }}% · <span class="offcut-stat">{{ offcut(plan) }}</span></div></div><span class="badge {{ 'ready' if row.cut_done and row.edge_done else ('progress' if row.cut_done else '') }}">{{ 'Kroy va kromka tayyor' if row.cut_done and row.edge_done else ('Kroy kesildi' if row.cut_done else 'Yangi') }}</span></div>{{ svg|safe }}
<div class="legend"><b>Qizil qisqa chiziq = kromka.</b> Pushti maydon = foydalanish mumkin bo'lgan qoldiq. GUL ↕ = aylantirilmaydi.</div>
<form class="no-print" method="post" action="{{ url_for('worker_update',token=job.token,sheet_no=plan.number) }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><div class="checks"><label class="check"><input name="cut_done" type="checkbox" {{ 'checked disabled' if row.cut_done else '' }}>Kroy kesildi</label><label class="check"><input name="edge_done" type="checkbox" {{ 'checked disabled' if row.edge_done else '' }}>Kromka tayyor</label></div>{% if row.cut_done %}<input type="hidden" name="cut_done" value="1">{% endif %}{% if row.edge_done %}<input type="hidden" name="edge_done" value="1">{% endif %}<label>Izoh<textarea name="note" placeholder="Kamchilik yoki izoh">{{ row.note }}</textarea></label><button class="ok" type="submit">Saqlash va mijozga bildirish</button></form></div>{% endfor %}</div>
"""


DASHBOARD_BODY = r"""
<div class="panel dashboard-hero"><h1>Mebel360° boshqaruv markazi</h1><p>Har bir xodim o'z bo'limida ishlaydi. Buyurtmadagi bajarilgan bosqichlar mijozga avtomatik ko'rinadi.</p></div>
<div class="kpi-grid"><div class="kpi"><b>{{ stats.total }}</b><span>Jami buyurtma</span></div><div class="kpi"><b>{{ stats.process }}</b><span>Jarayonda</span></div><div class="kpi"><b>{{ stats.ready }}</b><span>Kroy va kromka tayyor</span></div><div class="kpi"><b>{{ stats.delivered }}</b><span>Yetkazilgan</span></div></div>
<div class="role-grid">
{% if role in ['admin','constructor'] %}<a class="role-card" href="{{ url_for('constructor') }}"><div class="role-icon">📐</div><h3>Konstruktor</h3><p>Detal, kroy, kromka, PRO100 fayli va ishchi topshirig'i.</p><div class="role-arrow">Bo'limga kirish →</div></a>{% endif %}
{% if role in ['admin','manager'] %}<a class="role-card" href="{{ url_for('manager_dashboard') }}"><div class="role-icon">📋</div><h3>Menejer</h3><p>Buyurtmalar, tayyorlik foizi, mijoz havolasi va xabarlar.</p><div class="role-arrow">Bo'limga kirish →</div></a>{% endif %}
{% if role in ['admin','manager','constructor','worker'] %}<a class="role-card" href="{{ url_for('worker_center') }}"><div class="role-icon">🛠️</div><h3>Ishchi</h3><p>Kroy va kromka holatini belgilash, A4 topshiriq va izohlar.</p><div class="role-arrow">Bo'limga kirish →</div></a>{% endif %}
{% if role in ['admin','manager','driver'] %}<a class="role-card" href="{{ url_for('driver_dashboard') }}"><div class="role-icon">🚚</div><h3>Shafyor</h3><p>Yetkazishga tayyor, yo'lda va yetkazildi holatlari.</p><div class="role-arrow">Bo'limga kirish →</div></a>{% endif %}
</div>
<div class="panel"><h3>So'nggi buyurtmalar</h3>{% if jobs %}<div class="tablewrap"><table><thead><tr><th>Kod</th><th>Mijoz</th><th>Jarayon</th><th>Tayyorlik</th><th></th></tr></thead><tbody>{% for j in jobs %}<tr><td><b>{{ j.order_code }}</b></td><td>{{ j.customer or '-' }}</td><td>{{ j.stage }}</td><td><div class="progress-track"><div class="progress-fill" style="width:{{ j.progress }}%"></div></div><div class="progress-label">{{ j.progress }}%</div></td><td><a class="btn light" href="{{ url_for('job_view',job_id=j.id) }}">Ochish</a></td></tr>{% endfor %}</tbody></table></div>{% else %}<div class="empty">Hozircha buyurtma yo'q</div>{% endif %}</div>
"""

MANAGER_BODY = r"""
<div class="panel hero"><div class="hero-grid"><div><h1>Menejer boshqaruvi</h1><p>Mijoz buyurtmasi qayerga yetganini bitta oynada ko'ring va kuzatuv havolasini yuboring.</p></div><div class="hero-actions">{% if session.get('user_role') in ['admin','constructor'] %}<a class="btn" href="{{ url_for('constructor') }}">Yangi kroy</a>{% endif %}</div></div></div>
<div class="kpi-grid"><div class="kpi"><b>{{ stats.total }}</b><span>Buyurtmalar</span></div><div class="kpi"><b>{{ stats.process }}</b><span>Jarayonda</span></div><div class="kpi"><b>{{ stats.ready }}</b><span>Tayyor</span></div><div class="kpi"><b>{{ stats.delivered }}</b><span>Yetkazilgan</span></div></div>
<div class="panel"><h3>Buyurtmalar nazorati</h3>{% for j in jobs %}<div class="order-card"><div class="order-head"><div><div class="order-code">{{ j.order_code }}</div><div class="muted">{{ j.customer or 'Mijoz yozilmagan' }} · {{ j.material }}</div></div><span class="badge {{ 'ready' if j.progress==100 else 'progress' }}">{{ j.progress }}%</span></div><div class="progress-track" style="margin-top:10px"><div class="progress-fill" style="width:{{ j.progress }}%"></div></div><div class="progress-label">{{ j.stage }} · Yetkazish: {{ j.delivery_status }}</div><div class="actions"><a class="btn light" href="{{ url_for('job_view',job_id=j.id) }}">Buyurtmani ochish</a><a class="btn" target="_blank" href="{{ j.customer_url }}">Mijoz oynasi</a><button class="btn2" type="button" onclick="navigator.clipboard.writeText('{{ j.customer_url }}').then(()=>alert('Mijoz havolasi nusxalandi'))">Havolani nusxalash</button></div></div>{% else %}<div class="empty">Hozircha buyurtma yo'q</div>{% endfor %}</div>
"""

WORKER_CENTER_BODY = r"""
<div class="panel hero"><div class="hero-grid"><div><h1>Ishchi topshiriqlari</h1><p>Ishchi maxfiy havolani ochadi, kroy va kromka tugaganini belgilaydi. Mijozga xabar avtomatik tushadi.</p></div></div></div>
<div class="panel">{% for j in jobs %}<div class="order-card"><div class="order-head"><div><div class="order-code">{{ j.order_code }}</div><div class="muted">{{ j.worker_name or 'Ishchi belgilanmagan' }} · {{ j.material }} · {{ j.customer or '-' }}</div></div><span class="badge {{ 'ready' if j.progress>=80 else 'progress' }}">{{ j.progress }}%</span></div><div class="progress-track" style="margin-top:10px"><div class="progress-fill" style="width:{{ j.progress }}%"></div></div><div class="actions"><a class="btn ok" href="{{ j.worker_url }}" target="_blank">Topshiriqni ochish</a><button class="btn2" type="button" onclick="navigator.clipboard.writeText('{{ j.worker_url }}').then(()=>alert('Ishchi havolasi nusxalandi'))">Havolani olish</button><a class="btn light" href="{{ url_for('job_view',job_id=j.id) }}">Nazorat</a></div></div>{% else %}<div class="empty">Ishchi topshirig'i yo'q</div>{% endfor %}</div>
"""

DRIVER_BODY = r"""
<div class="panel hero"><div class="hero-grid"><div><h1>Shafyor boshqaruvi</h1><p>Yetkazish holatini yangilang. “Yo'lda” yoki “Yetkazildi” bosilishi bilan mijozning telefonida xabar chiqadi.</p></div></div></div>
<div class="panel">{% for j in jobs %}<div class="order-card"><div class="order-head"><div><div class="order-code">{{ j.order_code }}</div><div class="muted">{{ j.customer or '-' }} · {{ j.material }}</div></div><span class="delivery-status {{ 'road' if j.delivery_status=="Yo'lda" else ('done' if j.delivery_status=='Yetkazildi' else '') }}">{{ j.delivery_status }}</span></div><div class="progress-track" style="margin-top:10px"><div class="progress-fill" style="width:{{ j.progress }}%"></div></div><div class="progress-label">{{ j.progress }}% · {{ j.stage }}</div><form method="post" action="{{ url_for('driver_update',job_id=j.id) }}" style="margin-top:10px"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><label>Shafyor izohi<input name="note" value="{{ j.delivery_note or '' }}" placeholder="Masalan: 15:30 da yo'lga chiqildi"></label><div class="delivery-buttons">{% for st in delivery_statuses %}<button class="{{ 'ok' if st=='Yetkazildi' else ('warn' if st=="Yo'lda" else 'btn2') }}" type="submit" name="status" value="{{ st }}">{{ st }}</button>{% endfor %}</div></form></div>{% else %}<div class="empty">Buyurtma yo'q</div>{% endfor %}</div>
"""

CUSTOMER_BODY = r"""
<div class="customer-shell">
  <div class="customer-top"><div class="customer-code">BUYURTMA {{ job.order_code }}</div><div class="customer-title">{{ job.customer or 'Hurmatli mijoz' }}</div><div class="customer-stage">{{ snapshot.stage }}</div><div class="customer-progress"><div class="customer-progress-row"><span>Tayyorlik</span><span>{{ snapshot.progress }}%</span></div><div class="progress-track"><div class="progress-fill" style="width:{{ snapshot.progress }}%"></div></div></div><div class="stage-list">{% for st in stages %}<div class="stage-step {{ st.state }}">{{ st.name }}</div>{% endfor %}</div></div>
  <div class="panel" style="margin-top:14px"><div class="statusline"><span class="stat">Material: {{ job.material }}</span><span class="stat">Holat: {{ job.status }}</span><span class="stat">Yetkazish: {{ job.delivery_status or 'Kutilmoqda' }}</span></div><div class="notice-preview"><b>Jonli kuzatuv:</b> ishchi bajarilgan bosqichni belgilashi bilan bu sahifa avtomatik yangilanadi.</div></div>
  <div class="panel"><h3>Buyurtma yangiliklari</h3>{% if updates %}<div class="timeline">{% for u in updates %}<div class="timeline-item"><div class="timeline-title">{{ u.title }}</div><div class="timeline-time">{{ u.created_at }} · {{ u.progress }}%</div><div class="timeline-message">{{ u.message }}</div></div>{% endfor %}</div>{% else %}<div class="empty">Yangiliklar kutilmoqda</div>{% endif %}<div class="auto-note">Sahifa har 15 soniyada avtomatik yangilanadi.</div></div>
</div><script>setTimeout(()=>location.reload(),15000);</script>
"""

USERS_BODY = r"""
<div class="panel hero"><div class="hero-grid"><div><h1>Xodimlar va rollar</h1><p>Har bir xodimga alohida login bering. Parollar bazada shifrlangan holda saqlanadi.</p></div></div></div>
<div class="user-grid"><div class="panel"><h3>Yangi xodim</h3><form method="post" action="{{ url_for('user_create') }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><label>Login<input name="username" minlength="3" maxlength="40" required></label><label>Parol<input name="password" type="password" minlength="8" required></label><label>Roli<select name="role">{% for key,label in role_labels.items() %}{% if key!='admin' %}<option value="{{ key }}">{{ label }}</option>{% endif %}{% endfor %}</select></label><button class="ok" type="submit">Xodimni yaratish</button></form></div><div class="panel"><h3>Foydalanuvchilar</h3><div class="tablewrap"><table><thead><tr><th>Login</th><th>Roli</th><th>Holati</th><th></th></tr></thead><tbody>{% for u in users_list %}<tr><td><b>{{ u.username }}</b></td><td>{{ role_labels.get(u.role,u.role) }}</td><td>{{ 'Faol' if u.is_active else "O'chirilgan" }}</td><td>{% if u.id != session.get('admin_user_id') %}<form method="post" action="{{ url_for('user_toggle',user_id=u.id) }}"><input type="hidden" name="_csrf" value="{{ csrf_token() }}"><button class="{{ 'danger' if u.is_active else 'ok' }}" type="submit">{{ "O'chirish" if u.is_active else 'Faollashtirish' }}</button></form>{% endif %}</td></tr>{% endfor %}</tbody></table></div></div></div>
"""



def render_page(title: str, body_template: str, **context: Any) -> str:
    context.setdefault("role_labels", ROLE_LABELS)
    body = render_template_string(body_template, **context)
    return render_template_string(BASE_TEMPLATE, title=title, body=body, role_labels=ROLE_LABELS)


def current_draft() -> dict[str, Any]:
    defaults = {"order_code": "AD-001", "customer": "", "material": "LMDEF oq 16 mm", "sheet_length": 2800, "sheet_width": 2070, "kerf": 4, "trim": 10, "worker_name": "", "optimization_mode": "large"}
    saved = session.get("kroy_draft", {})
    if isinstance(saved, dict):
        defaults.update(saved)
    defaults.setdefault("optimization_mode", "large")
    return defaults


def _client_ip() -> str:
    forwarded = request.headers.get("X-Forwarded-For", "")
    if forwarded:
        return forwarded.split(",", 1)[0].strip()[:80]
    return (request.remote_addr or "unknown")[:80]


def _admin_exists() -> bool:
    conn = get_db()
    row = conn.execute(
        "SELECT 1 FROM app_users WHERE role='admin' AND is_active=1 LIMIT 1"
    ).fetchone()
    conn.close()
    return bool(row)


def _audit(action: str, user_id: int | None = None, username: str | None = None) -> None:
    try:
        conn = get_db()
        conn.execute(
            "INSERT INTO audit_log(user_id,username,action,ip,created_at) VALUES(?,?,?,?,?)",
            (
                user_id if user_id is not None else session.get("admin_user_id"),
                username if username is not None else session.get("admin_username", ""),
                action[:200],
                _client_ip(),
                now_iso(),
            ),
        )
        conn.commit()
        conn.close()
    except sqlite3.Error:
        # Audit xatosi asosiy ishni to'xtatmasin.
        pass


def _block_seconds_left(ip: str) -> int:
    conn = get_db()
    row = conn.execute(
        "SELECT blocked_until FROM login_attempts WHERE ip=?", (ip,)
    ).fetchone()
    conn.close()
    if not row:
        return 0
    return max(0, int(row["blocked_until"] or 0) - int(tashkent_now().timestamp()))


def _record_failed_login(ip: str) -> bool:
    now_ts = int(tashkent_now().timestamp())
    conn = get_db()
    row = conn.execute(
        "SELECT attempts,blocked_until FROM login_attempts WHERE ip=?", (ip,)
    ).fetchone()
    attempts = int(row["attempts"] or 0) + 1 if row else 1
    blocked_until = int(row["blocked_until"] or 0) if row else 0
    just_blocked = False
    if attempts >= MAX_LOGIN_ATTEMPTS:
        attempts = 0
        blocked_until = now_ts + LOGIN_BLOCK_SECONDS
        just_blocked = True
    conn.execute(
        """INSERT INTO login_attempts(ip,attempts,blocked_until,updated_at)
           VALUES(?,?,?,?)
           ON CONFLICT(ip) DO UPDATE SET
             attempts=excluded.attempts,
             blocked_until=excluded.blocked_until,
             updated_at=excluded.updated_at""",
        (ip, attempts, blocked_until, now_iso()),
    )
    conn.commit()
    conn.close()
    return just_blocked


def _clear_login_attempts(ip: str) -> None:
    conn = get_db()
    conn.execute("DELETE FROM login_attempts WHERE ip=?", (ip,))
    conn.commit()
    conn.close()


def login_required(view):
    @wraps(view)
    def wrapped(*args: Any, **kwargs: Any):
        if not _admin_exists():
            return redirect(url_for("setup"))
        if not session.get("admin_user_id"):
            flash("Avval tizimga kiring", "bad")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped


def roles_required(*allowed_roles: str):
    def decorator(view):
        @wraps(view)
        def wrapped(*args: Any, **kwargs: Any):
            if not _admin_exists():
                return redirect(url_for("setup"))
            if not session.get("admin_user_id"):
                flash("Avval tizimga kiring", "bad")
                return redirect(url_for("login"))
            role = session.get("user_role", "admin")
            if role not in allowed_roles:
                flash("Bu bo'limga kirish huquqingiz yo'q", "bad")
                endpoint = ROLE_HOME_ENDPOINTS.get(role, "dashboard")
                return redirect(url_for(endpoint))
            return view(*args, **kwargs)
        return wrapped
    return decorator


# Eski nom bilan yozilgan dekoratorlar ham ishlashda davom etadi.
admin_login_required = login_required


def _role_redirect(role: str) -> Response:
    return redirect(url_for(ROLE_HOME_ENDPOINTS.get(role, "dashboard")))


def _dashboard_rows(limit: int = 30) -> tuple[list[dict[str, Any]], dict[str, int]]:
    conn = get_db()
    rows = conn.execute("SELECT * FROM kroy_jobs ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
    result: list[dict[str, Any]] = []
    stats = {"total": 0, "process": 0, "ready": 0, "delivered": 0}
    for row in rows:
        snap = _progress_snapshot(conn, row["id"])
        item = dict(row)
        item.update(snap)
        item["customer_url"] = url_for("customer_view", token=row["customer_token"], _external=True)
        item["worker_url"] = url_for("worker_view", token=row["token"], _external=True)
        result.append(item)
    all_rows = conn.execute("SELECT id,delivery_status FROM kroy_jobs").fetchall()
    stats["total"] = len(all_rows)
    for row in all_rows:
        snap = _progress_snapshot(conn, row["id"])
        if snap["delivery_status"] == "Yetkazildi":
            stats["delivered"] += 1
        elif snap["edges"] == snap["total"] and snap["total"]:
            stats["ready"] += 1
        else:
            stats["process"] += 1
    conn.close()
    return result, stats


@app.route("/")
def home() -> Response:
    if not _admin_exists():
        return redirect(url_for("setup"))
    if not session.get("admin_user_id"):
        return redirect(url_for("login"))
    return _role_redirect(session.get("user_role", "admin"))


@app.route("/setup", methods=["GET", "POST"])
def setup() -> Response | str:
    if _admin_exists():
        return redirect(url_for("login"))
    if request.method == "POST":
        username = request.form.get("username", "admin").strip()
        password = request.form.get("password", "")
        password_confirm = request.form.get("password_confirm", "")
        if len(username) < 3 or len(username) > 40 or any(ch.isspace() for ch in username):
            flash("Login 3-40 belgi bo'lsin va bo'sh joy ishlatilmasin", "bad")
            return redirect(url_for("setup"))
        if len(password) < 8:
            flash("Parol kamida 8 ta belgidan iborat bo'lsin", "bad")
            return redirect(url_for("setup"))
        if password != password_confirm:
            flash("Ikki parol bir xil emas", "bad")
            return redirect(url_for("setup"))
        conn = get_db()
        try:
            cur = conn.execute(
                """INSERT INTO app_users(username,password_hash,role,is_active,created_at)
                   VALUES(?,?,'admin',1,?)""",
                (username, generate_password_hash(password), now_iso()),
            )
            conn.commit()
            user_id = int(cur.lastrowid)
        except sqlite3.IntegrityError:
            conn.close()
            flash("Bu login allaqachon mavjud", "bad")
            return redirect(url_for("setup"))
        conn.close()
        session.clear()
        session.permanent = True
        session["admin_user_id"] = user_id
        session["admin_username"] = username
        session["user_role"] = "admin"
        _audit("Birinchi rahbar yaratildi", user_id=user_id, username=username)
        create_database_backup()
        flash("Rahbar xavfsiz yaratildi")
        return redirect(url_for("dashboard"))
    return render_page("Birinchi sozlash", AUTH_BODY, setup=True)


@app.route("/login", methods=["GET", "POST"])
def login() -> Response | str:
    if not _admin_exists():
        return redirect(url_for("setup"))
    if session.get("admin_user_id"):
        return _role_redirect(session.get("user_role", "admin"))
    if request.method == "POST":
        ip = _client_ip()
        seconds_left = _block_seconds_left(ip)
        if seconds_left:
            minutes = max(1, (seconds_left + 59) // 60)
            flash(f"Ko'p noto'g'ri urinish bo'ldi. {minutes} daqiqadan keyin qayta urinib ko'ring.", "bad")
            return redirect(url_for("login"))
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        user = conn.execute(
            "SELECT id,username,password_hash,role FROM app_users WHERE username=? AND is_active=1",
            (username,),
        ).fetchone()
        conn.close()
        if not user or not check_password_hash(user["password_hash"], password):
            just_blocked = _record_failed_login(ip)
            _audit("Tizimga kirish muvaffaqiyatsiz", username=username)
            flash("5 marta noto'g'ri kiritildi. Kirish 15 daqiqaga bloklandi." if just_blocked else "Login yoki parol noto'g'ri", "bad")
            return redirect(url_for("login"))
        _clear_login_attempts(ip)
        session.clear()
        session.permanent = True
        session["admin_user_id"] = int(user["id"])
        session["admin_username"] = user["username"]
        session["user_role"] = user["role"]
        _audit(f"{ROLE_LABELS.get(user['role'], user['role'])} tizimga kirdi")
        flash(f"Xush kelibsiz, {ROLE_LABELS.get(user['role'], user['role'])}")
        return _role_redirect(user["role"])
    return render_page("Tizimga kirish", AUTH_BODY, setup=False)


@app.post("/logout")
@login_required
def logout() -> Response:
    _audit("Tizimdan chiqildi")
    session.clear()
    flash("Tizimdan chiqdingiz")
    return redirect(url_for("login"))


@app.get("/dashboard")
@login_required
def dashboard() -> str:
    jobs, stats = _dashboard_rows(8)
    return render_page("Boshqaruv markazi", DASHBOARD_BODY, jobs=jobs, stats=stats, role=session.get("user_role", "admin"))


@app.get("/manager")
@roles_required("admin", "manager")
def manager_dashboard() -> str:
    jobs, stats = _dashboard_rows(100)
    return render_page("Menejer", MANAGER_BODY, jobs=jobs, stats=stats)


@app.get("/workers")
@roles_required("admin", "manager", "constructor", "worker")
def worker_center() -> str:
    jobs, _ = _dashboard_rows(100)
    return render_page("Ishchi topshiriqlari", WORKER_CENTER_BODY, jobs=jobs)


@app.get("/driver")
@roles_required("admin", "manager", "driver")
def driver_dashboard() -> str:
    jobs, _ = _dashboard_rows(100)
    return render_page("Shafyor", DRIVER_BODY, jobs=jobs, delivery_statuses=DELIVERY_STATUSES)


@app.get("/users")
@roles_required("admin")
def users() -> str:
    conn = get_db()
    users_list = conn.execute("SELECT id,username,role,is_active,created_at FROM app_users ORDER BY id").fetchall()
    conn.close()
    return render_page("Xodimlar", USERS_BODY, users_list=users_list)


@app.post("/users/create")
@roles_required("admin")
def user_create() -> Response:
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    role = request.form.get("role", "worker")
    if role not in ROLE_LABELS or role == "admin":
        flash("Xodim roli noto'g'ri", "bad")
        return redirect(url_for("users"))
    if len(username) < 3 or len(username) > 40 or any(ch.isspace() for ch in username):
        flash("Login 3-40 belgi bo'lsin", "bad")
        return redirect(url_for("users"))
    if len(password) < 8:
        flash("Parol kamida 8 belgi bo'lsin", "bad")
        return redirect(url_for("users"))
    conn = get_db()
    try:
        conn.execute("INSERT INTO app_users(username,password_hash,role,is_active,created_at) VALUES(?,?,?,?,?)", (username, generate_password_hash(password), role, 1, now_iso()))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        flash("Bu login band", "bad")
        return redirect(url_for("users"))
    conn.close()
    create_database_backup()
    _audit(f"Xodim yaratildi: {username} / {role}")
    flash(f"{ROLE_LABELS[role]} uchun login yaratildi")
    return redirect(url_for("users"))


@app.post("/users/<int:user_id>/toggle")
@roles_required("admin")
def user_toggle(user_id: int) -> Response:
    if user_id == session.get("admin_user_id"):
        flash("O'zingizni o'chira olmaysiz", "bad")
        return redirect(url_for("users"))
    conn = get_db()
    row = conn.execute("SELECT username,is_active FROM app_users WHERE id=?", (user_id,)).fetchone()
    if not row:
        conn.close()
        abort(404)
    conn.execute("UPDATE app_users SET is_active=? WHERE id=?", (0 if row["is_active"] else 1, user_id))
    conn.commit()
    conn.close()
    create_database_backup()
    flash("Xodim holati yangilandi")
    return redirect(url_for("users"))


@app.route("/constructor")
@roles_required("admin", "constructor")
def constructor() -> str:
    draft = current_draft()
    parts: list[dict[str, Any]] = []
    conn = get_db()
    import_token = session.get("kroy_import_token")
    if import_token:
        draft_row = conn.execute("SELECT parts_json FROM kroy_import_drafts WHERE token=?", (import_token,)).fetchone()
        if draft_row:
            try:
                parts = json.loads(draft_row["parts_json"])
            except json.JSONDecodeError:
                parts = []
        else:
            session.pop("kroy_import_token", None)
    jobs = conn.execute("SELECT * FROM kroy_jobs ORDER BY id DESC LIMIT 20").fetchall()
    conn.close()
    return render_page("Konstruktor Kroy", CONSTRUCTOR_BODY, draft=draft, parts_json=json.dumps(parts, ensure_ascii=False), jobs=jobs)


@app.post("/constructor/dbf")
@roles_required("admin", "constructor")
def dbf_import() -> Response:
    upload = request.files.get("dbf")
    if not upload or not upload.filename:
        flash("DBF fayl tanlanmadi", "bad")
        return redirect(url_for("constructor"))
    data = upload.read(MAX_IMPORT_BYTES + 1)
    if len(data) > MAX_IMPORT_BYTES:
        flash("DBF fayl 5 MBdan katta bo'lmasligi kerak", "bad")
        return redirect(url_for("constructor"))
    try:
        parts = parse_dbf(data)
    except ValueError as exc:
        flash(str(exc), "bad")
        return redirect(url_for("constructor"))
    import_token = secrets.token_urlsafe(18)
    conn = get_db()
    conn.execute(
        "INSERT INTO kroy_import_drafts(token,source_name,parts_json,created_at) VALUES(?,?,?,?)",
        (import_token, upload.filename[:200], json.dumps([asdict(p) for p in parts], ensure_ascii=False), now_iso()),
    )
    conn.commit()
    conn.close()
    session["kroy_import_token"] = import_token
    draft = current_draft()
    draft["order_code"] = Path(upload.filename).stem[:50]
    session["kroy_draft"] = draft
    create_database_backup()
    _audit(f"DBFdan {len(parts)} tur detal yuklandi")
    flash(f"DBFdan {len(parts)} tur detal yuklandi")
    return redirect(url_for("constructor"))


@app.post("/constructor/table-import")
@roles_required("admin", "constructor")
def table_import() -> Response:
    upload = request.files.get("table_file")
    if not upload or not upload.filename:
        flash("CSV yoki TXT fayl tanlanmadi", "bad")
        return redirect(url_for("constructor"))
    suffix = Path(upload.filename).suffix.lower()
    if suffix not in {".csv", ".txt"}:
        flash("Faqat CSV yoki TXT fayl yuklang", "bad")
        return redirect(url_for("constructor"))
    data = upload.read(MAX_IMPORT_BYTES + 1)
    if len(data) > MAX_IMPORT_BYTES:
        flash("CSV/TXT fayl 5 MBdan katta bo'lmasligi kerak", "bad")
        return redirect(url_for("constructor"))
    try:
        parts = parse_table_parts(data)
    except ValueError as exc:
        flash(str(exc), "bad")
        return redirect(url_for("constructor"))
    import_token = secrets.token_urlsafe(18)
    conn = get_db()
    conn.execute(
        "INSERT INTO kroy_import_drafts(token,source_name,parts_json,created_at) VALUES(?,?,?,?)",
        (import_token, upload.filename[:200], json.dumps([asdict(p) for p in parts], ensure_ascii=False), now_iso()),
    )
    conn.commit()
    conn.close()
    session["kroy_import_token"] = import_token
    draft = current_draft()
    draft["order_code"] = Path(upload.filename).stem[:50]
    session["kroy_draft"] = draft
    create_database_backup()
    _audit(f"CSV/TXTdan {len(parts)} tur detal yuklandi")
    flash(f"CSV/TXTdan {len(parts)} tur detal yuklandi")
    return redirect(url_for("constructor"))


@app.post("/jobs")
@roles_required("admin", "constructor")
def job_create() -> Response:
    try:
        parts_data = json.loads(request.form.get("parts_json", "[]"))
        if not isinstance(parts_data, list):
            raise ValueError("Detallar ro'yxati noto'g'ri")
        parts = validate_parts([
            Part(
                uid=str(p.get("uid") or secrets.token_hex(8)),
                name=str(p.get("name", "")),
                length=int(p.get("length", 0)),
                width=int(p.get("width", 0)),
                qty=int(p.get("qty", 1)),
                rotate=_as_bool(p.get("rotate", True), True),
                edge_left=_as_bool(p.get("edge_left")),
                edge_right=_as_bool(p.get("edge_right")),
                edge_top=_as_bool(p.get("edge_top")),
                edge_bottom=_as_bool(p.get("edge_bottom")),
            )
            for p in parts_data
            if isinstance(p, dict)
        ])
        meta = {
            "order_code": request.form.get("order_code", "").strip()[:60],
            "customer": request.form.get("customer", "").strip()[:100],
            "material": request.form.get("material", "").strip()[:100],
            "sheet_length": int(request.form.get("sheet_length", 0)),
            "sheet_width": int(request.form.get("sheet_width", 0)),
            "kerf": int(request.form.get("kerf", 0)),
            "trim": int(request.form.get("trim", 0)),
            "worker_name": request.form.get("worker_name", "").strip()[:100],
            "optimization_mode": request.form.get("optimization_mode", "large") if request.form.get("optimization_mode", "large") in {"large", "full", "fast"} else "large",
        }
        if not meta["order_code"] or not meta["material"]:
            raise ValueError("Buyurtma kodi va materialni yozing")
        if not (300 <= meta["sheet_length"] <= 10000 and 300 <= meta["sheet_width"] <= 10000):
            raise ValueError("List o'lchami 300-10000 mm oralig'ida bo'lsin")
        if not (0 <= meta["kerf"] <= 30):
            raise ValueError("Arraning izi 0-30 mm oralig'ida bo'lsin")
        if not (0 <= meta["trim"] <= 300):
            raise ValueError("Chet kesimi 0-300 mm oralig'ida bo'lsin")
        if meta["sheet_length"] - 2 * meta["trim"] < MIN_PART_MM or meta["sheet_width"] - 2 * meta["trim"] < MIN_PART_MM:
            raise ValueError("Chet kesimidan keyin listning ishlatiladigan maydoni qolmadi")

        sto_upload = request.files.get("sto_file")
        sto_data: bytes | None = None
        sto_name = ""
        if sto_upload and sto_upload.filename:
            if Path(sto_upload.filename).suffix.lower() != ".sto":
                raise ValueError("PRO100 chizmasi faqat .STO formatida bo'lsin")
            sto_data = sto_upload.read(MAX_STO_BYTES + 1)
            if len(sto_data) > MAX_STO_BYTES:
                raise ValueError("STO fayl 20 MBdan katta bo'lmasligi kerak")
            if not sto_data:
                raise ValueError("STO fayl bo'sh")
            sto_name = sto_upload.filename

        plans, tested_variants, best_variant = pack_parts(
            parts, meta["sheet_length"], meta["sheet_width"], meta["kerf"], meta["trim"], meta["optimization_mode"]
        )
        if not plans:
            raise ValueError("Kroy listi hosil bo'lmadi")
        meta["tested_variants"] = tested_variants
        meta["best_variant"] = best_variant
        job_id, _ = save_job(meta, parts, plans)
        if sto_data is not None:
            save_sto_attachment(job_id, sto_name, sto_data)
        create_database_backup()
        _audit(f"Kroy yaratildi: {meta['order_code']} ({len(plans)} list)")
    except (KeyError, TypeError, ValueError, json.JSONDecodeError, OSError, sqlite3.Error) as exc:
        flash(f"Xato: {exc}", "bad")
        return redirect(url_for("constructor"))
    import_token = session.pop("kroy_import_token", None)
    if import_token:
        conn = get_db()
        conn.execute("DELETE FROM kroy_import_drafts WHERE token=?", (import_token,))
        conn.commit()
        conn.close()
    session["kroy_draft"] = meta
    flash(f"Kroy tayyor: {len(plans)} ta list. {tested_variants} ta variant tekshirildi.")
    return redirect(url_for("job_view", job_id=job_id))

@app.get("/jobs/<int:job_id>")
@roles_required("admin", "constructor", "manager", "worker", "driver")
def job_view(job_id: int) -> str:
    job, parts, sheets_raw = load_job(job_id=job_id)
    worker_url = url_for("worker_view", token=job["token"], _external=True) if job["worker_link_active"] else ""
    sheets = [(row, plan, svg_for_plan(plan)) for row, plan in sheets_raw]
    total_parts = sum(p["qty"] for p in parts)
    conn = get_db()
    attachments = conn.execute("SELECT * FROM kroy_attachments WHERE job_id=? ORDER BY id", (job_id,)).fetchall()
    updates = conn.execute("SELECT * FROM customer_updates WHERE job_id=? ORDER BY id DESC", (job_id,)).fetchall()
    snapshot = _progress_snapshot(conn, job_id)
    conn.close()
    customer_url = url_for("customer_view", token=job["customer_token"], _external=True)
    return render_page(
        f"Kroy {job['order_code']}", JOB_BODY, job=job, sheets=sheets, total_parts=total_parts,
        edge_m=parts_edge_m(parts), worker_url=worker_url, customer_url=customer_url,
        usage=sheet_usage, offcut=offcut_text, snapshot=snapshot, updates=updates,
        opt_label=optimization_label(job["optimization_mode"]), attachments=attachments,
    )


@app.post("/jobs/<int:job_id>/worker-link/revoke")
@roles_required("admin", "constructor", "manager")
def worker_link_revoke(job_id: int) -> Response:
    job, _, _ = load_job(job_id=job_id)
    conn = get_db()
    conn.execute("UPDATE kroy_jobs SET worker_link_active=0 WHERE id=?", (job_id,))
    conn.commit()
    conn.close()
    create_database_backup()
    _audit(f"Ishchi havolasi bekor qilindi: {job['order_code']}")
    flash("Eski ishchi havolasi bekor qilindi")
    return redirect(url_for("job_view", job_id=job_id))


@app.post("/jobs/<int:job_id>/worker-link/regenerate")
@roles_required("admin", "constructor", "manager")
def worker_link_regenerate(job_id: int) -> Response:
    job, _, _ = load_job(job_id=job_id)
    new_token = secrets.token_urlsafe(24)
    conn = get_db()
    conn.execute(
        "UPDATE kroy_jobs SET token=?,worker_link_active=1,worker_token_created_at=?,sent_at=? WHERE id=?",
        (new_token, now_iso(), now_iso(), job_id),
    )
    conn.commit()
    conn.close()
    create_database_backup()
    _audit(f"Yangi ishchi havolasi yaratildi: {job['order_code']}")
    flash("Yangi maxfiy ishchi havolasi yaratildi")
    return redirect(url_for("job_view", job_id=job_id))


@app.get("/attachments/<int:attachment_id>")
@roles_required("admin", "constructor", "manager")
def attachment_download(attachment_id: int) -> Response:
    conn = get_db()
    row = conn.execute("SELECT * FROM kroy_attachments WHERE id=?", (attachment_id,)).fetchone()
    conn.close()
    if not row:
        abort(404)
    path = UPLOAD_DIR / row["stored_name"]
    if not path.exists():
        abort(404, description="Biriktirilgan fayl diskda topilmadi")
    return send_file(path, as_attachment=True, download_name=row["original_name"], mimetype="application/octet-stream")


@app.get("/admin/backup.db")
@roles_required("admin")
def download_backup() -> Response:
    backup = create_database_backup()
    if not backup or not backup.exists():
        abort(500, description="Baza nusxasini yaratib bo'lmadi")
    _audit("Baza zaxira nusxasi yuklandi")
    return send_file(backup, as_attachment=True, download_name=backup.name, mimetype="application/x-sqlite3")


@app.get("/worker/<token>")
def worker_view(token: str) -> str:
    job, parts, sheets_raw = load_job(token=token)
    sheets = [(row, plan, svg_for_plan(plan)) for row, plan in sheets_raw]
    total_parts = sum(p["qty"] for p in parts)
    conn = get_db()
    snapshot = _progress_snapshot(conn, job["id"])
    conn.close()
    return render_page(
        f"Ishchi - {job['order_code']}", WORKER_BODY, job=job, sheets=sheets, total_parts=total_parts,
        edge_m=parts_edge_m(parts), usage=sheet_usage, offcut=offcut_text, snapshot=snapshot,
        opt_label=optimization_label(job["optimization_mode"]),
    )


@app.post("/worker/<token>/sheet/<int:sheet_no>")
def worker_update(token: str, sheet_no: int) -> Response:
    job, _, _ = load_job(token=token)
    requested_cut = 1 if request.form.get("cut_done") else 0
    requested_edge = 1 if request.form.get("edge_done") else 0
    note = request.form.get("note", "").strip()[:500]
    conn = get_db()
    old = conn.execute("SELECT cut_done,edge_done FROM kroy_sheets WHERE job_id=? AND sheet_no=?", (job["id"], sheet_no)).fetchone()
    if not old:
        conn.close()
        abort(404)
    # Bajarilgan ish ortga qaytmaydi: tasodifiy bosish mijozdagi xabarni buzmaydi.
    cut_done = max(int(old["cut_done"] or 0), requested_cut)
    edge_done = max(int(old["edge_done"] or 0), requested_edge)
    if edge_done:
        cut_done = 1
    conn.execute(
        "UPDATE kroy_sheets SET cut_done=?,edge_done=?,note=? WHERE job_id=? AND sheet_no=?",
        (cut_done, edge_done, note, job["id"], sheet_no),
    )
    stats = conn.execute(
        "SELECT COUNT(*) total,COALESCE(SUM(cut_done),0) cuts,COALESCE(SUM(edge_done),0) edges FROM kroy_sheets WHERE job_id=?",
        (job["id"],),
    ).fetchone()
    total = int(stats["total"] or 0)
    cuts = int(stats["cuts"] or 0)
    edges = int(stats["edges"] or 0)
    progress = _job_progress_from_counts(total, cuts, edges, job["delivery_status"] or "Kutilmoqda")
    if cut_done and not int(old["cut_done"] or 0):
        if total and cuts == total:
            _add_customer_update(conn, job["id"], "all-cut-done", "Kroy ishlari tugadi", f"{job['order_code']} buyurtmasining barcha detallari kesildi. Endi kromka ishlari davom etadi.", "Kroy", max(40, progress))
        else:
            _add_customer_update(conn, job["id"], f"sheet-{sheet_no}-cut", f"{sheet_no}-list kesildi", f"{job['order_code']} buyurtmasining {sheet_no}-listidagi detallar kesildi.", "Kroy", progress)
    if edge_done and not int(old["edge_done"] or 0):
        if total and edges == total:
            _add_customer_update(conn, job["id"], "all-edge-done", "Kromka ishlari tugadi", f"{job['order_code']} buyurtmasining barcha kromkalari urildi. Konstruktor bosqichi tayyor.", "Kromka", 80)
        else:
            _add_customer_update(conn, job["id"], f"sheet-{sheet_no}-edge", f"{sheet_no}-list kromkasi tayyor", f"{job['order_code']} buyurtmasining {sheet_no}-list kromkasi urildi.", "Kromka", progress)
    if total and cuts == total and edges == total:
        status, finished_at = "Tayyor", now_iso()
    elif cuts or edges:
        status, finished_at = "Jarayonda", ""
    else:
        status, finished_at = "Ishchiga yuborildi", ""
    conn.execute("UPDATE kroy_jobs SET status=?,finished_at=? WHERE id=?", (status, finished_at, job["id"]))
    conn.commit()
    conn.close()
    create_database_backup()
    _audit(f"Ishchi list {sheet_no} holatini yangiladi: {job['order_code']}", username=job["worker_name"] or "Ishchi")
    flash(f"List {sheet_no} saqlandi. Mijoz kuzatuvi avtomatik yangilandi.")
    return redirect(url_for("worker_view", token=token))


@app.post("/manager/jobs/<int:job_id>/customer-update")
@roles_required("admin", "manager", "constructor")
def manager_customer_update(job_id: int) -> Response:
    job, _, _ = load_job(job_id=job_id)
    title = request.form.get("title", "").strip()[:140]
    message = request.form.get("message", "").strip()[:500]
    try:
        progress = max(0, min(100, int(request.form.get("progress", 0))))
    except ValueError:
        progress = 0
    if not title or not message:
        flash("Sarlavha va xabarni yozing", "bad")
        return redirect(url_for("job_view", job_id=job_id))
    conn = get_db()
    _add_customer_update(conn, job_id, f"manual-{secrets.token_hex(6)}", title, message, "Menejer", progress)
    conn.commit()
    conn.close()
    create_database_backup()
    _audit(f"Mijozga xabar qo'shildi: {job['order_code']}")
    flash("Xabar mijoz kuzatuv sahifasiga qo'shildi")
    return redirect(url_for("job_view", job_id=job_id))


@app.post("/driver/jobs/<int:job_id>/status")
@roles_required("admin", "manager", "driver")
def driver_update(job_id: int) -> Response:
    job, _, _ = load_job(job_id=job_id)
    status = request.form.get("status", "Kutilmoqda")
    note = request.form.get("note", "").strip()[:500]
    if status not in DELIVERY_STATUSES:
        flash("Yetkazish holati noto'g'ri", "bad")
        return redirect(url_for("driver_dashboard"))
    delivered_at = now_iso() if status == "Yetkazildi" else ""
    conn = get_db()
    conn.execute("UPDATE kroy_jobs SET delivery_status=?,delivery_note=?,delivered_at=? WHERE id=?", (status, note, delivered_at, job_id))
    progress_map = {"Kutilmoqda": 80, "Yetkazishga tayyor": 85, "Yo'lda": 95, "Yetkazildi": 100}
    title_map = {"Kutilmoqda": "Yetkazish rejalashtirilmoqda", "Yetkazishga tayyor": "Buyurtma yetkazishga tayyor", "Yo'lda": "Buyurtma yo'lga chiqdi", "Yetkazildi": "Buyurtma yetkazildi"}
    message_map = {"Kutilmoqda": f"{job['order_code']} buyurtmasining yetkazish vaqti rejalashtirilmoqda.", "Yetkazishga tayyor": f"{job['order_code']} buyurtmasi yuklash va yetkazishga tayyor.", "Yo'lda": f"{job['order_code']} buyurtmasi siz tomon yo'lga chiqdi.", "Yetkazildi": f"{job['order_code']} buyurtmasi yetkazildi. Ishonchingiz uchun rahmat!"}
    if note:
        message_map[status] += " " + note
    _add_customer_update(conn, job_id, f"delivery-{status}-{secrets.token_hex(4)}", title_map[status], message_map[status], "Yetkazish", progress_map[status])
    conn.commit()
    conn.close()
    create_database_backup()
    _audit(f"Yetkazish holati: {job['order_code']} / {status}")
    flash("Yetkazish holati saqlandi va mijozga bildirildi")
    return redirect(url_for("driver_dashboard"))


@app.get("/customer/<token>")
def customer_view(token: str) -> str:
    conn = get_db()
    job = conn.execute("SELECT * FROM kroy_jobs WHERE customer_token=? AND customer_link_active=1", (token,)).fetchone()
    if not job:
        conn.close()
        abort(404)
    snapshot = _progress_snapshot(conn, job["id"])
    updates = conn.execute("SELECT * FROM customer_updates WHERE job_id=? ORDER BY id DESC", (job["id"],)).fetchall()
    conn.close()
    return render_page(f"Buyurtma {job['order_code']}", CUSTOMER_BODY, job=job, snapshot=snapshot, stages=_customer_stages(snapshot), updates=updates)


@app.get("/jobs/<int:job_id>/a4.pdf")
@roles_required("admin", "constructor", "manager", "worker")
def job_pdf(job_id: int) -> Response:
    job, _, sheets = load_job(job_id=job_id)
    out = build_pdf(job, sheets)
    filename = f"kroy_{job['order_code']}.pdf".replace("/", "-")
    return send_file(out, mimetype="application/pdf", as_attachment=False, download_name=filename)


@app.get("/worker/<token>/a4.pdf")
def worker_pdf(token: str) -> Response:
    job, _, sheets = load_job(token=token)
    out = build_pdf(job, sheets)
    filename = f"kroy_{job['order_code']}.pdf".replace("/", "-")
    return send_file(out, mimetype="application/pdf", as_attachment=False, download_name=filename)


@app.get("/api/health")
def health() -> Response:
    return jsonify({"ok": True, "module": "Mebel360 Boshqaruv Pro V9", "customer_updates": True, "roles": list(ROLE_LABELS)})


@app.errorhandler(400)
def bad_request(error: Exception) -> tuple[str, int]:
    description = getattr(error, "description", "So'rov noto'g'ri")
    return f"Xato: {description}", 400


@app.errorhandler(410)
def link_expired(_: Exception) -> tuple[str, int]:
    return "Bu maxfiy havola bekor qilingan. Rahbardan yangi havola oling.", 410


@app.errorhandler(413)
def too_large(_: Exception) -> tuple[str, int]:
    return "Yuklanayotgan fayl ruxsat etilgan hajmdan katta", 413


init_db()
create_database_backup()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
