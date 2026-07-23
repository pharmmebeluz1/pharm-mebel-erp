from __future__ import annotations

import os
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_from_directory, session
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("MEBEL360_DB", BASE_DIR / "mebel360_demo.db"))
UPLOAD_DIR = Path(os.environ.get("MEBEL360_UPLOADS", BASE_DIR / "uploads"))
ALLOWED_UPLOADS = {"png", "jpg", "jpeg", "webp", "pdf", "doc", "docx", "xls", "xlsx"}
MAX_UPLOAD_BYTES = 8 * 1024 * 1024
TASHKENT = timezone(timedelta(hours=5), name="Asia/Tashkent")

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "mebel360-demo-secret-change-me")
app.config["JSON_AS_ASCII"] = False
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES


def now_tashkent() -> datetime:
    return datetime.now(TASHKENT)


def get_db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def init_db() -> None:
    """Demo bazani yaratadi. Mavjud Mebel360 bazasiga ulash oson bo‘lishi uchun SQLite ishlatilgan."""
    with get_db() as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                full_name TEXT NOT NULL,
                phone TEXT NOT NULL DEFAULT '',
                passport_id TEXT NOT NULL DEFAULT '',
                address TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id INTEGER NOT NULL,
                order_code TEXT NOT NULL UNIQUE,
                product_name TEXT NOT NULL,
                specification TEXT NOT NULL DEFAULT '',
                material TEXT NOT NULL DEFAULT '',
                color TEXT NOT NULL DEFAULT '',
                quantity INTEGER NOT NULL DEFAULT 1,
                total_amount INTEGER NOT NULL DEFAULT 0,
                advance_amount INTEGER NOT NULL DEFAULT 0,
                interim_amount INTEGER NOT NULL DEFAULT 0,
                payment_method TEXT NOT NULL DEFAULT 'Naqd / karta / bank',
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                delivery_address TEXT NOT NULL DEFAULT '',
                landmark TEXT NOT NULL DEFAULT '',
                floor TEXT NOT NULL DEFAULT '',
                lift TEXT NOT NULL DEFAULT '',
                truck_access TEXT NOT NULL DEFAULT '',
                installation TEXT NOT NULL DEFAULT 'Kiritilgan',
                warranty_months INTEGER NOT NULL DEFAULT 6,
                status TEXT NOT NULL DEFAULT 'Tasdiqlash kutilmoqda',
                created_at TEXT NOT NULL,
                FOREIGN KEY (customer_id) REFERENCES customers(id)
            );

            CREATE TABLE IF NOT EXISTS contracts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL UNIQUE,
                contract_number TEXT NOT NULL UNIQUE,
                version INTEGER NOT NULL DEFAULT 1,
                generated_at TEXT NOT NULL,
                confirmed_at TEXT,
                confirmed_ip TEXT,
                confirmed_user_agent TEXT,
                status TEXT NOT NULL DEFAULT 'Tasdiqlash kutilmoqda',
                FOREIGN KEY (order_id) REFERENCES orders(id)
            );

            CREATE TABLE IF NOT EXISTS support_tickets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                customer_id INTEGER NOT NULL,
                topic TEXT NOT NULL DEFAULT 'Buyurtma jarayoni',
                subject TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'Yangi',
                priority TEXT NOT NULL DEFAULT 'Oddiy',
                assigned_role TEXT NOT NULL DEFAULT 'manager',
                escalated_to_admin INTEGER NOT NULL DEFAULT 0,
                customer_unread INTEGER NOT NULL DEFAULT 0,
                manager_unread INTEGER NOT NULL DEFAULT 1,
                admin_unread INTEGER NOT NULL DEFAULT 0,
                last_message_at TEXT NOT NULL,
                created_at TEXT NOT NULL,
                closed_at TEXT,
                FOREIGN KEY (order_id) REFERENCES orders(id),
                FOREIGN KEY (customer_id) REFERENCES customers(id)
            );

            CREATE TABLE IF NOT EXISTS support_messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                sender_role TEXT NOT NULL,
                sender_name TEXT NOT NULL,
                body TEXT NOT NULL DEFAULT '',
                attachment_name TEXT,
                attachment_path TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY (ticket_id) REFERENCES support_tickets(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_support_tickets_last_message
                ON support_tickets(last_message_at DESC);
            CREATE INDEX IF NOT EXISTS idx_support_messages_ticket
                ON support_messages(ticket_id, created_at);
            """
        )

        customer = db.execute("SELECT id FROM customers WHERE username = ?", ("mijoz",)).fetchone()
        if customer is None:
            cursor = db.execute(
                """INSERT INTO customers (username, full_name, phone, passport_id, address)
                   VALUES (?, ?, ?, ?, ?)""",
                (
                    "mijoz",
                    "Abdulla Karimov",
                    "+998 90 123 45 67",
                    "AA 1234567",
                    "Toshkent shahri, Chilonzor tumani",
                ),
            )
            customer_id = cursor.lastrowid
        else:
            customer_id = customer["id"]

        order = db.execute("SELECT id FROM orders WHERE order_code = ?", ("AD-001",)).fetchone()
        if order is None:
            cursor = db.execute(
                """
                INSERT INTO orders (
                    customer_id, order_code, product_name, specification, material, color, quantity,
                    total_amount, advance_amount, interim_amount, payment_method, start_date, end_date,
                    delivery_address, landmark, floor, lift, truck_access, installation,
                    warranty_months, status, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    customer_id,
                    "AD-001",
                    "Dorixona uchun mebel to‘plami",
                    "Savdo vitrinalari, kassa qismi, devor shkaflari, oynali polkalar va LED yoritish",
                    "MDF/LMDF, furnitura va oyna",
                    "Oq, yashil va mijoz tasdiqlagan ranglar",
                    1,
                    30_000_000,
                    15_000_000,
                    0,
                    "Naqd / Click / Payme / bank o‘tkazmasi",
                    "2026-07-22",
                    "2026-08-10",
                    "Toshkent shahri, Chilonzor tumani",
                    "Bunyodkor ko‘chasi yaqinida",
                    "1-qavat",
                    "Talab qilinmaydi",
                    "Katta mashina kira oladi",
                    "Yetkazish va montaj narxga kiritilgan",
                    6,
                    "Mijoz tasdig‘i kutilmoqda",
                    now_tashkent().isoformat(timespec="seconds"),
                ),
            )
            order_id = cursor.lastrowid
        else:
            order_id = order["id"]

        contract = db.execute("SELECT id FROM contracts WHERE order_id = ?", (order_id,)).fetchone()
        if contract is None:
            db.execute(
                """INSERT INTO contracts (order_id, contract_number, generated_at)
                   VALUES (?, ?, ?)""",
                (order_id, "M360-2026-AD001", now_tashkent().isoformat(timespec="seconds")),
            )


def money(value: int) -> str:
    return f"{int(value):,}".replace(",", " ") + " so‘m"


def pretty_date(value: str) -> str:
    try:
        return datetime.fromisoformat(value).strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return value or "—"


def get_customer_contract(username: str | None) -> dict | None:
    """Kirishdagi foydalanuvchining eng so‘nggi shartnomasini qaytaradi."""
    with get_db() as db:
        row = db.execute(
            """
            SELECT
                c.id AS contract_id, c.contract_number, c.version, c.generated_at,
                c.confirmed_at, c.status AS contract_status,
                o.id AS order_id, o.order_code, o.product_name, o.specification, o.material,
                o.color, o.quantity, o.total_amount, o.advance_amount, o.interim_amount,
                o.payment_method, o.start_date, o.end_date, o.delivery_address, o.landmark,
                o.floor, o.lift, o.truck_access, o.installation, o.warranty_months,
                o.status AS order_status, o.created_at,
                cu.full_name, cu.phone, cu.passport_id, cu.address
            FROM contracts c
            JOIN orders o ON o.id = c.order_id
            JOIN customers cu ON cu.id = o.customer_id
            WHERE cu.username = ?
            ORDER BY c.id DESC
            LIMIT 1
            """,
            (username or "",),
        ).fetchone()

        # Demo rejimida istalgan login bilan sinash uchun namuna mijoz shartnomasi ko‘rsatiladi.
        if row is None:
            row = db.execute(
                """
                SELECT
                    c.id AS contract_id, c.contract_number, c.version, c.generated_at,
                    c.confirmed_at, c.status AS contract_status,
                    o.id AS order_id, o.order_code, o.product_name, o.specification, o.material,
                    o.color, o.quantity, o.total_amount, o.advance_amount, o.interim_amount,
                    o.payment_method, o.start_date, o.end_date, o.delivery_address, o.landmark,
                    o.floor, o.lift, o.truck_access, o.installation, o.warranty_months,
                    o.status AS order_status, o.created_at,
                    cu.full_name, cu.phone, cu.passport_id, cu.address
                FROM contracts c
                JOIN orders o ON o.id = c.order_id
                JOIN customers cu ON cu.id = o.customer_id
                WHERE cu.username = 'mijoz'
                ORDER BY c.id DESC
                LIMIT 1
                """
            ).fetchone()

    if row is None:
        return None

    balance = int(row["total_amount"]) - int(row["advance_amount"]) - int(row["interim_amount"])
    generated_date = pretty_date(row["generated_at"])
    confirmed_at = row["confirmed_at"]

    return {
        "contract": {
            "id": row["contract_id"],
            "number": row["contract_number"],
            "version": row["version"],
            "generated_at": row["generated_at"],
            "generated_date": generated_date,
            "confirmed_at": confirmed_at,
            "confirmed_date": pretty_date(confirmed_at) if confirmed_at else None,
            "status": row["contract_status"],
        },
        "seller": {
            "name": "Mebel360° mebel ishlab chiqarish korxonasi",
            "representative": "Rahbar: Zuhriddin Ubaydullayev",
            "phone": "+998 __ ___ __ __",
            "address": "Toshkent shahri",
            "inn": "________________",
            "bank": "________________",
            "account": "________________",
        },
        "customer": {
            "full_name": row["full_name"],
            "phone": row["phone"],
            "passport_id": row["passport_id"],
            "address": row["address"],
        },
        "order": {
            "id": row["order_id"],
            "code": row["order_code"],
            "product_name": row["product_name"],
            "specification": row["specification"],
            "material": row["material"],
            "color": row["color"],
            "quantity": row["quantity"],
            "start_date": pretty_date(row["start_date"]),
            "end_date": pretty_date(row["end_date"]),
            "delivery_address": row["delivery_address"],
            "landmark": row["landmark"],
            "floor": row["floor"],
            "lift": row["lift"],
            "truck_access": row["truck_access"],
            "installation": row["installation"],
            "warranty_months": row["warranty_months"],
            "status": row["order_status"],
        },
        "payment": {
            "total": row["total_amount"],
            "total_text": money(row["total_amount"]),
            "advance": row["advance_amount"],
            "advance_text": money(row["advance_amount"]),
            "interim": row["interim_amount"],
            "interim_text": money(row["interim_amount"]),
            "balance": balance,
            "balance_text": money(balance),
            "method": row["payment_method"],
        },
    }


def current_customer_context() -> dict | None:
    contract = get_customer_contract(session.get("username"))
    if contract is None:
        return None
    return contract


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_UPLOADS


def save_attachment(file_storage) -> tuple[str | None, str | None]:
    if not file_storage or not file_storage.filename:
        return None, None
    original_name = secure_filename(file_storage.filename)
    if not original_name or not allowed_file(original_name):
        raise ValueError("Rasm, PDF yoki ofis hujjatini yuboring.")
    file_storage.stream.seek(0, 2)
    size = file_storage.stream.tell()
    file_storage.stream.seek(0)
    if size > MAX_UPLOAD_BYTES:
        raise ValueError("Fayl hajmi 8 MB dan oshmasligi kerak.")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    stamp = now_tashkent().strftime("%Y%m%d%H%M%S%f")
    saved_name = f"{stamp}_{original_name}"
    file_storage.save(UPLOAD_DIR / saved_name)
    return file_storage.filename, saved_name


def build_ai_acknowledgement(customer_name: str, topic: str, body: str) -> str:
    """Mijoz murojaatiga darhol, mazmuniga mos odobli avtomatik javob tayyorlaydi."""
    text = f"{topic} {body}".lower()
    text = (
        text.replace("’", "'")
        .replace("ʻ", "'")
        .replace("‘", "'")
        .replace("`", "'")
    )

    reply_topic = "Murojaatingiz"
    keyword_groups = [
        (("rang", "tus", "bo'yoq", "bo‘yoq", "color", "цвет"), "Rang masalasi bo‘yicha xabaringiz"),
        (("chizma", "o'lchov", "o‘lchov", "razmer", "dizayn", "eskiz"), "Chizma va o‘lchov masalasi bo‘yicha xabaringiz"),
        (("to'lov", "to‘lov", "pul", "avans", "qoldiq", "hisob", "payment"), "To‘lov masalasi bo‘yicha xabaringiz"),
        (("yetkaz", "montaj", "o'rnat", "o‘rnat", "manzil"), "Yetkazish va o‘rnatish bo‘yicha xabaringiz"),
        (("muddat", "qachon", "tayyor", "kechik", "vaqt", "sana"), "Buyurtma muddati bo‘yicha savolingiz"),
        (("o'zgart", "o‘zgart", "almashtir", "qo'sh", "qo‘sh"), "Buyurtmaga o‘zgartirish kiritish bo‘yicha xabaringiz"),
        (("shikoyat", "norozi", "muammo", "xato", "kamchilik"), "Muhim murojaatingiz"),
    ]
    for keywords, label in keyword_groups:
        if any(keyword in text for keyword in keywords):
            reply_topic = label
            break

    clean_name = " ".join(str(customer_name or "").split())
    greeting = f"Assalomu alaykum, hurmatli {clean_name}!" if clean_name else "Assalomu alaykum, hurmatli mijoz!"
    return (
        f"{greeting}\n\n"
        f"{reply_topic}ni qabul qildim. Ma’lumotlaringizni menejerga yetkazdim. "
        "Menejerimiz masalani ko‘rib chiqib, sizga tez orada shu chat orqali javob yozadi.\n\n"
        "Murojaatingiz e’tiborsiz qolmaydi."
    )


def serialize_message(row: sqlite3.Row) -> dict:
    return {
        "id": row["id"],
        "ticket_id": row["ticket_id"],
        "sender_role": row["sender_role"],
        "sender_name": row["sender_name"],
        "body": row["body"],
        "attachment_name": row["attachment_name"],
        "attachment_url": f"/uploads/{row['attachment_path']}" if row["attachment_path"] else None,
        "created_at": row["created_at"],
        "created_text": pretty_date_time(row["created_at"]),
    }


def pretty_date_time(value: str) -> str:
    try:
        return datetime.fromisoformat(value).strftime("%d.%m.%Y • %H:%M")
    except (TypeError, ValueError):
        return value or "—"


def serialize_ticket(row: sqlite3.Row) -> dict:
    return {
        "id": row["id"],
        "order_id": row["order_id"],
        "order_code": row["order_code"],
        "customer_name": row["customer_name"],
        "customer_phone": row["customer_phone"],
        "topic": row["topic"],
        "subject": row["subject"],
        "status": row["status"],
        "priority": row["priority"],
        "assigned_role": row["assigned_role"],
        "escalated_to_admin": bool(row["escalated_to_admin"]),
        "customer_unread": bool(row["customer_unread"]),
        "manager_unread": bool(row["manager_unread"]),
        "admin_unread": bool(row["admin_unread"]),
        "last_message_at": row["last_message_at"],
        "last_message_text": pretty_date_time(row["last_message_at"]),
        "created_at": row["created_at"],
        "created_text": pretty_date_time(row["created_at"]),
        "last_body": row["last_body"] or "",
        "message_count": row["message_count"],
    }


def get_ticket_rows(db: sqlite3.Connection, where: str = "", params: tuple = ()) -> list[sqlite3.Row]:
    query = f"""
        SELECT t.*, o.order_code,
               cu.full_name AS customer_name, cu.phone AS customer_phone,
               (SELECT body FROM support_messages sm WHERE sm.ticket_id=t.id AND sm.sender_role='customer' ORDER BY sm.id DESC LIMIT 1) AS last_body,
               (SELECT COUNT(*) FROM support_messages sm WHERE sm.ticket_id=t.id) AS message_count
        FROM support_tickets t
        JOIN orders o ON o.id=t.order_id
        JOIN customers cu ON cu.id=t.customer_id
        {where}
        ORDER BY t.escalated_to_admin DESC, t.last_message_at DESC
    """
    return db.execute(query, params).fetchall()


def require_staff() -> str | None:
    role = session.get("role")
    return role if role in {"manager", "admin"} else None


def apply_overdue_escalations(db: sqlite3.Connection) -> int:
    """Menejer 1 soat ichida javob bermagan murojaatni rahbar nazoratiga chiqaradi."""
    cutoff = (now_tashkent() - timedelta(hours=1)).isoformat(timespec="seconds")
    overdue = db.execute(
        """SELECT id FROM support_tickets
           WHERE escalated_to_admin=0
             AND status IN ('Yangi', 'Mijoz javob berdi')
             AND last_message_at <= ?""",
        (cutoff,),
    ).fetchall()
    if not overdue:
        return 0
    escalated_at = now_tashkent().isoformat(timespec="seconds")
    for row in overdue:
        db.execute(
            """UPDATE support_tickets SET escalated_to_admin=1, assigned_role='admin',
               priority='Muhim', status='Rahbarga yuborildi', admin_unread=1,
               last_message_at=? WHERE id=?""",
            (escalated_at, row["id"]),
        )
        db.execute(
            """INSERT INTO support_messages
               (ticket_id, sender_role, sender_name, body, created_at)
               VALUES (?, 'system', 'Mebel360°',
               'Menejer 1 soat ichida javob bermagani uchun murojaat avtomatik ravishda rahbar nazoratiga yuborildi.', ?)""",
            (row["id"], escalated_at),
        )
    return len(overdue)


@app.before_request
def ensure_database() -> None:
    init_db()


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/login")
def login():
    data = request.get_json(silent=True) or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not username or not password:
        return jsonify(ok=False, message="Login va parolni kiriting."), 400

    # Demo: istalgan bo‘sh bo‘lmagan login/parol. Mavjud tizimda haqiqiy tekshiruv bilan almashtiriladi.
    session["username"] = username
    session.permanent = bool(data.get("remember"))

    # Mijoz login qilsa shartnoma ro‘l tanlamasdan avtomatik ochiladi.
    with get_db() as db:
        is_customer = db.execute("SELECT 1 FROM customers WHERE username = ?", (username,)).fetchone() is not None
    if is_customer:
        session["role"] = "customer"

    return jsonify(
        ok=True,
        message="Kirish muvaffaqiyatli.",
        next_screen="contractScreen" if is_customer else "roleScreen",
        auto_contract=is_customer,
    )


@app.post("/api/select-role")
def select_role():
    allowed_roles = {"admin", "worker", "driver", "manager", "constructor", "customer"}
    data = request.get_json(silent=True) or {}
    role = str(data.get("role", "")).strip().lower()

    if role not in allowed_roles:
        return jsonify(ok=False, message="Rol noto‘g‘ri tanlandi."), 400

    session["role"] = role
    next_screen = None
    if role == "customer":
        next_screen = "contractScreen"
    elif role in {"manager", "admin"}:
        next_screen = "staffSupportScreen"

    return jsonify(ok=True, role=role, next_screen=next_screen, message="Rol tanlandi.")


@app.get("/api/customer/contract")
def customer_contract():
    contract = get_customer_contract(session.get("username"))
    if contract is None:
        return jsonify(ok=False, message="Mijozga biriktirilgan shartnoma topilmadi."), 404
    return jsonify(ok=True, data=contract)


@app.post("/api/customer/contract/confirm")
def confirm_customer_contract():
    data = request.get_json(silent=True) or {}
    contract_id = data.get("contract_id")
    accepted = bool(data.get("accepted"))
    if not accepted:
        return jsonify(ok=False, message="Avval shartnoma bilan tanishganingizni belgilang."), 400
    if not contract_id:
        return jsonify(ok=False, message="Shartnoma aniqlanmadi."), 400

    assigned_contract = get_customer_contract(session.get("username"))
    if assigned_contract is None or int(assigned_contract["contract"]["id"]) != int(contract_id):
        return jsonify(ok=False, message="Bu shartnomani tasdiqlash huquqi mavjud emas."), 403

    confirmed_at = now_tashkent().isoformat(timespec="seconds")
    forwarded_for = request.headers.get("X-Forwarded-For", "")
    client_ip = forwarded_for.split(",")[0].strip() if forwarded_for else request.remote_addr
    user_agent = request.headers.get("User-Agent", "")[:500]

    with get_db() as db:
        result = db.execute(
            """
            UPDATE contracts
            SET confirmed_at = COALESCE(confirmed_at, ?),
                confirmed_ip = COALESCE(confirmed_ip, ?),
                confirmed_user_agent = COALESCE(confirmed_user_agent, ?),
                status = 'Mijoz tasdiqladi'
            WHERE id = ?
            """,
            (confirmed_at, client_ip, user_agent, contract_id),
        )
        if result.rowcount == 0:
            return jsonify(ok=False, message="Shartnoma topilmadi."), 404
        db.execute(
            """UPDATE orders SET status = 'Mijoz tasdiqladi'
               WHERE id = (SELECT order_id FROM contracts WHERE id = ?)""",
            (contract_id,),
        )

    return jsonify(
        ok=True,
        message="Shartnoma tasdiqlandi.",
        confirmed_at=confirmed_at,
        confirmed_date=pretty_date(confirmed_at),
    )


@app.get("/uploads/<path:filename>")
def uploaded_file(filename: str):
    return send_from_directory(UPLOAD_DIR, filename, as_attachment=False)


@app.get("/api/customer/support")
def customer_support():
    contract = current_customer_context()
    if contract is None:
        return jsonify(ok=False, message="Mijoz buyurtmasi topilmadi."), 404
    order_id = int(contract["order"]["id"])
    customer_name = contract["customer"]["full_name"]
    requested_ticket = request.args.get("ticket_id", type=int)

    with get_db() as db:
        apply_overdue_escalations(db)
        rows = get_ticket_rows(db, "WHERE t.order_id = ?", (order_id,))
        ticket_id = requested_ticket or (rows[0]["id"] if rows else None)
        active = next((row for row in rows if row["id"] == ticket_id), None)
        messages = []
        if active is not None:
            db.execute("UPDATE support_tickets SET customer_unread=0 WHERE id=?", (active["id"],))
            rows = get_ticket_rows(db, "WHERE t.order_id = ?", (order_id,))
            active = next((row for row in rows if row["id"] == ticket_id), None)
            message_rows = db.execute(
                "SELECT * FROM support_messages WHERE ticket_id=? ORDER BY id", (ticket_id,)
            ).fetchall()
            messages = [serialize_message(row) for row in message_rows]

    return jsonify(
        ok=True,
        data={
            "customer_name": customer_name,
            "order": {
                "id": order_id,
                "code": contract["order"]["code"],
                "product_name": contract["order"]["product_name"],
                "status": contract["order"]["status"],
            },
            "tickets": [serialize_ticket(row) for row in rows],
            "active_ticket": serialize_ticket(active) if active else None,
            "messages": messages,
        },
    )


@app.post("/api/customer/support/messages")
def customer_support_message():
    contract = current_customer_context()
    if contract is None:
        return jsonify(ok=False, message="Mijoz buyurtmasi topilmadi."), 404

    payload = request.form if request.form else (request.get_json(silent=True) or {})
    body = str(payload.get("body", "")).strip()
    topic = str(payload.get("topic", "Buyurtma jarayoni")).strip() or "Buyurtma jarayoni"
    subject = str(payload.get("subject", "")).strip() or topic
    ticket_id = payload.get("ticket_id")
    try:
        ticket_id = int(ticket_id) if ticket_id else None
    except (TypeError, ValueError):
        ticket_id = None

    try:
        attachment_name, attachment_path = save_attachment(request.files.get("attachment"))
    except ValueError as exc:
        return jsonify(ok=False, message=str(exc)), 400
    if not body and not attachment_path:
        return jsonify(ok=False, message="Xabar matnini yozing yoki fayl biriktiring."), 400

    order_id = int(contract["order"]["id"])
    customer_id = None
    created_at = now_tashkent().isoformat(timespec="seconds")
    with get_db() as db:
        customer_row = db.execute("SELECT customer_id FROM orders WHERE id=?", (order_id,)).fetchone()
        customer_id = customer_row["customer_id"]
        if ticket_id:
            ticket = db.execute(
                "SELECT id, escalated_to_admin FROM support_tickets WHERE id=? AND order_id=?",
                (ticket_id, order_id),
            ).fetchone()
            if ticket is None:
                return jsonify(ok=False, message="Murojaat topilmadi."), 404
            admin_unread = 1 if ticket["escalated_to_admin"] else 0
            db.execute(
                """UPDATE support_tickets
                   SET status='Mijoz javob berdi', manager_unread=1, admin_unread=?,
                       customer_unread=0, last_message_at=?, closed_at=NULL
                   WHERE id=?""",
                (admin_unread, created_at, ticket_id),
            )
        else:
            priority = "Muhim" if topic in {"Shikoyat yoki taklif", "To‘lov masalasi"} else "Oddiy"
            admin_unread = 1 if topic == "Shikoyat yoki taklif" else 0
            cursor = db.execute(
                """INSERT INTO support_tickets
                   (order_id, customer_id, topic, subject, status, priority,
                    assigned_role, escalated_to_admin, manager_unread, admin_unread,
                    last_message_at, created_at)
                   VALUES (?, ?, ?, ?, 'Yangi', ?, 'manager', 0, 1, ?, ?, ?)""",
                (order_id, customer_id, topic, subject, priority, admin_unread, created_at, created_at),
            )
            ticket_id = cursor.lastrowid

        cursor = db.execute(
            """INSERT INTO support_messages
               (ticket_id, sender_role, sender_name, body, attachment_name, attachment_path, created_at)
               VALUES (?, 'customer', ?, ?, ?, ?, ?)""",
            (ticket_id, contract["customer"]["full_name"], body, attachment_name, attachment_path, created_at),
        )
        message_id = cursor.lastrowid
        message = db.execute("SELECT * FROM support_messages WHERE id=?", (message_id,)).fetchone()

        ai_created_at = (datetime.fromisoformat(created_at) + timedelta(seconds=1)).isoformat(timespec="seconds")
        ai_body = build_ai_acknowledgement(contract["customer"]["full_name"], topic, body)
        ai_cursor = db.execute(
            """INSERT INTO support_messages
               (ticket_id, sender_role, sender_name, body, created_at)
               VALUES (?, 'ai', 'Mebel360 AI', ?, ?)""",
            (ticket_id, ai_body, ai_created_at),
        )
        ai_message = db.execute(
            "SELECT * FROM support_messages WHERE id=?", (ai_cursor.lastrowid,)
        ).fetchone()
        db.execute(
            "UPDATE support_tickets SET last_message_at=? WHERE id=?",
            (ai_created_at, ticket_id),
        )

    return jsonify(
        ok=True,
        message="Mebel360 AI xabaringizni qabul qildi va menejerga yetkazdi.",
        ticket_id=ticket_id,
        data=serialize_message(message),
        ai_data=serialize_message(ai_message),
    )


@app.post("/api/customer/support/<int:ticket_id>/escalate")
def customer_support_escalate(ticket_id: int):
    contract = current_customer_context()
    if contract is None:
        return jsonify(ok=False, message="Mijoz buyurtmasi topilmadi."), 404
    order_id = int(contract["order"]["id"])
    created_at = now_tashkent().isoformat(timespec="seconds")
    with get_db() as db:
        ticket = db.execute("SELECT id, escalated_to_admin FROM support_tickets WHERE id=? AND order_id=?", (ticket_id, order_id)).fetchone()
        if ticket is None:
            return jsonify(ok=False, message="Murojaat topilmadi."), 404
        if ticket["escalated_to_admin"]:
            return jsonify(ok=True, message="Murojaat allaqachon rahbarga yuborilgan.")
        db.execute(
            """UPDATE support_tickets SET escalated_to_admin=1, assigned_role='admin',
               priority='Muhim', status='Rahbarga yuborildi', admin_unread=1,
               last_message_at=? WHERE id=?""",
            (created_at, ticket_id),
        )
        db.execute(
            """INSERT INTO support_messages
               (ticket_id, sender_role, sender_name, body, created_at)
               VALUES (?, 'system', 'Mebel360°', 'Mijoz murojaatni rahbar ko‘rib chiqishini so‘radi.', ?)""",
            (ticket_id, created_at),
        )
    return jsonify(ok=True, message="Murojaat rahbar kabinetiga yuborildi.")


@app.get("/api/staff/support/inbox")
def staff_support_inbox():
    role = require_staff()
    if role is None:
        return jsonify(ok=False, message="Bu bo‘limga kirish huquqi mavjud emas."), 403
    requested_ticket = request.args.get("ticket_id", type=int)
    with get_db() as db:
        apply_overdue_escalations(db)
        rows = get_ticket_rows(db)
        ticket_id = requested_ticket or (rows[0]["id"] if rows else None)
        active = next((row for row in rows if row["id"] == ticket_id), None)
        messages = []
        if active is not None:
            unread_column = "admin_unread" if role == "admin" else "manager_unread"
            db.execute(f"UPDATE support_tickets SET {unread_column}=0 WHERE id=?", (active["id"],))
            rows = get_ticket_rows(db)
            active = next((row for row in rows if row["id"] == ticket_id), None)
            message_rows = db.execute("SELECT * FROM support_messages WHERE ticket_id=? ORDER BY id", (ticket_id,)).fetchall()
            messages = [serialize_message(row) for row in message_rows]

        counts = {
            "all": len(rows),
            "new": sum(1 for row in rows if row["status"] in {"Yangi", "Mijoz javob berdi", "Rahbarga yuborildi"}),
            "open": sum(1 for row in rows if row["status"] not in {"Yopildi", "Hal qilindi"}),
            "escalated": sum(1 for row in rows if row["escalated_to_admin"]),
            "unread": sum(1 for row in rows if row["admin_unread" if role == "admin" else "manager_unread"]),
        }
    return jsonify(ok=True, data={"role": role, "counts": counts, "tickets": [serialize_ticket(row) for row in rows], "active_ticket": serialize_ticket(active) if active else None, "messages": messages})


@app.post("/api/staff/support/<int:ticket_id>/messages")
def staff_support_message(ticket_id: int):
    role = require_staff()
    if role is None:
        return jsonify(ok=False, message="Bu bo‘limga kirish huquqi mavjud emas."), 403
    payload = request.form if request.form else (request.get_json(silent=True) or {})
    body = str(payload.get("body", "")).strip()
    try:
        attachment_name, attachment_path = save_attachment(request.files.get("attachment"))
    except ValueError as exc:
        return jsonify(ok=False, message=str(exc)), 400
    if not body and not attachment_path:
        return jsonify(ok=False, message="Javob matnini yozing yoki fayl biriktiring."), 400
    created_at = now_tashkent().isoformat(timespec="seconds")
    sender_name = "Rahbar" if role == "admin" else "Menejer"
    with get_db() as db:
        ticket = db.execute("SELECT id FROM support_tickets WHERE id=?", (ticket_id,)).fetchone()
        if ticket is None:
            return jsonify(ok=False, message="Murojaat topilmadi."), 404
        db.execute(
            """UPDATE support_tickets SET status='Javob berildi', customer_unread=1,
               manager_unread=0, admin_unread=0, last_message_at=?, closed_at=NULL WHERE id=?""",
            (created_at, ticket_id),
        )
        cursor = db.execute(
            """INSERT INTO support_messages
               (ticket_id, sender_role, sender_name, body, attachment_name, attachment_path, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (ticket_id, role, sender_name, body, attachment_name, attachment_path, created_at),
        )
        row = db.execute("SELECT * FROM support_messages WHERE id=?", (cursor.lastrowid,)).fetchone()
    return jsonify(ok=True, message="Javob mijozga yuborildi.", data=serialize_message(row))


@app.post("/api/staff/support/<int:ticket_id>/status")
def staff_support_status(ticket_id: int):
    role = require_staff()
    if role is None:
        return jsonify(ok=False, message="Bu bo‘limga kirish huquqi mavjud emas."), 403
    data = request.get_json(silent=True) or {}
    status = str(data.get("status", "")).strip()
    allowed = {"Yangi", "Ko‘rib chiqilmoqda", "Javob berildi", "Hal qilindi", "Yopildi", "Rahbarga yuborildi"}
    if status not in allowed:
        return jsonify(ok=False, message="Holat noto‘g‘ri tanlandi."), 400
    closed_at = now_tashkent().isoformat(timespec="seconds") if status in {"Hal qilindi", "Yopildi"} else None
    with get_db() as db:
        result = db.execute("UPDATE support_tickets SET status=?, closed_at=? WHERE id=?", (status, closed_at, ticket_id))
        if result.rowcount == 0:
            return jsonify(ok=False, message="Murojaat topilmadi."), 404
    return jsonify(ok=True, message="Murojaat holati yangilandi.")


@app.post("/api/staff/support/<int:ticket_id>/escalate")
def staff_support_escalate(ticket_id: int):
    role = require_staff()
    if role is None:
        return jsonify(ok=False, message="Bu bo‘limga kirish huquqi mavjud emas."), 403
    created_at = now_tashkent().isoformat(timespec="seconds")
    with get_db() as db:
        result = db.execute(
            """UPDATE support_tickets SET escalated_to_admin=1, assigned_role='admin',
               priority='Muhim', status='Rahbarga yuborildi', admin_unread=1,
               last_message_at=? WHERE id=?""",
            (created_at, ticket_id),
        )
        if result.rowcount == 0:
            return jsonify(ok=False, message="Murojaat topilmadi."), 404
    return jsonify(ok=True, message="Murojaat rahbarga yuborildi.")


@app.get("/health")
def health():
    return jsonify(status="ok", app="Mebel360", contract_module=True, support_module=True)


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
