from __future__ import annotations

import os
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

from flask import Flask, jsonify, render_template, request, session

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("MEBEL360_DB", BASE_DIR / "mebel360_demo.db"))
TASHKENT = timezone(timedelta(hours=5), name="Asia/Tashkent")

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "mebel360-demo-secret-change-me")
app.config["JSON_AS_ASCII"] = False


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
    return jsonify(
        ok=True,
        role=role,
        next_screen="contractScreen" if role == "customer" else None,
        message="Rol tanlandi.",
    )


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


@app.get("/health")
def health():
    return jsonify(status="ok", app="Mebel360", contract_module=True)


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
