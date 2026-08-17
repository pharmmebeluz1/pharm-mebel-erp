from flask import Flask, render_template, request, redirect, url_for, flash
import sqlite3
from datetime import datetime

app = Flask(__name__)
app.secret_key = "mebel360-rating-secret"
DB = "mebel360.db"

GOOD_TAGS = [
    ("muomala", "Xushmuomala xodim"),
    ("sifat", "Ish sifati"),
    ("tozalik", "Ishning tozaligi"),
    ("vaqt", "Vaqtida kelishi"),
    ("etibor", "Ishga e’tibori"),
    ("tezkorlik", "Tezkorligi"),
]

BAD_TAGS = [
    ("qopol", "Qo‘pol muomala"),
    ("kechikish", "Kechikib kelish"),
    ("sifat_past", "Ish sifati past"),
    ("tartibsizlik", "Ishdagi tartibsizlik"),
    ("aloqa", "Bog‘lanish qiyin"),
    ("boshqa", "Boshqa muammo"),
]

SUPPORT_TOPICS = [
    ("buyurtma", "Buyurtma holati"),
    ("tolov", "To‘lov muammosi"),
    ("yetkazish", "Yetkazib berish"),
    ("ornatish", "O‘rnatish xizmati"),
    ("shikoyat", "Shikoyat / taklif"),
    ("operator", "Operator bilan bog‘lanish"),
]

def con():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    c = con()
    c.execute("""
        CREATE TABLE IF NOT EXISTS employees(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS orders(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_code TEXT UNIQUE NOT NULL,
            customer_name TEXT NOT NULL,
            employee_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'Jarayonda',
            visit_at TEXT,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS ratings(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER UNIQUE NOT NULL,
            employee_id INTEGER NOT NULL,
            stars INTEGER NOT NULL,
            tags TEXT,
            comment TEXT,
            score INTEGER NOT NULL,
            created_at TEXT NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS support_requests(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER,
            topic TEXT,
            message TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Yangi',
            created_at TEXT NOT NULL
        )
    """)

    if c.execute("SELECT COUNT(*) AS n FROM employees").fetchone()["n"] == 0:
        c.executemany("INSERT INTO employees(name) VALUES(?)", [
            ("Jamshid",), ("Sardor",), ("Akmal",)
        ])

    if c.execute("SELECT COUNT(*) AS n FROM orders").fetchone()["n"] == 0:
        c.executemany("""
            INSERT INTO orders(order_code, customer_name, employee_id, status, visit_at)
            VALUES(?,?,?,?,?)
        """, [
            ("AD-101", "Mijoz", 1, "Yakunlandi", "13 avgust 13:36"),
            ("AD-102", "Mijoz 2", 2, "Yakunlandi", "14 avgust 15:20"),
            ("AD-103", "Mijoz 3", 3, "Jarayonda", "15 avgust 11:10"),
        ])

    c.commit()
    c.close()

def score_for(stars, tag_count):
    base = stars * 20
    if stars >= 4:
        return base + tag_count * 2
    # past bahoda bonus yo‘q; reyting yulduzga qarab tushadi
    return base

@app.route("/")
def home():
    return redirect(url_for("rate_order", order_code="AD-101"))

@app.route("/baho/<order_code>", methods=["GET", "POST"])
def rate_order(order_code):
    init_db()
    c = con()
    order = c.execute("""
        SELECT o.*, e.name AS employee_name
        FROM orders o
        JOIN employees e ON e.id=o.employee_id
        WHERE o.order_code=?
    """, (order_code,)).fetchone()

    if not order:
        c.close()
        return "Buyurtma topilmadi", 404

    if order["status"] != "Yakunlandi":
        c.close()
        return render_template(
            "message.html",
            title="Baholash hali ochilmagan",
            message="Buyurtma yakunlangandan keyin baholash mumkin."
        )

    existing = c.execute("SELECT * FROM ratings WHERE order_id=?", (order["id"],)).fetchone()
    if existing:
        c.close()
        return render_template(
            "message.html",
            title="Rahmat! ⭐",
            message="Bu buyurtma avval baholangan. Har bir buyurtmaga faqat bir marta baho beriladi."
        )

    if request.method == "POST":
        try:
            stars = int(request.form.get("stars", "0"))
        except ValueError:
            stars = 0

        if stars < 1 or stars > 5:
            flash("Iltimos, yulduzcha tanlang.")
            c.close()
            return redirect(url_for("rate_order", order_code=order_code))

        selected = request.form.getlist("tags")
        allowed = {k for k, _ in (GOOD_TAGS if stars >= 4 else BAD_TAGS)}
        selected = [x for x in selected if x in allowed]

        comment = request.form.get("comment", "").strip()[:500]
        score = score_for(stars, len(selected))

        c.execute("""
            INSERT INTO ratings(order_id, employee_id, stars, tags, comment, score, created_at)
            VALUES(?,?,?,?,?,?,?)
        """, (
            order["id"], order["employee_id"], stars,
            ",".join(selected), comment, score,
            datetime.now().isoformat(timespec="seconds")
        ))
        c.commit()
        c.close()

        return render_template(
            "message.html",
            title="Bahoyingiz qabul qilindi ⭐",
            message=f"Rahmat! Xodim reytingiga {score} ball qo‘shildi."
        )

    c.close()
    return render_template(
        "rating.html",
        order=order,
        good_tags=GOOD_TAGS,
        bad_tags=BAD_TAGS
    )

@app.route("/support/<order_code>", methods=["GET", "POST"])
def support(order_code):
    c = con()
    order = c.execute("SELECT * FROM orders WHERE order_code=?", (order_code,)).fetchone()
    if not order:
        c.close()
        return "Buyurtma topilmadi", 404

    if request.method == "POST":
        topic = request.form.get("topic", "")
        allowed = {k for k, _ in SUPPORT_TOPICS}
        if topic not in allowed:
            topic = "shikoyat"

        message = request.form.get("message", "").strip()[:500]
        if not message:
            flash("Murojaat matnini yozing.")
            c.close()
            return redirect(url_for("support", order_code=order_code))

        c.execute("""
            INSERT INTO support_requests(order_id, topic, message, created_at)
            VALUES(?,?,?,?)
        """, (
            order["id"], topic, message,
            datetime.now().isoformat(timespec="seconds")
        ))
        c.commit()
        c.close()

        return render_template(
            "message.html",
            title="Murojaat yuborildi",
            message="Qo‘llab-quvvatlash markazi murojaatingizni qabul qildi."
        )

    c.close()
    return render_template("support.html", order=order, topics=SUPPORT_TOPICS)

@app.route("/rahbar/reytin")
def leaderboard():
    c = con()
    rows = c.execute("""
        SELECT
            e.id,
            e.name,
            COUNT(r.id) AS reviews,
            ROUND(COALESCE(AVG(r.stars),0),2) AS avg_stars,
            COALESCE(SUM(r.score),0) AS score
        FROM employees e
        LEFT JOIN ratings r ON r.employee_id=e.id
        GROUP BY e.id,e.name
        ORDER BY score DESC, avg_stars DESC
    """).fetchall()
    c.close()
    return render_template("leaderboard.html", rows=rows)

@app.route("/rahbar/murojaatlar")
def support_admin():
    c = con()
    rows = c.execute("""
        SELECT s.*, o.order_code
        FROM support_requests s
        LEFT JOIN orders o ON o.id=s.order_id
        ORDER BY s.id DESC
    """).fetchall()
    c.close()
    return render_template("support_admin.html", rows=rows)

init_db()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
