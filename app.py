# -*- coding: utf-8 -*-
import csv
import io
import os
import sqlite3
import secrets
from datetime import datetime, date
from flask import Flask, jsonify, request, render_template_string, Response, session, redirect, url_for, send_file, flash
import qrcode
from werkzeug.security import generate_password_hash, check_password_hash
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.secret_key = os.environ.get("PHARM_ERP_SECRET", "pharm-mebel-change-this-secret")
DB_NAME = os.environ.get("PHARM_ERP_DB", "pharm_mebel_erp_pro.db")


def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS ishchilar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ism TEXT NOT NULL,
        familiya TEXT DEFAULT '',
        telefon TEXT DEFAULT '',
        lavozim TEXT DEFAULT '',
        ishga_kirgan_sana TEXT DEFAULT '',
        staj_yil REAL DEFAULT 0,
        kunlik_stavka REAL DEFAULT 0,
        oylik_maosh REAL DEFAULT 0,
        sifat_ball REAL DEFAULT 5,
        tezlik_ball REAL DEFAULT 5,
        intizom_ball REAL DEFAULT 5,
        izoh TEXT DEFAULT '',
        faol INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS ish_turlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomi TEXT UNIQUE NOT NULL,
        kategoriya TEXT DEFAULT 'Ish',
        birlik TEXT DEFAULT 'dona',
        standart_narx REAL DEFAULT 0,
        faol INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS keldi_ketdi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        keldi_vaqti TEXT NOT NULL,
        ketdi_vaqti TEXT NOT NULL,
        ish_soatlari REAL DEFAULT 0,
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS ish_natijalari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        ish_turi_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        miqdor REAL DEFAULT 0,
        birlik_narxi REAL DEFAULT 0,
        jami_haq REAL DEFAULT 0,
        buyurtma_kodi TEXT DEFAULT '',
        izoh TEXT DEFAULT '',
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE,
        FOREIGN KEY (ish_turi_id) REFERENCES ish_turlari(id)
    );

    CREATE TABLE IF NOT EXISTS tolovlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        miqdor REAL DEFAULT 0,
        turi TEXT DEFAULT 'Avans',
        tavsifi TEXT DEFAULT '',
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS jarimalar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        miqdor REAL DEFAULT 0,
        sababi TEXT DEFAULT '',
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS buyurtmalar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kod TEXT UNIQUE NOT NULL,
        mijoz TEXT NOT NULL,
        telefon TEXT DEFAULT '',
        manzil TEXT DEFAULT '',
        mahsulot TEXT DEFAULT '',
        umumiy_narx REAL DEFAULT 0,
        oldindan_tolov REAL DEFAULT 0,
        tugash_sana TEXT DEFAULT '',
        holat TEXT DEFAULT 'Yangi',
        izoh TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS buyurtma_bosqichlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        bosqich TEXT NOT NULL,
        bajarildi INTEGER DEFAULT 0,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE,
        UNIQUE(buyurtma_id, bosqich)
    );

    CREATE TABLE IF NOT EXISTS ombor (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomi TEXT UNIQUE NOT NULL,
        kategoriya TEXT DEFAULT '',
        birlik TEXT DEFAULT 'dona',
        qoldiq REAL DEFAULT 0,
        min_qoldiq REAL DEFAULT 0,
        narx REAL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS ombor_harakat (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        material_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        turi TEXT NOT NULL,
        miqdor REAL NOT NULL,
        buyurtma_kodi TEXT DEFAULT '',
        izoh TEXT DEFAULT '',
        FOREIGN KEY (material_id) REFERENCES ombor(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS safarlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        mashina TEXT DEFAULT '',
        qayerdan TEXT DEFAULT '',
        qayerga TEXT DEFAULT '',
        masofa_km REAL DEFAULT 0,
        sabab TEXT DEFAULT '',
        yonilgi REAL DEFAULT 0,
        xarajat REAL DEFAULT 0,
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS xarajatlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sana TEXT NOT NULL,
        kategoriya TEXT NOT NULL,
        miqdor REAL DEFAULT 0,
        tavsifi TEXT DEFAULT '',
        buyurtma_kodi TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS bonuslar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        miqdor REAL DEFAULT 0,
        sababi TEXT DEFAULT '',
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS ishchi_holatlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        turi TEXT NOT NULL,
        izoh TEXT DEFAULT '',
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS buyurtma_tolovlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        sana TEXT NOT NULL,
        miqdor REAL DEFAULT 0,
        turi TEXT DEFAULT 'To‘lov',
        izoh TEXT DEFAULT '',
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS tayyor_mahsulot (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomi TEXT NOT NULL,
        kodi TEXT DEFAULT '',
        rang TEXT DEFAULT '',
        miqdor REAL DEFAULT 0,
        birlik TEXT DEFAULT 'dona',
        narx REAL DEFAULT 0,
        izoh TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sana_vaqt TEXT DEFAULT CURRENT_TIMESTAMP,
        amal TEXT NOT NULL,
        tafsilot TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS buyurtma_hujjatlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        nomi TEXT NOT NULL,
        fayl_nomi TEXT NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS ishchi_akkauntlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER UNIQUE,
        telefon TEXT UNIQUE NOT NULL,
        login TEXT UNIQUE,
        parol_hash TEXT DEFAULT '',
        tasdiqlangan INTEGER DEFAULT 0,
        admin_tasdiq INTEGER DEFAULT 0,
        faol INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS ishchi_otp (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        telefon TEXT NOT NULL,
        kod_hash TEXT NOT NULL,
        muddati TEXT NOT NULL,
        ishlatildi INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS ishchi_topshiriqlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER NOT NULL,
        buyurtma_kodi TEXT DEFAULT '',
        ish_turi TEXT NOT NULL,
        tavsif TEXT DEFAULT '',
        holat TEXT DEFAULT 'Yangi',
        progress INTEGER DEFAULT 0,
        sana TEXT NOT NULL,
        tugash_sana TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE CASCADE
    );
    """)

    ishlar = [
        ("Kesish","Ishlab chiqarish","dona"),("Bo‘yash","Ishlab chiqarish","dona"),
        ("Yig‘ish","Ishlab chiqarish","dona"),("Roverda ishlash","Stanok","soat"),
        ("Razmer olish","Xizmat","buyurtma"),("Kromka urish","Ishlab chiqarish","metr"),
        ("Sayqalash","Ishlab chiqarish","dona"),("Teshish","Ishlab chiqarish","dona"),
        ("Lazer","Stanok","soat"),("Sexda yig‘ish","Ishlab chiqarish","dona"),
        ("Sex ishlari","Umumiy","soat"),("Yangi loyihalar","Loyiha","loyiha"),
        ("Frezalash","Stanok","soat"),("Yumshoq mebel","Ishlab chiqarish","dona"),
        ("Shofyorlik","Transport","safar"),("Ish vaqtida hordiq","Tanaffus","daqiqa"),
        ("Bozorga tushish","Xizmat","safar"),("Qadoqlash","Ishlab chiqarish","dona"),
        ("Oyna o‘rnatish","Ishlab chiqarish","dona"),("Oyna teshish","Ishlab chiqarish","dona"),
        ("Tumba yig‘ish","Mebel turi","dona"),("Oyoq kiyim shkafi yig‘ish","Mebel turi","dona"),
        ("Kravot yig‘ish","Mebel turi","dona"),("Shkaf yig‘ish","Mebel turi","dona"),
        ("Kuxniy yig‘ish","Mebel turi","dona"),("Kamod yig‘ish","Mebel turi","dona"),
        ("Buyurtmani raspildan olib kelish","Transport","safar"),
        ("Abed","Tanaffus","daqiqa"),("Tanaffus","Tanaffus","daqiqa"),
        ("Namoz vaqti","Tanaffus","daqiqa")
    ]
    conn.executemany(
        "INSERT OR IGNORE INTO ish_turlari(nomi,kategoriya,birlik) VALUES(?,?,?)", ishlar
    )

    materials = [
        ("MDF 18 mm","Plita","list",0,5,0),
        ("MDF 16 mm","Plita","list",0,5,0),
        ("Kromka","Kromka","metr",0,100,0),
        ("Furnitura","Furnitura","dona",0,50,0),
        ("Bo‘yoq","Bo‘yoq","kg",0,10,0),
        ("Oyna","Oyna","m²",0,5,0),
        ("Akril","Plita","list",0,3,0)
    ]
    conn.executemany(
        "INSERT OR IGNORE INTO ombor(nomi,kategoriya,birlik,qoldiq,min_qoldiq,narx) VALUES(?,?,?,?,?,?)",
        materials
    )
    # V3 migratsiya: mijoz kuzatuv tokeni
    cols=[r[1] for r in conn.execute("PRAGMA table_info(buyurtmalar)").fetchall()]
    if "tracking_token" not in cols:
        conn.execute("ALTER TABLE buyurtmalar ADD COLUMN tracking_token TEXT DEFAULT ''")
    for row in conn.execute("SELECT id FROM buyurtmalar WHERE tracking_token IS NULL OR tracking_token='' ").fetchall():
        conn.execute("UPDATE buyurtmalar SET tracking_token=? WHERE id=?",(secrets.token_urlsafe(8),row[0]))
    conn.commit()
    conn.close()


def jdata():
    d = request.get_json(silent=True)
    return d if isinstance(d, dict) else {}


def calc_hours(start_text, end_text):
    s = datetime.strptime(start_text, "%H:%M")
    e = datetime.strptime(end_text, "%H:%M")
    sec = (e - s).total_seconds()
    if sec < 0:
        sec += 86400
    return round(sec / 3600, 2)


def log_action(amal, tafsilot=''):
    try:
        c=get_db(); c.execute("INSERT INTO audit_log(amal,tafsilot) VALUES(?,?)",(amal,tafsilot)); c.commit(); c.close()
    except Exception:
        pass


@app.before_request
def require_login():
    public_endpoints = {
        "login", "static", "public_track", "order_qr",
        "worker_register", "worker_verify", "worker_login", "worker_logout"
    }
    if request.endpoint in public_endpoints or request.path.startswith('/ishchi/public/'):
        return None
    if request.path.startswith('/ishchi/'):
        if not session.get("worker_account_id"):
            return redirect(url_for("worker_login"))
        return None
    if not session.get("logged_in"):
        return redirect(url_for("login"))


@app.route("/login", methods=["GET","POST"])
def login():
    error=''
    if request.method=='POST':
        user=request.form.get('user','')
        password=request.form.get('password','')
        if user==os.environ.get('PHARM_ERP_USER','admin') and password==os.environ.get('PHARM_ERP_PASSWORD','12345'):
            session['logged_in']=True
            session['user']=user
            return redirect(url_for('home'))
        error='Login yoki parol xato'
    return render_template_string(LOGIN_HTML,error=error)


@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for('login'))


@app.route("/")
def home():
    return render_template_string(HTML)


# ---------- ISHCHILAR ----------
@app.route("/api/ishchilar", methods=["GET"])
def workers_get():
    c = get_db()
    rows = c.execute("SELECT * FROM ishchilar WHERE faol=1 ORDER BY ism,familiya").fetchall()
    c.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/ishchilar", methods=["POST"])
def workers_add():
    d = jdata()
    if not str(d.get("ism","")).strip():
        return jsonify({"message":"Ism kiritilmagan"}),400
    c = get_db()
    c.execute("""INSERT INTO ishchilar
    (ism,familiya,telefon,lavozim,ishga_kirgan_sana,staj_yil,kunlik_stavka,oylik_maosh,
     sifat_ball,tezlik_ball,intizom_ball,izoh)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",(
        d.get("ism","").strip(),d.get("familiya","").strip(),d.get("telefon","").strip(),
        d.get("lavozim","").strip(),d.get("ishga_kirgan_sana",""),
        float(d.get("staj_yil") or 0),float(d.get("kunlik_stavka") or 0),
        float(d.get("oylik_maosh") or 0),float(d.get("sifat_ball") or 5),
        float(d.get("tezlik_ball") or 5),float(d.get("intizom_ball") or 5),
        d.get("izoh","").strip()
    ))
    c.commit(); c.close()
    return jsonify({"status":"ok"})


@app.route("/api/ishchilar/<int:i>", methods=["DELETE"])
def workers_delete(i):
    c=get_db(); c.execute("UPDATE ishchilar SET faol=0 WHERE id=?",(i,)); c.commit(); c.close()
    return jsonify({"status":"ok"})


# ---------- KELDI KETDI ----------
@app.route("/api/keldi-ketdi", methods=["GET"])
def attendance_get():
    c=get_db()
    rows=c.execute("""SELECT k.*,i.ism,i.familiya FROM keldi_ketdi k
    JOIN ishchilar i ON i.id=k.ishchi_id ORDER BY sana DESC,id DESC LIMIT 200""").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/keldi-ketdi", methods=["POST"])
def attendance_add():
    d=jdata()
    try:
        hours=calc_hours(d["keldi_vaqti"],d["ketdi_vaqti"])
        c=get_db()
        c.execute("""INSERT INTO keldi_ketdi(ishchi_id,sana,keldi_vaqti,ketdi_vaqti,ish_soatlari)
        VALUES(?,?,?,?,?)""",(int(d["ishchi_id"]),d["sana"],d["keldi_vaqti"],d["ketdi_vaqti"],hours))
        c.commit(); c.close()
        return jsonify({"status":"ok"})
    except Exception as e:
        return jsonify({"message":str(e)}),400


# ---------- ISH NATIJALARI ----------
@app.route("/api/ish-turlari")
def work_types_get():
    c=get_db(); rows=c.execute("SELECT * FROM ish_turlari WHERE faol=1 ORDER BY kategoriya,nomi").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/natijalar", methods=["GET"])
def results_get():
    c=get_db()
    rows=c.execute("""SELECT n.*,i.ism,i.familiya,t.nomi ish_turi,t.birlik
    FROM ish_natijalari n JOIN ishchilar i ON i.id=n.ishchi_id
    JOIN ish_turlari t ON t.id=n.ish_turi_id
    ORDER BY sana DESC,n.id DESC LIMIT 250""").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/natijalar", methods=["POST"])
def results_add():
    d=jdata()
    try:
        qty=float(d.get("miqdor") or 0); price=float(d.get("birlik_narxi") or 0)
        c=get_db()
        c.execute("""INSERT INTO ish_natijalari
        (ishchi_id,ish_turi_id,sana,miqdor,birlik_narxi,jami_haq,buyurtma_kodi,izoh)
        VALUES(?,?,?,?,?,?,?,?)""",(int(d["ishchi_id"]),int(d["ish_turi_id"]),d["sana"],
        qty,price,round(qty*price,2),d.get("buyurtma_kodi",""),d.get("izoh","")))
        c.commit(); c.close()
        return jsonify({"status":"ok"})
    except Exception as e:
        return jsonify({"message":str(e)}),400


# ---------- TOLOV VA JARIMA ----------
@app.route("/api/tolovlar", methods=["GET","POST"])
def payments():
    if request.method=="POST":
        d=jdata(); c=get_db()
        c.execute("INSERT INTO tolovlar(ishchi_id,sana,miqdor,turi,tavsifi) VALUES(?,?,?,?,?)",
                  (int(d["ishchi_id"]),d["sana"],float(d.get("miqdor") or 0),d.get("turi","Avans"),d.get("tavsifi","")))
        c.commit(); c.close(); return jsonify({"status":"ok"})
    c=get_db(); rows=c.execute("""SELECT t.*,i.ism,i.familiya FROM tolovlar t
    JOIN ishchilar i ON i.id=t.ishchi_id ORDER BY sana DESC,t.id DESC LIMIT 200""").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/jarimalar", methods=["GET","POST"])
def penalties():
    if request.method=="POST":
        d=jdata(); c=get_db()
        c.execute("INSERT INTO jarimalar(ishchi_id,sana,miqdor,sababi) VALUES(?,?,?,?)",
                  (int(d["ishchi_id"]),d["sana"],float(d.get("miqdor") or 0),d.get("sababi","")))
        c.commit(); c.close(); return jsonify({"status":"ok"})
    c=get_db(); rows=c.execute("""SELECT j.*,i.ism,i.familiya FROM jarimalar j
    JOIN ishchilar i ON i.id=j.ishchi_id ORDER BY sana DESC,j.id DESC LIMIT 200""").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


# ---------- BUYURTMALAR ----------
STAGES=["Razmer","Loyiha","Kesish","Rover","Kromka","Teshish","Bo‘yash","Yig‘ish","Oyna","Qadoqlash","Yetkazildi"]

@app.route("/api/buyurtmalar", methods=["GET","POST"])
def orders():
    if request.method=="POST":
        d=jdata(); c=get_db()
        try:
            cur=c.execute("""INSERT INTO buyurtmalar(kod,mijoz,telefon,manzil,mahsulot,umumiy_narx,
            oldindan_tolov,tugash_sana,holat,izoh,tracking_token) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            (d["kod"],d["mijoz"],d.get("telefon",""),d.get("manzil",""),d.get("mahsulot",""),
             float(d.get("umumiy_narx") or 0),float(d.get("oldindan_tolov") or 0),
             d.get("tugash_sana",""),d.get("holat","Yangi"),d.get("izoh",""),secrets.token_urlsafe(8)))
            oid=cur.lastrowid
            c.executemany("INSERT INTO buyurtma_bosqichlari(buyurtma_id,bosqich) VALUES(?,?)",
                          [(oid,s) for s in STAGES])
            c.commit(); c.close(); return jsonify({"status":"ok"})
        except Exception as e:
            c.close(); return jsonify({"message":str(e)}),400
    c=get_db()
    rows=c.execute("""SELECT *,ROUND(umumiy_narx-oldindan_tolov,2) qoldiq
    FROM buyurtmalar ORDER BY id DESC""").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/buyurtma/<int:oid>/bosqichlar")
def order_stages(oid):
    c=get_db(); rows=c.execute("SELECT * FROM buyurtma_bosqichlari WHERE buyurtma_id=? ORDER BY id",(oid,)).fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/buyurtma-bosqich/<int:sid>", methods=["POST"])
def stage_toggle(sid):
    d=jdata(); c=get_db()
    c.execute("UPDATE buyurtma_bosqichlari SET bajarildi=? WHERE id=?",(1 if d.get("bajarildi") else 0,sid))
    c.commit(); c.close(); return jsonify({"status":"ok"})


# ---------- OMBOR ----------
@app.route("/api/ombor", methods=["GET"])
def stock_get():
    c=get_db(); rows=c.execute("SELECT * FROM ombor ORDER BY kategoriya,nomi").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/ombor-harakat", methods=["POST"])
def stock_move():
    d=jdata()
    try:
        qty=float(d.get("miqdor") or 0); typ=d.get("turi","Kirim")
        delta=qty if typ=="Kirim" else -qty
        c=get_db()
        c.execute("UPDATE ombor SET qoldiq=qoldiq+? WHERE id=?",(delta,int(d["material_id"])))
        c.execute("""INSERT INTO ombor_harakat(material_id,sana,turi,miqdor,buyurtma_kodi,izoh)
        VALUES(?,?,?,?,?,?)""",(int(d["material_id"]),d["sana"],typ,qty,d.get("buyurtma_kodi",""),d.get("izoh","")))
        c.commit(); c.close(); return jsonify({"status":"ok"})
    except Exception as e:
        return jsonify({"message":str(e)}),400


# ---------- SAFAR ----------
@app.route("/api/safarlar", methods=["GET","POST"])
def trips():
    if request.method=="POST":
        d=jdata(); c=get_db()
        c.execute("""INSERT INTO safarlar(ishchi_id,sana,mashina,qayerdan,qayerga,masofa_km,sabab,yonilgi,xarajat)
        VALUES(?,?,?,?,?,?,?,?,?)""",(int(d["ishchi_id"]),d["sana"],d.get("mashina",""),d.get("qayerdan",""),
        d.get("qayerga",""),float(d.get("masofa_km") or 0),d.get("sabab",""),
        float(d.get("yonilgi") or 0),float(d.get("xarajat") or 0)))
        c.commit(); c.close(); return jsonify({"status":"ok"})
    c=get_db(); rows=c.execute("""SELECT s.*,i.ism,i.familiya FROM safarlar s
    JOIN ishchilar i ON i.id=s.ishchi_id ORDER BY sana DESC,s.id DESC LIMIT 200""").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


# ---------- QO‘SHIMCHA MODULLAR ----------
@app.route("/api/xarajatlar", methods=["GET","POST"])
def expenses():
    if request.method=="POST":
        d=jdata(); c=get_db()
        c.execute("INSERT INTO xarajatlar(sana,kategoriya,miqdor,tavsifi,buyurtma_kodi) VALUES(?,?,?,?,?)",
                  (d["sana"],d.get("kategoriya","Boshqa"),float(d.get("miqdor") or 0),d.get("tavsifi",""),d.get("buyurtma_kodi","")))
        c.commit(); c.close(); log_action("Xarajat qo‘shildi", d.get("kategoriya","")); return jsonify({"status":"ok"})
    c=get_db(); rows=c.execute("SELECT * FROM xarajatlar ORDER BY sana DESC,id DESC LIMIT 300").fetchall(); c.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/bonuslar", methods=["GET","POST"])
def bonuses():
    if request.method=="POST":
        d=jdata(); c=get_db(); c.execute("INSERT INTO bonuslar(ishchi_id,sana,miqdor,sababi) VALUES(?,?,?,?)",
            (int(d["ishchi_id"]),d["sana"],float(d.get("miqdor") or 0),d.get("sababi","")))
        c.commit(); c.close(); log_action("Bonus qo‘shildi", str(d.get("ishchi_id"))); return jsonify({"status":"ok"})
    c=get_db(); rows=c.execute("""SELECT b.*,i.ism,i.familiya FROM bonuslar b JOIN ishchilar i ON i.id=b.ishchi_id
        ORDER BY sana DESC,b.id DESC LIMIT 200""").fetchall(); c.close(); return jsonify([dict(r) for r in rows])

@app.route("/api/ishchi-holatlari", methods=["GET","POST"])
def worker_statuses():
    if request.method=="POST":
        d=jdata(); c=get_db(); c.execute("INSERT INTO ishchi_holatlari(ishchi_id,sana,turi,izoh) VALUES(?,?,?,?)",
            (int(d["ishchi_id"]),d["sana"],d.get("turi","Dam olish"),d.get("izoh","")))
        c.commit(); c.close(); return jsonify({"status":"ok"})
    c=get_db(); rows=c.execute("""SELECT h.*,i.ism,i.familiya FROM ishchi_holatlari h JOIN ishchilar i ON i.id=h.ishchi_id
        ORDER BY sana DESC,h.id DESC LIMIT 200""").fetchall(); c.close(); return jsonify([dict(r) for r in rows])

@app.route("/api/buyurtma/<int:oid>/tolovlar", methods=["GET","POST"])
def order_payments(oid):
    c=get_db()
    if request.method=="POST":
        d=jdata(); amount=float(d.get("miqdor") or 0)
        c.execute("INSERT INTO buyurtma_tolovlari(buyurtma_id,sana,miqdor,turi,izoh) VALUES(?,?,?,?,?)",
                  (oid,d["sana"],amount,d.get("turi","To‘lov"),d.get("izoh","")))
        c.execute("UPDATE buyurtmalar SET oldindan_tolov=oldindan_tolov+? WHERE id=?",(amount,oid))
        c.commit(); c.close(); return jsonify({"status":"ok"})
    rows=c.execute("SELECT * FROM buyurtma_tolovlari WHERE buyurtma_id=? ORDER BY sana DESC,id DESC",(oid,)).fetchall(); c.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/tayyor-mahsulot", methods=["GET","POST"])
def finished_goods():
    if request.method=="POST":
        d=jdata(); c=get_db(); c.execute("INSERT INTO tayyor_mahsulot(nomi,kodi,rang,miqdor,birlik,narx,izoh) VALUES(?,?,?,?,?,?,?)",
            (d["nomi"],d.get("kodi",""),d.get("rang",""),float(d.get("miqdor") or 0),d.get("birlik","dona"),float(d.get("narx") or 0),d.get("izoh","")))
        c.commit(); c.close(); return jsonify({"status":"ok"})
    c=get_db(); rows=c.execute("SELECT * FROM tayyor_mahsulot ORDER BY id DESC").fetchall(); c.close(); return jsonify([dict(r) for r in rows])

@app.route("/api/moliyaviy-xulosa")
def finance_summary():
    start=request.args.get("start") or "1900-01-01"; end=request.args.get("end") or "2999-12-31"
    c=get_db()
    income=c.execute("SELECT COALESCE(SUM(oldindan_tolov),0) FROM buyurtmalar WHERE date(created_at) BETWEEN ? AND ?",(start,end)).fetchone()[0]
    expense=c.execute("SELECT COALESCE(SUM(miqdor),0) FROM xarajatlar WHERE sana BETWEEN ? AND ?",(start,end)).fetchone()[0]
    salary=c.execute("SELECT COALESCE(SUM(miqdor),0) FROM tolovlar WHERE sana BETWEEN ? AND ?",(start,end)).fetchone()[0]
    bonus=c.execute("SELECT COALESCE(SUM(miqdor),0) FROM bonuslar WHERE sana BETWEEN ? AND ?",(start,end)).fetchone()[0]
    c.close(); return jsonify({"kirim":income,"xarajat":expense,"ishchi_tolov":salary,"bonus":bonus,"sof_foyda":income-expense-salary-bonus})

@app.route("/api/buyurtma-progress/<int:oid>")
def order_progress(oid):
    c=get_db(); row=c.execute("SELECT COUNT(*) jami,SUM(bajarildi) bajarildi FROM buyurtma_bosqichlari WHERE buyurtma_id=?",(oid,)).fetchone(); c.close()
    jami=row['jami'] or 0; done=row['bajarildi'] or 0; return jsonify({"foiz": round(done*100/jami,1) if jami else 0})

def _order_bundle(oid):
    c=get_db()
    order=c.execute("SELECT * FROM buyurtmalar WHERE id=?",(oid,)).fetchone()
    stages=c.execute("SELECT * FROM buyurtma_bosqichlari WHERE buyurtma_id=? ORDER BY id",(oid,)).fetchall()
    pays=c.execute("SELECT * FROM buyurtma_tolovlari WHERE buyurtma_id=? ORDER BY sana,id",(oid,)).fetchall()
    c.close()
    return order,stages,pays

@app.route("/kuzatuv/<token>")
def public_track(token):
    c=get_db(); order=c.execute("SELECT * FROM buyurtmalar WHERE tracking_token=?",(token,)).fetchone()
    if not order:
        c.close(); return "Buyurtma topilmadi",404
    stages=c.execute("SELECT * FROM buyurtma_bosqichlari WHERE buyurtma_id=? ORDER BY id",(order['id'],)).fetchall(); c.close()
    done=sum(int(x['bajarildi']) for x in stages); pct=round(done*100/len(stages)) if stages else 0
    html='''<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1"><title>Buyurtma kuzatuvi</title>
    <style>body{font-family:Arial;background:#eef3f8;margin:0;padding:18px;color:#172033}.box{max-width:650px;margin:auto;background:white;border-radius:18px;padding:20px;box-shadow:0 8px 30px #0002}h1{color:#1d4ed8}.bar{height:18px;background:#e2e8f0;border-radius:20px;overflow:hidden}.fill{height:100%;background:#16a34a}.stage{padding:10px;border-bottom:1px solid #eee}.ok{color:#15803d}.wait{color:#64748b}</style>
    <div class="box"><h1>Pharm Mebel</h1><h2>Buyurtma: {{o['kod']}}</h2><p><b>Mijoz:</b> {{o['mijoz']}}</p><p><b>Mahsulot:</b> {{o['mahsulot']}}</p><p><b>Tugash sanasi:</b> {{o['tugash_sana']}}</p><h3>Tayyorlik: {{pct}}%</h3><div class="bar"><div class="fill" style="width:{{pct}}%"></div></div>{% for s in stages %}<div class="stage {{'ok' if s['bajarildi'] else 'wait'}}">{{'✅' if s['bajarildi'] else '⬜'}} {{s['bosqich']}}</div>{% endfor %}</div>'''
    return render_template_string(html,o=order,stages=stages,pct=pct)

@app.route("/buyurtma/<int:oid>/qr.png")
def order_qr(oid):
    c=get_db(); row=c.execute("SELECT tracking_token FROM buyurtmalar WHERE id=?",(oid,)).fetchone(); c.close()
    if not row:return "Topilmadi",404
    url=request.url_root.rstrip('/')+url_for('public_track',token=row['tracking_token'])
    img=qrcode.make(url); out=io.BytesIO(); img.save(out,format='PNG'); out.seek(0)
    return send_file(out,mimetype='image/png',download_name=f'buyurtma_{oid}_qr.png')

@app.route("/buyurtma/<int:oid>/shartnoma.pdf")
def order_contract_pdf(oid):
    order,stages,pays=_order_bundle(oid)
    if not order:return "Topilmadi",404
    out=io.BytesIO(); p=canvas.Canvas(out,pagesize=A4); w,h=A4
    p.setFont('Helvetica-Bold',18); p.drawString(50,h-55,'PHARM MEBEL - BUYURTMA SHARTNOMASI')
    p.setFont('Helvetica',11); y=h-95
    lines=[f"Buyurtma kodi: {order['kod']}",f"Sana: {str(order['created_at'])[:10]}",f"Mijoz: {order['mijoz']}",f"Telefon: {order['telefon']}",f"Manzil: {order['manzil']}",f"Mahsulot: {order['mahsulot']}",f"Umumiy narx: {order['umumiy_narx']:,.0f} so'm",f"To'langan: {order['oldindan_tolov']:,.0f} so'm",f"Qoldiq: {order['umumiy_narx']-order['oldindan_tolov']:,.0f} so'm",f"Tugash sanasi: {order['tugash_sana']}",f"Izoh: {order['izoh']}"]
    for line in lines:p.drawString(55,y,line); y-=22
    y-=20; p.line(55,y,250,y); p.line(340,y,535,y); p.drawString(90,y-18,'Buyurtmachi imzosi'); p.drawString(385,y-18,'Ijrochi imzosi')
    p.save(); out.seek(0); return send_file(out,mimetype='application/pdf',as_attachment=True,download_name=f"shartnoma_{order['kod']}.pdf")

@app.route("/buyurtma/<int:oid>/chek.pdf")
def order_receipt_pdf(oid):
    order,stages,pays=_order_bundle(oid)
    if not order:return "Topilmadi",404
    out=io.BytesIO(); p=canvas.Canvas(out,pagesize=A4); w,h=A4
    p.setFont('Helvetica-Bold',20); p.drawCentredString(w/2,h-60,'PHARM MEBEL - TOLOV CHEKI')
    p.setFont('Helvetica',12); y=h-110
    paid=float(order['oldindan_tolov'] or 0); remaining=float(order['umumiy_narx'] or 0)-paid
    for line in [f"Buyurtma: {order['kod']}",f"Mijoz: {order['mijoz']}",f"Mahsulot: {order['mahsulot']}",f"Umumiy summa: {order['umumiy_narx']:,.0f} so'm",f"Jami to'langan: {paid:,.0f} so'm",f"Qoldiq: {remaining:,.0f} so'm",f"Chek sanasi: {date.today().isoformat()}"]:
        p.drawString(70,y,line); y-=25
    p.drawString(70,y-20,'Rahmat! Pharm Mebel xizmatidan foydalanganingiz uchun.')
    p.save(); out.seek(0); return send_file(out,mimetype='application/pdf',as_attachment=True,download_name=f"chek_{order['kod']}.pdf")

@app.route("/api/buyurtma/<int:oid>/link")
def order_public_link(oid):
    c=get_db(); row=c.execute("SELECT tracking_token FROM buyurtmalar WHERE id=?",(oid,)).fetchone(); c.close()
    if not row:return jsonify({'message':'Topilmadi'}),404
    return jsonify({'url':request.url_root.rstrip('/')+url_for('public_track',token=row['tracking_token'])})

@app.route("/backup")
def backup_db():
    if not os.path.exists(DB_NAME):
        return jsonify({"message":"Baza topilmadi"}),404
    return send_file(DB_NAME, as_attachment=True, download_name=f"pharm_mebel_backup_{date.today().isoformat()}.db")

@app.route("/api/audit")
def audit_get():
    c=get_db(); rows=c.execute("SELECT * FROM audit_log ORDER BY id DESC LIMIT 200").fetchall(); c.close(); return jsonify([dict(r) for r in rows])

# ---------- JAMLANMA / REYTING ----------
@app.route("/api/jami")
def totals():
    start=request.args.get("start") or "1900-01-01"
    end=request.args.get("end") or "2999-12-31"
    c=get_db()
    rows=c.execute("""
    SELECT i.id,i.ism,i.familiya,i.lavozim,i.sifat_ball,i.tezlik_ball,i.intizom_ball,
    COALESCE(a.kun,0) ish_kunlari,COALESCE(a.soat,0) jami_soat,
    COALESCE(n.miqdor,0) jami_miqdor,COALESCE(n.haq,0) ish_haqi,
    COALESCE(j.jarima,0) jarima,COALESCE(t.tolov,0) tolangan,COALESCE(b.bonus,0) bonus,
    ROUND(COALESCE(n.haq,0)+COALESCE(b.bonus,0)-COALESCE(j.jarima,0)-COALESCE(t.tolov,0),2) qoldiq,
    ROUND((i.sifat_ball+i.tezlik_ball+i.intizom_ball)/3.0 +
          MIN(COALESCE(n.miqdor,0)/100.0,5),2) reyting
    FROM ishchilar i
    LEFT JOIN (SELECT ishchi_id,COUNT(DISTINCT sana) kun,ROUND(SUM(ish_soatlari),2) soat
               FROM keldi_ketdi WHERE sana BETWEEN ? AND ? GROUP BY ishchi_id) a ON a.ishchi_id=i.id
    LEFT JOIN (SELECT ishchi_id,ROUND(SUM(miqdor),2) miqdor,ROUND(SUM(jami_haq),2) haq
               FROM ish_natijalari WHERE sana BETWEEN ? AND ? GROUP BY ishchi_id) n ON n.ishchi_id=i.id
    LEFT JOIN (SELECT ishchi_id,ROUND(SUM(miqdor),2) jarima
               FROM jarimalar WHERE sana BETWEEN ? AND ? GROUP BY ishchi_id) j ON j.ishchi_id=i.id
    LEFT JOIN (SELECT ishchi_id,ROUND(SUM(miqdor),2) tolov
               FROM tolovlar WHERE sana BETWEEN ? AND ? GROUP BY ishchi_id) t ON t.ishchi_id=i.id
    LEFT JOIN (SELECT ishchi_id,ROUND(SUM(miqdor),2) bonus
               FROM bonuslar WHERE sana BETWEEN ? AND ? GROUP BY ishchi_id) b ON b.ishchi_id=i.id
    WHERE i.faol=1 ORDER BY reyting DESC,i.ism
    """,(start,end,start,end,start,end,start,end,start,end)).fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/dashboard")
def dashboard():
    c=get_db(); today=date.today().isoformat()
    data={
      "workers":c.execute("SELECT COUNT(*) FROM ishchilar WHERE faol=1").fetchone()[0],
      "orders":c.execute("SELECT COUNT(*) FROM buyurtmalar WHERE holat!='Yetkazildi'").fetchone()[0],
      "hours":c.execute("SELECT COALESCE(SUM(ish_soatlari),0) FROM keldi_ketdi WHERE sana=?",(today,)).fetchone()[0],
      "production":c.execute("SELECT COALESCE(SUM(miqdor),0) FROM ish_natijalari WHERE sana=?",(today,)).fetchone()[0],
      "km":c.execute("SELECT COALESCE(SUM(masofa_km),0) FROM safarlar WHERE sana=?",(today,)).fetchone()[0],
      "low_stock":c.execute("SELECT COUNT(*) FROM ombor WHERE qoldiq<=min_qoldiq").fetchone()[0]
    }
    c.close(); return jsonify(data)


@app.route("/export/jami.csv")
def export_csv():
    start=request.args.get("start") or "1900-01-01"; end=request.args.get("end") or "2999-12-31"
    c=get_db()
    rows=c.execute("""SELECT i.ism||' '||i.familiya ishchi,i.lavozim,
    COALESCE(SUM(n.miqdor),0) miqdor,COALESCE(SUM(n.jami_haq),0) ish_haqi
    FROM ishchilar i LEFT JOIN ish_natijalari n ON n.ishchi_id=i.id AND n.sana BETWEEN ? AND ?
    WHERE i.faol=1 GROUP BY i.id ORDER BY ishchi""",(start,end)).fetchall()
    c.close()
    out=io.StringIO(); w=csv.writer(out); w.writerow(["Ishchi","Lavozim","Miqdor","Ish haqi"])
    for r in rows:w.writerow(list(r))
    return Response(out.getvalue(),mimetype="text/csv",
                    headers={"Content-Disposition":"attachment; filename=jami_hisob.csv"})


# ---------- ISHCHI RO'YXATDAN O'TISH VA SHAXSIY KABINET ----------
def normalize_phone(value):
    digits=''.join(ch for ch in str(value or '') if ch.isdigit())
    if digits.startswith('998') and len(digits)==12:
        return '+'+digits
    if len(digits)==9:
        return '+998'+digits
    return '+'+digits if digits else ''


def make_otp():
    return f"{secrets.randbelow(1000000):06d}"


@app.route('/ishchi/royxat', methods=['GET','POST'])
def worker_register():
    msg=''; error=''; demo_code=''
    if request.method=='POST':
        phone=normalize_phone(request.form.get('telefon'))
        if len(phone)<10:
            error='Telefon raqamini to‘g‘ri kiriting.'
        else:
            c=get_db()
            existing=c.execute('SELECT id FROM ishchi_akkauntlari WHERE telefon=?',(phone,)).fetchone()
            if existing:
                error='Bu telefon raqami avval ro‘yxatdan o‘tgan.'
            else:
                code=make_otp()
                expires=datetime.now().timestamp()+600
                c.execute('DELETE FROM ishchi_otp WHERE telefon=?',(phone,))
                c.execute('INSERT INTO ishchi_otp(telefon,kod_hash,muddati) VALUES(?,?,?)',
                          (phone,generate_password_hash(code),str(expires)))
                c.commit(); c.close()
                session['worker_pending_phone']=phone
                # Haqiqiy SMS provayder ulanmaguncha kod admin panelida ko‘rinadi.
                if os.environ.get('DEV_SHOW_OTP','1')=='1':
                    demo_code=code
                msg='Tasdiqlash kodi yaratildi. Kodni kiriting.'
    return render_template_string(WORKER_REGISTER_HTML,msg=msg,error=error,demo_code=demo_code)


@app.route('/ishchi/kod', methods=['GET','POST'])
def worker_verify():
    phone=session.get('worker_pending_phone','')
    if not phone:
        return redirect(url_for('worker_register'))
    error=''
    if request.method=='POST':
        code=request.form.get('kod','').strip()
        login=request.form.get('login','').strip()
        password=request.form.get('password','')
        ism=request.form.get('ism','').strip()
        familiya=request.form.get('familiya','').strip()
        if len(login)<3 or len(password)<6 or not ism:
            error='Ism, kamida 3 belgili login va 6 belgili parol kiriting.'
        else:
            c=get_db()
            row=c.execute('SELECT * FROM ishchi_otp WHERE telefon=? AND ishlatildi=0 ORDER BY id DESC LIMIT 1',(phone,)).fetchone()
            if not row or float(row['muddati']) < datetime.now().timestamp() or not check_password_hash(row['kod_hash'],code):
                error='Kod noto‘g‘ri yoki muddati tugagan.'
            elif c.execute('SELECT 1 FROM ishchi_akkauntlari WHERE login=?',(login,)).fetchone():
                error='Bu login band. Boshqa login tanlang.'
            else:
                # telefon bazadagi ishchiga mos tushsa bog‘laymiz, aks holda yangi ishchi yozuvi ochiladi
                worker=c.execute('SELECT id FROM ishchilar WHERE telefon=? AND faol=1 LIMIT 1',(phone,)).fetchone()
                if worker:
                    worker_id=worker['id']
                else:
                    cur=c.execute('INSERT INTO ishchilar(ism,familiya,telefon,lavozim) VALUES(?,?,?,?)',
                                  (ism,familiya,phone,'Ishchi'))
                    worker_id=cur.lastrowid
                c.execute('INSERT INTO ishchi_akkauntlari(ishchi_id,telefon,login,parol_hash,tasdiqlangan,admin_tasdiq) VALUES(?,?,?,?,1,0)',
                          (worker_id,phone,login,generate_password_hash(password)))
                c.execute('UPDATE ishchi_otp SET ishlatildi=1 WHERE id=?',(row['id'],))
                c.commit(); c.close()
                session.pop('worker_pending_phone',None)
                return render_template_string(WORKER_WAIT_HTML)
            c.close()
    return render_template_string(WORKER_VERIFY_HTML,telefon=phone,error=error)


@app.route('/ishchi/login', methods=['GET','POST'])
def worker_login():
    error=''
    if request.method=='POST':
        login=request.form.get('login','').strip()
        password=request.form.get('password','')
        c=get_db(); row=c.execute('SELECT * FROM ishchi_akkauntlari WHERE login=? AND faol=1',(login,)).fetchone(); c.close()
        if not row or not check_password_hash(row['parol_hash'],password):
            error='Login yoki parol xato.'
        elif not row['admin_tasdiq']:
            error='Admin hali akkauntingizni tasdiqlamagan.'
        else:
            session.clear(); session['worker_account_id']=row['id']; session['worker_id']=row['ishchi_id']
            return redirect(url_for('worker_dashboard'))
    return render_template_string(WORKER_LOGIN_HTML,error=error)


@app.route('/ishchi/logout')
def worker_logout():
    session.clear(); return redirect(url_for('worker_login'))


@app.route('/ishchi/kabinet')
def worker_dashboard():
    wid=session.get('worker_id')
    c=get_db()
    worker=c.execute('SELECT * FROM ishchilar WHERE id=?',(wid,)).fetchone()
    tasks=c.execute('SELECT * FROM ishchi_topshiriqlari WHERE ishchi_id=? ORDER BY sana DESC,id DESC LIMIT 50',(wid,)).fetchall()
    stats=c.execute('''SELECT COUNT(DISTINCT sana) kun,COALESCE(SUM(ish_soatlari),0) soat FROM keldi_ketdi WHERE ishchi_id=?''',(wid,)).fetchone()
    result=c.execute('SELECT COALESCE(SUM(miqdor),0) miqdor FROM ish_natijalari WHERE ishchi_id=?',(wid,)).fetchone()
    rating=round(((worker['sifat_ball']+worker['tezlik_ball']+worker['intizom_ball'])/3.0)+min(float(result['miqdor'] or 0)/100.0,5),2)
    c.close()
    return render_template_string(WORKER_DASHBOARD_HTML,worker=worker,tasks=tasks,stats=stats,rating=rating)


@app.route('/ishchi/topshiriq/<int:tid>', methods=['POST'])
def worker_task_update(tid):
    wid=session.get('worker_id')
    progress=max(0,min(100,int(request.form.get('progress') or 0)))
    status=request.form.get('holat','Jarayonda')
    c=get_db(); c.execute('UPDATE ishchi_topshiriqlari SET progress=?,holat=? WHERE id=? AND ishchi_id=?',(progress,status,tid,wid)); c.commit(); c.close()
    return redirect(url_for('worker_dashboard'))


@app.route('/ishchi-boshqaruv', methods=['GET','POST'])
def worker_admin():
    c=get_db()
    if request.method=='POST':
        action=request.form.get('action')
        if action=='approve':
            aid=int(request.form['account_id']); c.execute('UPDATE ishchi_akkauntlari SET admin_tasdiq=1 WHERE id=?',(aid,))
        elif action=='block':
            aid=int(request.form['account_id']); c.execute('UPDATE ishchi_akkauntlari SET faol=0 WHERE id=?',(aid,))
        elif action=='task':
            c.execute('''INSERT INTO ishchi_topshiriqlari(ishchi_id,buyurtma_kodi,ish_turi,tavsif,holat,progress,sana,tugash_sana)
                         VALUES(?,?,?,?,?,?,?,?)''',(
                int(request.form['ishchi_id']),request.form.get('buyurtma_kodi','').strip(),request.form.get('ish_turi','').strip(),
                request.form.get('tavsif','').strip(),'Yangi',0,request.form.get('sana') or date.today().isoformat(),request.form.get('tugash_sana','')
            ))
        c.commit()
    accounts=c.execute('''SELECT a.*,i.ism,i.familiya,i.lavozim FROM ishchi_akkauntlari a LEFT JOIN ishchilar i ON i.id=a.ishchi_id ORDER BY a.id DESC''').fetchall()
    workers=c.execute('SELECT id,ism,familiya,lavozim FROM ishchilar WHERE faol=1 ORDER BY ism').fetchall()
    tasks=c.execute('''SELECT t.*,i.ism,i.familiya FROM ishchi_topshiriqlari t JOIN ishchilar i ON i.id=t.ishchi_id ORDER BY t.id DESC LIMIT 100''').fetchall()
    otp_rows=c.execute('SELECT telefon,created_at FROM ishchi_otp WHERE ishlatildi=0 ORDER BY id DESC LIMIT 20').fetchall()
    c.close()
    return render_template_string(WORKER_ADMIN_HTML,accounts=accounts,workers=workers,tasks=tasks,otp_rows=otp_rows,today=date.today().isoformat())


LOGIN_HTML = r"""
<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pharm Mebel ERP - Kirish</title><style>body{margin:0;background:linear-gradient(135deg,#0f1b33,#2563eb);font-family:Arial;display:grid;place-items:center;min-height:100vh}.box{background:#fff;padding:28px;border-radius:18px;width:min(360px,92%);box-shadow:0 20px 50px #0005}h2{margin-top:0}input{width:100%;padding:12px;margin:8px 0;border:1px solid #cbd5e1;border-radius:9px;box-sizing:border-box}button{width:100%;padding:12px;border:0;border-radius:9px;background:#2563eb;color:#fff;font-weight:700}.err{color:#b91c1c;font-size:13px}</style></head><body><form class="box" method="post"><h2>🏭 Pharm Mebel ERP</h2><input name="user" placeholder="Login" value="admin"><input name="password" type="password" placeholder="Parol"><button>Kirish</button><div class="err">{{error}}</div><p style="font-size:12px;color:#64748b">Standart: admin / 12345</p><hr><p style="text-align:center"><a href="/ishchi/login">👷 Ishchi kirishi</a> · <a href="/ishchi/royxat">Ro‘yxatdan o‘tish</a></p></form></body></html>
"""

WORKER_BASE_STYLE = """
<style>
*{box-sizing:border-box}body{margin:0;font-family:Arial;background:#eef3f8;color:#182235}.head{background:linear-gradient(135deg,#0f1b33,#2563eb);color:white;padding:18px}.wrap{max-width:1050px;margin:auto;padding:16px}.box,.card{background:white;border-radius:16px;padding:18px;box-shadow:0 8px 24px #0f172a18;margin-bottom:14px}input,select,textarea{width:100%;padding:11px;border:1px solid #cbd5e1;border-radius:9px;margin:5px 0 10px}button,.btn{display:inline-block;border:0;border-radius:9px;padding:10px 14px;background:#2563eb;color:white;font-weight:700;text-decoration:none;cursor:pointer}.green{background:#16a34a}.red{background:#dc2626}.muted{color:#64748b;font-size:13px}.err{color:#b91c1c}.ok{color:#166534}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.stat b{font-size:25px;color:#2563eb}.task{border-left:5px solid #2563eb}.bar{height:10px;background:#e2e8f0;border-radius:20px;overflow:hidden}.bar i{display:block;height:100%;background:#16a34a}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left}@media(max-width:700px){.grid{grid-template-columns:1fr}.wrap{padding:9px}}
</style>
"""

WORKER_REGISTER_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi ro‘yxati</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>👷 Ishchi ro‘yxatdan o‘tishi</b></div><div class="wrap"><form class="box" method="post"><h2>Telefon raqamingiz</h2><p class="muted">Masalan: +998 90 123 45 67</p><input name="telefon" required placeholder="+998901234567"><button>Kod olish</button><p class="ok">{{msg}}</p><p class="err">{{error}}</p>{% if demo_code %}<div class="card"><b>Sinov kodi: {{demo_code}}</b><p class="muted">Haqiqiy SMS xizmati ulanmaguncha shu koddan foydalaning.</p><a class="btn green" href="/ishchi/kod">Kodni kiritish</a></div>{% endif %}<p><a href="/ishchi/login">Akkauntim bor — kirish</a></p></form></div></body></html>"""

WORKER_VERIFY_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kodni tasdiqlash</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🔐 Telefonni tasdiqlash</b></div><div class="wrap"><form class="box" method="post"><p>{{telefon}}</p><label>Kod</label><input name="kod" inputmode="numeric" maxlength="6" required><label>Ism</label><input name="ism" required><label>Familiya</label><input name="familiya"><label>Yangi login</label><input name="login" required><label>Yangi parol</label><input name="password" type="password" minlength="6" required><button>Ro‘yxatdan o‘tish</button><p class="err">{{error}}</p></form></div></body></html>"""

WORKER_WAIT_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kutilmoqda</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="wrap"><div class="box"><h2>✅ Ro‘yxatdan o‘tdingiz</h2><p>Endi administrator akkauntingizni tasdiqlaydi. Tasdiqlangach login va parolingiz bilan kirasiz.</p><a class="btn" href="/ishchi/login">Kirish sahifasi</a></div></div></body></html>"""

WORKER_LOGIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi kirishi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🏭 Pharm Mebel — Ishchi kabineti</b></div><div class="wrap"><form class="box" method="post"><h2>Kirish</h2><input name="login" placeholder="Login" required><input name="password" type="password" placeholder="Parol" required><button>Kirish</button><p class="err">{{error}}</p><p><a href="/ishchi/royxat">Yangi ro‘yxatdan o‘tish</a></p></form></div></body></html>"""

WORKER_DASHBOARD_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi kabineti</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><div class="wrap" style="padding:0"><b>👷 {{worker['ism']}} {{worker['familiya']}}</b> — {{worker['lavozim']}} <a class="btn red" style="float:right" href="/ishchi/logout">Chiqish</a></div></div><div class="wrap"><div class="grid"><div class="card stat"><span>REYTING</span><br><b>{{rating}} / 10</b></div><div class="card stat"><span>ISHLAGAN KUN</span><br><b>{{stats['kun'] or 0}}</b></div><div class="card stat"><span>JAMI SOAT</span><br><b>{{'%.1f'|format(stats['soat'] or 0)}}</b></div></div><div class="card"><h3>Ballarim</h3><p>Sifat: <b>{{worker['sifat_ball']}}</b> · Tezlik: <b>{{worker['tezlik_ball']}}</b> · Intizom: <b>{{worker['intizom_ball']}}</b></p><p class="muted">Pul summalari va korxona moliyasi bu kabinetda ko‘rsatilmaydi.</p></div><h2>Topshiriqlarim</h2>{% for t in tasks %}<div class="card task"><b>{{t['ish_turi']}}</b> {% if t['buyurtma_kodi'] %}<span class="muted">— {{t['buyurtma_kodi']}}</span>{% endif %}<p>{{t['tavsif']}}</p><div class="bar"><i style="width:{{t['progress']}}%"></i></div><p><b>{{t['progress']}}%</b> · {{t['holat']}} · {{t['sana']}}</p><form method="post" action="/ishchi/topshiriq/{{t['id']}}"><select name="holat"><option {% if t['holat']=='Yangi' %}selected{% endif %}>Yangi</option><option {% if t['holat']=='Jarayonda' %}selected{% endif %}>Jarayonda</option><option {% if t['holat']=='Tayyor' %}selected{% endif %}>Tayyor</option></select><input type="number" name="progress" min="0" max="100" value="{{t['progress']}}"><button>Yangilash</button></form></div>{% else %}<div class="card">Hozircha topshiriq yo‘q.</div>{% endfor %}</div></body></html>"""

WORKER_ADMIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi boshqaruvi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>👥 Ishchi kabinetlarini boshqarish</b> <a class="btn" style="float:right" href="/">ERP bosh sahifa</a></div><div class="wrap"><div class="grid"><form class="box" method="post"><h3>Yangi topshiriq</h3><input type="hidden" name="action" value="task"><label>Ishchi</label><select name="ishchi_id" required>{% for w in workers %}<option value="{{w['id']}}">{{w['ism']}} {{w['familiya']}} — {{w['lavozim']}}</option>{% endfor %}</select><label>Buyurtma kodi</label><input name="buyurtma_kodi" placeholder="AB-001"><label>Ish turi</label><input name="ish_turi" required placeholder="Kesish / Rover / Yig‘ish"><label>Topshiriq</label><textarea name="tavsif"></textarea><label>Boshlanish</label><input type="date" name="sana" value="{{today}}"><label>Tugash</label><input type="date" name="tugash_sana"><button>Topshiriq berish</button></form><div class="box" style="grid-column:span 2;overflow:auto"><h3>Ro‘yxatdan o‘tganlar</h3><table><tr><th>Ishchi</th><th>Telefon</th><th>Login</th><th>Holat</th><th>Amal</th></tr>{% for a in accounts %}<tr><td>{{a['ism']}} {{a['familiya']}}</td><td>{{a['telefon']}}</td><td>{{a['login'] or ''}}</td><td>{% if a['admin_tasdiq'] %}✅ Tasdiqlangan{% else %}⏳ Kutilmoqda{% endif %}</td><td>{% if not a['admin_tasdiq'] %}<form method="post" style="display:inline"><input type="hidden" name="action" value="approve"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="green">Tasdiqlash</button></form>{% endif %}<form method="post" style="display:inline"><input type="hidden" name="action" value="block"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="red">Bloklash</button></form></td></tr>{% endfor %}</table></div></div><div class="box" style="overflow:auto"><h3>Topshiriqlar</h3><table><tr><th>Ishchi</th><th>Buyurtma</th><th>Ish</th><th>Holat</th><th>Progress</th><th>Sana</th></tr>{% for t in tasks %}<tr><td>{{t['ism']}} {{t['familiya']}}</td><td>{{t['buyurtma_kodi']}}</td><td>{{t['ish_turi']}}</td><td>{{t['holat']}}</td><td>{{t['progress']}}%</td><td>{{t['sana']}}</td></tr>{% endfor %}</table></div><div class="box"><h3>Telefon kodi haqida</h3><p class="muted">Hozir sinov rejimida kod ishchining ekranida ko‘rinadi. Haqiqiy SMS yuborish uchun Eskiz.uz yoki boshqa SMS provayder API kaliti kerak bo‘ladi.</p></div></div></body></html>"""


HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pharm Mebel ERP Enterprise V3.1</title>
<style>
:root{--nav:#0f1b33;--blue:#2563eb;--bg:#eef3f8;--card:#fff;--text:#182235;--muted:#64748b;--danger:#dc2626;--ok:#16a34a}
*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;background:var(--bg);color:var(--text)}
header{background:linear-gradient(135deg,#0f1b33,#1d4ed8);color:#fff;padding:20px;position:sticky;top:0;z-index:5}
.top{max-width:1400px;margin:auto;display:flex;justify-content:space-between;align-items:center;gap:12px}
h1{margin:0;font-size:27px}.sub{opacity:.85;font-size:13px}.wrap{max-width:1400px;margin:auto;padding:16px}
.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:15px}
.card,.panel{background:var(--card);border-radius:14px;box-shadow:0 7px 20px #0f172a12;padding:14px}
.card span{font-size:11px;color:var(--muted);font-weight:700}.card b{display:block;font-size:25px;color:var(--blue);margin-top:5px}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:13px}
button{border:0;border-radius:9px;padding:9px 13px;background:var(--blue);color:#fff;font-weight:700;cursor:pointer}
button:hover{filter:brightness(.94)}.tabs button{background:#dbe5f4;color:#24334d}.tabs button.active{background:var(--nav);color:#fff}
.tab{display:none}.tab.active{display:block}.grid{display:grid;grid-template-columns:360px 1fr;gap:14px}
h3{margin:0 0 11px}label{display:block;font-size:12px;font-weight:700;margin-top:8px;color:#334155}
input,select,textarea{width:100%;margin-top:4px;padding:9px;border:1px solid #cbd5e1;border-radius:8px;background:#fff}
textarea{min-height:65px}form button{width:100%;margin-top:12px}table{width:100%;border-collapse:collapse;font-size:12px}
th,td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left;white-space:nowrap}th{background:#f8fafc;color:#475569}
.tablewrap{overflow:auto;max-height:590px}.danger{background:var(--danger);padding:6px 9px}.ok{background:var(--ok)}
.msg{min-height:18px;margin-top:7px;font-size:12px;color:#166534}.badge{padding:3px 7px;border-radius:20px;background:#e0e7ff;color:#3730a3;font-size:11px}
.money{font-weight:700;color:#166534}.minus{font-weight:700;color:#b91c1c}.balance{font-weight:800;color:#1d4ed8}
.stage{display:flex;align-items:center;gap:6px;padding:6px 0}.stage input{width:auto;margin:0}.low{background:#fee2e2!important}
@media(max-width:1000px){.cards{grid-template-columns:repeat(3,1fr)}.grid{grid-template-columns:1fr}}
@media(max-width:600px){.cards{grid-template-columns:repeat(2,1fr)}header{position:static}.wrap{padding:9px}}
</style>
</head>
<body>
<header><div class="top"><div><h1>🏭 Pharm Mebel ERP Enterprise V3.1</h1><div class="sub">Ishchilar, buyurtmalar, ombor, ishlab chiqarish va moliya</div></div><div><a href="/backup"><button style="background:#16a34a">Backup</button></a> <a href="/logout"><button style="background:#dc2626">Chiqish</button></a></div></div></header>
<main class="wrap">
<div class="cards">
 <div class="card"><span>ISHCHILAR</span><b id="dWorkers">0</b></div>
 <div class="card"><span>FAOL BUYURTMALAR</span><b id="dOrders">0</b></div>
 <div class="card"><span>BUGUNGI SOAT</span><b id="dHours">0</b></div>
 <div class="card"><span>BUGUNGI ISH</span><b id="dProduction">0</b></div>
 <div class="card"><span>BUGUNGI KM</span><b id="dKm">0</b></div>
 <div class="card"><span>KAM QOLDIQ</span><b id="dLow">0</b></div>
</div>

<div class="tabs">
 <button class="active" data-tab="workers">Ishchilar</button>
 <button data-tab="attendance">Keldi-ketdi</button>
 <button data-tab="results">Ish natijasi</button>
 <button data-tab="orders">Buyurtmalar</button>
 <button data-tab="stock">Ombor</button>
 <button data-tab="trips">Shofyor</button>
 <button data-tab="payments">To‘lov</button>
 <button data-tab="penalties">Jarima</button>
 <button data-tab="totals">Jami/Reyting</button>
 <button data-tab="expenses">Xarajat</button>
 <button data-tab="bonuses">Bonus/Ta’til</button>
 <button data-tab="finished">Tayyor mahsulot</button>
 <button data-tab="finance">Foyda</button>
</div>

<section id="workers" class="tab active"><div class="grid">
<div class="panel"><h3>Yangi ishchi</h3><form id="workerForm">
<label>Ism<input name="ism" required></label><label>Familiya<input name="familiya"></label>
<label>Telefon<input name="telefon"></label><label>Lavozim<input name="lavozim"></label>
<label>Ishga kirgan sana<input type="date" name="ishga_kirgan_sana"></label>
<label>Staj (yil)<input type="number" step="0.1" name="staj_yil" value="0"></label>
<label>Kunlik stavka<input type="number" name="kunlik_stavka" value="0"></label>
<label>Oylik maosh<input type="number" name="oylik_maosh" value="0"></label>
<label>Sifat balli (1-10)<input type="number" min="1" max="10" name="sifat_ball" value="5"></label>
<label>Tezlik balli (1-10)<input type="number" min="1" max="10" name="tezlik_ball" value="5"></label>
<label>Intizom balli (1-10)<input type="number" min="1" max="10" name="intizom_ball" value="5"></label>
<label>Izoh<textarea name="izoh"></textarea></label><button>Saqlash</button><div class="msg"></div>
</form></div><div class="panel tablewrap"><h3>Ishchilar</h3><table><thead><tr><th>Ism</th><th>Lavozim</th><th>Staj</th><th>Kunlik</th><th>Oylik</th><th></th></tr></thead><tbody id="workersBody"></tbody></table></div>
</div></section>

<section id="attendance" class="tab"><div class="grid"><div class="panel"><h3>Keldi-ketdi</h3><form id="attendanceForm">
<label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label>
<label>Keldi<input type="time" name="keldi_vaqti" required></label><label>Ketdi<input type="time" name="ketdi_vaqti" required></label>
<button>Saqlash</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Ishchi</th><th>Sana</th><th>Keldi</th><th>Ketdi</th><th>Soat</th></tr></thead><tbody id="attendanceBody"></tbody></table></div></div></section>

<section id="results" class="tab"><div class="grid"><div class="panel"><h3>Ish natijasi</h3><form id="resultForm">
<label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label>
<label>Ish turi<select id="workTypeSelect" name="ish_turi_id" required></select></label>
<label>Sana<input type="date" name="sana" required></label><label>Miqdor<input type="number" step="0.1" name="miqdor" required></label>
<label>Birlik narxi<input type="number" name="birlik_narxi" value="0"></label>
<label>Buyurtma kodi<input name="buyurtma_kodi"></label><label>Izoh<textarea name="izoh"></textarea></label>
<button>Saqlash</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Ishchi</th><th>Ish turi</th><th>Sana</th><th>Miqdor</th><th>Haq</th><th>Buyurtma</th></tr></thead><tbody id="resultsBody"></tbody></table></div></div></section>

<section id="orders" class="tab"><div class="grid"><div class="panel"><h3>Yangi buyurtma</h3><form id="orderForm">
<label>Kod<input name="kod" required placeholder="AB-001"></label><label>Mijoz<input name="mijoz" required></label><label>Telefon<input name="telefon"></label>
<label>Manzil<input name="manzil"></label><label>Mahsulot<input name="mahsulot"></label>
<label>Umumiy narx<input type="number" name="umumiy_narx" value="0"></label><label>Oldindan to‘lov<input type="number" name="oldindan_tolov" value="0"></label>
<label>Tugash sana<input type="date" name="tugash_sana"></label><label>Holat<select name="holat"><option>Yangi</option><option>Jarayonda</option><option>Tayyor</option><option>Yetkazildi</option></select></label>
<label>Izoh<textarea name="izoh"></textarea></label><button>Saqlash</button><div class="msg"></div></form></div>
<div class="panel tablewrap"><h3>Buyurtmalar</h3><table><thead><tr><th>Kod</th><th>Mijoz</th><th>Mahsulot</th><th>Narx</th><th>To‘lov</th><th>Qoldiq</th><th>Holat</th><th>Jarayon</th><th>Bosqich</th><th>To‘lov</th></tr></thead><tbody id="ordersBody"></tbody></table></div></div></section>

<section id="stock" class="tab"><div class="grid"><div class="panel"><h3>Ombor harakati</h3><form id="stockForm">
<label>Material<select id="stockSelect" name="material_id" required></select></label><label>Sana<input type="date" name="sana" required></label>
<label>Turi<select name="turi"><option>Kirim</option><option>Chiqim</option></select></label><label>Miqdor<input type="number" step="0.1" name="miqdor" required></label>
<label>Buyurtma kodi<input name="buyurtma_kodi"></label><label>Izoh<input name="izoh"></label><button>Saqlash</button><div class="msg"></div></form></div>
<div class="panel tablewrap"><h3>Ombor qoldig‘i</h3><table><thead><tr><th>Material</th><th>Kategoriya</th><th>Qoldiq</th><th>Birlik</th><th>Min.</th></tr></thead><tbody id="stockBody"></tbody></table></div></div></section>

<section id="trips" class="tab"><div class="grid"><div class="panel"><h3>Shofyor safari</h3><form id="tripForm">
<label>Shofyor<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label>
<label>Mashina<input name="mashina"></label><label>Qayerdan<input name="qayerdan"></label><label>Qayerga<input name="qayerga"></label>
<label>Masofa km<input type="number" step="0.1" name="masofa_km"></label><label>Sabab<input name="sabab"></label>
<label>Yoqilg‘i litr<input type="number" step="0.1" name="yonilgi"></label><label>Xarajat<input type="number" name="xarajat"></label>
<button>Saqlash</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Shofyor</th><th>Sana</th><th>Mashina</th><th>Yo‘nalish</th><th>Km</th><th>Sabab</th><th>Xarajat</th></tr></thead><tbody id="tripsBody"></tbody></table></div></div></section>

<section id="payments" class="tab"><div class="grid"><div class="panel"><h3>To‘lov</h3><form id="paymentForm">
<label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label>
<label>Summa<input type="number" name="miqdor" required></label><label>Turi<select name="turi"><option>Avans</option><option>Oylik</option><option>Bonus</option></select></label>
<label>Izoh<input name="tavsifi"></label><button>Saqlash</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Ishchi</th><th>Sana</th><th>Summa</th><th>Turi</th></tr></thead><tbody id="paymentsBody"></tbody></table></div></div></section>

<section id="penalties" class="tab"><div class="grid"><div class="panel"><h3>Jarima</h3><form id="penaltyForm">
<label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label>
<label>Summa<input type="number" name="miqdor" required></label><label>Sababi<input name="sababi"></label><button>Saqlash</button><div class="msg"></div></form></div>
<div class="panel tablewrap"><table><thead><tr><th>Ishchi</th><th>Sana</th><th>Summa</th><th>Sabab</th></tr></thead><tbody id="penaltiesBody"></tbody></table></div></div></section>

<section id="totals" class="tab"><div class="panel">
<div style="display:flex;gap:8px;align-items:end;flex-wrap:wrap;margin-bottom:12px">
<label style="margin:0">Boshlanish<input id="totalStart" type="date"></label><label style="margin:0">Tugash<input id="totalEnd" type="date"></label>
<button onclick="loadTotals()">Hisoblash</button><button onclick="setMonth()" style="background:#475569">Shu oy</button>
<a id="csvLink"><button style="background:#16a34a">Excel/CSV</button></a></div>
<div class="tablewrap"><table><thead><tr><th>O‘rin</th><th>Ishchi</th><th>Lavozim</th><th>Kun</th><th>Soat</th><th>Miqdor</th><th>Ish haqi</th><th>Jarima</th><th>Bonus</th><th>To‘langan</th><th>Qoldiq</th><th>Reyting</th></tr></thead><tbody id="totalsBody"></tbody></table></div>
</div></section>

<section id="expenses" class="tab"><div class="grid"><div class="panel"><h3>Yangi xarajat</h3><form id="expenseForm">
<label>Sana<input type="date" name="sana" required></label><label>Kategoriya<select name="kategoriya"><option>Material</option><option>Transport</option><option>Ijara</option><option>Kommunal</option><option>Stanok</option><option>Uy</option><option>Seh uchun</option><option>Boshqa</option></select></label>
<label>Summa<input type="number" name="miqdor" required></label><label>Buyurtma kodi<input name="buyurtma_kodi"></label><label>Izoh<input name="tavsifi"></label><button>Saqlash</button><div class="msg"></div></form></div>
<div class="panel tablewrap"><table><thead><tr><th>Sana</th><th>Kategoriya</th><th>Summa</th><th>Buyurtma</th><th>Izoh</th></tr></thead><tbody id="expensesBody"></tbody></table></div></div></section>

<section id="bonuses" class="tab"><div class="grid"><div class="panel"><h3>Bonus</h3><form id="bonusForm"><label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label><label>Summa<input type="number" name="miqdor" required></label><label>Sabab<input name="sababi"></label><button>Bonus saqlash</button><div class="msg"></div></form><hr><h3>Ta’til / holat</h3><form id="statusForm"><label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label><label>Turi<select name="turi"><option>Dam olish</option><option>Ta’til</option><option>Kasallik</option><option>Sababsiz</option></select></label><label>Izoh<input name="izoh"></label><button>Saqlash</button><div class="msg"></div></form></div>
<div class="panel tablewrap"><h3>Bonuslar</h3><table><thead><tr><th>Ishchi</th><th>Sana</th><th>Summa</th><th>Sabab</th></tr></thead><tbody id="bonusesBody"></tbody></table><h3 style="margin-top:20px">Ta’til va holatlar</h3><table><thead><tr><th>Ishchi</th><th>Sana</th><th>Turi</th><th>Izoh</th></tr></thead><tbody id="statusesBody"></tbody></table></div></div></section>

<section id="finished" class="tab"><div class="grid"><div class="panel"><h3>Tayyor mahsulot</h3><form id="finishedForm"><label>Nomi<input name="nomi" required></label><label>Kodi<input name="kodi"></label><label>Rang<input name="rang"></label><label>Miqdor<input type="number" step="0.1" name="miqdor" required></label><label>Birlik<select name="birlik"><option>dona</option><option>komplekt</option></select></label><label>Narx<input type="number" name="narx"></label><label>Izoh<input name="izoh"></label><button>Saqlash</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Nomi</th><th>Kod</th><th>Rang</th><th>Miqdor</th><th>Narx</th></tr></thead><tbody id="finishedBody"></tbody></table></div></div></section>

<section id="finance" class="tab"><div class="panel"><h3>Moliyaviy xulosa</h3><div style="display:flex;gap:8px;flex-wrap:wrap;align-items:end"><label>Boshlanish<input id="finStart" type="date"></label><label>Tugash<input id="finEnd" type="date"></label><button onclick="loadFinance()">Hisoblash</button></div><div class="cards" style="margin-top:15px"><div class="card"><span>KIRIM</span><b id="fIncome">0</b></div><div class="card"><span>XARAJAT</span><b id="fExpense">0</b></div><div class="card"><span>ISHCHI TO‘LOV</span><b id="fSalary">0</b></div><div class="card"><span>BONUS</span><b id="fBonus">0</b></div><div class="card"><span>SOF FOYDA</span><b id="fProfit">0</b></div></div></div></section>

</main>

<div id="stageModal" style="display:none;position:fixed;inset:0;background:#0009;z-index:20;place-items:center">
<div class="panel" style="width:min(420px,92%);max-height:85vh;overflow:auto"><h3>Buyurtma bosqichlari</h3><div id="stageList"></div><button onclick="closeStage()" style="margin-top:10px">Yopish</button></div>
</div>

<script>
const $=s=>document.querySelector(s),$$=s=>document.querySelectorAll(s);
const today=new Date().toISOString().slice(0,10);$$('input[type=date]').forEach(x=>x.value=today);
$$('.tabs button').forEach(b=>b.onclick=()=>{$$('.tabs button').forEach(x=>x.classList.remove('active'));$$('.tab').forEach(x=>x.classList.remove('active'));b.classList.add('active');$('#'+b.dataset.tab).classList.add('active')});
async function api(u,o){const r=await fetch(u,o),j=await r.json();if(!r.ok)throw new Error(j.message||'Xato');return j}
function fj(f){return Object.fromEntries(new FormData(f).entries())}
function bind(id,url){const f=$(id);f.onsubmit=async e=>{e.preventDefault();const m=f.querySelector('.msg');try{await api(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(fj(f))});m.textContent='✅ Saqlandi';await refresh()}catch(x){m.textContent='❌ '+x.message}}}
function money(v){return Number(v||0).toLocaleString('uz-UZ',{maximumFractionDigits:2})}

async function loadWorkers(){const a=await api('/api/ishchilar');$('#workersBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.lavozim||''}</td><td>${x.staj_yil} yil</td><td>${money(x.kunlik_stavka)}</td><td>${money(x.oylik_maosh)}</td><td><button class="danger" onclick="delWorker(${x.id})">O‘chirish</button></td></tr>`).join('');$$('.workerSelect').forEach(s=>s.innerHTML='<option value="">Tanlang</option>'+a.map(x=>`<option value="${x.id}">${x.ism} ${x.familiya||''}</option>`).join(''))}
async function delWorker(i){if(confirm('O‘chirasizmi?')){await api('/api/ishchilar/'+i,{method:'DELETE'});refresh()}}
async function loadTypes(){const a=await api('/api/ish-turlari');$('#workTypeSelect').innerHTML=a.map(x=>`<option value="${x.id}">${x.kategoriya} — ${x.nomi} (${x.birlik})</option>`).join('')}
async function loadAttendance(){const a=await api('/api/keldi-ketdi');$('#attendanceBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.sana}</td><td>${x.keldi_vaqti}</td><td>${x.ketdi_vaqti}</td><td>${x.ish_soatlari}</td></tr>`).join('')}
async function loadResults(){const a=await api('/api/natijalar');$('#resultsBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.ish_turi}</td><td>${x.sana}</td><td>${x.miqdor} ${x.birlik}</td><td>${money(x.jami_haq)}</td><td>${x.buyurtma_kodi||''}</td></tr>`).join('')}
async function loadOrders(){const a=await api('/api/buyurtmalar');$('#ordersBody').innerHTML=a.map(x=>`<tr><td>${x.kod}</td><td>${x.mijoz}</td><td>${x.mahsulot||''}</td><td>${money(x.umumiy_narx)}</td><td>${money(x.oldindan_tolov)}</td><td class="balance">${money(x.qoldiq)}</td><td><span class="badge">${x.holat}</span></td><td><span id="pr${x.id}">0%</span></td><td><button onclick="openStage(${x.id})">Ko‘rish</button></td><td><button onclick="addOrderPayment(${x.id})">To‘lov</button></td><td><a href="/buyurtma/${x.id}/shartnoma.pdf" target="_blank"><button>Shartnoma</button></a> <a href="/buyurtma/${x.id}/chek.pdf" target="_blank"><button>Chek</button></a> <a href="/buyurtma/${x.id}/qr.png" target="_blank"><button class="ok">QR</button></a> <button onclick="copyTrack(${x.id})">Link</button></td></tr>`).join('')}
async function openStage(id){const a=await api('/api/buyurtma/'+id+'/bosqichlar');$('#stageList').innerHTML=a.map(x=>`<label class="stage"><input type="checkbox" ${x.bajarildi?'checked':''} onchange="toggleStage(${x.id},this.checked)"> ${x.bosqich}</label>`).join('');$('#stageModal').style.display='grid'}
function closeStage(){$('#stageModal').style.display='none'}
async function toggleStage(id,v){await api('/api/buyurtma-bosqich/'+id,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({bajarildi:v})})}
async function loadStock(){const a=await api('/api/ombor');$('#stockSelect').innerHTML=a.map(x=>`<option value="${x.id}">${x.nomi} (${x.birlik})</option>`).join('');$('#stockBody').innerHTML=a.map(x=>`<tr class="${Number(x.qoldiq)<=Number(x.min_qoldiq)?'low':''}"><td>${x.nomi}</td><td>${x.kategoriya}</td><td>${x.qoldiq}</td><td>${x.birlik}</td><td>${x.min_qoldiq}</td></tr>`).join('')}
async function loadTrips(){const a=await api('/api/safarlar');$('#tripsBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.sana}</td><td>${x.mashina||''}</td><td>${x.qayerdan||''} → ${x.qayerga||''}</td><td>${x.masofa_km}</td><td>${x.sabab||''}</td><td>${money(x.xarajat)}</td></tr>`).join('')}
async function loadPayments(){const a=await api('/api/tolovlar');$('#paymentsBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.sana}</td><td>${money(x.miqdor)}</td><td>${x.turi}</td></tr>`).join('')}
async function loadPenalties(){const a=await api('/api/jarimalar');$('#penaltiesBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.sana}</td><td class="minus">${money(x.miqdor)}</td><td>${x.sababi||''}</td></tr>`).join('')}
async function loadDashboard(){const x=await api('/api/dashboard');$('#dWorkers').textContent=x.workers;$('#dOrders').textContent=x.orders;$('#dHours').textContent=x.hours;$('#dProduction').textContent=x.production;$('#dKm').textContent=x.km;$('#dLow').textContent=x.low_stock}
function setMonth(){const d=new Date(),y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0');$('#totalStart').value=`${y}-${m}-01`;$('#totalEnd').value=new Date(y,d.getMonth()+1,0).toISOString().slice(0,10);loadTotals()}
async function loadTotals(){const s=$('#totalStart').value||'1900-01-01',e=$('#totalEnd').value||'2999-12-31';$('#csvLink').href=`/export/jami.csv?start=${s}&end=${e}`;const a=await api(`/api/jami?start=${s}&end=${e}`);$('#totalsBody').innerHTML=a.map((x,i)=>`<tr><td>${i+1}</td><td>${x.ism} ${x.familiya||''}</td><td>${x.lavozim||''}</td><td>${x.ish_kunlari}</td><td>${x.jami_soat}</td><td>${x.jami_miqdor}</td><td class="money">${money(x.ish_haqi)}</td><td class="minus">${money(x.jarima)}</td><td class="money">${money(x.bonus)}</td><td>${money(x.tolangan)}</td><td class="balance">${money(x.qoldiq)}</td><td>${x.reyting}</td></tr>`).join('')}

async function loadExpenses(){const a=await api('/api/xarajatlar');$('#expensesBody').innerHTML=a.map(x=>`<tr><td>${x.sana}</td><td>${x.kategoriya}</td><td class="minus">${money(x.miqdor)}</td><td>${x.buyurtma_kodi||''}</td><td>${x.tavsifi||''}</td></tr>`).join('')}
async function loadBonuses(){const a=await api('/api/bonuslar');$('#bonusesBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.sana}</td><td class="money">${money(x.miqdor)}</td><td>${x.sababi||''}</td></tr>`).join('')}
async function loadStatuses(){const a=await api('/api/ishchi-holatlari');$('#statusesBody').innerHTML=a.map(x=>`<tr><td>${x.ism} ${x.familiya||''}</td><td>${x.sana}</td><td>${x.turi}</td><td>${x.izoh||''}</td></tr>`).join('')}
async function loadFinished(){const a=await api('/api/tayyor-mahsulot');$('#finishedBody').innerHTML=a.map(x=>`<tr><td>${x.nomi}</td><td>${x.kodi||''}</td><td>${x.rang||''}</td><td>${x.miqdor} ${x.birlik}</td><td>${money(x.narx)}</td></tr>`).join('')}
async function loadFinance(){const s=$('#finStart').value||'1900-01-01',e=$('#finEnd').value||'2999-12-31',x=await api(`/api/moliyaviy-xulosa?start=${s}&end=${e}`);$('#fIncome').textContent=money(x.kirim);$('#fExpense').textContent=money(x.xarajat);$('#fSalary').textContent=money(x.ishchi_tolov);$('#fBonus').textContent=money(x.bonus);$('#fProfit').textContent=money(x.sof_foyda)}
async function addOrderPayment(id){const miqdor=prompt('To‘lov summasi');if(!miqdor)return;await api(`/api/buyurtma/${id}/tolovlar`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({sana:today,miqdor})});refresh()}
async function copyTrack(id){const x=await api(`/api/buyurtma/${id}/link`);try{await navigator.clipboard.writeText(x.url);alert('Mijoz kuzatuv havolasi nusxalandi')}catch(e){prompt('Havolani nusxalang:',x.url)}}
async function loadProgress(){const rows=await api('/api/buyurtmalar');for(const x of rows){try{const p=await api(`/api/buyurtma-progress/${x.id}`),el=$(`#pr${x.id}`);if(el)el.textContent=p.foiz+'%'}catch(e){}}}

async function refresh(){await Promise.all([loadWorkers(),loadTypes(),loadAttendance(),loadResults(),loadOrders(),loadStock(),loadTrips(),loadPayments(),loadPenalties(),loadDashboard(),loadTotals(),loadExpenses(),loadBonuses(),loadStatuses(),loadFinished(),loadFinance(),loadProgress()])}
bind('#workerForm','/api/ishchilar');bind('#attendanceForm','/api/keldi-ketdi');bind('#resultForm','/api/natijalar');bind('#orderForm','/api/buyurtmalar');bind('#stockForm','/api/ombor-harakat');bind('#tripForm','/api/safarlar');bind('#paymentForm','/api/tolovlar');bind('#penaltyForm','/api/jarimalar');bind('#expenseForm','/api/xarajatlar');bind('#bonusForm','/api/bonuslar');bind('#statusForm','/api/ishchi-holatlari');bind('#finishedForm','/api/tayyor-mahsulot');setMonth();$('#finStart').value=$('#totalStart').value;$('#finEnd').value=$('#totalEnd').value;refresh();
</script>
</body></html>
"""

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)
