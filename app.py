# -*- coding: utf-8 -*-
import csv
import io
import os
import sqlite3
import secrets
from datetime import datetime, date, timedelta
from flask import Flask, jsonify, request, render_template_string, Response, session, redirect, url_for, send_file, send_from_directory, flash
import qrcode
from werkzeug.security import generate_password_hash, check_password_hash
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
import hashlib
_original_md5 = hashlib.md5
def _md5_windows7_compat(*args, **kwargs):
    kwargs.pop("usedforsecurity", None)
    return _original_md5(*args, **kwargs)
hashlib.md5 = _md5_windows7_compat
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
import re

app = Flask(__name__)
DB_NAME = os.environ.get("PHARM_ERP_DB", "pharm_mebel_erp_pro.db")
_APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _get_or_create_secret_key():
    """PHARM_ERP_SECRET muhit o'zgaruvchisi bo'lmasa, bir martalik tasodifiy
    kalit yaratib, faylga saqlaydi (har safar ilova qayta ishga tushganda
    bir xil kalit ishlatilishi uchun, aks holda barcha sessiyalar uziladi)."""
    env_key = os.environ.get("PHARM_ERP_SECRET")
    if env_key:
        return env_key
    key_path = os.path.join(_APP_DIR, ".secret_key")
    if os.path.exists(key_path):
        with open(key_path, "r", encoding="utf-8") as f:
            saved = f.read().strip()
            if saved:
                return saved
    new_key = secrets.token_hex(32)
    try:
        with open(key_path, "w", encoding="utf-8") as f:
            f.write(new_key)
    except Exception:
        pass
    return new_key


def _get_or_create_admin_password():
    """PHARM_ERP_PASSWORD muhit o'zgaruvchisi bo'lmasa, standart "12345" o'rniga
    tasodifiy parol yaratib, admin_parol.txt fayliga yozadi va konsolda ko'rsatadi."""
    env_pwd = os.environ.get("PHARM_ERP_PASSWORD")
    if env_pwd:
        return env_pwd
    pwd_path = os.path.join(_APP_DIR, "admin_parol.txt")
    if os.path.exists(pwd_path):
        with open(pwd_path, "r", encoding="utf-8") as f:
            saved = f.read().strip()
            if saved:
                return saved
    new_pwd = secrets.token_urlsafe(9)
    try:
        with open(pwd_path, "w", encoding="utf-8") as f:
            f.write(new_pwd)
    except Exception:
        pass
    print("=" * 60)
    print(f"  BIRINCHI ISHGA TUSHIRISH: Rahbar (admin) paroli yaratildi:")
    print(f"  Login: admin")
    print(f"  Parol: {new_pwd}")
    print(f"  (Bu parol '{pwd_path}' fayliga ham yozildi.)")
    print("=" * 60)
    return new_pwd


app.secret_key = _get_or_create_secret_key()
_ADMIN_PASSWORD = _get_or_create_admin_password()


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

    CREATE TABLE IF NOT EXISTS shartnoma_versiyalari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        versiya INTEGER DEFAULT 1,
        docx_fayl TEXT DEFAULT '',
        pdf_fayl TEXT DEFAULT '',
        yaratilgan_vaqt TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
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

    CREATE TABLE IF NOT EXISTS shofyor_akkauntlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ishchi_id INTEGER UNIQUE NOT NULL,
        login TEXT UNIQUE NOT NULL,
        parol_hash TEXT NOT NULL,
        faol INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
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


    CREATE TABLE IF NOT EXISTS mijozlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fio TEXT NOT NULL,
        telefon TEXT DEFAULT '',
        pasport_id TEXT DEFAULT '',
        manzil TEXT DEFAULT '',
        rozilik_reklama INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS mijoz_akkauntlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mijoz_id INTEGER,
        telefon TEXT NOT NULL,
        parol_hash TEXT DEFAULT '',
        sms_tasdiq INTEGER DEFAULT 0,
        faol INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (mijoz_id) REFERENCES mijozlar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS buyurtma_bosqich_hodisalari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        bosqich TEXT NOT NULL,
        boshlandi TEXT DEFAULT '',
        tugadi TEXT DEFAULT '',
        ishchi_id INTEGER,
        izoh TEXT DEFAULT '',
        media_havola TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE,
        FOREIGN KEY (ishchi_id) REFERENCES ishchilar(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS buyurtma_media (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        turi TEXT DEFAULT 'Rasm',
        havola TEXT NOT NULL,
        izoh TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS buyurtma_tasdiqlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        turi TEXT NOT NULL,
        holat TEXT DEFAULT 'Kutilmoqda',
        izoh TEXT DEFAULT '',
        tasdiqlagan TEXT DEFAULT '',
        tasdiqlangan_vaqt TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS keshbek_harakatlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mijoz_telefon TEXT NOT NULL,
        buyurtma_id INTEGER,
        turi TEXT NOT NULL,
        miqdor REAL DEFAULT 0,
        amal_muddati TEXT DEFAULT '',
        izoh TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS baholar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        sifat INTEGER DEFAULT 5,
        muddat INTEGER DEFAULT 5,
        muomala INTEGER DEFAULT 5,
        yetkazish INTEGER DEFAULT 5,
        montaj INTEGER DEFAULT 5,
        umumiy INTEGER DEFAULT 5,
        izoh TEXT DEFAULT '',
        rasm_havola TEXT DEFAULT '',
        reklama_ruxsat INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS servis_murojaatlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        turi TEXT DEFAULT 'Servis',
        muammo TEXT NOT NULL,
        media_havola TEXT DEFAULT '',
        holat TEXT DEFAULT 'Qabul qilindi',
        usta_id INTEGER,
        servis_sana TEXT DEFAULT '',
        pullik INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE,
        FOREIGN KEY (usta_id) REFERENCES ishchilar(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS kafolatlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER UNIQUE NOT NULL,
        boshlanish TEXT DEFAULT '',
        tugash TEXT DEFAULT '',
        shartlar TEXT DEFAULT '',
        qr_token TEXT DEFAULT '',
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS qoshimcha_ishlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        nomi TEXT NOT NULL,
        summa REAL DEFAULT 0,
        qoshimcha_kun INTEGER DEFAULT 0,
        mijoz_tasdiq INTEGER DEFAULT 0,
        kim_kiritdi TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS ombor_rezervlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        material_id INTEGER NOT NULL,
        miqdor REAL DEFAULT 0,
        holat TEXT DEFAULT 'Rezerv',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE,
        FOREIGN KEY (material_id) REFERENCES ombor(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS buyurtma_tannarxi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER UNIQUE NOT NULL,
        material REAL DEFAULT 0,
        ishchi REAL DEFAULT 0,
        boyoq REAL DEFAULT 0,
        oyna REAL DEFAULT 0,
        furnitura REAL DEFAULT 0,
        transport REAL DEFAULT 0,
        elektr REAL DEFAULT 0,
        tashqi_xizmat REAL DEFAULT 0,
        boshqa REAL DEFAULT 0,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS tizim_xabarlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        foydalanuvchi_turi TEXT DEFAULT '',
        foydalanuvchi_id INTEGER,
        buyurtma_id INTEGER,
        mavzu TEXT NOT NULL,
        matn TEXT NOT NULL,
        kanal TEXT DEFAULT 'Sayt',
        oqildi INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS foydalanuvchi_rollari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomi TEXT UNIQUE NOT NULL,
        huquqlar TEXT DEFAULT ''
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

    CREATE TABLE IF NOT EXISTS mijoz_baholari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        sifat INTEGER DEFAULT 5, muddat INTEGER DEFAULT 5, muomala INTEGER DEFAULT 5,
        yetkazish INTEGER DEFAULT 5, montaj INTEGER DEFAULT 5, izoh TEXT DEFAULT '',
        reklama_ruxsat INTEGER DEFAULT 0, created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS servis_murojaatlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT, buyurtma_id INTEGER NOT NULL,
        turi TEXT DEFAULT 'Servis', muammo TEXT NOT NULL, holat TEXT DEFAULT 'Qabul qilindi',
        servis_sana TEXT DEFAULT '', usta TEXT DEFAULT '', izoh TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS qoshimcha_ishlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT, buyurtma_id INTEGER NOT NULL,
        nomi TEXT NOT NULL, summa REAL DEFAULT 0, qoshimcha_kun INTEGER DEFAULT 0,
        tasdiq INTEGER DEFAULT 0, izoh TEXT DEFAULT '', created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS yetkazishlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        buyurtma_id INTEGER NOT NULL,
        haydovchi_id INTEGER,
        sana TEXT DEFAULT '',
        navbat INTEGER DEFAULT 1,
        mashina TEXT DEFAULT '',
        holat TEXT DEFAULT 'Rejalashtirilgan',
        yolga_chiqdi TEXT DEFAULT '',
        yetib_keldi TEXT DEFAULT '',
        topshirildi TEXT DEFAULT '',
        yetkazildi TEXT DEFAULT '',
        benzin REAL DEFAULT 0,
        yol_xarajati REAL DEFAULT 0,
        izoh TEXT DEFAULT '',
        qadoq_soni INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE,
        FOREIGN KEY (haydovchi_id) REFERENCES ishchilar(id) ON DELETE SET NULL
    );
    """)

    roles=[
        ("Rahbar","all"),("Administrator","orders,customers,workers"),
        ("Hisobchi","finance,payments,expenses"),("Omborchi","stock"),
        ("Ishchi","own_tasks"),("Bo‘lim boshlig‘i","department"),
        ("Shofyor","own_deliveries"),("Mijoz","own_orders")
    ]
    conn.executemany("INSERT OR IGNORE INTO foydalanuvchi_rollari(nomi,huquqlar) VALUES(?,?)",roles)

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
    # V5 migratsiya: ishchining maxfiy pasport ma'lumotlari
    worker_cols=[r[1] for r in conn.execute("PRAGMA table_info(ishchilar)").fetchall()]
    worker_new_cols={
        "pasport":"TEXT DEFAULT ''",
        "jshshir":"TEXT DEFAULT ''",
        "tugilgan_sana":"TEXT DEFAULT ''",
        "yashash_manzil":"TEXT DEFAULT ''",
        "pasport_berilgan_sana":"TEXT DEFAULT ''",
        "pasport_bergan":"TEXT DEFAULT ''",
        "favqulodda_telefon":"TEXT DEFAULT ''"
    }
    for col,ctype in worker_new_cols.items():
        if col not in worker_cols:
            conn.execute(f"ALTER TABLE ishchilar ADD COLUMN {col} {ctype}")

    # V5 migratsiya: xarajat tafsilotlari
    expense_cols=[r[1] for r in conn.execute("PRAGMA table_info(xarajatlar)").fetchall()]
    expense_new_cols={
        "xarajat_nomi":"TEXT DEFAULT ''",
        "kimga_berildi":"TEXT DEFAULT ''",
        "chek_havola":"TEXT DEFAULT ''"
    }
    for col,ctype in expense_new_cols.items():
        if col not in expense_cols:
            conn.execute(f"ALTER TABLE xarajatlar ADD COLUMN {col} {ctype}")

    # V4.5 migratsiya: shofyor ro'yxatdan o'tishi va admin tasdig'i
    driver_cols=[r[1] for r in conn.execute("PRAGMA table_info(shofyor_akkauntlari)").fetchall()]
    if "telefon" not in driver_cols:
        conn.execute("ALTER TABLE shofyor_akkauntlari ADD COLUMN telefon TEXT DEFAULT ''")
    if "admin_tasdiq" not in driver_cols:
        conn.execute("ALTER TABLE shofyor_akkauntlari ADD COLUMN admin_tasdiq INTEGER DEFAULT 0")

    # V4.4.2 migratsiya: yetkazishlar jadvalini eski va yangi tizimga moslash
    y_cols=[r[1] for r in conn.execute("PRAGMA table_info(yetkazishlar)").fetchall()]
    required_y_cols={"haydovchi_id","sana","navbat","mashina","holat","yolga_chiqdi",
                     "yetib_keldi","topshirildi","yetkazildi","benzin","yol_xarajati",
                     "izoh","qadoq_soni","created_at"}
    if not required_y_cols.issubset(set(y_cols)):
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute("""
            CREATE TABLE yetkazishlar_yangi (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                buyurtma_id INTEGER NOT NULL,
                haydovchi_id INTEGER,
                sana TEXT DEFAULT '',
                navbat INTEGER DEFAULT 1,
                mashina TEXT DEFAULT '',
                holat TEXT DEFAULT 'Rejalashtirilgan',
                yolga_chiqdi TEXT DEFAULT '',
                yetib_keldi TEXT DEFAULT '',
                topshirildi TEXT DEFAULT '',
                yetkazildi TEXT DEFAULT '',
                benzin REAL DEFAULT 0,
                yol_xarajati REAL DEFAULT 0,
                izoh TEXT DEFAULT '',
                qadoq_soni INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE,
                FOREIGN KEY (haydovchi_id) REFERENCES ishchilar(id) ON DELETE SET NULL
            )
        """)
        old_cols=set(y_cols)
        rows=conn.execute("SELECT * FROM yetkazishlar").fetchall()
        for r in rows:
            def rv(name, default=None):
                return r[name] if name in old_cols else default
            driver_id=rv("haydovchi_id")
            if driver_id is None:
                driver_id=rv("shofyor_id")
            conn.execute("""
                INSERT INTO yetkazishlar_yangi
                (id,buyurtma_id,haydovchi_id,sana,navbat,mashina,holat,yolga_chiqdi,
                 yetib_keldi,topshirildi,yetkazildi,benzin,yol_xarajati,izoh,qadoq_soni,created_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,(
                rv("id"),rv("buyurtma_id"),driver_id,
                rv("sana",date.today().isoformat()) or date.today().isoformat(),
                rv("navbat",1) or 1,rv("mashina","") or "",
                rv("holat","Rejalashtirilgan") or "Rejalashtirilgan",
                rv("yolga_chiqdi","") or "",rv("yetib_keldi","") or "",
                rv("topshirildi","") or "",rv("yetkazildi","") or "",
                rv("benzin",0) or 0,rv("yol_xarajati",0) or 0,
                rv("izoh","") or "",rv("qadoq_soni",1) or 1,
                rv("created_at",datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ))
        conn.execute("DROP TABLE yetkazishlar")
        conn.execute("ALTER TABLE yetkazishlar_yangi RENAME TO yetkazishlar")
        conn.execute("PRAGMA foreign_keys = ON")

    # V4.4 migratsiya: ishchi topshirig'ining haqiqiy boshlanish va tugash vaqti
    task_cols=[r[1] for r in conn.execute("PRAGMA table_info(ishchi_topshiriqlari)").fetchall()]
    if "boshlandi_vaqt" not in task_cols:
        conn.execute("ALTER TABLE ishchi_topshiriqlari ADD COLUMN boshlandi_vaqt TEXT DEFAULT ''")
    if "tugadi_vaqt" not in task_cols:
        conn.execute("ALTER TABLE ishchi_topshiriqlari ADD COLUMN tugadi_vaqt TEXT DEFAULT ''")

    # V4.2 migratsiya: xarajat to'lov usuli
    xarajat_cols=[r[1] for r in conn.execute("PRAGMA table_info(xarajatlar)").fetchall()]
    if "tolov_usuli" not in xarajat_cols:
        conn.execute("ALTER TABLE xarajatlar ADD COLUMN tolov_usuli TEXT DEFAULT 'Naqd'")

    # V5.0.6 migratsiya: shartnoma uchun qo'shimcha buyurtma ma'lumotlari
    order_contract_cols=[r[1] for r in conn.execute("PRAGMA table_info(buyurtmalar)").fetchall()]
    order_contract_new={
        "pasport_id":"TEXT DEFAULT ''",
        "olcham":"TEXT DEFAULT ''",
        "soni":"INTEGER DEFAULT 1",
        "material":"TEXT DEFAULT ''",
        "rang":"TEXT DEFAULT ''",
        "tolov_usuli":"TEXT DEFAULT 'Naqd'",
        "oraliq_tolov":"REAL DEFAULT 0",
        "montaj":"TEXT DEFAULT 'Kiritilgan'",
        "yetkazish":"TEXT DEFAULT 'Kiritilgan'",
        "kafolat_muddati":"TEXT DEFAULT '12 oy'"
    }
    for col,ctype in order_contract_new.items():
        if col not in order_contract_cols:
            conn.execute(f"ALTER TABLE buyurtmalar ADD COLUMN {col} {ctype}")

    # V3 migratsiya: mijoz kuzatuv tokeni
    cols=[r[1] for r in conn.execute("PRAGMA table_info(buyurtmalar)").fetchall()]
    if "tracking_token" not in cols:
        conn.execute("ALTER TABLE buyurtmalar ADD COLUMN tracking_token TEXT DEFAULT ''")
    for row in conn.execute("SELECT id FROM buyurtmalar WHERE tracking_token IS NULL OR tracking_token='' ").fetchall():
        conn.execute("UPDATE buyurtmalar SET tracking_token=? WHERE id=?",(secrets.token_urlsafe(8),row[0]))
    # V4 migratsiya: mijoz, chegirma, keshbek, kafolat va lokatsiya ustunlari
    order_cols={r[1] for r in conn.execute("PRAGMA table_info(buyurtmalar)").fetchall()}
    migrations={
      "boshlanish_sana":"TEXT DEFAULT ''", "taxminiy_sana":"TEXT DEFAULT ''",
      "kechikish_foiz":"REAL DEFAULT 0", "maks_chegirma_foiz":"REAL DEFAULT 20",
      "kechikish_turi":"TEXT DEFAULT 'Korxona'", "keshbek_foiz":"REAL DEFAULT 0",
      "keshbek_summa":"REAL DEFAULT 0", "keshbek_amal_sana":"TEXT DEFAULT ''",
      "kafolat_boshlanish":"TEXT DEFAULT ''", "kafolat_tugash":"TEXT DEFAULT ''",
      "kafolat_sharti":"TEXT DEFAULT ''", "lokatsiya":"TEXT DEFAULT ''",
      "moljal":"TEXT DEFAULT ''", "qavat":"TEXT DEFAULT ''", "lift":"TEXT DEFAULT ''",
      "katta_mashina":"TEXT DEFAULT ''", "masul_xodim":"TEXT DEFAULT ''",
      "bloklangan":"INTEGER DEFAULT 0"
    }
    for col,typ in migrations.items():
        if col not in order_cols:
            conn.execute(f"ALTER TABLE buyurtmalar ADD COLUMN {col} {typ}")
    stage_cols={r[1] for r in conn.execute("PRAGMA table_info(buyurtma_bosqichlari)").fetchall()}
    for col,typ in {"boshlanish_vaqti":"TEXT DEFAULT ''","tugash_vaqti":"TEXT DEFAULT ''","ishchi":"TEXT DEFAULT ''","izoh":"TEXT DEFAULT ''","media_url":"TEXT DEFAULT ''"}.items():
        if col not in stage_cols: conn.execute(f"ALTER TABLE buyurtma_bosqichlari ADD COLUMN {col} {typ}")
    conn.commit()
    conn.close()



def _safe_filename(value):
    value=str(value or '').strip()
    value=re.sub(r'[^A-Za-z0-9_-]+','_',value)
    return value.strip('_') or 'buyurtma'


def _contract_dir():
    path=os.path.join(os.path.dirname(os.path.abspath(__file__)),'shartnomalar')
    os.makedirs(path,exist_ok=True)
    return path


def _fmt_money(value):
    try:
        return f"{float(value or 0):,.0f}".replace(","," ")
    except Exception:
        return "0"


def _add_docx_field(document,label,value):
    p=document.add_paragraph()
    p.paragraph_format.space_after=Pt(3)
    r=p.add_run(label+": ")
    r.bold=True
    p.add_run(str(value or "________________"))


def generate_order_contract(oid):
    c=get_db()
    order=c.execute("SELECT * FROM buyurtmalar WHERE id=?",(oid,)).fetchone()
    if not order:
        c.close()
        raise ValueError("Buyurtma topilmadi")
    last=c.execute("SELECT COALESCE(MAX(versiya),0) FROM shartnoma_versiyalari WHERE buyurtma_id=?",(oid,)).fetchone()[0]
    version=int(last or 0)+1
    base=f"{_safe_filename(order['kod'])}_shartnoma_v{version}"
    docx_path=os.path.join(_contract_dir(),base+".docx")
    pdf_path=os.path.join(_contract_dir(),base+".pdf")

    # DOCX
    doc=Document()
    section=doc.sections[0]
    section.top_margin=Cm(1.5); section.bottom_margin=Cm(1.5)
    section.left_margin=Cm(1.7); section.right_margin=Cm(1.7)
    styles=doc.styles
    styles['Normal'].font.name='Arial'
    styles['Normal'].font.size=Pt(10)

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("MEBEL360\nMEBEL ISHLAB CHIQARISH (BUYURTMA) SHARTNOMASI")
    r.bold=True; r.font.size=Pt(14)
    p2=doc.add_paragraph()
    p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"Shartnoma № {order['kod']}  |  Sana: {str(order['created_at'])[:10]}").bold=True

    doc.add_heading("1. TOMONLAR VA BUYURTMA", level=1)
    _add_docx_field(doc,"Buyurtmachi",order['mijoz'])
    _add_docx_field(doc,"Pasport/ID",order['pasport_id'])
    _add_docx_field(doc,"Telefon",order['telefon'])
    _add_docx_field(doc,"Manzil",order['manzil'])

    doc.add_heading("2. BUYURTMA TARKIBI", level=1)
    table=doc.add_table(rows=2,cols=7)
    table.alignment=WD_TABLE_ALIGNMENT.CENTER
    table.style='Table Grid'
    headers=["Kod","Mahsulot","O‘lcham","Soni","Material","Rang","Jami narx"]
    for i,h in enumerate(headers):
        cell=table.rows[0].cells[i]
        cell.text=h
        cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    vals=[order['kod'],order['mahsulot'],order['olcham'],order['soni'],
          order['material'],order['rang'],_fmt_money(order['umumiy_narx'])+" so‘m"]
    for i,v in enumerate(vals):
        table.rows[1].cells[i].text=str(v or "")

    doc.add_heading("3. TO‘LOV TARTIBI", level=1)
    _add_docx_field(doc,"Umumiy summa",_fmt_money(order['umumiy_narx'])+" so‘m")
    _add_docx_field(doc,"Avans",_fmt_money(order['oldindan_tolov'])+" so‘m")
    _add_docx_field(doc,"Oraliq to‘lov",_fmt_money(order['oraliq_tolov'])+" so‘m")
    qoldiq=float(order['umumiy_narx'] or 0)-float(order['oldindan_tolov'] or 0)-float(order['oraliq_tolov'] or 0)
    _add_docx_field(doc,"Qoldiq",_fmt_money(max(0,qoldiq))+" so‘m")
    _add_docx_field(doc,"To‘lov usuli",order['tolov_usuli'])

    doc.add_heading("4. MUDDAT, YETKAZISH VA MONTAJ", level=1)
    _add_docx_field(doc,"Boshlanish sanasi",order['boshlanish_sana'])
    _add_docx_field(doc,"Rejalashtirilgan tugash sanasi",order['tugash_sana'] or order['taxminiy_sana'])
    _add_docx_field(doc,"Yetkazish",order['yetkazish'])
    _add_docx_field(doc,"Montaj",order['montaj'])
    _add_docx_field(doc,"Yetkazish manzili",order['manzil'])
    _add_docx_field(doc,"Mo‘ljal",order['moljal'])
    _add_docx_field(doc,"Qavat",order['qavat'])
    _add_docx_field(doc,"Lift",order['lift'])
    _add_docx_field(doc,"Katta mashina kirishi",order['katta_mashina'])

    doc.add_heading("5. SIFAT, TASDIQLASH VA KAFOLAT", level=1)
    paragraphs=[
        "Mebel tasdiqlangan chizma, o‘lcham, material va rang asosida ishlab chiqariladi.",
        "Buyurtmachi chizma yoki materialni tasdiqlagandan keyingi o‘zgarishlar qo‘shimcha narx va muddat bilan bajariladi.",
        f"Kafolat muddati: {order['kafolat_muddati'] or '12 oy'}. "
        f"Kafolat sharti: {order['kafolat_sharti'] or 'Ishlab chiqarish va montaj nuqsonlariga amal qiladi.'}",
        "Noto‘g‘ri foydalanish, namlik, mexanik shikast va uchinchi shaxs ta’miri kafolatga kirmaydi.",
        "Buyurtmachi sababli tasdiq, to‘lov yoki obyektga kirish kechiksa, ish muddati tegishlicha uzayadi."
    ]
    for s in paragraphs:
        doc.add_paragraph(s,style=None)

    doc.add_heading("6. QABUL QILISH VA YAKUNIY SHARTLAR", level=1)
    for s in [
        "Mebel topshirilganda qabul-topshirish dalolatnomasi imzolanadi.",
        "Yakuniy to‘lov to‘liq amalga oshirilmaguncha Pudratchi mahsulotni topshirmaslikka haqli.",
        "Buyurtmachi buyurtmani bekor qilsa, bajarilgan ishlar, sotib olingan materiallar va tasdiqlangan xarajatlar hisobdan ushlab qolinadi.",
        "Tomonlarning elektron xabar, SMS yoki dasturdagi tasdig‘i yozma kelishuv sifatida qabul qilinadi."
    ]:
        doc.add_paragraph(s)

    if order['izoh']:
        doc.add_heading("Qo‘shimcha izoh",level=2)
        doc.add_paragraph(order['izoh'])

    doc.add_paragraph("\n")
    sig=doc.add_table(rows=3,cols=2)
    sig.style='Table Grid'
    sig.cell(0,0).text="PUDRATCHI: Mebel360"
    sig.cell(0,1).text=f"BUYURTMACHI: {order['mijoz']}"
    sig.cell(1,0).text="Imzo: __________________"
    sig.cell(1,1).text="Imzo: __________________"
    sig.cell(2,0).text="Sana: __________________"
    sig.cell(2,1).text="Sana: __________________"
    doc.save(docx_path)

    # PDF
    out=canvas.Canvas(pdf_path,pagesize=A4)
    w,h=A4
    y=h-45
    def draw_wrapped(txt,bold=False,size=10,indent=45,space=15):
        nonlocal y
        out.setFont('Helvetica-Bold' if bold else 'Helvetica',size)
        ascii_txt=(str(txt).replace("‘","'").replace("’","'").replace("–","-")
                   .replace("—","-").replace("o‘","o'").replace("g‘","g'"))
        lines=simpleSplit(ascii_txt,'Helvetica-Bold' if bold else 'Helvetica',size,w-indent-45)
        for line in lines:
            if y<55:
                out.showPage(); y=h-45
            out.drawString(indent,y,line); y-=space

    draw_wrapped("MEBEL360",True,16,45,20)
    draw_wrapped("MEBEL ISHLAB CHIQARISH (BUYURTMA) SHARTNOMASI",True,13,45,20)
    draw_wrapped(f"Shartnoma № {order['kod']} | Sana: {str(order['created_at'])[:10]}",True,10)
    y-=5
    pdf_items=[
        ("Buyurtmachi",order['mijoz']),("Pasport/ID",order['pasport_id']),
        ("Telefon",order['telefon']),("Manzil",order['manzil']),
        ("Mahsulot",order['mahsulot']),("O'lcham",order['olcham']),
        ("Soni",order['soni']),("Material",order['material']),("Rang",order['rang']),
        ("Umumiy summa",_fmt_money(order['umumiy_narx'])+" so'm"),
        ("Avans",_fmt_money(order['oldindan_tolov'])+" so'm"),
        ("Oraliq to'lov",_fmt_money(order['oraliq_tolov'])+" so'm"),
        ("Qoldiq",_fmt_money(max(0,qoldiq))+" so'm"),
        ("To'lov usuli",order['tolov_usuli']),
        ("Boshlanish",order['boshlanish_sana']),("Tugash",order['tugash_sana'] or order['taxminiy_sana']),
        ("Yetkazish",order['yetkazish']),("Montaj",order['montaj']),
        ("Mo'ljal",order['moljal']),("Qavat",order['qavat']),("Lift",order['lift']),
        ("Katta mashina",order['katta_mashina']),("Kafolat",order['kafolat_muddati'])
    ]
    for label,val in pdf_items:
        draw_wrapped(f"{label}: {val or '-'}",False,10)
    y-=7
    for s in paragraphs:
        draw_wrapped("- "+s,False,9,50,13)
    if order['izoh']:
        draw_wrapped("Izoh: "+order['izoh'],False,9)
    y-=20
    draw_wrapped("Pudratchi imzosi: ____________________",False,10)
    draw_wrapped("Buyurtmachi imzosi: ____________________",False,10)
    out.save()

    c.execute("""INSERT INTO shartnoma_versiyalari
                 (buyurtma_id,versiya,docx_fayl,pdf_fayl)
                 VALUES(?,?,?,?)""",(oid,version,docx_path,pdf_path))
    c.commit(); c.close()
    return {"version":version,"docx":docx_path,"pdf":pdf_path}


def latest_order_contract(oid):
    c=get_db()
    row=c.execute("""SELECT * FROM shartnoma_versiyalari
                     WHERE buyurtma_id=? ORDER BY versiya DESC,id DESC LIMIT 1""",(oid,)).fetchone()
    c.close()
    return row


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


# ---------- CSRF HIMOYASI ----------
def get_csrf_token():
    if 'csrf_token' not in session:
        session['csrf_token'] = secrets.token_hex(16)
    return session['csrf_token']

app.jinja_env.globals['csrf_token'] = get_csrf_token


def _csrf_valid():
    token = session.get('csrf_token')
    sent = request.form.get('csrf_token', '')
    return bool(token) and secrets.compare_digest(token, sent)


# ---------- LOGIN URINISHLARINI CHEKLASH (BRUTE-FORCE HIMOYASI) ----------
_login_attempts = {}
LOGIN_MAX_ATTEMPTS = 5
LOGIN_LOCKOUT_SECONDS = 300

def _login_key(username):
    return f"{request.remote_addr}:{username}"

def _is_login_locked(username):
    now = datetime.now().timestamp()
    rec = _login_attempts.get(_login_key(username))
    if rec and rec[1] and now < rec[1]:
        return True, int(rec[1]-now)
    return False, 0

def _register_failed_login(username):
    now = datetime.now().timestamp()
    key = _login_key(username)
    count, locked_until = _login_attempts.get(key, (0, 0))
    count += 1
    locked_until = now + LOGIN_LOCKOUT_SECONDS if count >= LOGIN_MAX_ATTEMPTS else 0
    _login_attempts[key] = (count, locked_until)

def _clear_login_attempts(username):
    _login_attempts.pop(_login_key(username), None)


# ---------- PAROL MUSTAHKAMLIGI ----------
def _weak_password(pw):
    """True qaytaradi agar parol zaif bo'lsa: kamida 8 belgi, kamida bitta harf va bitta raqam kerak."""
    if not pw or len(pw) < 8:
        return True
    has_letter = any(ch.isalpha() for ch in pw)
    has_digit = any(ch.isdigit() for ch in pw)
    return not (has_letter and has_digit)


# ---------- AVTOMATIK KUNLIK ZAXIRA ----------
def _auto_backup_if_needed():
    try:
        backup_dir = os.path.join(_APP_DIR, 'backups')
        os.makedirs(backup_dir, exist_ok=True)
        today = date.today().isoformat()
        marker = os.path.join(backup_dir, f".last_backup_{today}")
        if os.path.exists(marker):
            return
        if not os.path.exists(DB_NAME):
            return
        import shutil
        dest = os.path.join(backup_dir, f"pharm_mebel_backup_{today}.db")
        shutil.copyfile(DB_NAME, dest)
        with open(marker, 'w') as f:
            f.write('1')
        # Oxirgi 30 kunlik zaxiradan ortig'ini o'chirib, joy tejash
        old_backups = sorted(
            [f for f in os.listdir(backup_dir) if f.startswith('pharm_mebel_backup_') and f.endswith('.db')]
        )
        for old in old_backups[:-30]:
            try:
                os.remove(os.path.join(backup_dir, old))
            except Exception:
                pass
    except Exception:
        pass


@app.before_request
def require_login():
    _auto_backup_if_needed()
    public_endpoints = {
        "login", "static", "public_track", "order_qr",
        "worker_register", "worker_verify", "worker_login", "worker_logout", "driver_login", "driver_logout", "driver_register"
    }
    # CSRF tekshiruvi: barcha oddiy HTML-forma orqali yuboriladigan POST so'rovlar uchun.
    # /api/ yo'nalishlari brauzerdagi JS orqali (fetch, o'sha domendan) chaqiriladi, shuning
    # uchun bu yerda tekshirilmaydi.
    if request.method == 'POST' and not request.path.startswith('/api/'):
        if not _csrf_valid():
            return "Xato: sessiya muddati tugagan yoki noto'g'ri so'rov (CSRF). Sahifani qayta yuklab, qayta urinib ko'ring.", 400
    if request.endpoint in public_endpoints or request.path.startswith('/ishchi/public/'):
        return None
    if request.path.startswith('/ishchi/'):
        if not session.get("worker_account_id"):
            return redirect(url_for("worker_login"))
        return None
    if request.path.startswith('/shofyor/'):
        if not session.get("driver_account_id"):
            return redirect(url_for("driver_login"))
        return None
    if not session.get("logged_in"):
        return redirect(url_for("login"))


@app.route("/login", methods=["GET","POST"])
def login():
    error=''
    if request.method=='POST':
        user=request.form.get('user','')
        password=request.form.get('password','')
        locked, wait_sec = _is_login_locked(user)
        if locked:
            error=f'Juda ko‘p xato urinish. {wait_sec} soniyadan so‘ng qayta urinib ko‘ring.'
        elif user==os.environ.get('PHARM_ERP_USER','admin') and password==_ADMIN_PASSWORD:
            _clear_login_attempts(user)
            session['logged_in']=True
            session['user']=user
            log_action('admin_login', f'user={user}')
            return redirect(url_for('home'))
        else:
            _register_failed_login(user)
            log_action('admin_login_failed', f'user={user}')
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
     sifat_ball,tezlik_ball,intizom_ball,izoh,pasport,jshshir,tugilgan_sana,yashash_manzil,
     pasport_berilgan_sana,pasport_bergan,favqulodda_telefon)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
        d.get("ism","").strip(),d.get("familiya","").strip(),d.get("telefon","").strip(),
        d.get("lavozim","").strip(),d.get("ishga_kirgan_sana",""),
        float(d.get("staj_yil") or 0),float(d.get("kunlik_stavka") or 0),
        float(d.get("oylik_maosh") or 0),float(d.get("sifat_ball") or 5),
        float(d.get("tezlik_ball") or 5),float(d.get("intizom_ball") or 5),
        d.get("izoh","").strip(),d.get("pasport","").strip(),d.get("jshshir","").strip(),
        d.get("tugilgan_sana",""),d.get("yashash_manzil","").strip(),
        d.get("pasport_berilgan_sana",""),d.get("pasport_bergan","").strip(),
        d.get("favqulodda_telefon","").strip()
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



@app.route("/api/davomat/keldi", methods=["POST"])
def attendance_clock_in():
    d=jdata()
    try:
        worker_id=int(d["ishchi_id"])
        now=datetime.now()
        sana=now.strftime("%Y-%m-%d")
        vaqt=now.strftime("%H:%M")
        c=get_db()
        row=c.execute("""SELECT id FROM keldi_ketdi
                         WHERE ishchi_id=? AND sana=? AND (ketdi_vaqti='' OR ketdi_vaqti IS NULL)
                         ORDER BY id DESC LIMIT 1""",(worker_id,sana)).fetchone()
        if row:
            c.close()
            return jsonify({"message":"Bu ishchi bugun allaqachon kelgan deb belgilangan."}),400
        c.execute("""INSERT INTO keldi_ketdi(ishchi_id,sana,keldi_vaqti,ketdi_vaqti,ish_soatlari)
                     VALUES(?,?,?,?,0)""",(worker_id,sana,vaqt,""))
        c.commit(); c.close()
        return jsonify({"status":"ok","sana":sana,"vaqt":vaqt})
    except Exception as e:
        return jsonify({"message":str(e)}),400


@app.route("/api/davomat/ketdi", methods=["POST"])
def attendance_clock_out():
    d=jdata()
    try:
        worker_id=int(d["ishchi_id"])
        now=datetime.now()
        sana=now.strftime("%Y-%m-%d")
        vaqt=now.strftime("%H:%M")
        c=get_db()
        row=c.execute("""SELECT * FROM keldi_ketdi
                         WHERE ishchi_id=? AND sana=? AND (ketdi_vaqti='' OR ketdi_vaqti IS NULL)
                         ORDER BY id DESC LIMIT 1""",(worker_id,sana)).fetchone()
        if not row:
            c.close()
            return jsonify({"message":"Avval “Hozir keldi” tugmasini bosing."}),400
        hours=calc_hours(row["keldi_vaqti"],vaqt)
        c.execute("UPDATE keldi_ketdi SET ketdi_vaqti=?, ish_soatlari=? WHERE id=?",
                  (vaqt,hours,row["id"]))
        c.commit(); c.close()
        return jsonify({"status":"ok","sana":sana,"vaqt":vaqt,"ish_soatlari":hours})
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
STAGES=["Buyurtma qabul qilindi","Razmer olindi","Chizma tayyorlanmoqda","Mijoz tasdiqlashi kutilmoqda","Material tayyorlanmoqda","Kesish","Kromka","Teshish","Frezalash","Lazer","Bo‘yash","Sayqalash","Yig‘ish","Sexda tekshirish","Qadoqlash","Yetkazishga tayyor","Haydovchiga topshirildi","Yetkazib berildi","Buyurtma yopildi"]

@app.route("/api/buyurtmalar", methods=["GET","POST"])
def orders():
    if request.method=="POST":
        d=jdata(); c=get_db()
        try:
            cur=c.execute("""INSERT INTO buyurtmalar(
            kod,mijoz,telefon,manzil,mahsulot,umumiy_narx,oldindan_tolov,
            boshlanish_sana,tugash_sana,taxminiy_sana,holat,izoh,tracking_token,
            kechikish_foiz,maks_chegirma_foiz,keshbek_foiz,keshbek_summa,
            kafolat_boshlanish,kafolat_tugash,lokatsiya,moljal,qavat,lift,
            katta_mashina,masul_xodim,pasport_id,olcham,soni,material,rang,
            tolov_usuli,oraliq_tolov,montaj,yetkazish,kafolat_muddati)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (d["kod"],d["mijoz"],d.get("telefon",""),d.get("manzil",""),d.get("mahsulot",""),
             float(d.get("umumiy_narx") or 0),float(d.get("oldindan_tolov") or 0),
             d.get("boshlanish_sana",""),d.get("tugash_sana",""),d.get("taxminiy_sana",""),
             d.get("holat","Yangi"),d.get("izoh",""),secrets.token_urlsafe(8),
             float(d.get("kechikish_foiz") or 0),float(d.get("maks_chegirma_foiz") or 20),
             float(d.get("keshbek_foiz") or 0),float(d.get("keshbek_summa") or 0),
             d.get("kafolat_boshlanish",""),d.get("kafolat_tugash",""),
             d.get("lokatsiya",""),d.get("moljal",""),d.get("qavat",""),
             d.get("lift",""),d.get("katta_mashina",""),d.get("masul_xodim",""),
             d.get("pasport_id",""),d.get("olcham",""),int(d.get("soni") or 1),
             d.get("material",""),d.get("rang",""),d.get("tolov_usuli","Naqd"),
             float(d.get("oraliq_tolov") or 0),d.get("montaj","Kiritilgan"),
             d.get("yetkazish","Kiritilgan"),d.get("kafolat_muddati","12 oy")))
            oid=cur.lastrowid
            c.executemany("INSERT INTO buyurtma_bosqichlari(buyurtma_id,bosqich) VALUES(?,?)",
                          [(oid,s) for s in STAGES])
            c.commit(); c.close()
            try:
                contract=generate_order_contract(oid)
                return jsonify({"status":"ok","buyurtma_id":oid,"shartnoma_versiya":contract["version"]})
            except Exception as contract_error:
                return jsonify({"status":"ok","buyurtma_id":oid,
                                "warning":"Buyurtma saqlandi, shartnoma yaratishda xato: "+str(contract_error)})
        except Exception as e:
            c.close(); return jsonify({"message":str(e)}),400
    c=get_db()
    rows=c.execute("""SELECT *,
    MAX(0,CAST(julianday('now')-julianday(tugash_sana) AS INTEGER)) kechikkan_kun,
    ROUND(MIN(CASE WHEN kechikish_turi='Korxona' THEN MAX(0,julianday('now')-julianday(tugash_sana))*kechikish_foiz ELSE 0 END,maks_chegirma_foiz),2) chegirma_foiz,
    ROUND(MAX(0,(umumiy_narx-oldindan_tolov)*(1-MIN(CASE WHEN kechikish_turi='Korxona' THEN MAX(0,julianday('now')-julianday(tugash_sana))*kechikish_foiz ELSE 0 END,maks_chegirma_foiz)/100.0)),2) qoldiq
    FROM buyurtmalar ORDER BY id DESC""").fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/buyurtma/<int:oid>/bosqichlar")
def order_stages(oid):
    c=get_db(); rows=c.execute("SELECT * FROM buyurtma_bosqichlari WHERE buyurtma_id=? ORDER BY id",(oid,)).fetchall()
    c.close(); return jsonify([dict(r) for r in rows])


@app.route("/api/buyurtma-bosqich/<int:sid>", methods=["POST"])
def stage_toggle(sid):
    d=jdata(); c=get_db()
    done=1 if d.get("bajarildi") else 0
    c.execute("""UPDATE buyurtma_bosqichlari SET bajarildi=?,
      boshlanish_vaqti=CASE WHEN ?=1 AND boshlanish_vaqti='' THEN CURRENT_TIMESTAMP ELSE boshlanish_vaqti END,
      tugash_vaqti=CASE WHEN ?=1 THEN CURRENT_TIMESTAMP ELSE '' END,
      ishchi=COALESCE(NULLIF(?,''),ishchi),izoh=COALESCE(NULLIF(?,''),izoh),media_url=COALESCE(NULLIF(?,''),media_url)
      WHERE id=?""",(done,done,done,d.get("ishchi",''),d.get("izoh",''),d.get("media_url",''),sid))
    c.commit(); c.close(); return jsonify({"status":"ok"})



# ---------- V5.1 PRO MODULLAR ----------
@app.route("/api/pro/bosqich-hodisa", methods=["POST"])
def pro_stage_event_add():
    d=jdata(); c=get_db()
    c.execute("""INSERT INTO buyurtma_bosqich_hodisalari
        (buyurtma_id,bosqich,boshlandi,tugadi,ishchi_id,izoh,media_havola)
        VALUES(?,?,?,?,?,?,?)""",
        (int(d["buyurtma_id"]),d.get("bosqich",""),d.get("boshlandi",""),
         d.get("tugadi",""),int(d["ishchi_id"]) if d.get("ishchi_id") else None,
         d.get("izoh",""),d.get("media_havola","")))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/pro/media", methods=["POST"])
def pro_media_add():
    d=jdata(); c=get_db()
    c.execute("INSERT INTO buyurtma_media(buyurtma_id,turi,havola,izoh) VALUES(?,?,?,?)",
              (int(d["buyurtma_id"]),d.get("turi","Rasm"),d["havola"],d.get("izoh","")))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/pro/tasdiq", methods=["POST"])
def pro_approval_add():
    d=jdata(); c=get_db()
    c.execute("""INSERT INTO buyurtma_tasdiqlari
        (buyurtma_id,turi,holat,izoh,tasdiqlagan,tasdiqlangan_vaqt)
        VALUES(?,?,?,?,?,?)""",
        (int(d["buyurtma_id"]),d.get("turi","Chizma"),d.get("holat","Tasdiqlandi"),
         d.get("izoh",""),d.get("tasdiqlagan",""),
         datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/pro/baho", methods=["POST"])
def pro_rating_add():
    d=jdata(); c=get_db()
    vals=[max(1,min(5,int(d.get(k,5)))) for k in ["sifat","muddat","muomala","yetkazish","montaj"]]
    umumiy=round(sum(vals)/len(vals))
    c.execute("""INSERT INTO baholar
        (buyurtma_id,sifat,muddat,muomala,yetkazish,montaj,umumiy,izoh,rasm_havola,reklama_ruxsat)
        VALUES(?,?,?,?,?,?,?,?,?,?)""",
        (int(d["buyurtma_id"]),*vals,umumiy,d.get("izoh",""),d.get("rasm_havola",""),
         1 if d.get("reklama_ruxsat") else 0))
    c.commit(); c.close(); return jsonify({"status":"ok","umumiy":umumiy})

@app.route("/api/pro/servis", methods=["POST"])
def pro_service_add():
    d=jdata(); c=get_db()
    c.execute("""INSERT INTO servis_murojaatlari
        (buyurtma_id,turi,muammo,media_havola,holat,servis_sana,pullik)
        VALUES(?,?,?,?,?,?,?)""",
        (int(d["buyurtma_id"]),d.get("turi","Servis"),d["muammo"],
         d.get("media_havola",""),d.get("holat","Qabul qilindi"),
         d.get("servis_sana",""),1 if d.get("pullik") else 0))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/pro/qoshimcha-ish", methods=["POST"])
def pro_extra_work_add():
    d=jdata(); c=get_db()
    c.execute("""INSERT INTO qoshimcha_ishlar
        (buyurtma_id,nomi,summa,qoshimcha_kun,mijoz_tasdiq,kim_kiritdi)
        VALUES(?,?,?,?,?,?)""",
        (int(d["buyurtma_id"]),d["nomi"],float(d.get("summa") or 0),
         int(d.get("qoshimcha_kun") or 0),1 if d.get("mijoz_tasdiq") else 0,
         d.get("kim_kiritdi","Rahbar")))
    c.execute("UPDATE buyurtmalar SET umumiy_narx=umumiy_narx+? WHERE id=?",
              (float(d.get("summa") or 0),int(d["buyurtma_id"])))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/pro/rezerv", methods=["POST"])
def pro_stock_reserve():
    d=jdata(); c=get_db()
    c.execute("""INSERT INTO ombor_rezervlari(buyurtma_id,material_id,miqdor)
                 VALUES(?,?,?)""",
              (int(d["buyurtma_id"]),int(d["material_id"]),float(d.get("miqdor") or 0)))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/pro/tannarx", methods=["POST"])
def pro_cost_save():
    d=jdata(); c=get_db()
    vals=[float(d.get(k) or 0) for k in ["material","ishchi","boyoq","oyna","furnitura","transport","elektr","tashqi_xizmat","boshqa"]]
    c.execute("""INSERT INTO buyurtma_tannarxi
        (buyurtma_id,material,ishchi,boyoq,oyna,furnitura,transport,elektr,tashqi_xizmat,boshqa)
        VALUES(?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(buyurtma_id) DO UPDATE SET
        material=excluded.material,ishchi=excluded.ishchi,boyoq=excluded.boyoq,
        oyna=excluded.oyna,furnitura=excluded.furnitura,transport=excluded.transport,
        elektr=excluded.elektr,tashqi_xizmat=excluded.tashqi_xizmat,boshqa=excluded.boshqa""",
        (int(d["buyurtma_id"]),*vals))
    c.commit(); c.close(); return jsonify({"status":"ok","jami":sum(vals)})

@app.route("/api/pro/xabar", methods=["POST"])
def pro_message_add():
    d=jdata(); c=get_db()
    c.execute("""INSERT INTO tizim_xabarlari
        (foydalanuvchi_turi,foydalanuvchi_id,buyurtma_id,mavzu,matn,kanal)
        VALUES(?,?,?,?,?,?)""",
        (d.get("foydalanuvchi_turi","Mijoz"),d.get("foydalanuvchi_id"),
         d.get("buyurtma_id"),d["mavzu"],d["matn"],d.get("kanal","Sayt")))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/pro/buyurtma/<int:oid>")
def pro_order_bundle(oid):
    c=get_db()
    order=c.execute("SELECT * FROM buyurtmalar WHERE id=?",(oid,)).fetchone()
    stages=c.execute("SELECT * FROM buyurtma_bosqich_hodisalari WHERE buyurtma_id=? ORDER BY id",(oid,)).fetchall()
    media=c.execute("SELECT * FROM buyurtma_media WHERE buyurtma_id=? ORDER BY id DESC",(oid,)).fetchall()
    approvals=c.execute("SELECT * FROM buyurtma_tasdiqlari WHERE buyurtma_id=? ORDER BY id DESC",(oid,)).fetchall()
    payments=c.execute("SELECT * FROM buyurtma_tolovlari WHERE buyurtma_id=? ORDER BY id DESC",(oid,)).fetchall()
    services=c.execute("SELECT * FROM servis_murojaatlari WHERE buyurtma_id=? ORDER BY id DESC",(oid,)).fetchall()
    extras=c.execute("SELECT * FROM qoshimcha_ishlar WHERE buyurtma_id=? ORDER BY id DESC",(oid,)).fetchall()
    rating=c.execute("SELECT * FROM baholar WHERE buyurtma_id=? ORDER BY id DESC LIMIT 1",(oid,)).fetchone()
    cost=c.execute("SELECT * FROM buyurtma_tannarxi WHERE buyurtma_id=?",(oid,)).fetchone()
    c.close()
    return jsonify({
        "order":dict(order) if order else None,
        "stages":[dict(x) for x in stages],
        "media":[dict(x) for x in media],
        "approvals":[dict(x) for x in approvals],
        "payments":[dict(x) for x in payments],
        "services":[dict(x) for x in services],
        "extras":[dict(x) for x in extras],
        "rating":dict(rating) if rating else None,
        "cost":dict(cost) if cost else None
    })

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
        c.execute("INSERT INTO xarajatlar(sana,kategoriya,miqdor,tavsifi,buyurtma_kodi,tolov_usuli,xarajat_nomi,kimga_berildi,chek_havola) VALUES(?,?,?,?,?,?,?,?,?)",
                  (d["sana"],d.get("kategoriya","Boshqa"),float(d.get("miqdor") or 0),
                   d.get("tavsifi",""),d.get("buyurtma_kodi",""),d.get("tolov_usuli","Naqd"),
                   d.get("xarajat_nomi",""),d.get("kimga_berildi",""),d.get("chek_havola","")))
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
    c=get_db()
    order=c.execute("SELECT * FROM buyurtmalar WHERE tracking_token=?",(token,)).fetchone()
    if not order:
        c.close()
        return "Buyurtma topilmadi",404

    stages=c.execute("SELECT * FROM buyurtma_bosqichlari WHERE buyurtma_id=? ORDER BY id",
                     (order['id'],)).fetchall()

    delivery=c.execute(
        "SELECT y.*, i.ism AS haydovchi_ism, i.familiya AS haydovchi_familiya, "
        "i.telefon AS haydovchi_telefon "
        "FROM yetkazishlar y "
        "LEFT JOIN ishchilar i ON i.id=y.haydovchi_id "
        "WHERE y.buyurtma_id=? ORDER BY y.id DESC LIMIT 1",
        (order['id'],)
    ).fetchone()
    c.close()

    done=sum(int(x['bajarildi']) for x in stages)
    pct=round(done*100/len(stages)) if stages else 0

    html='''<!doctype html>
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <meta http-equiv="refresh" content="15">
    <title>Buyurtma kuzatuvi</title>
    <style>
    body{font-family:Arial;background:#eef3f8;margin:0;padding:18px;color:#172033}
    .box{max-width:650px;margin:auto;background:white;border-radius:18px;padding:20px;box-shadow:0 8px 30px #0002}
    h1{color:#1d4ed8}.bar{height:18px;background:#e2e8f0;border-radius:20px;overflow:hidden}
    .fill{height:100%;background:#16a34a}.stage{padding:10px;border-bottom:1px solid #eee}
    .ok{color:#15803d}.wait{color:#64748b}
    .delivery{margin:18px 0;padding:16px;border-radius:15px;background:#eff6ff;border:2px solid #bfdbfe}
    .delivery h3{margin:0 0 10px;color:#1d4ed8}
    .status{font-size:21px;font-weight:bold;padding:10px;border-radius:10px;background:white}
    .road{color:#ea580c}.arrived{color:#7c3aed}.delivered{color:#15803d}.planned{color:#64748b}
    .small{font-size:13px;color:#64748b}
    </style>
    <div class="box">
      <h1>Mebel360</h1>
      <h2>Buyurtma: {{o['kod']}}</h2>
      <p><b>Mijoz:</b> {{o['mijoz']}}</p>
      <p><b>Mahsulot:</b> {{o['mahsulot']}}</p>
      <p><b>Mas’ul xodim:</b> {{o['masul_xodim'] or 'Belgilanmagan'}}</p>
      <p><b>Manzil:</b> {{o['manzil']}}</p>
      <p><b>Tugash sanasi:</b> {{o['tugash_sana']}}</p>
      <p><b>Kafolat:</b> {{o['kafolat_boshlanish']}} — {{o['kafolat_tugash']}}</p>

      <div class="delivery">
        <h3>🚚 Yetkazib berish holati</h3>
        {% if delivery %}
          {% set h=delivery['holat'] or 'Rejalashtirilgan' %}
          <div class="status
            {{'road' if h in ['Yo‘lga chiqdim','Yo‘lga chiqdi'] else
              'arrived' if h in ['Yetib keldim','Yetib keldi'] else
              'delivered' if h in ['Yetkazib berdim','Yetkazib berildi','Yetkazildi'] else
              'planned'}}">
            {% if h in ['Yo‘lga chiqdim','Yo‘lga chiqdi'] %}🚚 Shofyor yo‘lda
            {% elif h in ['Yetib keldim','Yetib keldi'] %}📍 Shofyor yetib keldi
            {% elif h in ['Yetkazib berdim','Yetkazib berildi','Yetkazildi'] %}✅ Buyurtma yetkazib berildi
            {% else %}🕒 Yetkazish rejalashtirilgan
            {% endif %}
          </div>
          <p><b>Shofyor:</b> {{(delivery['haydovchi_ism'] or '')+' '+(delivery['haydovchi_familiya'] or '')}}</p>
          <p><b>Telefon:</b> {{delivery['haydovchi_telefon'] or 'Ko‘rsatilmagan'}}</p>
          <p><b>Mashina:</b> {{delivery['mashina'] or 'Ko‘rsatilmagan'}}</p>
          {% if delivery['yolga_chiqdi'] %}<p><b>Yo‘lga chiqqan vaqt:</b> {{delivery['yolga_chiqdi']}}</p>{% endif %}
          {% if delivery['yetib_keldi'] %}<p><b>Yetib kelgan vaqt:</b> {{delivery['yetib_keldi']}}</p>{% endif %}
          {% if delivery['topshirildi'] %}<p><b>Topshirilgan vaqt:</b> {{delivery['topshirildi']}}</p>{% endif %}
        {% else %}
          <div class="status planned">🚛 Hali shofyor biriktirilmagan</div>
        {% endif %}
        <div class="small">Sahifa har 15 soniyada avtomatik yangilanadi.</div>
      </div>

      <h3>Tayyorlik: {{pct}}%</h3>
      <div class="bar"><div class="fill" style="width:{{pct}}%"></div></div>
      {% for s in stages %}
        <div class="stage {{'ok' if s['bajarildi'] else 'wait'}}">
          {{'✅' if s['bajarildi'] else '⬜'}} {{s['bosqich']}}
        </div>
      {% endfor %}
    </div>'''
    return render_template_string(html,o=order,stages=stages,pct=pct,delivery=delivery)


@app.route("/buyurtma/<int:oid>/qr.png")
def order_qr(oid):
    c=get_db(); row=c.execute("SELECT tracking_token FROM buyurtmalar WHERE id=?",(oid,)).fetchone(); c.close()
    if not row:return "Topilmadi",404
    url=request.url_root.rstrip('/')+url_for('public_track',token=row['tracking_token'])
    img=qrcode.make(url); out=io.BytesIO(); img.save(out,format='PNG'); out.seek(0)
    return send_file(out,mimetype='image/png',download_name=f'buyurtma_{oid}_qr.png')

@app.route("/buyurtma/<int:oid>/shartnoma-yaratish", methods=["POST"])
def order_contract_regenerate(oid):
    try:
        result=generate_order_contract(oid)
        return jsonify({"status":"ok","versiya":result["version"]})
    except Exception as e:
        return jsonify({"message":str(e)}),400


@app.route("/buyurtma/<int:oid>/shartnoma.docx")
def order_contract_docx(oid):
    row=latest_order_contract(oid)
    if not row:
        try:
            generate_order_contract(oid)
            row=latest_order_contract(oid)
        except Exception as e:
            return "Shartnoma yaratishda xato: "+str(e),500
    if not row or not os.path.exists(row["docx_fayl"]):
        return "Shartnoma Word fayli topilmadi",404
    return send_file(row["docx_fayl"],as_attachment=True,
                     download_name=f"shartnoma_{oid}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/buyurtma/<int:oid>/shartnoma.pdf")
def order_contract_pdf(oid):
    row=latest_order_contract(oid)
    if not row:
        try:
            generate_order_contract(oid)
            row=latest_order_contract(oid)
        except Exception as e:
            return "Shartnoma yaratishda xato: "+str(e),500
    if not row or not os.path.exists(row["pdf_fayl"]):
        return "Shartnoma PDF fayli topilmadi",404
    return send_file(row["pdf_fayl"],mimetype="application/pdf",
                     as_attachment=False,download_name=f"shartnoma_{oid}.pdf")


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
    p.drawString(70,y-20,'Rahmat! Mebel360 xizmatidan foydalanganingiz uchun.')
    p.save(); out.seek(0); return send_file(out,mimetype='application/pdf',as_attachment=True,download_name=f"chek_{order['kod']}.pdf")

@app.route("/api/buyurtma/<int:oid>/link")
def order_public_link(oid):
    c=get_db(); row=c.execute("SELECT tracking_token FROM buyurtmalar WHERE id=?",(oid,)).fetchone(); c.close()
    if not row:return jsonify({'message':'Topilmadi'}),404
    return jsonify({'url':request.url_root.rstrip('/')+url_for('public_track',token=row['tracking_token'])})

@app.route("/shartnoma-namuna")
def contract_template_download():
    try:
        base_dir=os.path.dirname(os.path.abspath(__file__))
        file_path=os.path.join(base_dir,"Mebel_Shartnoma_Tuzatilgan.docx")
        if not os.path.exists(file_path):
            return "Word shartnoma fayli topilmadi: "+file_path,404
        with open(file_path,"rb") as f:
            data=f.read()
        return Response(
            data,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition":"attachment; filename=Mebel_Shartnoma.docx"}
        )
    except Exception as e:
        return "Word shartnomani ochishda xato: "+str(e),500

@app.route("/shartnoma-pdf")
def contract_pdf_download():
    try:
        base_dir=os.path.dirname(os.path.abspath(__file__))
        file_path=os.path.join(base_dir,"Mebel_Shartnoma_Yakuniy.pdf")
        if not os.path.exists(file_path):
            return "PDF shartnoma fayli topilmadi: "+file_path,404
        with open(file_path,"rb") as f:
            data=f.read()
        return Response(
            data,
            mimetype="application/pdf",
            headers={"Content-Disposition":"inline; filename=Mebel_Shartnoma.pdf"}
        )
    except Exception as e:
        return "PDF shartnomani ochishda xato: "+str(e),500

@app.route("/backup")
def backup_db():
    if not os.path.exists(DB_NAME):
        return jsonify({"message":"Baza topilmadi"}),404
    log_action('backup_downloaded', f'user={session.get("user","")}')
    return send_file(DB_NAME, as_attachment=True, download_name=f"pharm_mebel_backup_{date.today().isoformat()}.db")

@app.route("/api/audit")
def audit_get():
    c=get_db(); rows=c.execute("SELECT * FROM audit_log ORDER BY id DESC LIMIT 200").fetchall(); c.close(); return jsonify([dict(r) for r in rows])


# ---------- MIJOZ SERVISI / KAFOLAT / YETKAZISH ----------
@app.route("/api/buyurtma/<int:oid>/baho", methods=["GET","POST"])
def order_rating(oid):
    c=get_db()
    if request.method=="POST":
        d=jdata(); vals=[max(1,min(5,int(d.get(k) or 5))) for k in ("sifat","muddat","muomala","yetkazish","montaj")]
        c.execute("INSERT INTO mijoz_baholari(buyurtma_id,sifat,muddat,muomala,yetkazish,montaj,izoh,reklama_ruxsat) VALUES(?,?,?,?,?,?,?,?,?)",
                  (oid,*vals,d.get("izoh",""),1 if d.get("reklama_ruxsat") else 0)); c.commit(); c.close(); return jsonify({"status":"ok"})
    rows=c.execute("SELECT *,ROUND((sifat+muddat+muomala+yetkazish+montaj)/5.0,1) umumiy FROM mijoz_baholari WHERE buyurtma_id=? ORDER BY id DESC",(oid,)).fetchall(); c.close(); return jsonify([dict(r) for r in rows])

@app.route("/api/servis", methods=["GET","POST"])
def service_api():
    c=get_db()
    if request.method=="POST":
        d=jdata(); c.execute("INSERT INTO servis_murojaatlari(buyurtma_id,turi,muammo,holat,servis_sana,usta,izoh) VALUES(?,?,?,?,?,?,?)",
          (int(d["buyurtma_id"]),d.get("turi","Servis"),d.get("muammo",""),d.get("holat","Qabul qilindi"),d.get("servis_sana",""),d.get("usta",""),d.get("izoh",""))); c.commit(); c.close(); return jsonify({"status":"ok"})
    rows=c.execute("""SELECT s.*,b.kod,b.mijoz FROM servis_murojaatlari s JOIN buyurtmalar b ON b.id=s.buyurtma_id ORDER BY s.id DESC""").fetchall(); c.close(); return jsonify([dict(r) for r in rows])

@app.route("/api/servis/<int:sid>", methods=["POST"])
def service_update(sid):
    d=jdata(); c=get_db(); c.execute("UPDATE servis_murojaatlari SET holat=?,servis_sana=?,usta=?,izoh=? WHERE id=?",(d.get("holat","Ko‘rib chiqilmoqda"),d.get("servis_sana",""),d.get("usta",""),d.get("izoh",""),sid)); c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/qoshimcha-ish", methods=["GET","POST"])
def extras_api():
    c=get_db()
    if request.method=="POST":
        d=jdata(); c.execute("INSERT INTO qoshimcha_ishlar(buyurtma_id,nomi,summa,qoshimcha_kun,tasdiq,izoh) VALUES(?,?,?,?,?,?)",(int(d["buyurtma_id"]),d.get("nomi",""),float(d.get("summa") or 0),int(d.get("qoshimcha_kun") or 0),1 if d.get("tasdiq") else 0,d.get("izoh",""))); c.commit(); c.close(); return jsonify({"status":"ok"})
    rows=c.execute("SELECT q.*,b.kod,b.mijoz FROM qoshimcha_ishlar q JOIN buyurtmalar b ON b.id=q.buyurtma_id ORDER BY q.id DESC").fetchall(); c.close(); return jsonify([dict(r) for r in rows])

@app.route("/api/yetkazish", methods=["GET","POST"])
def delivery_api():
    c=get_db()
    if request.method=="POST":
        d=jdata(); c.execute("INSERT INTO yetkazishlar(buyurtma_id,haydovchi_id,sana,navbat,mashina,holat,benzin,yol_xarajati,izoh) VALUES(?,?,?,?,?,?,?,?,?)",(int(d["buyurtma_id"]),int(d["haydovchi_id"]) if d.get("haydovchi_id") else None,d.get("sana") or date.today().isoformat(),int(d.get("navbat") or 1),d.get("mashina",""),d.get("holat","Rejalashtirilgan"),float(d.get("benzin") or 0),float(d.get("yol_xarajati") or 0),d.get("izoh",""))); c.commit(); c.close(); return jsonify({"status":"ok"})
    rows=c.execute("""SELECT y.*,b.kod,b.mijoz,b.telefon,b.manzil,b.lokatsiya,i.ism haydovchi_ism,i.familiya haydovchi_familiya FROM yetkazishlar y JOIN buyurtmalar b ON b.id=y.buyurtma_id LEFT JOIN ishchilar i ON i.id=y.haydovchi_id ORDER BY y.sana DESC,y.navbat""").fetchall(); c.close(); return jsonify([dict(r) for r in rows])

@app.route("/api/yetkazish/<int:yid>/holat", methods=["POST"])
def delivery_status(yid):
    d=jdata(); holat=d.get("holat","Rejalashtirilgan"); now=datetime.now().strftime('%Y-%m-%d %H:%M')
    field={"Yo‘lga chiqdim":"yolga_chiqdi","Yetib keldim":"yetib_keldi","Yetkazib berdim":"topshirildi"}.get(holat)
    c=get_db(); c.execute("UPDATE yetkazishlar SET holat=? WHERE id=?",(holat,yid))
    if field: c.execute(f"UPDATE yetkazishlar SET {field}=? WHERE id=?",(now,yid))
    c.commit(); c.close(); return jsonify({"status":"ok"})

@app.route("/api/mijoz-xulosa")
def customer_summary():
    c=get_db(); row=c.execute("""SELECT COUNT(*) buyurtmalar,COALESCE(SUM(umumiy_narx),0) jami_summa,COALESCE(SUM(oldindan_tolov),0) tushgan,
      COALESCE(AVG((sifat+muddat+muomala+yetkazish+montaj)/5.0),0) reyting FROM buyurtmalar LEFT JOIN mijoz_baholari ON mijoz_baholari.buyurtma_id=buyurtmalar.id""").fetchone(); c.close(); return jsonify(dict(row))

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
                if os.environ.get('DEV_SHOW_OTP','0')=='1':
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
        if len(login)<3 or not ism:
            error='Ism va kamida 3 belgili login kiriting.'
        elif _weak_password(password):
            error='Parol kamida 8 belgidan iborat bo‘lib, harf va raqamni birga o‘z ichiga olishi kerak.'
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
        locked, wait_sec = _is_login_locked('w:'+login)
        if locked:
            error=f'Juda ko‘p xato urinish. {wait_sec} soniyadan so‘ng qayta urinib ko‘ring.'
        else:
            c=get_db(); row=c.execute('SELECT * FROM ishchi_akkauntlari WHERE login=? AND faol=1',(login,)).fetchone(); c.close()
            if not row or not check_password_hash(row['parol_hash'],password):
                _register_failed_login('w:'+login)
                error='Login yoki parol xato.'
            elif not row['admin_tasdiq']:
                error='Admin hali akkauntingizni tasdiqlamagan.'
            else:
                _clear_login_attempts('w:'+login)
                session.clear(); session['worker_account_id']=row['id']; session['worker_id']=row['ishchi_id']
                log_action('worker_login', f'login={login}')
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
    tasks=c.execute('SELECT * FROM ishchi_topshiriqlari WHERE ishchi_id=? ORDER BY CASE WHEN holat="Ishlayapti" THEN 0 WHEN holat="Yangi" THEN 1 ELSE 2 END,sana DESC,id DESC LIMIT 50',(wid,)).fetchall()
    active_task=c.execute('SELECT * FROM ishchi_topshiriqlari WHERE ishchi_id=? AND holat="Ishlayapti" ORDER BY id DESC LIMIT 1',(wid,)).fetchone()
    worker_state='Ishlayapti' if active_task else 'Bo‘sh'
    stats=c.execute('''SELECT COUNT(DISTINCT sana) kun,COALESCE(SUM(ish_soatlari),0) soat FROM keldi_ketdi WHERE ishchi_id=?''',(wid,)).fetchone()
    result=c.execute('SELECT COALESCE(SUM(miqdor),0) miqdor FROM ish_natijalari WHERE ishchi_id=?',(wid,)).fetchone()
    rating=round(((worker['sifat_ball']+worker['tezlik_ball']+worker['intizom_ball'])/3.0)+min(float(result['miqdor'] or 0)/100.0,5),2)
    c.close()
    return render_template_string(WORKER_DASHBOARD_HTML,worker=worker,tasks=tasks,stats=stats,rating=rating,worker_state=worker_state,active_task=active_task)


@app.route('/ishchi/topshiriq/<int:tid>/boshlash', methods=['POST'])
def worker_task_start(tid):
    wid=session.get('worker_id')
    now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c=get_db()
    other=c.execute('SELECT id FROM ishchi_topshiriqlari WHERE ishchi_id=? AND holat="Ishlayapti" AND id<>? LIMIT 1',(wid,tid)).fetchone()
    if other:
        c.close()
        flash('Avval ishlayotgan topshiriqni tugating.')
        return redirect(url_for('worker_dashboard'))
    c.execute('''UPDATE ishchi_topshiriqlari
                 SET holat='Ishlayapti',progress=CASE WHEN progress<1 THEN 1 ELSE progress END,
                     boshlandi_vaqt=CASE WHEN boshlandi_vaqt='' OR boshlandi_vaqt IS NULL THEN ? ELSE boshlandi_vaqt END
                 WHERE id=? AND ishchi_id=?''',(now,tid,wid))
    c.commit(); c.close()
    return redirect(url_for('worker_dashboard'))


@app.route('/ishchi/topshiriq/<int:tid>/tugatish', methods=['POST'])
def worker_task_finish(tid):
    wid=session.get('worker_id')
    now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c=get_db()
    c.execute('''UPDATE ishchi_topshiriqlari
                 SET holat='Tayyor',progress=100,tugadi_vaqt=?
                 WHERE id=? AND ishchi_id=?''',(now,tid,wid))
    c.commit(); c.close()
    return redirect(url_for('worker_dashboard'))


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
            aid=int(request.form['account_id']); c.execute('UPDATE ishchi_akkauntlari SET admin_tasdiq=1 WHERE id=?',(aid,)); log_action('worker_approved', f'account_id={aid}')
        elif action=='block':
            aid=int(request.form['account_id']); c.execute('UPDATE ishchi_akkauntlari SET faol=0 WHERE id=?',(aid,)); log_action('worker_blocked', f'account_id={aid}')
        elif action=='task':
            c.execute('''INSERT INTO ishchi_topshiriqlari(ishchi_id,buyurtma_kodi,ish_turi,tavsif,holat,progress,sana,tugash_sana)
                         VALUES(?,?,?,?,?,?,?,?)''',(
                int(request.form['ishchi_id']),request.form.get('buyurtma_kodi','').strip(),request.form.get('ish_turi','').strip(),
                request.form.get('tavsif','').strip(),'Yangi',0,request.form.get('sana') or date.today().isoformat(),request.form.get('tugash_sana','')
            ))
        c.commit()
    accounts=c.execute('''SELECT a.*,i.ism,i.familiya,i.lavozim FROM ishchi_akkauntlari a LEFT JOIN ishchilar i ON i.id=a.ishchi_id ORDER BY a.id DESC''').fetchall()
    workers=c.execute('SELECT id,ism,familiya,lavozim FROM ishchilar WHERE faol=1 ORDER BY ism').fetchall()
    worker_states=c.execute('''SELECT i.id,i.ism,i.familiya,i.lavozim,
        CASE WHEN t.id IS NULL THEN 'Bo‘sh' ELSE 'Ishlayapti' END ish_holati,
        COALESCE(t.buyurtma_kodi,'') buyurtma_kodi,
        COALESCE(t.ish_turi,'') ish_turi,
        COALESCE(t.boshlandi_vaqt,'') boshlandi_vaqt
        FROM ishchilar i
        LEFT JOIN ishchi_topshiriqlari t ON t.id=(
            SELECT t2.id FROM ishchi_topshiriqlari t2
            WHERE t2.ishchi_id=i.id AND t2.holat='Ishlayapti'
            ORDER BY t2.id DESC LIMIT 1
        )
        WHERE i.faol=1 ORDER BY ish_holati DESC,i.ism''').fetchall()
    tasks=c.execute('''SELECT t.*,i.ism,i.familiya FROM ishchi_topshiriqlari t JOIN ishchilar i ON i.id=t.ishchi_id ORDER BY t.id DESC LIMIT 100''').fetchall()
    otp_rows=c.execute('SELECT telefon,created_at FROM ishchi_otp WHERE ishlatildi=0 ORDER BY id DESC LIMIT 20').fetchall()
    c.close()
    return render_template_string(WORKER_ADMIN_HTML,accounts=accounts,workers=workers,worker_states=worker_states,tasks=tasks,otp_rows=otp_rows,today=date.today().isoformat())



# ---------- SHOFYOR KABINETI ----------
@app.route('/shofyor/royxat', methods=['GET','POST'])
def driver_register():
    error=''
    msg=''
    if request.method=='POST':
        ism=request.form.get('ism','').strip()
        familiya=request.form.get('familiya','').strip()
        telefon=normalize_phone(request.form.get('telefon',''))
        login=request.form.get('login','').strip()
        password=request.form.get('password','')
        if not ism or len(login)<3 or len(telefon)<10:
            error='Ism, telefon va kamida 3 belgili login kiriting.'
        elif _weak_password(password):
            error='Parol kamida 8 belgidan iborat bo‘lib, harf va raqamni birga o‘z ichiga olishi kerak.'
        else:
            c=get_db()
            if c.execute("SELECT 1 FROM shofyor_akkauntlari WHERE login=?",(login,)).fetchone():
                error='Bu login band.'
            elif c.execute("SELECT 1 FROM shofyor_akkauntlari WHERE telefon=?",(telefon,)).fetchone():
                error='Bu telefon avval ro‘yxatdan o‘tgan.'
            else:
                worker=c.execute("SELECT id FROM ishchilar WHERE telefon=? AND faol=1",(telefon,)).fetchone()
                if worker:
                    wid=worker['id']
                    c.execute("UPDATE ishchilar SET lavozim='Shofyor' WHERE id=?",(wid,))
                else:
                    cur=c.execute("INSERT INTO ishchilar(ism,familiya,telefon,lavozim) VALUES(?,?,?,'Shofyor')",
                                  (ism,familiya,telefon))
                    wid=cur.lastrowid
                c.execute("INSERT INTO shofyor_akkauntlari(ishchi_id,login,parol_hash,telefon,faol,admin_tasdiq) VALUES(?,?,?,?,1,0)",
                          (wid,login,generate_password_hash(password),telefon))
                c.commit()
                msg='Ro‘yxatdan o‘tdingiz. Endi rahbar tasdiqlaydi.'
            c.close()
    return render_template_string(DRIVER_REGISTER_HTML,error=error,msg=msg)

@app.route('/shofyor/login', methods=['GET','POST'])
def driver_login():
    error=''
    if request.method=='POST':
        login=request.form.get('login','').strip()
        password=request.form.get('password','')
        locked, wait_sec = _is_login_locked('d:'+login)
        if locked:
            error=f'Juda ko‘p xato urinish. {wait_sec} soniyadan so‘ng qayta urinib ko‘ring.'
        else:
            c=get_db()
            row=c.execute("SELECT a.*,i.ism,i.familiya FROM shofyor_akkauntlari a JOIN ishchilar i ON i.id=a.ishchi_id WHERE a.login=? AND a.faol=1",(login,)).fetchone()
            c.close()
            if not row or not check_password_hash(row['parol_hash'],password):
                _register_failed_login('d:'+login)
                error='Login yoki parol xato.'
            elif not row['admin_tasdiq']:
                error='Rahbar hali akkauntingizni tasdiqlamagan.'
            else:
                _clear_login_attempts('d:'+login)
                session.clear()
                session['driver_account_id']=row['id']
                session['driver_id']=row['ishchi_id']
                log_action('driver_login', f'login={login}')
                return redirect(url_for('driver_dashboard'))
    return render_template_string(DRIVER_LOGIN_HTML,error=error)

@app.route('/shofyor/logout')
def driver_logout():
    session.clear()
    return redirect(url_for('driver_login'))

@app.route('/shofyor/kabinet')
def driver_dashboard():
    did=session.get('driver_id')
    c=get_db()
    driver=c.execute("SELECT * FROM ishchilar WHERE id=?",(did,)).fetchone()
    deliveries=c.execute("SELECT y.*,b.kod,b.mijoz,b.telefon,b.manzil,b.mahsulot,b.lokatsiya,b.moljal,b.qavat,b.lift,b.katta_mashina,b.izoh buyurtma_izoh FROM yetkazishlar y JOIN buyurtmalar b ON b.id=y.buyurtma_id WHERE y.haydovchi_id=? ORDER BY CASE WHEN y.holat='Yetkazib berildi' THEN 1 ELSE 0 END,y.navbat,y.id DESC",(did,)).fetchall()
    c.close()
    return render_template_string(DRIVER_DASHBOARD_HTML,driver=driver,deliveries=deliveries)

@app.route('/shofyor/yetkazish/<int:yid>')
def driver_delivery_detail(yid):
    did=session.get('driver_id')
    c=get_db()
    row=c.execute("SELECT y.*,b.kod,b.mijoz,b.telefon,b.manzil,b.mahsulot,b.lokatsiya,b.moljal,b.qavat,b.lift,b.katta_mashina,b.izoh buyurtma_izoh FROM yetkazishlar y JOIN buyurtmalar b ON b.id=y.buyurtma_id WHERE y.id=? AND y.haydovchi_id=?",(yid,did)).fetchone()
    c.close()
    if not row:
        return 'Yetkazish topilmadi',404
    return render_template_string(DRIVER_DETAIL_HTML,x=row)

@app.route('/shofyor/yetkazish/<int:yid>/holat', methods=['POST'])
def driver_delivery_status(yid):
    did=session.get('driver_id')
    action=request.form.get('action','')
    now=datetime.now().strftime('%Y-%m-%d %H:%M')
    mapping={'yolga':('Yo‘lga chiqdi','yolga_chiqdi'),'yetib':('Yetib keldi','yetib_keldi'),'yetkazildi':('Yetkazib berildi','yetkazildi')}
    if action not in mapping:
        return redirect(url_for('driver_delivery_detail',yid=yid))
    status,col=mapping[action]
    c=get_db()
    c.execute(f"UPDATE yetkazishlar SET holat=?, {col}=? WHERE id=? AND haydovchi_id=?",(status,now,yid,did))
    c.commit(); c.close()
    return redirect(url_for('driver_delivery_detail',yid=yid))

@app.route('/shofyor-boshqaruv', methods=['GET','POST'])
def driver_admin():
    c=get_db()
    msg=''
    if request.method=='POST':
        action=request.form.get('action')
        try:
            if action=='account':
                wid=int(request.form['ishchi_id'])
                login=request.form.get('login','').strip()
                password=request.form.get('password','')
                if _weak_password(password):
                    msg='Xato: parol kamida 8 belgidan iborat bo‘lib, harf va raqamni birga o‘z ichiga olishi kerak.'
                else:
                    existing=c.execute("SELECT id FROM shofyor_akkauntlari WHERE ishchi_id=?",(wid,)).fetchone()
                    if existing:
                        c.execute("UPDATE shofyor_akkauntlari SET login=?,parol_hash=?,faol=1,admin_tasdiq=1 WHERE ishchi_id=?",(login,generate_password_hash(password),wid))
                    else:
                        c.execute("INSERT INTO shofyor_akkauntlari(ishchi_id,login,parol_hash,faol,admin_tasdiq) VALUES(?,?,?,1,1)",(wid,login,generate_password_hash(password)))
                    msg='Shofyor akkaunti saqlandi.'
                    log_action('driver_account_set', f'ishchi_id={wid},login={login}')
            elif action=='approve_driver':
                aid=int(request.form['account_id'])
                c.execute("UPDATE shofyor_akkauntlari SET admin_tasdiq=1,faol=1 WHERE id=?",(aid,))
                log_action('driver_approved', f'account_id={aid}')
                msg='Shofyor tasdiqlandi.'
            elif action=='block_driver':
                aid=int(request.form['account_id'])
                c.execute("UPDATE shofyor_akkauntlari SET faol=0 WHERE id=?",(aid,))
                log_action('driver_blocked', f'account_id={aid}')
                msg='Shofyor bloklandi.'
            elif action=='delivery':
                c.execute("INSERT INTO yetkazishlar(buyurtma_id,haydovchi_id,sana,qadoq_soni,navbat,izoh) VALUES(?,?,?,?,?,?)",(int(request.form['buyurtma_id']),int(request.form['shofyor_id']),date.today().isoformat(),int(request.form.get('qadoq_soni') or 1),int(request.form.get('navbat') or 1),request.form.get('izoh','')))
                log_action('delivery_assigned', f"buyurtma_id={request.form['buyurtma_id']},shofyor_id={request.form['shofyor_id']}")
                msg='Yetkazish shofyorga biriktirildi.'
            c.commit()
        except Exception as e:
            c.rollback()
            msg='Xato: '+str(e)
    drivers=c.execute("SELECT id,ism,familiya,lavozim FROM ishchilar WHERE faol=1 AND (lower(lavozim) LIKE '%shof%' OR lower(lavozim) LIKE '%haydov%') ORDER BY ism").fetchall()
    orders=c.execute("SELECT id,kod,mijoz,mahsulot FROM buyurtmalar WHERE holat!='Yetkazildi' ORDER BY id DESC").fetchall()
    assigned=c.execute("SELECT y.*,b.kod,b.mahsulot,i.ism,i.familiya FROM yetkazishlar y JOIN buyurtmalar b ON b.id=y.buyurtma_id JOIN ishchilar i ON i.id=y.haydovchi_id ORDER BY y.id DESC LIMIT 100").fetchall()
    driver_accounts=c.execute("SELECT a.*,i.ism,i.familiya FROM shofyor_akkauntlari a JOIN ishchilar i ON i.id=a.ishchi_id ORDER BY a.id DESC").fetchall()
    c.close()
    return render_template_string(DRIVER_ADMIN_HTML,drivers=drivers,orders=orders,assigned=assigned,driver_accounts=driver_accounts,msg=msg)





@app.route("/pro-boshqaruv")
def pro_admin_page():
    c=get_db()
    stats={
        "orders":c.execute("SELECT COUNT(*) FROM buyurtmalar").fetchone()[0],
        "services":c.execute("SELECT COUNT(*) FROM servis_murojaatlari WHERE holat!='Yopildi'").fetchone()[0],
        "ratings":c.execute("SELECT COUNT(*) FROM baholar").fetchone()[0],
        "messages":c.execute("SELECT COUNT(*) FROM tizim_xabarlari WHERE oqildi=0").fetchone()[0],
        "reservations":c.execute("SELECT COUNT(*) FROM ombor_rezervlari WHERE holat='Rezerv'").fetchone()[0]
    }
    orders=c.execute("SELECT id,kod,mijoz,mahsulot,holat,umumiy_narx,oldindan_tolov FROM buyurtmalar ORDER BY id DESC LIMIT 30").fetchall()
    c.close()
    return render_template_string(PRO_ADMIN_HTML,stats=stats,orders=orders)


LOGIN_HTML = r"""
<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mebel360 - Kirish</title><style>body{margin:0;background:linear-gradient(135deg,#0f1b33,#2563eb);font-family:Arial;display:grid;place-items:center;min-height:100vh}.box{background:#fff;padding:28px;border-radius:18px;width:min(360px,92%);box-shadow:0 20px 50px #0005}h2{margin-top:0}input{width:100%;padding:12px;margin:8px 0;border:1px solid #cbd5e1;border-radius:9px;box-sizing:border-box}button{width:100%;padding:12px;border:0;border-radius:9px;background:#2563eb;color:#fff;font-weight:700}.err{color:#b91c1c;font-size:13px}</style></head><body><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h2>🏭 Mebel360</h2><input name="user" placeholder="Login" value="admin"><input name="password" type="password" placeholder="Parol"><button>Kirish</button><div class="err">{{error}}</div><p style="font-size:12px;color:#64748b">Parolni bilmasangiz, papkadagi <b>admin_parol.txt</b> faylini oching yoki konsol oynasidagi (CMD) chiqishga qarang.</p><hr><p style="text-align:center"><a href="/ishchi/login">👷 Ishchi kirishi</a> · <a href="/ishchi/royxat">Ro‘yxatdan o‘tish</a><br><a href="/shofyor/login">🚚 Shofyor kirishi</a> · <a href="/shofyor/royxat">Ro‘yxatdan o‘tish</a></p></form></body></html>
"""

WORKER_BASE_STYLE = """
<style>
*{box-sizing:border-box}body{margin:0;font-family:Arial;background:#eef3f8;color:#182235}.head{background:linear-gradient(135deg,#0f1b33,#2563eb);color:white;padding:18px}.wrap{max-width:1050px;margin:auto;padding:16px}.box,.card{background:white;border-radius:16px;padding:18px;box-shadow:0 8px 24px #0f172a18;margin-bottom:14px}input,select,textarea{width:100%;padding:11px;border:1px solid #cbd5e1;border-radius:9px;margin:5px 0 10px}button,.btn{display:inline-block;border:0;border-radius:9px;padding:10px 14px;background:#2563eb;color:white;font-weight:700;text-decoration:none;cursor:pointer}.green{background:#16a34a}.red{background:#dc2626}.muted{color:#64748b;font-size:13px}.err{color:#b91c1c}.ok{color:#166534}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.stat b{font-size:25px;color:#2563eb}.task{border-left:5px solid #2563eb}.bar{height:10px;background:#e2e8f0;border-radius:20px;overflow:hidden}.bar i{display:block;height:100%;background:#16a34a}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left}@media(max-width:700px){.grid{grid-template-columns:1fr}.wrap{padding:9px}}
</style>
"""

DRIVER_REGISTER_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor ro‘yxatdan o‘tishi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🚚 Shofyor ro‘yxatdan o‘tishi</b></div><div class="wrap"><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h2>Yangi akkaunt</h2><label>Ism</label><input name="ism" required><label>Familiya</label><input name="familiya"><label>Telefon</label><input name="telefon" placeholder="+998901234567" required><label>Login</label><input name="login" required><label>Parol</label><input name="password" type="password" minlength="8" required><button>Ro‘yxatdan o‘tish</button><p class="ok">{{msg}}</p><p class="err">{{error}}</p><p><a href="/shofyor/login">Akkauntim bor — kirish</a></p></form></div></body></html>"""

DRIVER_LOGIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor kirishi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🚚 Mebel360 — Shofyor kabineti</b></div><div class="wrap"><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h2>Kirish</h2><input name="login" placeholder="Login" required><input name="password" type="password" placeholder="Parol" required><button>Kirish</button><p class="err">{{error}}</p><p><a href="/shofyor/royxat">Yangi shofyor — ro‘yxatdan o‘tish</a></p><p><a href="/login">Rahbar kirishi</a></p></form></div></body></html>"""

DRIVER_DASHBOARD_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor kabineti</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><div class="wrap" style="padding:0"><b>🚚 {{driver['ism']}} {{driver['familiya']}}</b><a class="btn red" style="float:right" href="/shofyor/logout">Chiqish</a></div></div><div class="wrap"><h2>Menga biriktirilgan yuklar</h2>{% for x in deliveries %}<a href="/shofyor/yetkazish/{{x['id']}}" style="text-decoration:none;color:inherit"><div class="card task"><div style="display:flex;justify-content:space-between;gap:10px"><div><b style="font-size:21px">{{x['kod']}}</b><p>{{x['mahsulot']}}</p><p class="muted">{{x['manzil']}} · {{x['qavat'] or '-'}}-qavat · Lift: {{x['lift'] or '-'}}</p></div><div><span class="btn {% if x['holat']=='Yetkazib berildi' %}green{% endif %}">{{x['holat']}}</span></div></div></div></a>{% else %}<div class="card">Hozircha yuk biriktirilmagan.</div>{% endfor %}</div></body></html>"""

DRIVER_DETAIL_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{x['kod']}}</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><div class="wrap" style="padding:0"><b>📦 {{x['kod']}}</b><a class="btn" style="float:right" href="/shofyor/kabinet">Orqaga</a></div></div><div class="wrap"><div class="card"><h2>{{x['mahsulot']}}</h2><p><b>Mijoz:</b> {{x['mijoz']}}</p><p><b>Telefon:</b> <a href="tel:{{x['telefon']}}">{{x['telefon']}}</a></p><p><b>Manzil:</b> {{x['manzil']}}</p><p><b>Mo‘ljal:</b> {{x['moljal'] or '-'}}</p><p><b>Qavat:</b> {{x['qavat'] or '-'}}</p><p><b>Lift:</b> {{x['lift'] or '-'}}</p><p><b>Katta mashina:</b> {{x['katta_mashina'] or '-'}}</p><p><b>Qadoqlar:</b> {{x['qadoq_soni']}} ta</p><p><b>Yetkazish navbati:</b> {{x['navbat']}}</p><p><b>Izoh:</b> {{x['izoh'] or x['buyurtma_izoh'] or '-'}}</p>{% if x['lokatsiya'] %}
<div class="card" style="background:#f8fafc">
<h3>📍 Xarita tanlang</h3>
<p class="muted">Lokatsiyani o‘zingiz xohlagan xaritada oching.</p>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
<a class="btn green" href="https://www.google.com/maps/search/?api=1&query={{x['lokatsiya']|urlencode}}" target="_blank">Google Maps</a>
<a class="btn" style="background:#ef4444" href="https://yandex.com/maps/?text={{x['lokatsiya']|urlencode}}" target="_blank">Yandex Maps</a>
</div></div>
{% endif %}</div><div class="card"><h3>Holat: {{x['holat']}}</h3><form method="post" action="/shofyor/yetkazish/{{x['id']}}/holat"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><button name="action" value="yolga">🚚 Yo‘lga chiqdim</button><button name="action" value="yetib" class="green">📍 Yetib keldim</button><button name="action" value="yetkazildi" style="background:#7c3aed">✅ Yetkazib berdim</button></form><p class="muted">Yo‘lga chiqdi: {{x['yolga_chiqdi'] or '-'}}</p><p class="muted">Yetib keldi: {{x['yetib_keldi'] or '-'}}</p><p class="muted">Yetkazildi: {{x['yetkazildi'] or '-'}}</p></div></div></body></html>"""

DRIVER_ADMIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor boshqaruvi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🚚 Shofyor boshqaruvi</b><a class="btn" style="float:right" href="/">ERP bosh sahifa</a></div><div class="wrap"><p class="ok">{{msg}}</p><div class="grid"><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h3>Shofyor login yaratish</h3><input type="hidden" name="action" value="account"><label>Shofyor</label><select name="ishchi_id" required>{% for d in drivers %}<option value="{{d['id']}}">{{d['ism']}} {{d['familiya']}}</option>{% endfor %}</select><label>Login</label><input name="login" required><label>Parol</label><input name="password" type="password" minlength="8" required><button>Akkaunt yaratish</button></form><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h3>Yuk biriktirish</h3><input type="hidden" name="action" value="delivery"><label>Buyurtma</label><select name="buyurtma_id" required>{% for o in orders %}<option value="{{o['id']}}">{{o['kod']}} — {{o['mahsulot']}}</option>{% endfor %}</select><label>Shofyor</label><select name="shofyor_id" required>{% for d in drivers %}<option value="{{d['id']}}">{{d['ism']}} {{d['familiya']}}</option>{% endfor %}</select><label>Qadoq soni</label><input type="number" name="qadoq_soni" value="1" min="1"><label>Yetkazish navbati</label><input type="number" name="navbat" value="1" min="1"><label>Izoh</label><textarea name="izoh"></textarea><button>Biriktirish</button></form></div><div class="box" style="overflow:auto"><h3>Shofyor akkauntlari</h3><table><tr><th>Shofyor</th><th>Telefon</th><th>Login</th><th>Holat</th><th>Amal</th></tr>{% for a in driver_accounts %}<tr><td>{{a['ism']}} {{a['familiya']}}</td><td>{{a['telefon'] or '-'}}</td><td>{{a['login']}}</td><td>{% if a['admin_tasdiq'] %}✅ Tasdiqlangan{% else %}⏳ Kutilmoqda{% endif %}</td><td>{% if not a['admin_tasdiq'] %}<form method="post" style="display:inline"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><input type="hidden" name="action" value="approve_driver"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="green">Tasdiqlash</button></form>{% endif %}<form method="post" style="display:inline"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><input type="hidden" name="action" value="block_driver"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="red">Bloklash</button></form></td></tr>{% endfor %}</table></div><div class="box" style="overflow:auto"><h3>Biriktirilgan yuklar</h3><table><tr><th>Buyurtma</th><th>Mahsulot</th><th>Shofyor</th><th>Qadoq</th><th>Holat</th></tr>{% for a in assigned %}<tr><td>{{a['kod']}}</td><td>{{a['mahsulot']}}</td><td>{{a['ism']}} {{a['familiya']}}</td><td>{{a['qadoq_soni']}}</td><td>{{a['holat']}}</td></tr>{% endfor %}</table></div></div></body></html>"""


WORKER_REGISTER_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi ro‘yxati</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>👷 Ishchi ro‘yxatdan o‘tishi</b></div><div class="wrap"><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h2>Telefon raqamingiz</h2><p class="muted">Masalan: +998 90 123 45 67</p><input name="telefon" required placeholder="+998901234567"><button>Kod olish</button><p class="ok">{{msg}}</p><p class="err">{{error}}</p>{% if demo_code %}<div class="card"><b>Sinov kodi: {{demo_code}}</b><p class="muted">Haqiqiy SMS xizmati ulanmaguncha shu koddan foydalaning.</p><a class="btn green" href="/ishchi/kod">Kodni kiritish</a></div>{% endif %}<p><a href="/ishchi/login">Akkauntim bor — kirish</a></p></form></div></body></html>"""

WORKER_VERIFY_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kodni tasdiqlash</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🔐 Telefonni tasdiqlash</b></div><div class="wrap"><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><p>{{telefon}}</p><label>Kod</label><input name="kod" inputmode="numeric" maxlength="6" required><label>Ism</label><input name="ism" required><label>Familiya</label><input name="familiya"><label>Yangi login</label><input name="login" required><label>Yangi parol</label><input name="password" type="password" minlength="8" required><button>Ro‘yxatdan o‘tish</button><p class="err">{{error}}</p></form></div></body></html>"""

WORKER_WAIT_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kutilmoqda</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="wrap"><div class="box"><h2>✅ Ro‘yxatdan o‘tdingiz</h2><p>Endi administrator akkauntingizni tasdiqlaydi. Tasdiqlangach login va parolingiz bilan kirasiz.</p><a class="btn" href="/ishchi/login">Kirish sahifasi</a></div></div></body></html>"""

WORKER_LOGIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi kirishi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🏭 Mebel360 — Ishchi kabineti</b></div><div class="wrap"><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h2>Kirish</h2><input name="login" placeholder="Login" required><input name="password" type="password" placeholder="Parol" required><button>Kirish</button><p class="err">{{error}}</p><p><a href="/ishchi/royxat">Yangi ro‘yxatdan o‘tish</a></p></form></div></body></html>"""

WORKER_DASHBOARD_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi kabineti</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><div class="wrap" style="padding:0"><b>👷 {{worker['ism']}} {{worker['familiya']}}</b> — {{worker['lavozim']}} <a class="btn red" style="float:right" href="/ishchi/logout">Chiqish</a></div></div><div class="wrap">
{% with messages = get_flashed_messages() %}{% if messages %}<div class="card err">{{messages[0]}}</div>{% endif %}{% endwith %}
<div class="grid"><div class="card stat"><span>HOZIRGI HOLAT</span><br><b style="color:{% if worker_state=='Ishlayapti' %}#16a34a{% else %}#64748b{% endif %}">{{worker_state}}</b></div><div class="card stat"><span>ISHLAGAN KUN</span><br><b>{{stats['kun'] or 0}}</b></div><div class="card stat"><span>JAMI SOAT</span><br><b>{{'%.1f'|format(stats['soat'] or 0)}}</b></div></div>
{% if active_task %}<div class="card" style="border-left:7px solid #16a34a"><h3>🟢 Hozir ishlayotgan ishim</h3><p><b>{{active_task['ish_turi']}}</b> — {{active_task['buyurtma_kodi'] or 'Buyurtmasiz'}}</p><p>{{active_task['tavsif']}}</p><p>Boshlangan vaqt: <b>{{active_task['boshlandi_vaqt']}}</b></p></div>{% endif %}
<h2>Topshiriqlarim</h2>{% for t in tasks %}<div class="card task"><b>{{t['ish_turi']}}</b> {% if t['buyurtma_kodi'] %}<span class="muted">— {{t['buyurtma_kodi']}}</span>{% endif %}<p>{{t['tavsif']}}</p><div class="bar"><i style="width:{{t['progress']}}%"></i></div><p><b>{{t['progress']}}%</b> · {{t['holat']}} · Reja: {{t['sana']}}{% if t['tugash_sana'] %} — {{t['tugash_sana']}}{% endif %}</p><p class="muted">Boshladi: {{t['boshlandi_vaqt'] or '-'}} · Tugatdi: {{t['tugadi_vaqt'] or '-'}}</p>
{% if t['holat']=='Yangi' or t['holat']=='Jarayonda' %}<form method="post" action="/ishchi/topshiriq/{{t['id']}}/boshlash"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><button class="green">▶ Ishni boshladim</button></form>{% elif t['holat']=='Ishlayapti' %}<form method="post" action="/ishchi/topshiriq/{{t['id']}}/tugatish"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><button style="background:#7c3aed">✅ Ishni tugatdim</button></form>{% endif %}
{% if t['holat']!='Tayyor' %}<form method="post" action="/ishchi/topshiriq/{{t['id']}}"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><input type="hidden" name="holat" value="{{t['holat']}}"><label>Jarayon foizi</label><input type="number" name="progress" min="0" max="100" value="{{t['progress']}}"><button>Foizni yangilash</button></form>{% endif %}</div>{% else %}<div class="card">Hozircha topshiriq yo‘q. Holatingiz: <b>Bo‘sh</b>.</div>{% endfor %}</div></body></html>"""

WORKER_ADMIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi boshqaruvi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>👥 Ishchi kabinetlarini boshqarish</b> <a class="btn" style="float:right" href="/">ERP bosh sahifa</a></div><div class="wrap">
<div class="box" style="overflow:auto"><h3>Ishchilarning hozirgi holati</h3><table><tr><th>Ishchi</th><th>Holat</th><th>Buyurtma</th><th>Ish turi</th><th>Boshlagan vaqt</th></tr>{% for w in worker_states %}<tr><td>{{w['ism']}} {{w['familiya']}}</td><td>{% if w['ish_holati']=='Ishlayapti' %}<b style="color:#16a34a">🟢 Ishlayapti</b>{% else %}<b style="color:#64748b">⚪ Bo‘sh</b>{% endif %}</td><td>{{w['buyurtma_kodi'] or '-'}}</td><td>{{w['ish_turi'] or '-'}}</td><td>{{w['boshlandi_vaqt'] or '-'}}</td></tr>{% endfor %}</table></div>
<div class="grid"><form class="box" method="post"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><h3>Yangi topshiriq</h3><input type="hidden" name="action" value="task"><label>Ishchi</label><select name="ishchi_id" required>{% for w in workers %}<option value="{{w['id']}}">{{w['ism']}} {{w['familiya']}} — {{w['lavozim']}}</option>{% endfor %}</select><label>Buyurtma kodi</label><input name="buyurtma_kodi" placeholder="AB 007"><label>Ish turi</label><input name="ish_turi" required placeholder="Kesish / Rover / Yig‘ish"><label>Topshiriq</label><textarea name="tavsif" placeholder="Nima ish qilishi kerakligini batafsil yozing"></textarea><label>Boshlanish rejasi</label><input type="date" name="sana" value="{{today}}"><label>Tugash rejasi</label><input type="date" name="tugash_sana"><button>Topshiriq berish</button></form><div class="box" style="grid-column:span 2;overflow:auto"><h3>Ro‘yxatdan o‘tganlar</h3><table><tr><th>Ishchi</th><th>Telefon</th><th>Login</th><th>Holat</th><th>Amal</th></tr>{% for a in accounts %}<tr><td>{{a['ism']}} {{a['familiya']}}</td><td>{{a['telefon']}}</td><td>{{a['login'] or ''}}</td><td>{% if a['admin_tasdiq'] %}✅ Tasdiqlangan{% else %}⏳ Kutilmoqda{% endif %}</td><td>{% if not a['admin_tasdiq'] %}<form method="post" style="display:inline"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><input type="hidden" name="action" value="approve"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="green">Tasdiqlash</button></form>{% endif %}<form method="post" style="display:inline"><input type=\"hidden\" name=\"csrf_token\" value=\"{{csrf_token()}}\"><input type="hidden" name="action" value="block"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="red">Bloklash</button></form></td></tr>{% endfor %}</table></div></div>
<div class="box" style="overflow:auto"><h3>Topshiriqlar tarixi</h3><table><tr><th>Ishchi</th><th>Buyurtma</th><th>Ish</th><th>Holat</th><th>Progress</th><th>Boshladi</th><th>Tugatdi</th></tr>{% for t in tasks %}<tr><td>{{t['ism']}} {{t['familiya']}}</td><td>{{t['buyurtma_kodi']}}</td><td>{{t['ish_turi']}}</td><td>{{t['holat']}}</td><td>{{t['progress']}}%</td><td>{{t['boshlandi_vaqt'] or '-'}}</td><td>{{t['tugadi_vaqt'] or '-'}}</td></tr>{% endfor %}</table></div></div></body></html>"""


HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mebel360 V5.1.4</title>
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
<header><div class="top"><div><h1>🏭 Mebel360 V5.1.4</h1><div class="sub">Ishchilar, buyurtmalar, ombor, ishlab chiqarish va moliya</div></div><div><a href="/pro-boshqaruv"><button style="background:#0f766e">PRO boshqaruv</button></a> <a href="/shofyor-boshqaruv"><button style="background:#7c3aed">Shofyor boshqaruvi</button></a> <a href="/shartnoma-namuna"><button style="background:#0f766e">Shartnoma Word</button></a> <a href="/shartnoma-pdf" target="_blank"><button style="background:#0369a1">Shartnoma PDF</button></a> <a href="/backup"><button style="background:#16a34a">Backup</button></a> <a href="/logout"><button style="background:#dc2626">Chiqish</button></a></div></div></header>
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
 <button data-tab="customerPro">Mijoz PRO</button>
 <button data-tab="service">Servis/Kafolat</button>
 <button data-tab="delivery">Yetkazish</button>
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
<hr><h3>Maxfiy hujjatlar</h3>
<p style="font-size:12px;color:#64748b">Bu ma’lumotlarni faqat rahbar ko‘radi.</p>
<label>Pasport/ID seriya-raqami<input name="pasport"></label>
<label>JSHSHIR<input name="jshshir"></label>
<label>Tug‘ilgan sana<input type="date" name="tugilgan_sana"></label>
<label>Yashash manzili<textarea name="yashash_manzil"></textarea></label>
<label>Pasport berilgan sana<input type="date" name="pasport_berilgan_sana"></label>
<label>Pasportni bergan tashkilot<input name="pasport_bergan"></label>
<label>Favqulodda aloqa telefoni<input name="favqulodda_telefon"></label>
<label>Izoh<textarea name="izoh"></textarea></label><button>Saqlash</button><div class="msg"></div>
</form></div><div class="panel tablewrap"><h3>Ishchilar</h3><table><thead><tr><th>Ism</th><th>Lavozim</th><th>Staj</th><th>Kunlik</th><th>Oylik</th><th></th></tr></thead><tbody id="workersBody"></tbody></table></div>
</div></section>

<section id="attendance" class="tab"><div class="grid"><div class="panel"><h3>Avtomatik keldi-ketdi</h3>
<label>Ishchi<select class="workerSelect" id="autoWorkerSelect" required></select></label>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
<button type="button" class="ok" onclick="clockIn()">Hozir keldi</button>
<button type="button" style="background:#dc2626" onclick="clockOut()">Hozir ketdi</button>
</div><div id="attendanceAutoMsg" class="msg"></div>
<hr><h3>Qo‘lda kiritish</h3><form id="attendanceForm">
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
<label>Manzil<input name="manzil"></label>
<label>Pasport/ID<input name="pasport_id"></label>
<label>Mahsulot<input name="mahsulot"></label>
<label>O‘lcham<input name="olcham" placeholder="2000x600x2400 mm"></label>
<label>Soni<input type="number" name="soni" value="1" min="1"></label>
<label>Material<input name="material" placeholder="MDF / LDSP / Akril"></label>
<label>Rang<input name="rang" placeholder="Rang yoki kod"></label>
<label>Umumiy narx<input type="number" name="umumiy_narx" value="0"></label>
<label>Avans<input type="number" name="oldindan_tolov" value="0"></label>
<label>Oraliq to‘lov<input type="number" name="oraliq_tolov" value="0"></label>
<label>To‘lov usuli<select name="tolov_usuli"><option>Naqd</option><option>Click</option><option>Payme</option><option>Bank orqali</option><option>Plastik karta</option><option>Qarz</option></select></label>
<label>Yetkazish<select name="yetkazish"><option>Kiritilgan</option><option>Kiritilmagan</option><option>Alohida haq</option></select></label>
<label>Montaj<select name="montaj"><option>Kiritilgan</option><option>Kiritilmagan</option><option>Alohida haq</option></select></label>
<label>Kafolat muddati<input name="kafolat_muddati" value="12 oy"></label>
<label>Boshlanish sana<input type="date" name="boshlanish_sana"></label>
<label>Tugash sana<input type="date" name="tugash_sana"></label>
<label>Taxminiy tayyor sana<input type="date" name="taxminiy_sana"></label>
<label>Mas’ul xodim<input name="masul_xodim"></label>
<label>Holat<select name="holat"><option>Yangi</option><option>Jarayonda</option><option>Tayyor</option><option>Yetkazildi</option></select></label>
<label>Kechikish sababi<select name="kechikish_turi"><option value="">Yo‘q</option><option>Korxona</option><option>Mijoz</option><option>Material</option><option>Favqulodda holat</option></select></label>
<label>Kechikish chegirmasi (%/kun)<input type="number" step="0.1" name="kechikish_foiz" value="0"></label>
<label>Maksimal chegirma %<input type="number" step="0.1" name="maks_chegirma_foiz" value="20"></label>
<label>Keshbek %<input type="number" step="0.1" name="keshbek_foiz" value="0"></label>
<label>Keshbek summasi<input type="number" name="keshbek_summa" value="0"></label>
<label>Kafolat boshlanish<input type="date" name="kafolat_boshlanish"></label>
<label>Kafolat tugash<input type="date" name="kafolat_tugash"></label>
<label>Kafolat sharti<textarea name="kafolat_sharti"></textarea></label>
<label>Lokatsiya havolasi<input name="lokatsiya" placeholder="Google Maps havolasi"></label>
<label>Mo‘ljal<input name="moljal"></label>
<label>Qavat<input name="qavat"></label>
<label>Lift<select name="lift"><option></option><option>Bor</option><option>Yo‘q</option></select></label>
<label>Katta mashina<select name="katta_mashina"><option></option><option>Kira oladi</option><option>Kira olmaydi</option></select></label>
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

<section id="expenses" class="tab">
<div class="grid">
<div class="panel"><h3>Yangi xarajat</h3><form id="expenseForm">
<label>Sana<input type="date" name="sana" required></label>
<label>Kategoriya<select name="kategoriya">
<option>Material</option><option>Furnitura</option><option>Ishlab chiqarish</option>
<option>Ishchi</option><option>Transport</option><option>Sex</option>
<option>Ofis va reklama</option><option>Mijoz va buyurtma</option>
<option>Uy</option><option>Akam</option><option>Ota uchun</option><option>Ona uchun</option>
<option>Favqulodda</option><option>Boshqa</option>
</select></label>
<label>Xarajat nomi<select name="xarajat_nomi">
<option>MDF</option><option>DSP</option><option>LDSP</option><option>HDF</option><option>DVP</option>
<option>Akril</option><option>Fanera</option><option>Oyna</option><option>Shisha</option>
<option>Kromka</option><option>Profil</option><option>Bo‘yoq</option><option>Lak</option>
<option>Grunt</option><option>Shpaklyovka</option><option>Yelim</option><option>Silikon</option>
<option>Porolon</option><option>Mato</option><option>Ekokoja</option><option>Qadoqlash materiali</option>
<option>Petlya</option><option>Ruchka</option><option>Napravlyayushiy</option><option>Gazlift</option>
<option>Oyoq</option><option>G‘ildirak</option><option>Zamok</option><option>Magnit</option>
<option>Samorez</option><option>Bolt va gayka</option><option>Tortma mexanizmi</option>
<option>Ko‘tarma mexanizm</option><option>LED lenta</option><option>Blok pitaniya</option>
<option>Sim va elektr jihozlari</option><option>Raspil</option><option>Kromka urish</option>
<option>CNC Rover</option><option>Frezalash</option><option>Lazer</option><option>Teshish</option>
<option>Bo‘yash</option><option>Sayqalash</option><option>Oyna teshish</option>
<option>Oyna kesish</option><option>Payvandlash</option><option>Tashqi ustaga berilgan ish</option>
<option>Qayta ishlash xarajati</option><option>Yaroqsiz mahsulot</option><option>Material chiqindisi</option>
<option>Kunlik ish haqi</option><option>Oylik maosh</option><option>Soatbay haq</option>
<option>Ish hajmiga haq</option><option>Avans</option><option>Bonus</option>
<option>Qo‘shimcha ish haqi</option><option>Ortib ishlagan vaqt</option>
<option>Ovqat puli</option><option>Yo‘l puli</option><option>Telefon puli</option>
<option>Xizmat safari</option><option>Ishchi kiyimi</option><option>Himoya vositalari</option>
<option>Benzin</option><option>Gaz</option><option>Dizel</option><option>Moy almashtirish</option>
<option>Mashina ta’miri</option><option>Ehtiyot qismlar</option><option>Shina</option>
<option>Yuvish</option><option>Jarima</option><option>Yo‘l haqi</option><option>Parking</option>
<option>Yuk tashish</option><option>Yetkazib berish</option><option>Bozorga borish</option>
<option>Raspildan olib kelish</option><option>Mijoz uyiga borish</option>
<option>Elektr</option><option>Gaz</option><option>Suv</option><option>Ijara</option>
<option>Internet</option><option>Telefon</option><option>Qo‘riqlash</option><option>Tozalash</option>
<option>Chiqindi olib ketish</option><option>Stanok ta’miri</option><option>Asbob-uskunalar</option>
<option>Freza</option><option>Disk</option><option>Sverlo</option><option>Zımpara</option>
<option>Kompressor xarajati</option><option>Moylash vositalari</option><option>Texnika xavfsizligi</option>
<option>Kantselyariya</option><option>Printer qog‘ozi</option><option>Kartrij</option>
<option>Kompyuter ta’miri</option><option>Dastur xarajati</option><option>Hosting</option>
<option>Domen</option><option>SMS xizmati</option><option>Reklama</option>
<option>Instagram reklama</option><option>Telegram reklama</option><option>OLX reklama</option>
<option>Foto-video</option><option>Dizayn</option><option>Bank xizmati</option>
<option>Soliq</option><option>Buxgalteriya</option><option>Razmer olish</option>
<option>Chizma tayyorlash</option><option>Namuna tayyorlash</option><option>Montaj</option>
<option>Qavatga ko‘tarish</option><option>Servis</option><option>Kafolat ta’miri</option>
<option>Mijozga chegirma</option><option>Keshbek</option><option>Qaytarilgan pul</option>
<option>Kechikish kompensatsiyasi</option><option>Qo‘shimcha ish</option>
<option>Uy uchun</option><option>Akam</option><option>Ota uchun</option><option>Ona uchun</option>
<option>Sex uchun</option><option>Qarzdorlik yopish</option><option>Favqulodda xarajat</option>
<option>Boshqa xarajat</option>
</select></label>
<label>Summa<input type="number" name="miqdor" required></label>
<label>To‘lov usuli<select name="tolov_usuli">
<option>Naqd</option><option>Click</option><option>Payme</option>
<option>Bank orqali</option><option>Plastik karta</option><option>Qarz</option>
</select></label>
<label>Buyurtma kodi<input name="buyurtma_kodi" placeholder="AB 007"></label>
<label>Kimga berildi<input name="kimga_berildi"></label>
<label>Chek rasmi yoki havolasi<input name="chek_havola"></label>
<label>Izoh<input name="tavsifi"></label>
<button>Saqlash</button><div class="msg"></div>
</form></div>
<div class="panel tablewrap"><table><thead><tr>
<th>Sana</th><th>Kategoriya</th><th>Xarajat nomi</th><th>Summa</th>
<th>To‘lov usuli</th><th>Buyurtma</th><th>Kimga</th><th>Izoh</th>
</tr></thead><tbody id="expensesBody"></tbody></table></div>
</div></section>

<section id="bonuses" class="tab"><div class="grid"><div class="panel"><h3>Bonus</h3><form id="bonusForm"><label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label><label>Summa<input type="number" name="miqdor" required></label><label>Sabab<input name="sababi"></label><button>Bonus saqlash</button><div class="msg"></div></form><hr><h3>Ta’til / holat</h3><form id="statusForm"><label>Ishchi<select class="workerSelect" name="ishchi_id" required></select></label><label>Sana<input type="date" name="sana" required></label><label>Turi<select name="turi"><option>Dam olish</option><option>Ta’til</option><option>Kasallik</option><option>Sababsiz</option></select></label><label>Izoh<input name="izoh"></label><button>Saqlash</button><div class="msg"></div></form></div>
<div class="panel tablewrap"><h3>Bonuslar</h3><table><thead><tr><th>Ishchi</th><th>Sana</th><th>Summa</th><th>Sabab</th></tr></thead><tbody id="bonusesBody"></tbody></table><h3 style="margin-top:20px">Ta’til va holatlar</h3><table><thead><tr><th>Ishchi</th><th>Sana</th><th>Turi</th><th>Izoh</th></tr></thead><tbody id="statusesBody"></tbody></table></div></div></section>

<section id="finished" class="tab"><div class="grid"><div class="panel"><h3>Tayyor mahsulot</h3><form id="finishedForm"><label>Nomi<input name="nomi" required></label><label>Kodi<input name="kodi"></label><label>Rang<input name="rang"></label><label>Miqdor<input type="number" step="0.1" name="miqdor" required></label><label>Birlik<select name="birlik"><option>dona</option><option>komplekt</option></select></label><label>Narx<input type="number" name="narx"></label><label>Izoh<input name="izoh"></label><button>Saqlash</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Nomi</th><th>Kod</th><th>Rang</th><th>Miqdor</th><th>Narx</th></tr></thead><tbody id="finishedBody"></tbody></table></div></div></section>

<section id="finance" class="tab"><div class="panel"><h3>Moliyaviy xulosa</h3><div style="display:flex;gap:8px;flex-wrap:wrap;align-items:end"><label>Boshlanish<input id="finStart" type="date"></label><label>Tugash<input id="finEnd" type="date"></label><button onclick="loadFinance()">Hisoblash</button></div><div class="cards" style="margin-top:15px"><div class="card"><span>KIRIM</span><b id="fIncome">0</b></div><div class="card"><span>XARAJAT</span><b id="fExpense">0</b></div><div class="card"><span>ISHCHI TO‘LOV</span><b id="fSalary">0</b></div><div class="card"><span>BONUS</span><b id="fBonus">0</b></div><div class="card"><span>SOF FOYDA</span><b id="fProfit">0</b></div></div></div></section>

<section id="customerPro" class="tab"><div class="grid"><div class="panel"><h3>Mijoz bahosi</h3><form id="ratingForm"><label>Buyurtma<select class="orderSelect" name="buyurtma_id" required></select></label><label>Mebel sifati (1-5)<input type="number" min="1" max="5" name="sifat" value="5"></label><label>Muddat (1-5)<input type="number" min="1" max="5" name="muddat" value="5"></label><label>Muomala (1-5)<input type="number" min="1" max="5" name="muomala" value="5"></label><label>Yetkazish (1-5)<input type="number" min="1" max="5" name="yetkazish" value="5"></label><label>Montaj (1-5)<input type="number" min="1" max="5" name="montaj" value="5"></label><label>Izoh<textarea name="izoh"></textarea></label><button>Bahoni saqlash</button><div class="msg"></div></form><hr><h3>Qo‘shimcha ish</h3><form id="extraForm"><label>Buyurtma<select class="orderSelect" name="buyurtma_id" required></select></label><label>Ish nomi<input name="nomi" required></label><label>Qo‘shimcha summa<input type="number" name="summa"></label><label>Qo‘shimcha kun<input type="number" name="qoshimcha_kun"></label><label>Izoh<input name="izoh"></label><button>Qo‘shish</button><div class="msg"></div></form></div><div class="panel tablewrap"><h3>Qo‘shimcha ishlar</h3><table><thead><tr><th>Buyurtma</th><th>Mijoz</th><th>Ish</th><th>Summa</th><th>Kun</th></tr></thead><tbody id="extrasBody"></tbody></table></div></div></section>
<section id="service" class="tab"><div class="grid"><div class="panel"><h3>Servis / kafolat murojaati</h3><form id="serviceForm"><label>Buyurtma<select class="orderSelect" name="buyurtma_id" required></select></label><label>Turi<select name="turi"><option>Kafolat</option><option>Servis</option><option>Pullik xizmat</option></select></label><label>Muammo<textarea name="muammo" required></textarea></label><label>Servis sanasi<input type="date" name="servis_sana"></label><label>Usta<input name="usta"></label><button>Qabul qilish</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Buyurtma</th><th>Mijoz</th><th>Turi</th><th>Muammo</th><th>Holat</th><th>Sana</th><th>Usta</th></tr></thead><tbody id="serviceBody"></tbody></table></div></div></section>
<section id="delivery" class="tab"><div class="grid"><div class="panel"><h3>Yetkazishni rejalash</h3><form id="deliveryForm"><label>Buyurtma<select class="orderSelect" name="buyurtma_id" required></select></label><label>Haydovchi<select class="workerSelect" name="haydovchi_id"></select></label><label>Sana<input type="date" name="sana" required></label><label>Navbat<input type="number" name="navbat" value="1"></label><label>Mashina<input name="mashina"></label><label>Benzin<input type="number" name="benzin"></label><label>Yo‘l xarajati<input type="number" name="yol_xarajati"></label><label>Izoh<input name="izoh"></label><button>Rejalash</button><div class="msg"></div></form></div><div class="panel tablewrap"><table><thead><tr><th>Sana</th><th>Navbat</th><th>Buyurtma</th><th>Mijoz</th><th>Haydovchi</th><th>Manzil</th><th>Holat</th><th>Amal</th></tr></thead><tbody id="deliveryBody"></tbody></table></div></div></section>

</main>

<div id="stageModal" style="display:none;position:fixed;inset:0;background:#0009;z-index:20;place-items:center">
<div class="panel" style="width:min(420px,92%);max-height:85vh;overflow:auto"><h3>Buyurtma bosqichlari</h3><div id="stageList"></div><button onclick="closeStage()" style="margin-top:10px">Yopish</button></div>
</div>

<script>
const $=s=>document.querySelector(s),$$=s=>document.querySelectorAll(s);
const today=new Date().toISOString().slice(0,10);$$('input[type=date]').forEach(x=>x.value=today);
$$('.tabs button').forEach(b=>b.onclick=()=>{$$('.tabs button').forEach(x=>x.classList.remove('active'));$$('.tab').forEach(x=>x.classList.remove('active'));b.classList.add('active');$('#'+b.dataset.tab).classList.add('active')});
function esc(s){return String(s===undefined||s===null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
async function api(u,o){const r=await fetch(u,o),j=await r.json();if(!r.ok)throw new Error(j.message||'Xato');return j}
function fj(f){return Object.fromEntries(new FormData(f).entries())}
function bind(id,url){const f=$(id);f.onsubmit=async e=>{e.preventDefault();const m=f.querySelector('.msg');try{await api(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(fj(f))});m.textContent='✅ Saqlandi';await refresh()}catch(x){m.textContent='❌ '+x.message}}}
function money(v){return Number(v||0).toLocaleString('uz-UZ',{maximumFractionDigits:2})}

async function loadWorkers(){const a=await api('/api/ishchilar');$('#workersBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.lavozim||'')}</td><td>${esc(x.staj_yil)} yil</td><td>${money(x.kunlik_stavka)}</td><td>${money(x.oylik_maosh)}</td><td><button class="danger" onclick="delWorker(${esc(x.id)})">O‘chirish</button></td></tr>`).join('');$$('.workerSelect').forEach(s=>s.innerHTML='<option value="">Tanlang</option>'+a.map(x=>`<option value="${esc(x.id)}">${esc(x.ism)} ${esc(x.familiya||'')}</option>`).join(''))}
async function delWorker(i){if(confirm('O‘chirasizmi?')){await api('/api/ishchilar/'+i,{method:'DELETE'});refresh()}}
async function loadTypes(){const a=await api('/api/ish-turlari');$('#workTypeSelect').innerHTML=a.map(x=>`<option value="${esc(x.id)}">${esc(x.kategoriya)} — ${esc(x.nomi)} (${esc(x.birlik)})</option>`).join('')}
async function loadAttendance(){const a=await api('/api/keldi-ketdi');$('#attendanceBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.sana)}</td><td>${esc(x.keldi_vaqti)}</td><td>${esc(x.ketdi_vaqti)}</td><td>${esc(x.ish_soatlari)}</td></tr>`).join('')}
async function loadResults(){const a=await api('/api/natijalar');$('#resultsBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.ish_turi)}</td><td>${esc(x.sana)}</td><td>${esc(x.miqdor)} ${esc(x.birlik)}</td><td>${money(x.jami_haq)}</td><td>${esc(x.buyurtma_kodi||'')}</td></tr>`).join('')}
async function loadOrders(){const a=await api('/api/buyurtmalar');$$('.orderSelect').forEach(s=>s.innerHTML='<option value="">Tanlang</option>'+a.map(x=>`<option value="${esc(x.id)}">${esc(x.kod)} — ${esc(x.mijoz)}</option>`).join(''));$('#ordersBody').innerHTML=a.map(x=>`<tr><td>${esc(x.kod)}</td><td>${esc(x.mijoz)}</td><td>${esc(x.mahsulot||'')}</td><td>${money(x.umumiy_narx)}</td><td>${money(x.oldindan_tolov)}</td><td class="balance">${money(x.qoldiq)}</td><td><span class="badge">${esc(x.holat)}</span></td><td><span id="pr${esc(x.id)}">0%</span></td><td><button onclick="openStage(${esc(x.id)})">Ko‘rish</button></td><td><button onclick="addOrderPayment(${esc(x.id)})">To‘lov</button></td><td><a href="/buyurtma/${esc(x.id)}/shartnoma.docx"><button>Word</button></a> <a href="/buyurtma/${esc(x.id)}/shartnoma.pdf" target="_blank"><button>PDF</button></a> <button onclick="regenContract(${esc(x.id)})">Yangilash</button> <a href="/buyurtma/${esc(x.id)}/chek.pdf" target="_blank"><button>Chek</button></a> <a href="/buyurtma/${esc(x.id)}/qr.png" target="_blank"><button class="ok">QR</button></a> <button onclick="copyTrack(${esc(x.id)})">Link</button></td></tr>`).join('')}
async function openStage(id){const a=await api('/api/buyurtma/'+id+'/bosqichlar');$('#stageList').innerHTML=a.map(x=>`<label class="stage"><input type="checkbox" ${x.bajarildi?'checked':''} onchange="toggleStage(${esc(x.id)},this.checked)"> ${esc(x.bosqich)}</label>`).join('');$('#stageModal').style.display='grid'}
function closeStage(){$('#stageModal').style.display='none'}
async function toggleStage(id,v){await api('/api/buyurtma-bosqich/'+id,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({bajarildi:v})})}
async function loadStock(){const a=await api('/api/ombor');$('#stockSelect').innerHTML=a.map(x=>`<option value="${esc(x.id)}">${esc(x.nomi)} (${esc(x.birlik)})</option>`).join('');$('#stockBody').innerHTML=a.map(x=>`<tr class="${Number(x.qoldiq)<=Number(x.min_qoldiq)?'low':''}"><td>${esc(x.nomi)}</td><td>${esc(x.kategoriya)}</td><td>${esc(x.qoldiq)}</td><td>${esc(x.birlik)}</td><td>${esc(x.min_qoldiq)}</td></tr>`).join('')}
async function loadTrips(){const a=await api('/api/safarlar');$('#tripsBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.sana)}</td><td>${esc(x.mashina||'')}</td><td>${esc(x.qayerdan||'')} → ${esc(x.qayerga||'')}</td><td>${esc(x.masofa_km)}</td><td>${esc(x.sabab||'')}</td><td>${money(x.xarajat)}</td></tr>`).join('')}
async function loadPayments(){const a=await api('/api/tolovlar');$('#paymentsBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.sana)}</td><td>${money(x.miqdor)}</td><td>${esc(x.turi)}</td></tr>`).join('')}
async function loadPenalties(){const a=await api('/api/jarimalar');$('#penaltiesBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.sana)}</td><td class="minus">${money(x.miqdor)}</td><td>${esc(x.sababi||'')}</td></tr>`).join('')}
async function loadDashboard(){const x=await api('/api/dashboard');$('#dWorkers').textContent=x.workers;$('#dOrders').textContent=x.orders;$('#dHours').textContent=x.hours;$('#dProduction').textContent=x.production;$('#dKm').textContent=x.km;$('#dLow').textContent=x.low_stock}
function setMonth(){const d=new Date(),y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0');$('#totalStart').value=`${y}-${m}-01`;$('#totalEnd').value=new Date(y,d.getMonth()+1,0).toISOString().slice(0,10);loadTotals()}
async function loadTotals(){const s=$('#totalStart').value||'1900-01-01',e=$('#totalEnd').value||'2999-12-31';$('#csvLink').href=`/export/jami.csv?start=${s}&end=${e}`;const a=await api(`/api/jami?start=${s}&end=${e}`);$('#totalsBody').innerHTML=a.map((x,i)=>`<tr><td>${i+1}</td><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.lavozim||'')}</td><td>${esc(x.ish_kunlari)}</td><td>${esc(x.jami_soat)}</td><td>${esc(x.jami_miqdor)}</td><td class="money">${money(x.ish_haqi)}</td><td class="minus">${money(x.jarima)}</td><td class="money">${money(x.bonus)}</td><td>${money(x.tolangan)}</td><td class="balance">${money(x.qoldiq)}</td><td>${esc(x.reyting)}</td></tr>`).join('')}

async function loadExpenses(){const a=await api('/api/xarajatlar');$('#expensesBody').innerHTML=a.map(x=>`<tr><td>${esc(x.sana)}</td><td>${esc(x.kategoriya)}</td><td>${esc(x.xarajat_nomi||'')}</td><td class="minus">${money(x.miqdor)}</td><td>${esc(x.tolov_usuli||'Naqd')}</td><td>${esc(x.buyurtma_kodi||'')}</td><td>${esc(x.kimga_berildi||'')}</td><td>${esc(x.tavsifi||'')}</td></tr>`).join('')}
async function clockIn(){const id=$('#autoWorkerSelect').value,m=$('#attendanceAutoMsg');if(!id){m.textContent='Ishchini tanlang';return}try{const x=await api('/api/davomat/keldi',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ishchi_id:id})});m.textContent=`✅ Keldi: ${esc(x.sana)} ${esc(x.vaqt)}`;loadAttendance();loadDashboard()}catch(e){m.textContent='❌ '+e.message}}
async function clockOut(){const id=$('#autoWorkerSelect').value,m=$('#attendanceAutoMsg');if(!id){m.textContent='Ishchini tanlang';return}try{const x=await api('/api/davomat/ketdi',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ishchi_id:id})});m.textContent=`✅ Ketdi: ${esc(x.sana)} ${esc(x.vaqt)} — ${esc(x.ish_soatlari)} soat`;loadAttendance();loadDashboard()}catch(e){m.textContent='❌ '+e.message}}
async function loadBonuses(){const a=await api('/api/bonuslar');$('#bonusesBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.sana)}</td><td class="money">${money(x.miqdor)}</td><td>${esc(x.sababi||'')}</td></tr>`).join('')}
async function loadStatuses(){const a=await api('/api/ishchi-holatlari');$('#statusesBody').innerHTML=a.map(x=>`<tr><td>${esc(x.ism)} ${esc(x.familiya||'')}</td><td>${esc(x.sana)}</td><td>${esc(x.turi)}</td><td>${esc(x.izoh||'')}</td></tr>`).join('')}
async function loadFinished(){const a=await api('/api/tayyor-mahsulot');$('#finishedBody').innerHTML=a.map(x=>`<tr><td>${esc(x.nomi)}</td><td>${esc(x.kodi||'')}</td><td>${esc(x.rang||'')}</td><td>${esc(x.miqdor)} ${esc(x.birlik)}</td><td>${money(x.narx)}</td></tr>`).join('')}
async function loadFinance(){const s=$('#finStart').value||'1900-01-01',e=$('#finEnd').value||'2999-12-31',x=await api(`/api/moliyaviy-xulosa?start=${s}&end=${e}`);$('#fIncome').textContent=money(x.kirim);$('#fExpense').textContent=money(x.xarajat);$('#fSalary').textContent=money(x.ishchi_tolov);$('#fBonus').textContent=money(x.bonus);$('#fProfit').textContent=money(x.sof_foyda)}
async function addOrderPayment(id){const miqdor=prompt('To‘lov summasi');if(!miqdor)return;await api(`/api/buyurtma/${id}/tolovlar`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({sana:today,miqdor})});refresh()}
async function regenContract(id){try{const x=await api(`/buyurtma/${id}/shartnoma-yaratish`,{method:'POST'});alert('Shartnoma yangilandi. Versiya: '+x.versiya)}catch(e){alert('Xato: '+e.message)}}
async function copyTrack(id){const x=await api(`/api/buyurtma/${id}/link`);try{await navigator.clipboard.writeText(x.url);alert('Mijoz kuzatuv havolasi nusxalandi')}catch(e){prompt('Havolani nusxalang:',x.url)}}
async function loadProgress(){const rows=await api('/api/buyurtmalar');for(const x of rows){try{const p=await api(`/api/buyurtma-progress/${esc(x.id)}`),el=$(`#pr${esc(x.id)}`);if(el)el.textContent=p.foiz+'%'}catch(e){}}}


async function loadExtras(){const a=await api('/api/qoshimcha-ish');$('#extrasBody').innerHTML=a.map(x=>`<tr><td>${esc(x.kod)}</td><td>${esc(x.mijoz)}</td><td>${esc(x.nomi)}</td><td>${money(x.summa)}</td><td>${esc(x.qoshimcha_kun)}</td></tr>`).join('')}
async function loadService(){const a=await api('/api/servis');$('#serviceBody').innerHTML=a.map(x=>`<tr><td>${esc(x.kod)}</td><td>${esc(x.mijoz)}</td><td>${esc(x.turi)}</td><td>${esc(x.muammo)}</td><td><span class="badge">${esc(x.holat)}</span></td><td>${esc(x.servis_sana||'')}</td><td>${esc(x.usta||'')}</td></tr>`).join('')}
async function loadDelivery(){const a=await api('/api/yetkazish');$('#deliveryBody').innerHTML=a.map(x=>`<tr><td>${esc(x.sana)}</td><td>${esc(x.navbat)}</td><td>${esc(x.kod)}</td><td>${esc(x.mijoz)}</td><td>${esc(x.haydovchi_ism||'')} ${esc(x.haydovchi_familiya||'')}</td><td>${x.lokatsiya?`<a href="${esc(x.lokatsiya)}" target="_blank">Xarita</a>`:esc(x.manzil||'')}</td><td>${esc(x.holat)}</td><td><button onclick="deliveryState(${esc(x.id)},'Yo‘lga chiqdim')">Yo‘lga</button> <button class="ok" onclick="deliveryState(${esc(x.id)},'Yetkazib berdim')">Topshirildi</button></td></tr>`).join('')}
async function deliveryState(id,holat){await api(`/api/yetkazish/${id}/holat`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({holat})});loadDelivery()}

async function refresh(){await Promise.all([loadWorkers(),loadTypes(),loadAttendance(),loadResults(),loadOrders(),loadStock(),loadTrips(),loadPayments(),loadPenalties(),loadDashboard(),loadTotals(),loadExpenses(),loadBonuses(),loadStatuses(),loadFinished(),loadFinance(),loadProgress(),loadExtras(),loadService(),loadDelivery()])}
const rf=$('#ratingForm');rf.onsubmit=async e=>{e.preventDefault();const d=fj(rf),id=d.buyurtma_id;delete d.buyurtma_id;const m=rf.querySelector('.msg');try{await api(`/api/buyurtma/${id}/baho`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});m.textContent='✅ Baho saqlandi'}catch(x){m.textContent='❌ '+x.message}};
bind('#extraForm','/api/qoshimcha-ish');bind('#serviceForm','/api/servis');bind('#deliveryForm','/api/yetkazish');
bind('#workerForm','/api/ishchilar');bind('#attendanceForm','/api/keldi-ketdi');bind('#resultForm','/api/natijalar');bind('#orderForm','/api/buyurtmalar');bind('#stockForm','/api/ombor-harakat');bind('#tripForm','/api/safarlar');bind('#paymentForm','/api/tolovlar');bind('#penaltyForm','/api/jarimalar');bind('#expenseForm','/api/xarajatlar');bind('#bonusForm','/api/bonuslar');bind('#statusForm','/api/ishchi-holatlari');bind('#finishedForm','/api/tayyor-mahsulot');setMonth();$('#finStart').value=$('#totalStart').value;$('#finEnd').value=$('#totalEnd').value;refresh();
</script>
</body></html>
"""

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)
