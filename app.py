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


app.secret_key = _get_or_create_secret_key()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("PHARM_ERP_HTTPS", "0") == "1",
)


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


    CREATE TABLE IF NOT EXISTS admin_akkauntlari (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        login TEXT UNIQUE NOT NULL,
        parol_hash TEXT NOT NULL,
        faol INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
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
        ("Lazer","Stanok","soat"),("Sehda yig‘ish","Ishlab chiqarish","dona"),
        ("Seh ishlari","Umumiy","soat"),("Yangi loyihalar","Loyiha","loyiha"),
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


    # PHARM_MEBEL_SEH_MIGRATION_V2_SAFE
    # Eski bazadagi "Sex" yozuvlarini ma'lumot yo'qotmasdan "Seh" ga o‘tkazadi.
    # Agar yangi nom avvaldan mavjud bo‘lsa, takroriy yozuvlar xavfsiz birlashtiriladi.

    old_work_types = conn.execute("""
        SELECT id, nomi FROM ish_turlari
        WHERE nomi LIKE '%Sex%' OR nomi LIKE '%sex%' OR nomi LIKE '%SEX%'
    """).fetchall()

    for old_work in old_work_types:
        old_id = int(old_work["id"])
        old_name = str(old_work["nomi"] or "")
        new_name = (old_name.replace("Sex", "Seh")
                            .replace("sex", "seh")
                            .replace("SEX", "SEH"))

        existing_work = conn.execute(
            "SELECT id FROM ish_turlari WHERE nomi=? AND id<>?",
            (new_name, old_id)
        ).fetchone()

        if existing_work:
            new_id = int(existing_work["id"])
            # Eski ish natijalarini yangi ish turiga bog‘laymiz.
            conn.execute(
                "UPDATE ish_natijalari SET ish_turi_id=? WHERE ish_turi_id=?",
                (new_id, old_id)
            )
            conn.execute("DELETE FROM ish_turlari WHERE id=?", (old_id,))
        else:
            conn.execute(
                "UPDATE ish_turlari SET nomi=? WHERE id=?",
                (new_name, old_id)
            )

    # Matn ko‘rinishida saqlangan topshiriqlar.
    conn.execute("""
        UPDATE ishchi_topshiriqlari
        SET ish_turi=REPLACE(REPLACE(REPLACE(ish_turi,'Sex','Seh'),'sex','seh'),'SEX','SEH')
        WHERE ish_turi LIKE '%Sex%' OR ish_turi LIKE '%sex%' OR ish_turi LIKE '%SEX%'
    """)

    # Buyurtma bosqichlarida bir xil yangi nom mavjud bo‘lsa, bajarilgan holatini birlashtiramiz.
    old_stages = conn.execute("""
        SELECT id, buyurtma_id, bosqich, bajarildi
        FROM buyurtma_bosqichlari
        WHERE bosqich LIKE '%Sex%' OR bosqich LIKE '%sex%' OR bosqich LIKE '%SEX%'
    """).fetchall()

    for old_stage in old_stages:
        old_stage_id = int(old_stage["id"])
        order_id = int(old_stage["buyurtma_id"])
        old_stage_name = str(old_stage["bosqich"] or "")
        new_stage_name = (old_stage_name.replace("Sex", "Seh")
                                        .replace("sex", "seh")
                                        .replace("SEX", "SEH"))

        existing_stage = conn.execute("""
            SELECT id, bajarildi FROM buyurtma_bosqichlari
            WHERE buyurtma_id=? AND bosqich=? AND id<>?
        """, (order_id, new_stage_name, old_stage_id)).fetchone()

        if existing_stage:
            merged_done = max(
                int(existing_stage["bajarildi"] or 0),
                int(old_stage["bajarildi"] or 0)
            )
            conn.execute(
                "UPDATE buyurtma_bosqichlari SET bajarildi=? WHERE id=?",
                (merged_done, int(existing_stage["id"]))
            )
            conn.execute(
                "DELETE FROM buyurtma_bosqichlari WHERE id=?",
                (old_stage_id,)
            )
        else:
            conn.execute(
                "UPDATE buyurtma_bosqichlari SET bosqich=? WHERE id=?",
                (new_stage_name, old_stage_id)
            )

    conn.execute("""
        UPDATE buyurtma_bosqich_hodisalari
        SET bosqich=REPLACE(REPLACE(REPLACE(bosqich,'Sex','Seh'),'sex','seh'),'SEX','SEH')
        WHERE bosqich LIKE '%Sex%' OR bosqich LIKE '%sex%' OR bosqich LIKE '%SEX%'
    """)

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
    # ADMIN_XAVFSIZLIK_V1
    # Eski ochiq admin parolini bir marta shifrlab bazaga ko'chiradi.
    # Keyingi ishga tushirishlarda parol hech qachon TXT yoki CMDda ko'rsatilmaydi.
    admin_login = (os.environ.get("PHARM_ERP_USER", "admin") or "admin").strip()
    admin_row = conn.execute(
        "SELECT id FROM admin_akkauntlari WHERE login=? LIMIT 1",
        (admin_login,),
    ).fetchone()
    old_password_path = os.path.join(_APP_DIR, "admin_parol.txt")
    if not admin_row:
        bootstrap_password = (os.environ.get("PHARM_ERP_PASSWORD") or "").strip()
        migrated_from_txt = False
        if not bootstrap_password and os.path.exists(old_password_path):
            try:
                with open(old_password_path, "r", encoding="utf-8") as f:
                    bootstrap_password = f.read().strip()
                migrated_from_txt = bool(bootstrap_password)
            except OSError:
                bootstrap_password = ""
        if bootstrap_password:
            conn.execute(
                "INSERT INTO admin_akkauntlari(login,parol_hash,faol) VALUES(?,?,1)",
                (admin_login, generate_password_hash(bootstrap_password)),
            )
            if migrated_from_txt:
                try:
                    os.remove(old_password_path)
                except OSError:
                    pass
    elif os.path.exists(old_password_path):
        # Bazada xavfsiz akkaunt mavjud bo'lsa, eski ochiq parol fayli kerak emas.
        try:
            os.remove(old_password_path)
        except OSError:
            pass

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
    r=p.add_run("PHARM MEBEL\nMEBEL ISHLAB CHIQARISH (BUYURTMA) SHARTNOMASI")
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
    sig.cell(0,0).text="PUDRATCHI: Pharm Mebel"
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

    draw_wrapped("PHARM MEBEL",True,16,45,20)
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
LOGIN_LOCKOUT_SECONDS = 900

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




# PHARM_MEBEL_SPLASH_LOGO_CLOCK_V1
SPLASH_HTML = '<!doctype html>\n<html lang="uz">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width,initial-scale=1">\n<title>Pharm Mebel</title>\n<style>\n*{box-sizing:border-box}\nhtml,body{margin:0;min-height:100%;font-family:Arial,sans-serif}\nbody{\n  min-height:100vh;display:grid;place-items:center;overflow:hidden;color:#14532d;\n  background:\n    radial-gradient(circle at 15% 18%,rgba(134,239,172,.50),transparent 34%),\n    radial-gradient(circle at 85% 82%,rgba(187,247,208,.78),transparent 38%),\n    linear-gradient(145deg,#fbfffc,#dcfce7);\n}\n.splash{\n  width:min(830px,94vw);min-height:min(680px,94vh);padding:24px 30px;\n  border:1px solid rgba(34,197,94,.25);border-radius:30px;\n  background:rgba(255,255,255,.84);box-shadow:0 28px 80px rgba(20,83,45,.18);\n  backdrop-filter:blur(13px);display:flex;flex-direction:column;align-items:center;\n  justify-content:center;text-align:center;animation:appear .65s ease both;\n}\n.logo{\n  width:min(680px,96%);max-height:385px;object-fit:contain;border-radius:22px;\n  filter:drop-shadow(0 14px 20px rgba(20,83,45,.13));animation:logoIn .9s ease both;\n}\n.clock{\n  margin-top:5px;font-size:clamp(48px,9vw,86px);line-height:1;font-weight:900;\n  letter-spacing:4px;color:#166534;text-shadow:0 4px 18px rgba(22,101,52,.13);\n}\n.date{margin-top:11px;font-size:clamp(16px,3vw,23px);font-weight:700;color:#3f7a52}\n.welcome{margin-top:17px;font-size:clamp(18px,3vw,23px);font-weight:900}\n.status{margin-top:7px;font-size:14px;color:#568467}\n.progress{\n  width:min(500px,84%);height:8px;margin-top:18px;overflow:hidden;\n  border-radius:999px;background:#d1fae5;\n}\n.bar{\n  width:0;height:100%;border-radius:999px;\n  background:linear-gradient(90deg,#4ade80,#15803d);animation:loading 5s linear forwards;\n}\n.enter{\n  margin-top:16px;padding:12px 28px;border:0;border-radius:13px;background:#16a34a;\n  color:#fff;font-size:16px;font-weight:800;cursor:pointer;\n  box-shadow:0 9px 24px rgba(22,163,74,.25);\n}\n.enter:hover{background:#15803d}\n@keyframes appear{from{opacity:0;transform:scale(.97)}to{opacity:1;transform:scale(1)}}\n@keyframes logoIn{from{opacity:0;transform:translateY(-15px)}to{opacity:1;transform:translateY(0)}}\n@keyframes loading{to{width:100%}}\n@media(max-width:600px){\n  .splash{min-height:94vh;padding:17px 12px;border-radius:22px}\n  .logo{width:100%;max-height:300px}\n  .clock{letter-spacing:2px}\n}\n</style>\n</head>\n<body>\n<section class="splash">\n  <img class="logo" src="data:image/jpeg;base64,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ZdVmJ36wSATABEgAABJSVIAAAExMDmnSua05+f1U18vnzj8h4K3xe2antldO5VRVfvqmKosiRFRKEiEoRJIkUplMJhACJQppqiFFFdKLdFy2rZ8/o89WE0reNM5+W1h87g8ObCbNrWy9+nU9q1bae7umJjTREwSBEwBEYb5o+mfmenNRGQ2TPLE/RXy11i+vV4NemjhO/wDDsudO0YymOF+hPnr6EvptkTGvTwPSt20nDj6d2PjnY79CYX1mEiJI+bcBn9fx4ux9N5l03XrmHgm3H9GzGLx4+t9K17YNOpxnsvGkc137Qt6z5u7jbtQRBKbPnf6I+daY6x1zkfW65dYmJ26QWAAAAAAAmmYJAAATBMSKebdK5rTDn1dNXL51Xg93grfFbbqW2V07nXTXfvmqJixHjiPTHz74cuX6QcGu2nujhVw7lXw/ZV+nKatdgWRJEAiYIU1UQportlNuqiK2vPf81Yxmn7hqPPzW8HnsBhz4PZNb2P0L9T2zUdu7e6RppCRAETAERifmX6a+Zac/u+hvnn6cR8yUda45Wn0dkuG5+ddJx1jqsZbtw36P+braYH6G+ePoWK7fExr1cF0nddJw4crnNcys2yE+CpPQOl8u6hrvMTC/zhr2xa7jx57N6/lbW9WAy2MrW113im0J79OMyevXHGuzcajPmnu8Prx5NieVfX053V87a3cpprv0x86/RXztTHV+uci65XLrExO3WCUwCYAAAAAAAAAJiYJIHNelc0pz6BVRVy+fcx2Rx1b4ratU2uunda6a9PQmqmqsx4fb4M6cV1Db9O5PHiY6H19HP3d5ttwncMpiJp3uaard0krwkimSSFuIWbemxXc6eD11x7nRpWy1t6rFdlbwajtGq4YU4LNYHHnw2xa5sXfbqm3aht/b3SNNETAABAiMT8yfTnzHTm93078xfTyY4L37BW0+brty3jyZn6G1zbdOq383fSXzdFdf+hfnn6HU24adXBNJ3bSMeLpnZeOdj06UwtomJkiaYfOWvZ/X8eHsnS+ZdOv1rVyLX45zbvfBM+Pdu7/NP0nptc432PjZzXedF3rLDusw17ZEwmJS+dvor52pjqnXeSdbrj1hDbslELVolEkEwAAAAESAAAExMExIp5p0vm1Ofn1VFfL51eOyOOrfF7Xq21027nVTVr31liEaRrGic/L78f7a6cHm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" alt="Pharm Mebel logosi">\n  <div id="clock" class="clock">00:00:00</div>\n  <div id="date" class="date"></div>\n  <div class="welcome">Pharm Mebel dasturiga xush kelibsiz</div>\n  <div class="status">Dastur ochilmoqda...</div>\n  <div class="progress"><div class="bar"></div></div>\n  <button class="enter" type="button" onclick="goNext()">Kirish</button>\n</section>\n<script>\nconst NEXT_URL = {{ next_url|tojson }};\nconst DAYS = ["Yakshanba","Dushanba","Seshanba","Chorshanba","Payshanba","Juma","Shanba"];\nconst MONTHS = ["yanvar","fevral","mart","aprel","may","iyun",\n                "iyul","avgust","sentabr","oktabr","noyabr","dekabr"];\nfunction two(n){ return String(n).padStart(2,"0"); }\nfunction updateClock(){\n  const d = new Date();\n  document.getElementById("clock").textContent =\n    two(d.getHours()) + ":" + two(d.getMinutes()) + ":" + two(d.getSeconds());\n  document.getElementById("date").textContent =\n    DAYS[d.getDay()] + ", " + d.getDate() + " " +\n    MONTHS[d.getMonth()] + " " + d.getFullYear();\n}\nlet leaving = false;\nfunction goNext(){\n  if(leaving) return;\n  leaving = true;\n  document.body.style.transition = "opacity .35s ease";\n  document.body.style.opacity = "0";\n  setTimeout(() => window.location.replace(NEXT_URL), 330);\n}\nupdateClock();\nsetInterval(updateClock,1000);\nsetTimeout(goNext,5000);\n</script>\n</body>\n</html>'

@app.before_request
def require_login():
    _auto_backup_if_needed()
    public_endpoints = {
        "splash",
        "login", "admin_setup", "static", "public_track", "order_qr",
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


def _admin_account(login):
    c = get_db()
    row = c.execute(
        "SELECT id,login,parol_hash,faol FROM admin_akkauntlari WHERE login=? LIMIT 1",
        ((login or "").strip(),),
    ).fetchone()
    c.close()
    return row


def _admin_is_configured():
    c = get_db()
    row = c.execute("SELECT id FROM admin_akkauntlari WHERE faol=1 LIMIT 1").fetchone()
    c.close()
    return bool(row)


@app.route("/admin/setup", methods=["GET", "POST"])
def admin_setup():
    if _admin_is_configured():
        return redirect(url_for("login"))
    error = ""
    if request.method == "POST":
        user = (request.form.get("user") or "admin").strip()
        password = request.form.get("password") or ""
        confirm = request.form.get("confirm") or ""
        if len(user) < 3:
            error = "Login kamida 3 ta belgidan iborat bo‘lsin."
        elif _weak_password(password):
            error = "Parol kamida 8 belgi, kamida bitta harf va bitta raqamdan iborat bo‘lsin."
        elif password != confirm:
            error = "Ikki parol bir xil emas."
        else:
            c = get_db()
            try:
                c.execute(
                    "INSERT INTO admin_akkauntlari(login,parol_hash,faol) VALUES(?,?,1)",
                    (user, generate_password_hash(password)),
                )
                c.commit()
            except sqlite3.IntegrityError:
                c.rollback()
                error = "Bu login avval ishlatilgan."
            finally:
                c.close()
            if not error:
                session.clear()
                session["logged_in"] = True
                session["user"] = user
                log_action("admin_first_setup", f"user={user}")
                return redirect(url_for("home"))
    return render_template_string(ADMIN_SETUP_HTML, error=error)


@app.route("/login", methods=["GET","POST"])
def login():
    if not _admin_is_configured():
        return redirect(url_for("admin_setup"))
    error=''
    if request.method=='POST':
        user=(request.form.get('user') or '').strip()
        password=request.form.get('password') or ''
        locked, wait_sec = _is_login_locked(user)
        account = _admin_account(user)
        if locked:
            minutes = max(1, (wait_sec + 59) // 60)
            error=f'Juda ko‘p xato urinish. Taxminan {minutes} daqiqadan so‘ng qayta urinib ko‘ring.'
        elif account and int(account['faol'] or 0) == 1 and check_password_hash(account['parol_hash'], password):
            _clear_login_attempts(user)
            session.clear()
            session['logged_in']=True
            session['user']=account['login']
            log_action('admin_login', f"user={account['login']}")
            return redirect(url_for('home'))
        else:
            _register_failed_login(user)
            log_action('admin_login_failed', f'user={user}')
            error='Login yoki parol xato'
    return render_template_string(LOGIN_HTML,error=error)


@app.route("/logout")
def logout():
    user = session.get("user", "")
    log_action("admin_logout", f"user={user}")
    session.clear()
    return redirect(url_for('login'))


@app.route("/")
def splash():
    next_url = url_for("home") if session.get("logged_in") else url_for("login")
    return render_template_string(SPLASH_HTML, next_url=next_url)


@app.route("/dashboard")
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
STAGES=["Buyurtma qabul qilindi","Razmer olindi","Chizma tayyorlanmoqda","Mijoz tasdiqlashi kutilmoqda","Material tayyorlanmoqda","Kesish","Kromka","Teshish","Frezalash","Lazer","Bo‘yash","Sayqalash","Yig‘ish","Sehda tekshirish","Qadoqlash","Yetkazishga tayyor","Haydovchiga topshirildi","Yetkazib berildi","Buyurtma yopildi"]

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
      <h1>Pharm Mebel</h1>
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
    p.setFont('Helvetica-Bold',20); p.drawCentredString(w/2,h-60,'Pharm Mebel - TOLOV CHEKI')
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


ADMIN_SETUP_HTML = r"""
<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pharm Mebel - Birinchi xavfsiz sozlash</title><style>body{margin:0;background:linear-gradient(135deg,#0f1b33,#16a34a);font-family:Arial;display:grid;place-items:center;min-height:100vh}.box{background:#fff;padding:28px;border-radius:18px;width:min(410px,92%);box-shadow:0 20px 50px #0005}h2{margin-top:0}input{width:100%;padding:12px;margin:8px 0;border:1px solid #cbd5e1;border-radius:9px;box-sizing:border-box}button{width:100%;padding:12px;border:0;border-radius:9px;background:#16a34a;color:#fff;font-weight:700}.err{color:#b91c1c;font-size:13px}.note{font-size:13px;color:#475569;line-height:1.45}</style></head><body><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h2>🔐 Birinchi xavfsiz sozlash</h2><input name="user" placeholder="Admin login" value="admin" minlength="3" required><input name="password" type="password" placeholder="Yangi parol" minlength="8" required><input name="confirm" type="password" placeholder="Parolni takrorlang" minlength="8" required><button>Admin akkauntini yaratish</button><div class="err">{{error}}</div></form></body></html>
"""

LOGIN_HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pharm Mebel — Rahbar kirishi</title>
<style>
:root{
  --navy:#0b1736;
  --blue:#2563eb;
  --blue2:#38bdf8;
  --green:#16a34a;
  --text:#172033;
  --muted:#64748b;
  --line:#dbe5f0;
  --danger:#b91c1c;
}
*{box-sizing:border-box}
html,body{margin:0;min-height:100%;font-family:Arial,Helvetica,sans-serif;color:var(--text)}
body{
  min-height:100vh;
  display:grid;
  place-items:center;
  padding:24px;
  background:
    radial-gradient(circle at 12% 16%,rgba(56,189,248,.28),transparent 31%),
    radial-gradient(circle at 88% 82%,rgba(37,99,235,.30),transparent 34%),
    linear-gradient(145deg,#07142f 0%,#123b8f 50%,#0b5fa8 100%);
  overflow-x:hidden;
}
body:before,body:after{
  content:"";position:fixed;border-radius:999px;filter:blur(2px);pointer-events:none;
  background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12)
}
body:before{width:320px;height:320px;left:-100px;bottom:-120px}
body:after{width:230px;height:230px;right:-65px;top:-60px}
.login-shell{
  width:min(900px,100%);
  display:grid;
  grid-template-columns:1.05fr .95fr;
  border:1px solid rgba(255,255,255,.22);
  border-radius:28px;
  overflow:hidden;
  background:rgba(255,255,255,.96);
  box-shadow:0 32px 90px rgba(2,8,23,.38);
  animation:show .55s ease both;
}
.brand-panel{
  position:relative;
  min-height:560px;
  padding:48px 42px;
  display:flex;
  flex-direction:column;
  justify-content:space-between;
  color:#fff;
  background:
    radial-gradient(circle at 78% 20%,rgba(56,189,248,.34),transparent 30%),
    linear-gradient(150deg,#0b1736,#154aa8 62%,#0784c8);
  overflow:hidden;
}
.brand-panel:after{
  content:"360°";
  position:absolute;
  right:-18px;
  bottom:18px;
  font-size:132px;
  line-height:1;
  font-weight:900;
  color:rgba(255,255,255,.055);
  letter-spacing:-10px;
}
.brand-top{position:relative;z-index:1}
.logo-mark{
  width:66px;height:66px;border-radius:20px;
  display:grid;place-items:center;
  background:linear-gradient(145deg,#fff,#dbeafe);
  color:#154aa8;font-size:31px;font-weight:900;
  box-shadow:0 15px 35px rgba(2,8,23,.25)
}
.brand-title{margin:28px 0 8px;font-size:42px;line-height:1.05;font-weight:900;letter-spacing:-1px}
.brand-sub{font-size:18px;line-height:1.55;color:#dbeafe;max-width:390px}
.brand-features{position:relative;z-index:1;display:grid;gap:11px;margin-top:28px}
.feature{display:flex;align-items:center;gap:10px;font-size:14px;color:#e0f2fe}
.feature i{width:27px;height:27px;border-radius:9px;display:grid;place-items:center;background:rgba(255,255,255,.13);font-style:normal}
.form-panel{padding:48px 44px;display:flex;align-items:center;background:#fff}
.form-wrap{width:100%;max-width:360px;margin:auto}
.role-badge{display:inline-flex;align-items:center;gap:7px;padding:7px 11px;border-radius:999px;background:#eff6ff;color:#1d4ed8;font-size:12px;font-weight:800}
h1{font-size:30px;margin:17px 0 7px;letter-spacing:-.5px}
.intro{margin:0 0 25px;color:var(--muted);font-size:14px;line-height:1.5}
label{display:block;margin:13px 0 6px;font-size:13px;font-weight:800;color:#334155}
.field{position:relative}
.field input{
  width:100%;height:50px;padding:0 46px 0 43px;
  border:1px solid var(--line);border-radius:13px;
  background:#f8fafc;color:#0f172a;font-size:15px;outline:none;
  transition:.18s ease;
}
.field input:focus{border-color:#60a5fa;background:#fff;box-shadow:0 0 0 4px rgba(37,99,235,.11)}
.field-icon{position:absolute;left:15px;top:50%;transform:translateY(-50%);font-size:17px;opacity:.72}
.show-pass{position:absolute;right:8px;top:50%;transform:translateY(-50%);border:0;background:transparent;color:#475569;font-size:12px;font-weight:800;padding:8px;cursor:pointer}
.login-btn{
  width:100%;height:51px;margin-top:20px;border:0;border-radius:13px;
  color:#fff;font-size:15px;font-weight:900;cursor:pointer;
  background:linear-gradient(100deg,#2563eb,#0284c7);
  box-shadow:0 13px 28px rgba(37,99,235,.25);
  transition:.18s ease;
}
.login-btn:hover{transform:translateY(-1px);box-shadow:0 16px 32px rgba(37,99,235,.30)}
.login-btn:active{transform:translateY(0)}
.error-box{margin-top:13px;padding:10px 12px;border-radius:10px;background:#fef2f2;color:var(--danger);font-size:13px;font-weight:700;border:1px solid #fecaca}
.other-title{display:flex;align-items:center;gap:10px;margin:24px 0 13px;color:#94a3b8;font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.8px}
.other-title:before,.other-title:after{content:"";height:1px;background:#e2e8f0;flex:1}
.role-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px}
.role-link{
  min-height:67px;padding:10px;border:1px solid #e2e8f0;border-radius:13px;
  text-decoration:none;color:#334155;background:#f8fafc;
  display:flex;align-items:center;gap:9px;transition:.18s ease
}
.role-link:hover{border-color:#93c5fd;background:#eff6ff;transform:translateY(-1px)}
.role-icon{width:35px;height:35px;border-radius:10px;display:grid;place-items:center;background:#dbeafe;font-size:17px;flex:none}
.role-link.driver .role-icon{background:#dcfce7}
.role-link b{display:block;font-size:12px}.role-link small{display:block;color:#64748b;font-size:10px;margin-top:3px}
.footer-note{text-align:center;color:#94a3b8;font-size:11px;margin-top:18px}
@keyframes show{from{opacity:0;transform:translateY(12px) scale(.985)}to{opacity:1;transform:none}}
@media(max-width:760px){
  body{padding:13px;display:block}
  .login-shell{grid-template-columns:1fr;max-width:480px;margin:12px auto;border-radius:22px}
  .brand-panel{min-height:auto;padding:26px 25px 24px}
  .brand-title{font-size:29px;margin-top:18px}.brand-sub{font-size:14px}
  .brand-features{display:none}.brand-panel:after{font-size:85px;bottom:-5px}
  .logo-mark{width:52px;height:52px;border-radius:16px;font-size:24px}
  .form-panel{padding:29px 22px 25px}
  h1{font-size:26px}.role-grid{grid-template-columns:1fr 1fr}
}
@media(max-width:390px){.role-grid{grid-template-columns:1fr}.role-link{min-height:58px}}
</style>
</head>
<body>
<main class="login-shell">
  <section class="brand-panel">
    <div class="brand-top">
      <div class="logo-mark">M</div>
      <div class="brand-title">Pharm Mebel</div>
      <div class="brand-sub">Korxona, buyurtmalar, ishchilar va omborni yagona tizimda boshqaring.</div>
    </div>
    <div class="brand-features">
      <div class="feature"><i>✓</i><span>Buyurtmalar va ishlab chiqarish nazorati</span></div>
      <div class="feature"><i>✓</i><span>Ishchi, Shofyor va Mijoz kabinetlari</span></div>
      <div class="feature"><i>✓</i><span>Ombor, xarajat va hisobotlar</span></div>
    </div>
  </section>

  <section class="form-panel">
    <form class="form-wrap" method="post">
      <input type="hidden" name="csrf_token" value="{{csrf_token()}}">
      <span class="role-badge">◆ Rahbar kabineti</span>
      <h1>Xush kelibsiz</h1>
      <p class="intro">Dastur boshqaruv paneliga kirish uchun login va parolingizni kiriting.</p>

      <label for="user">Login</label>
      <div class="field">
        <span class="field-icon">👤</span>
        <input id="user" name="user" placeholder="Login" value="admin" autocomplete="username" required autofocus>
      </div>

      <label for="password">Parol</label>
      <div class="field">
        <span class="field-icon">🔑</span>
        <input id="password" name="password" type="password" placeholder="Parol" autocomplete="current-password" required>
        <button class="show-pass" type="button" onclick="togglePassword()" id="showPassword">Ko‘rsatish</button>
      </div>

      <button class="login-btn" type="submit">Kirish →</button>
      {% if error %}<div class="error-box">⚠ {{error}}</div>{% endif %}

      <div class="other-title">Boshqa kabinetlar</div>
      <div class="role-grid">
        <a class="role-link" href="/ishchi/login"><span class="role-icon">👷</span><span><b>Ishchi kirishi</b><small>Kabinetga kirish</small></span></a>
        <a class="role-link" href="/ishchi/royxat"><span class="role-icon">＋</span><span><b>Ishchi ro‘yxati</b><small>Yangi akkaunt</small></span></a>
        <a class="role-link driver" href="/shofyor/login"><span class="role-icon">🚚</span><span><b>Shofyor kirishi</b><small>Kabinetga kirish</small></span></a>
        <a class="role-link driver" href="/shofyor/royxat"><span class="role-icon">＋</span><span><b>Shofyor ro‘yxati</b><small>Yangi akkaunt</small></span></a>
      </div>
      <div class="footer-note">Mebel360° boshqaruv tizimi</div>
    </form>
  </section>
</main>
<script>
function togglePassword(){
  const input=document.getElementById('password');
  const button=document.getElementById('showPassword');
  const visible=input.type==='text';
  input.type=visible?'password':'text';
  button.textContent=visible?'Ko‘rsatish':'Yashirish';
}
</script>
</body>
</html>
"""


WORKER_BASE_STYLE = """
<style>
*{box-sizing:border-box}body{margin:0;font-family:Arial;background:#eef3f8;color:#182235}.head{background:linear-gradient(135deg,#0f1b33,#2563eb);color:white;padding:18px}.wrap{max-width:1050px;margin:auto;padding:16px}.box,.card{background:white;border-radius:16px;padding:18px;box-shadow:0 8px 24px #0f172a18;margin-bottom:14px}input,select,textarea{width:100%;padding:11px;border:1px solid #cbd5e1;border-radius:9px;margin:5px 0 10px}button,.btn{display:inline-block;border:0;border-radius:9px;padding:10px 14px;background:#2563eb;color:white;font-weight:700;text-decoration:none;cursor:pointer}.green{background:#16a34a}.red{background:#dc2626}.muted{color:#64748b;font-size:13px}.err{color:#b91c1c}.ok{color:#166534}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.stat b{font-size:25px;color:#2563eb}.task{border-left:5px solid #2563eb}.bar{height:10px;background:#e2e8f0;border-radius:20px;overflow:hidden}.bar i{display:block;height:100%;background:#16a34a}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:8px;border-bottom:1px solid #e5e7eb;text-align:left}@media(max-width:700px){.grid{grid-template-columns:1fr}.wrap{padding:9px}}
</style>
"""

DRIVER_REGISTER_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor ro‘yxatdan o‘tishi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🚚 Shofyor ro‘yxatdan o‘tishi</b></div><div class="wrap"><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h2>Yangi akkaunt</h2><label>Ism</label><input name="ism" required><label>Familiya</label><input name="familiya"><label>Telefon</label><input name="telefon" placeholder="+998901234567" required><label>Login</label><input name="login" required><label>Parol</label><input name="password" type="password" minlength="8" required><button>Ro‘yxatdan o‘tish</button><p class="ok">{{msg}}</p><p class="err">{{error}}</p><p><a href="/shofyor/login">Akkauntim bor — kirish</a></p></form></div></body></html>"""

DRIVER_LOGIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor kirishi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🚚 Pharm Mebel — Shofyor kabineti</b></div><div class="wrap"><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h2>Kirish</h2><input name="login" placeholder="Login" required><input name="password" type="password" placeholder="Parol" required><button>Kirish</button><p class="err">{{error}}</p><p><a href="/shofyor/royxat">Yangi shofyor — ro‘yxatdan o‘tish</a></p><p><a href="/login">Rahbar kirishi</a></p></form></div></body></html>"""

DRIVER_DASHBOARD_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor kabineti</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><div class="wrap" style="padding:0"><b>🚚 {{driver['ism']}} {{driver['familiya']}}</b><a class="btn red" style="float:right" href="/shofyor/logout">Chiqish</a></div></div><div class="wrap"><h2>Menga biriktirilgan yuklar</h2>{% for x in deliveries %}<a href="/shofyor/yetkazish/{{x['id']}}" style="text-decoration:none;color:inherit"><div class="card task"><div style="display:flex;justify-content:space-between;gap:10px"><div><b style="font-size:21px">{{x['kod']}}</b><p>{{x['mahsulot']}}</p><p class="muted">{{x['manzil']}} · {{x['qavat'] or '-'}}-qavat · Lift: {{x['lift'] or '-'}}</p></div><div><span class="btn {% if x['holat']=='Yetkazib berildi' %}green{% endif %}">{{x['holat']}}</span></div></div></div></a>{% else %}<div class="card">Hozircha yuk biriktirilmagan.</div>{% endfor %}</div></body></html>"""

DRIVER_DETAIL_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{x['kod']}}</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><div class="wrap" style="padding:0"><b>📦 {{x['kod']}}</b><a class="btn" style="float:right" href="/shofyor/kabinet">Orqaga</a></div></div><div class="wrap"><div class="card"><h2>{{x['mahsulot']}}</h2><p><b>Mijoz:</b> {{x['mijoz']}}</p><p><b>Telefon:</b> <a href="tel:{{x['telefon']}}">{{x['telefon']}}</a></p><p><b>Manzil:</b> {{x['manzil']}}</p><p><b>Mo‘ljal:</b> {{x['moljal'] or '-'}}</p><p><b>Qavat:</b> {{x['qavat'] or '-'}}</p><p><b>Lift:</b> {{x['lift'] or '-'}}</p><p><b>Katta mashina:</b> {{x['katta_mashina'] or '-'}}</p><p><b>Qadoqlar:</b> {{x['qadoq_soni']}} ta</p><p><b>Yetkazish navbati:</b> {{x['navbat']}}</p><p><b>Izoh:</b> {{x['izoh'] or x['buyurtma_izoh'] or '-'}}</p>{% if x['lokatsiya'] %}
<div class="card" style="background:#f8fafc">
<h3>📍 Xarita tanlang</h3>
<p class="muted">Lokatsiyani o‘zingiz xohlagan xaritada oching.</p>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
<a class="btn green" href="https://www.google.com/maps/search/?api=1&query={{x['lokatsiya']|urlencode}}" target="_blank">Google Maps</a>
<a class="btn" style="background:#ef4444" href="https://yandex.com/maps/?text={{x['lokatsiya']|urlencode}}" target="_blank">Yandex Maps</a>
</div></div>
{% endif %}</div><div class="card"><h3>Holat: {{x['holat']}}</h3><form method="post" action="/shofyor/yetkazish/{{x['id']}}/holat"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><button name="action" value="yolga">🚚 Yo‘lga chiqdim</button><button name="action" value="yetib" class="green">📍 Yetib keldim</button><button name="action" value="yetkazildi" style="background:#7c3aed">✅ Yetkazib berdim</button></form><p class="muted">Yo‘lga chiqdi: {{x['yolga_chiqdi'] or '-'}}</p><p class="muted">Yetib keldi: {{x['yetib_keldi'] or '-'}}</p><p class="muted">Yetkazildi: {{x['yetkazildi'] or '-'}}</p></div></div></body></html>"""

DRIVER_ADMIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shofyor boshqaruvi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🚚 Shofyor boshqaruvi</b><a class="btn" style="float:right" href="/dashboard">ERP bosh sahifa</a></div><div class="wrap"><p class="ok">{{msg}}</p><div class="grid"><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h3>Shofyor login yaratish</h3><input type="hidden" name="action" value="account"><label>Shofyor</label><select name="ishchi_id" required>{% for d in drivers %}<option value="{{d['id']}}">{{d['ism']}} {{d['familiya']}}</option>{% endfor %}</select><label>Login</label><input name="login" required><label>Parol</label><input name="password" type="password" minlength="8" required><button>Akkaunt yaratish</button></form><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h3>Yuk biriktirish</h3><input type="hidden" name="action" value="delivery"><label>Buyurtma</label><select name="buyurtma_id" required>{% for o in orders %}<option value="{{o['id']}}">{{o['kod']}} — {{o['mahsulot']}}</option>{% endfor %}</select><label>Shofyor</label><select name="shofyor_id" required>{% for d in drivers %}<option value="{{d['id']}}">{{d['ism']}} {{d['familiya']}}</option>{% endfor %}</select><label>Qadoq soni</label><input type="number" name="qadoq_soni" value="1" min="1"><label>Yetkazish navbati</label><input type="number" name="navbat" value="1" min="1"><label>Izoh</label><textarea name="izoh"></textarea><button>Biriktirish</button></form></div><div class="box" style="overflow:auto"><h3>Shofyor akkauntlari</h3><table><tr><th>Shofyor</th><th>Telefon</th><th>Login</th><th>Holat</th><th>Amal</th></tr>{% for a in driver_accounts %}<tr><td>{{a['ism']}} {{a['familiya']}}</td><td>{{a['telefon'] or '-'}}</td><td>{{a['login']}}</td><td>{% if a['admin_tasdiq'] %}✅ Tasdiqlangan{% else %}⏳ Kutilmoqda{% endif %}</td><td>{% if not a['admin_tasdiq'] %}<form method="post" style="display:inline"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><input type="hidden" name="action" value="approve_driver"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="green">Tasdiqlash</button></form>{% endif %}<form method="post" style="display:inline"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><input type="hidden" name="action" value="block_driver"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="red">Bloklash</button></form></td></tr>{% endfor %}</table></div><div class="box" style="overflow:auto"><h3>Biriktirilgan yuklar</h3><table><tr><th>Buyurtma</th><th>Mahsulot</th><th>Shofyor</th><th>Qadoq</th><th>Holat</th></tr>{% for a in assigned %}<tr><td>{{a['kod']}}</td><td>{{a['mahsulot']}}</td><td>{{a['ism']}} {{a['familiya']}}</td><td>{{a['qadoq_soni']}}</td><td>{{a['holat']}}</td></tr>{% endfor %}</table></div></div></body></html>"""


WORKER_REGISTER_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi ro‘yxati</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>👷 Ishchi ro‘yxatdan o‘tishi</b></div><div class="wrap"><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h2>Telefon raqamingiz</h2><p class="muted">Masalan: +998 90 123 45 67</p><input name="telefon" required placeholder="+998901234567"><button>Kod olish</button><p class="ok">{{msg}}</p><p class="err">{{error}}</p>{% if demo_code %}<div class="card"><b>Sinov kodi: {{demo_code}}</b><p class="muted">Haqiqiy SMS xizmati ulanmaguncha shu koddan foydalaning.</p><a class="btn green" href="/ishchi/kod">Kodni kiritish</a></div>{% endif %}<p><a href="/ishchi/login">Akkauntim bor — kirish</a></p></form></div></body></html>"""

WORKER_VERIFY_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kodni tasdiqlash</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🔐 Telefonni tasdiqlash</b></div><div class="wrap"><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><p>{{telefon}}</p><label>Kod</label><input name="kod" inputmode="numeric" maxlength="6" required><label>Ism</label><input name="ism" required><label>Familiya</label><input name="familiya"><label>Yangi login</label><input name="login" required><label>Yangi parol</label><input name="password" type="password" minlength="8" required><button>Ro‘yxatdan o‘tish</button><p class="err">{{error}}</p></form></div></body></html>"""

WORKER_WAIT_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kutilmoqda</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="wrap"><div class="box"><h2>✅ Ro‘yxatdan o‘tdingiz</h2><p>Endi administrator akkauntingizni tasdiqlaydi. Tasdiqlangach login va parolingiz bilan kirasiz.</p><a class="btn" href="/ishchi/login">Kirish sahifasi</a></div></div></body></html>"""

WORKER_LOGIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi kirishi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>🏭 Pharm Mebel — Ishchi kabineti</b></div><div class="wrap"><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h2>Kirish</h2><input name="login" placeholder="Login" required><input name="password" type="password" placeholder="Parol" required><button>Kirish</button><p class="err">{{error}}</p><p><a href="/ishchi/royxat">Yangi ro‘yxatdan o‘tish</a></p></form></div></body></html>"""

WORKER_DASHBOARD_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi kabineti</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><div class="wrap" style="padding:0"><b>👷 {{worker['ism']}} {{worker['familiya']}}</b> — {{worker['lavozim']}} <a class="btn red" style="float:right" href="/ishchi/logout">Chiqish</a></div></div><div class="wrap">
{% with messages = get_flashed_messages() %}{% if messages %}<div class="card err">{{messages[0]}}</div>{% endif %}{% endwith %}
<div class="grid"><div class="card stat"><span>HOZIRGI HOLAT</span><br><b style="color:{% if worker_state=='Ishlayapti' %}#16a34a{% else %}#64748b{% endif %}">{{worker_state}}</b></div><div class="card stat"><span>ISHLAGAN KUN</span><br><b>{{stats['kun'] or 0}}</b></div><div class="card stat"><span>JAMI SOAT</span><br><b>{{'%.1f'|format(stats['soat'] or 0)}}</b></div></div>
{% if active_task %}<div class="card" style="border-left:7px solid #16a34a"><h3>🟢 Hozir ishlayotgan ishim</h3><p><b>{{active_task['ish_turi']}}</b> — {{active_task['buyurtma_kodi'] or 'Buyurtmasiz'}}</p><p>{{active_task['tavsif']}}</p><p>Boshlangan vaqt: <b>{{active_task['boshlandi_vaqt']}}</b></p></div>{% endif %}
<h2>Topshiriqlarim</h2>{% for t in tasks %}<div class="card task"><b>{{t['ish_turi']}}</b> {% if t['buyurtma_kodi'] %}<span class="muted">— {{t['buyurtma_kodi']}}</span>{% endif %}<p>{{t['tavsif']}}</p><div class="bar"><i style="width:{{t['progress']}}%"></i></div><p><b>{{t['progress']}}%</b> · {{t['holat']}} · Reja: {{t['sana']}}{% if t['tugash_sana'] %} — {{t['tugash_sana']}}{% endif %}</p><p class="muted">Boshladi: {{t['boshlandi_vaqt'] or '-'}} · Tugatdi: {{t['tugadi_vaqt'] or '-'}}</p>
{% if t['holat']=='Yangi' or t['holat']=='Jarayonda' %}<form method="post" action="/ishchi/topshiriq/{{t['id']}}/boshlash"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><button class="green">▶ Ishni boshladim</button></form>{% elif t['holat']=='Ishlayapti' %}<form method="post" action="/ishchi/topshiriq/{{t['id']}}/tugatish"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><button style="background:#7c3aed">✅ Ishni tugatdim</button></form>{% endif %}
{% if t['holat']!='Tayyor' %}<form method="post" action="/ishchi/topshiriq/{{t['id']}}"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><input type="hidden" name="holat" value="{{t['holat']}}"><label>Jarayon foizi</label><input type="number" name="progress" min="0" max="100" value="{{t['progress']}}"><button>Foizni yangilash</button></form>{% endif %}</div>{% else %}<div class="card">Hozircha topshiriq yo‘q. Holatingiz: <b>Bo‘sh</b>.</div>{% endfor %}</div></body></html>"""

WORKER_ADMIN_HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ishchi boshqaruvi</title>"""+WORKER_BASE_STYLE+r"""</head><body><div class="head"><b>👥 Ishchi kabinetlarini boshqarish</b> <a class="btn" style="float:right" href="/dashboard">ERP bosh sahifa</a></div><div class="wrap">
<div class="box" style="overflow:auto"><h3>Ishchilarning hozirgi holati</h3><table><tr><th>Ishchi</th><th>Holat</th><th>Buyurtma</th><th>Ish turi</th><th>Boshlagan vaqt</th></tr>{% for w in worker_states %}<tr><td>{{w['ism']}} {{w['familiya']}}</td><td>{% if w['ish_holati']=='Ishlayapti' %}<b style="color:#16a34a">🟢 Ishlayapti</b>{% else %}<b style="color:#64748b">⚪ Bo‘sh</b>{% endif %}</td><td>{{w['buyurtma_kodi'] or '-'}}</td><td>{{w['ish_turi'] or '-'}}</td><td>{{w['boshlandi_vaqt'] or '-'}}</td></tr>{% endfor %}</table></div>
<div class="grid"><form class="box" method="post"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><h3>Yangi topshiriq</h3><input type="hidden" name="action" value="task"><label>Ishchi</label><select name="ishchi_id" required>{% for w in workers %}<option value="{{w['id']}}">{{w['ism']}} {{w['familiya']}} — {{w['lavozim']}}</option>{% endfor %}</select><label>Buyurtma kodi</label><input name="buyurtma_kodi" placeholder="AB 007"><label>Ish turi</label><input name="ish_turi" required placeholder="Kesish / Rover / Yig‘ish"><label>Topshiriq</label><textarea name="tavsif" placeholder="Nima ish qilishi kerakligini batafsil yozing"></textarea><label>Boshlanish rejasi</label><input type="date" name="sana" value="{{today}}"><label>Tugash rejasi</label><input type="date" name="tugash_sana"><button>Topshiriq berish</button></form><div class="box" style="grid-column:span 2;overflow:auto"><h3>Ro‘yxatdan o‘tganlar</h3><table><tr><th>Ishchi</th><th>Telefon</th><th>Login</th><th>Holat</th><th>Amal</th></tr>{% for a in accounts %}<tr><td>{{a['ism']}} {{a['familiya']}}</td><td>{{a['telefon']}}</td><td>{{a['login'] or ''}}</td><td>{% if a['admin_tasdiq'] %}✅ Tasdiqlangan{% else %}⏳ Kutilmoqda{% endif %}</td><td>{% if not a['admin_tasdiq'] %}<form method="post" style="display:inline"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><input type="hidden" name="action" value="approve"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="green">Tasdiqlash</button></form>{% endif %}<form method="post" style="display:inline"><input type="hidden" name="csrf_token" value="{{csrf_token()}}"><input type="hidden" name="action" value="block"><input type="hidden" name="account_id" value="{{a['id']}}"><button class="red">Bloklash</button></form></td></tr>{% endfor %}</table></div></div>
<div class="box" style="overflow:auto"><h3>Topshiriqlar tarixi</h3><table><tr><th>Ishchi</th><th>Buyurtma</th><th>Ish</th><th>Holat</th><th>Progress</th><th>Boshladi</th><th>Tugatdi</th></tr>{% for t in tasks %}<tr><td>{{t['ism']}} {{t['familiya']}}</td><td>{{t['buyurtma_kodi']}}</td><td>{{t['ish_turi']}}</td><td>{{t['holat']}}</td><td>{{t['progress']}}%</td><td>{{t['boshlandi_vaqt'] or '-'}}</td><td>{{t['tugadi_vaqt'] or '-'}}</td></tr>{% endfor %}</table></div></div></body></html>"""


HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pharm Mebel V5.1.4</title>
<style>
:root{--nav:#0f1b33;--blue:#2563eb;--bg:#eef3f8;--card:#fff;--text:#182235;--muted:#64748b;--danger:#dc2626;--ok:#16a34a}
*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;background:var(--bg);color:var(--text)}
header{background:linear-gradient(135deg,#0f1b33,#1d4ed8);color:#fff;padding:20px;position:sticky;top:0;z-index:5}
.top{max-width:1400px;margin:auto;display:grid;grid-template-columns:auto auto 1fr;align-items:center;gap:16px}
.brand{display:flex;align-items:center;gap:12px;min-width:310px}.brand-logo{width:105px;height:66px;object-fit:contain;background:#fff;border-radius:12px;padding:4px;box-shadow:0 5px 16px #0003}.brand-text h1{margin:0;font-size:27px}.sub{opacity:.88;font-size:13px;margin-top:3px}.live-clock{min-width:155px;text-align:center;background:#ffffff18;border:1px solid #ffffff35;border-radius:14px;padding:8px 13px;box-shadow:0 6px 18px #0002}.live-clock-time{font-size:25px;font-weight:900;letter-spacing:2px;line-height:1.1}.live-clock-date{font-size:11px;opacity:.92;margin-top:4px;white-space:nowrap}.header-actions{display:flex;flex-wrap:wrap;justify-content:flex-end;gap:6px}.header-actions a{display:inline-flex;text-decoration:none}.header-actions button{white-space:nowrap}h1{margin:0;font-size:27px}.wrap{max-width:1400px;margin:auto;padding:16px}
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
@media(max-width:1150px){.top{grid-template-columns:1fr auto}.header-actions{grid-column:1/-1;justify-content:flex-start}.cards{grid-template-columns:repeat(3,1fr)}.grid{grid-template-columns:1fr}}
@media(max-width:600px){.top{display:flex;flex-direction:column;align-items:stretch}.brand{min-width:0}.brand-logo{width:88px;height:58px}.brand-text h1{font-size:22px}.live-clock{width:100%}.header-actions{justify-content:flex-start}.cards{grid-template-columns:repeat(2,1fr)}header{position:static;padding:13px}.wrap{padding:9px}}
</style>
</head>
<body>
<header><div class="top">
<div class="brand">
  <img class="brand-logo" src="data:image/jpeg;base64,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ZdVmJ36wSATABEgAABJSVIAAAExMDmnSua05+f1U18vnzj8h4K3xe2antldO5VRVfvqmKosiRFRKEiEoRJIkUplMJhACJQppqiFFFdKLdFy2rZ8/o89WE0reNM5+W1h87g8ObCbNrWy9+nU9q1bae7umJjTREwSBEwBEYb5o+mfmenNRGQ2TPLE/RXy11i+vV4NemjhO/wDDsudO0YymOF+hPnr6EvptkTGvTwPSt20nDj6d2PjnY79CYX1mEiJI+bcBn9fx4ux9N5l03XrmHgm3H9GzGLx4+t9K17YNOpxnsvGkc137Qt6z5u7jbtQRBKbPnf6I+daY6x1zkfW65dYmJ26QWAAAAAAAmmYJAAATBMSKebdK5rTDn1dNXL51Xg93grfFbbqW2V07nXTXfvmqJixHjiPTHz74cuX6QcGu2nujhVw7lXw/ZV+nKatdgWRJEAiYIU1UQportlNuqiK2vPf81Yxmn7hqPPzW8HnsBhz4PZNb2P0L9T2zUdu7e6RppCRAETAERifmX6a+Zac/u+hvnn6cR8yUda45Wn0dkuG5+ddJx1jqsZbtw36P+braYH6G+ePoWK7fExr1cF0nddJw4crnNcys2yE+CpPQOl8u6hrvMTC/zhr2xa7jx57N6/lbW9WAy2MrW113im0J79OMyevXHGuzcajPmnu8Prx5NieVfX053V87a3cpprv0x86/RXztTHV+uci65XLrExO3WCUwCYAAAAAAAAAJiYJIHNelc0pz6BVRVy+fcx2Rx1b4ratU2uunda6a9PQmqmqsx4fb4M6cV1Db9O5PHiY6H19HP3d5ttwncMpiJp3uaard0krwkimSSFuIWbemxXc6eD11x7nRpWy1t6rFdlbwajtGq4YU4LNYHHnw2xa5sXfbqm3aht/b3SNNETAABAiMT8yfTnzHTm93078xfTyY4L37BW0+brty3jyZn6G1zbdOq383fSXzdFdf+hfnn6HU24adXBNJ3bSMeLpnZeOdj06UwtomJkiaYfOWvZ/X8eHsnS+ZdOv1rVyLX45zbvfBM+Pdu7/NP0nptc432PjZzXedF3rLDusw17ZEwmJS+dvor52pjqnXeSdbrj1hDbslELVolEkEwAAAAESAAAExMExIp5p0vm1Ofn1VFfL51eOyOOrfF7Xq21027nVTVr31liEaRrGic/L78f7a6cHm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" alt="Pharm Mebel logosi">
  <div class="brand-text">
    <h1>Pharm Mebel V5.1.4</h1>
    <div class="sub">Ishchilar, buyurtmalar, ombor, ishlab chiqarish va moliya</div>
  </div>
</div>
<div class="live-clock">
  <div id="pharmClock" class="live-clock-time">00:00:00</div>
  <div id="pharmDate" class="live-clock-date">Toshkent vaqti</div>
</div>
<div class="header-actions">
  <a href="/pro-boshqaruv"><button style="background:#0f766e">PRO boshqaruv</button></a>
  <a href="/shofyor-boshqaruv"><button style="background:#7c3aed">Shofyor boshqaruvi</button></a>
  <a href="/shartnoma-namuna"><button style="background:#0f766e">Shartnoma Word</button></a>
  <a href="/shartnoma-pdf" target="_blank"><button style="background:#0369a1">Shartnoma PDF</button></a>
  <a href="/backup"><button style="background:#16a34a">Backup</button></a>
  <a href="/logout"><button style="background:#dc2626">Chiqish</button></a>
</div>
</div></header>
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
<option>Ishchi</option><option>Transport</option><option>Seh</option>
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
<option>Seh uchun</option><option>Qarzdorlik yopish</option><option>Favqulodda xarajat</option>
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
<script>
(function(){
  const DAYS=["Yakshanba","Dushanba","Seshanba","Chorshanba","Payshanba","Juma","Shanba"];
  const MONTHS=["yanvar","fevral","mart","aprel","may","iyun","iyul","avgust","sentabr","oktabr","noyabr","dekabr"];
  function two(n){return String(n).padStart(2,"0")}
  function updatePharmClock(){
    const d=new Date(Date.now()+5*60*60*1000);
    const time=two(d.getUTCHours())+":"+two(d.getUTCMinutes())+":"+two(d.getUTCSeconds());
    const date=DAYS[d.getUTCDay()]+", "+d.getUTCDate()+" "+MONTHS[d.getUTCMonth()]+" "+d.getUTCFullYear();
    const clock=document.getElementById("pharmClock");
    const dateBox=document.getElementById("pharmDate");
    if(clock) clock.textContent=time;
    if(dateBox) dateBox.textContent=date+" · Toshkent";
  }
  updatePharmClock();
  setInterval(updatePharmClock,1000);
})();
</script>
</body></html>
"""

# Gunicorn modulni import qilganda ham bazani tayyorlash
init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)
