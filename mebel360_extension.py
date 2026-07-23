# -*- coding: utf-8 -*-
"""
Mebel360° xavfsiz qo‘shimcha modul.

Bu fayl app.py ni almashtirmaydi. Gunicorn ishga tushganda gunicorn.conf.py
orqali mavjud Flask ilovasiga quyidagilarni qo‘shadi:
  • buyurtmani ikki bosqichli tasdiq bilan o‘chirish;
  • o‘chirilgan buyurtmaning JSON zaxira nusxasi;
  • har bir buyurtma uchun avans siyosati;
  • mijoz kuzatuv sahifasida avans holati kartasi.
"""
from __future__ import annotations

import json
import os
import secrets
import sqlite3
import sys
import traceback
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple


_EXTENSION_VERSION = "1.0.0"
_PATCHED_ATTR = "_mebel360_order_policy_extension_patched"


def _log(message: str) -> None:
    try:
        print(f"[Mebel360 qo‘shimcha modul] {message}", file=sys.stderr, flush=True)
    except Exception:
        pass


def _safe_identifier(name: str) -> str:
    return '"' + str(name).replace('"', '""') + '"'


def _as_dict(row: Any) -> Dict[str, Any]:
    if row is None:
        return {}
    if isinstance(row, dict):
        return dict(row)
    try:
        return {key: row[key] for key in row.keys()}
    except Exception:
        return {}


def _json_default(value: Any) -> Any:
    if isinstance(value, (bytes, bytearray)):
        return value.hex()
    if isinstance(value, (datetime,)):
        return value.isoformat()
    return str(value)


def _num(value: Any) -> float:
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _first(data: Dict[str, Any], names: Iterable[str], default: Any = "") -> Any:
    lowered = {str(k).lower(): v for k, v in data.items()}
    for name in names:
        if name in data and data[name] not in (None, ""):
            return data[name]
        value = lowered.get(str(name).lower())
        if value not in (None, ""):
            return value
    return default


def _table_names(conn: sqlite3.Connection) -> List[str]:
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
    ).fetchall()
    return [str(_as_dict(r).get("name") or r[0]) for r in rows]


def _columns(conn: sqlite3.Connection, table: str) -> List[str]:
    rows = conn.execute(f"PRAGMA table_info({_safe_identifier(table)})").fetchall()
    result: List[str] = []
    for row in rows:
        data = _as_dict(row)
        result.append(str(data.get("name") if data else row[1]))
    return result


def _connect(module: Any) -> sqlite3.Connection:
    get_db = getattr(module, "get_db", None)
    if callable(get_db):
        conn = get_db()
        try:
            conn.row_factory = sqlite3.Row
        except Exception:
            pass
        try:
            conn.execute("PRAGMA foreign_keys = ON")
        except Exception:
            pass
        return conn

    db_name = getattr(module, "DB_NAME", os.environ.get("PHARM_ERP_DB", "pharm_mebel_erp_pro.db"))
    conn = sqlite3.connect(db_name)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _ensure_tables(module: Any) -> None:
    conn = _connect(module)
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS mebel360_avans_siyosati (
                buyurtma_id INTEGER PRIMARY KEY,
                turi TEXT NOT NULL DEFAULT 'summa',
                talab_summa REAL NOT NULL DEFAULT 2500000,
                talab_foiz REAL NOT NULL DEFAULT 0,
                izoh TEXT NOT NULL DEFAULT '',
                rahbar_tasdiq INTEGER NOT NULL DEFAULT 0,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (buyurtma_id) REFERENCES buyurtmalar(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS mebel360_ochirilgan_buyurtmalar (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                asl_buyurtma_id INTEGER,
                kod TEXT,
                mijoz TEXT,
                sabab TEXT NOT NULL,
                snapshot_json TEXT NOT NULL,
                ochirilgan_vaqt TEXT NOT NULL,
                ochirgan TEXT NOT NULL DEFAULT 'Rahbar'
            );
            """
        )
        conn.commit()
    finally:
        conn.close()


def _all_orders(module: Any) -> List[Dict[str, Any]]:
    conn = _connect(module)
    try:
        rows = conn.execute("SELECT * FROM buyurtmalar ORDER BY id DESC").fetchall()
        policies = {
            int(_as_dict(r).get("buyurtma_id")): _as_dict(r)
            for r in conn.execute("SELECT * FROM mebel360_avans_siyosati").fetchall()
        }
        return [_public_order(module, conn, _as_dict(row), policies.get(int(_as_dict(row).get("id") or 0))) for row in rows]
    finally:
        conn.close()


def _payments_total(conn: sqlite3.Connection, order: Dict[str, Any]) -> float:
    base = _num(_first(order, ["oldindan_tolov", "avans", "tolangan", "jami_tolangan"], 0))
    order_id = int(_first(order, ["id"], 0) or 0)
    if not order_id:
        return base
    try:
        tables = set(_table_names(conn))
        if "buyurtma_tolovlari" not in tables:
            return base
        cols = _columns(conn, "buyurtma_tolovlari")
        if "buyurtma_id" not in cols:
            return base
        amount_col = next(
            (c for c in ["miqdor", "summa", "tolov", "amount", "jami"] if c in cols),
            None,
        )
        if not amount_col:
            return base
        row = conn.execute(
            f"SELECT COALESCE(SUM({_safe_identifier(amount_col)}), 0) AS jami "
            "FROM buyurtma_tolovlari WHERE buyurtma_id=?",
            (order_id,),
        ).fetchone()
        summed = _num(_as_dict(row).get("jami") if row else 0)
        # Eski dastur ayrim versiyalarida oldindan_tolov jami to‘lov bilan yangilanadi,
        # ayrimlarida esa alohida tarix yuritiladi. Ikki marta qo‘shib yubormaslik uchun max olinadi.
        return max(base, summed)
    except Exception:
        return base


def _policy_row(conn: sqlite3.Connection, order_id: int) -> Dict[str, Any]:
    row = conn.execute(
        "SELECT * FROM mebel360_avans_siyosati WHERE buyurtma_id=?", (order_id,)
    ).fetchone()
    if row:
        return _as_dict(row)
    return {
        "buyurtma_id": order_id,
        "turi": "summa",
        "talab_summa": 2500000,
        "talab_foiz": 0,
        "izoh": "",
        "rahbar_tasdiq": 0,
    }


def _policy_status(order: Dict[str, Any], policy: Dict[str, Any], paid: float) -> Dict[str, Any]:
    total = _num(_first(order, ["umumiy_narx", "narx", "summa", "jami_narx"], 0))
    kind = str(policy.get("turi") or "summa").strip().lower()
    amount = max(0.0, _num(policy.get("talab_summa")))
    percent = max(0.0, min(100.0, _num(policy.get("talab_foiz"))))
    approved = bool(int(_num(policy.get("rahbar_tasdiq"))))

    if kind == "ishonchli":
        required = 0.0
        can_start = True
        title = "Doimiy mijoz — avans talab qilinmaydi"
        message = (
            "Siz doimiy va ishonchli mijoz sifatida belgilandingiz. "
            "Ushbu buyurtma uchun avans talab qilinmaydi va ishni boshlash mumkin."
        )
        badge = "success"
    elif kind == "rahbar":
        required = 0.0
        can_start = approved
        if approved:
            title = "Rahbar ruxsati berildi"
            message = "Avanssiz ish boshlash rahbar tomonidan tasdiqlandi."
            badge = "success"
        else:
            title = "Rahbar tasdig‘i kutilmoqda"
            message = "Ish boshlanishi uchun rahbarning avanssiz ishlash tasdig‘i kutilmoqda."
            badge = "waiting"
    else:
        if kind == "foiz":
            required = round(total * percent / 100.0, 2)
        else:
            required = amount
        remaining = max(0.0, required - paid)
        can_start = paid >= required and required >= 0
        if can_start:
            title = "Avans yetarli — ishni boshlash mumkin"
            message = "Belgilangan avans miqdori yetarli. Buyurtma ishlab chiqarishga tayyor."
            badge = "success"
        else:
            title = "Avans hali yetarli emas"
            message = f"Ishni boshlash uchun yana {remaining:,.0f} so‘m avans kerak."
            badge = "waiting"

    remaining = max(0.0, required - paid)
    return {
        "turi": kind,
        "required": required,
        "paid": paid,
        "remaining": remaining,
        "can_start": can_start,
        "title": title,
        "message": message,
        "badge": badge,
        "percent": percent,
        "note": str(policy.get("izoh") or ""),
        "rahbar_tasdiq": approved,
    }


def _public_order(module: Any, conn: sqlite3.Connection, order: Dict[str, Any], policy: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    order_id = int(_first(order, ["id"], 0) or 0)
    policy = policy or _policy_row(conn, order_id)
    paid = _payments_total(conn, order)
    status = _policy_status(order, policy, paid)
    return {
        "id": order_id,
        "kod": str(_first(order, ["kod", "buyurtma_kodi", "code"], "")),
        "mijoz": str(_first(order, ["mijoz", "mijoz_ismi", "client"], "")),
        "mahsulot": str(_first(order, ["mahsulot", "nomi", "product"], "")),
        "umumiy_narx": _num(_first(order, ["umumiy_narx", "narx", "summa", "jami_narx"], 0)),
        "tolangan": paid,
        "policy": {
            "turi": str(policy.get("turi") or "summa"),
            "talab_summa": _num(policy.get("talab_summa")),
            "talab_foiz": _num(policy.get("talab_foiz")),
            "izoh": str(policy.get("izoh") or ""),
            "rahbar_tasdiq": bool(int(_num(policy.get("rahbar_tasdiq")))),
        },
        "status": status,
    }


def _find_order(module: Any, *, order_id: Optional[int] = None, kod: str = "", token: str = "", q: str = "") -> Optional[Dict[str, Any]]:
    conn = _connect(module)
    try:
        if order_id:
            row = conn.execute("SELECT * FROM buyurtmalar WHERE id=?", (order_id,)).fetchone()
            return _as_dict(row) if row else None

        if kod:
            row = conn.execute("SELECT * FROM buyurtmalar WHERE TRIM(kod)=TRIM(?)", (kod,)).fetchone()
            if row:
                return _as_dict(row)

        token = (token or "").strip()
        if token:
            order_cols = _columns(conn, "buyurtmalar")
            candidate_cols = [
                c for c in order_cols
                if any(word in c.lower() for word in ("token", "havola", "link", "qr", "kuzatuv"))
            ]
            for col in candidate_cols:
                try:
                    row = conn.execute(
                        f"SELECT * FROM buyurtmalar WHERE {_safe_identifier(col)}=? LIMIT 1", (token,)
                    ).fetchone()
                    if row:
                        return _as_dict(row)
                except Exception:
                    continue

            # Token alohida jadvalda saqlanadigan versiyalar uchun umumiy qidiruv.
            for table in _table_names(conn):
                if table in {"buyurtmalar", "mebel360_ochirilgan_buyurtmalar"}:
                    continue
                cols = _columns(conn, table)
                if "buyurtma_id" not in cols:
                    continue
                token_cols = [
                    c for c in cols
                    if any(word in c.lower() for word in ("token", "havola", "link", "qr", "kuzatuv"))
                ]
                for col in token_cols:
                    try:
                        found = conn.execute(
                            f"SELECT buyurtma_id FROM {_safe_identifier(table)} "
                            f"WHERE {_safe_identifier(col)}=? LIMIT 1",
                            (token,),
                        ).fetchone()
                        if found:
                            found_id = int(_as_dict(found).get("buyurtma_id") or found[0])
                            row = conn.execute("SELECT * FROM buyurtmalar WHERE id=?", (found_id,)).fetchone()
                            if row:
                                return _as_dict(row)
                    except Exception:
                        continue

        q_norm = " ".join((q or "").upper().split())
        if q_norm:
            rows = conn.execute("SELECT * FROM buyurtmalar ORDER BY LENGTH(kod) DESC").fetchall()
            for row in rows:
                data = _as_dict(row)
                code = " ".join(str(_first(data, ["kod"], "")).upper().split())
                if code and code in q_norm:
                    return data
        return None
    finally:
        conn.close()


def _snapshot_related(conn: sqlite3.Connection, order: Dict[str, Any]) -> Dict[str, Any]:
    order_id = int(_first(order, ["id"], 0) or 0)
    order_code = str(_first(order, ["kod", "buyurtma_kodi"], ""))
    snapshot: Dict[str, Any] = {"buyurtmalar": [order], "related": {}}
    for table in _table_names(conn):
        if table in {"buyurtmalar", "mebel360_ochirilgan_buyurtmalar"}:
            continue
        try:
            cols = _columns(conn, table)
            rows: List[Any] = []
            if "buyurtma_id" in cols:
                rows = conn.execute(
                    f"SELECT * FROM {_safe_identifier(table)} WHERE buyurtma_id=?", (order_id,)
                ).fetchall()
            elif order_code and "buyurtma_kodi" in cols:
                rows = conn.execute(
                    f"SELECT * FROM {_safe_identifier(table)} WHERE buyurtma_kodi=?", (order_code,)
                ).fetchall()
            if rows:
                snapshot["related"][table] = [_as_dict(r) for r in rows]
        except Exception as exc:
            snapshot["related"][table] = {"snapshot_error": str(exc)}
    return snapshot


def _delete_order(module: Any, order_id: int, exact_code: str, reason: str) -> Tuple[bool, str, Optional[int]]:
    conn = _connect(module)
    try:
        row = conn.execute("SELECT * FROM buyurtmalar WHERE id=?", (order_id,)).fetchone()
        if not row:
            return False, "Buyurtma topilmadi.", None
        order = _as_dict(row)
        code = str(_first(order, ["kod", "buyurtma_kodi"], "")).strip()
        if exact_code.strip() != code:
            return False, "Tasdiqlash kodi buyurtma kodiga aynan mos emas.", None
        if len(reason.strip()) < 4:
            return False, "O‘chirish sababini kamida 4 ta harf bilan yozing.", None

        snapshot = _snapshot_related(conn, order)
        cur = conn.execute(
            """
            INSERT INTO mebel360_ochirilgan_buyurtmalar
                (asl_buyurtma_id, kod, mijoz, sabab, snapshot_json, ochirilgan_vaqt, ochirgan)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                order_id,
                code,
                str(_first(order, ["mijoz", "mijoz_ismi"], "")),
                reason.strip(),
                json.dumps(snapshot, ensure_ascii=False, default=_json_default),
                datetime.now().isoformat(timespec="seconds"),
                "Rahbar",
            ),
        )
        archive_id = int(cur.lastrowid)

        # FK bo‘lmagan eski jadvallar ham qolib ketmasligi uchun buyurtma_id bo‘yicha qo‘lda tozalanadi.
        for table in _table_names(conn):
            if table in {"buyurtmalar", "mebel360_ochirilgan_buyurtmalar"}:
                continue
            try:
                cols = _columns(conn, table)
                if "buyurtma_id" in cols:
                    conn.execute(
                        f"DELETE FROM {_safe_identifier(table)} WHERE buyurtma_id=?", (order_id,)
                    )
            except Exception:
                # Muhim moliyaviy tarixni tasodifan buzmaslik uchun xatoda davom etiladi;
                # asosiy DELETE FK cheklovi bilan xavfsiz to‘xtaydi.
                continue

        conn.execute("DELETE FROM buyurtmalar WHERE id=?", (order_id,))
        conn.commit()
        return True, "Buyurtma o‘chirildi. Zaxira nusxasi saqlandi.", archive_id
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return False, f"Buyurtma bog‘langan ma’lumotlar sabab o‘chirilmadi: {exc}", None
    except Exception as exc:
        conn.rollback()
        return False, f"O‘chirishda xato: {exc}", None
    finally:
        conn.close()


_DASHBOARD_INJECTION = r'''
<style id="m360-ext-style">
#m360ManageBtn{position:fixed;right:20px;bottom:20px;z-index:99990;border:0;border-radius:18px;padding:14px 18px;background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#fff;font:800 14px Arial;box-shadow:0 14px 35px rgba(79,70,229,.38);cursor:pointer}
#m360ManageBtn:hover{transform:translateY(-1px)}
#m360Overlay{position:fixed;inset:0;z-index:99998;background:rgba(15,23,42,.58);backdrop-filter:blur(7px);display:none;align-items:center;justify-content:center;padding:18px;font-family:Arial,sans-serif}
#m360Overlay.open{display:flex}
#m360Panel{width:min(1050px,100%);max-height:92vh;overflow:auto;background:#f8fafc;border-radius:26px;box-shadow:0 28px 80px rgba(15,23,42,.35)}
#m360Head{position:sticky;top:0;z-index:2;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:20px 22px;background:linear-gradient(135deg,#111827,#312e81);color:#fff;border-radius:26px 26px 0 0}
#m360Head h2{margin:0;font-size:22px}#m360Head p{margin:5px 0 0;color:#c7d2fe;font-size:13px}
#m360Close{border:0;border-radius:12px;width:42px;height:42px;background:rgba(255,255,255,.14);color:#fff;font-size:24px;cursor:pointer}
#m360Body{padding:18px}.m360-order{background:#fff;border:1px solid #e2e8f0;border-radius:20px;padding:17px;margin-bottom:14px;box-shadow:0 8px 24px rgba(15,23,42,.06)}
.m360-top{display:flex;justify-content:space-between;gap:14px;align-items:flex-start}.m360-code{font-size:18px;font-weight:900;color:#172554}.m360-client{margin-top:4px;color:#475569;font-weight:700}.m360-money{text-align:right;color:#334155;font-size:13px;line-height:1.55}.m360-status{margin:13px 0;padding:11px 13px;border-radius:13px;font-weight:800;font-size:13px}.m360-status.success{background:#dcfce7;color:#166534}.m360-status.waiting{background:#fff7ed;color:#9a3412}
.m360-grid{display:grid;grid-template-columns:1.2fr 1fr 1fr;gap:10px}.m360-field label{display:block;margin:0 0 5px;color:#475569;font-size:11px;font-weight:900;text-transform:uppercase}.m360-field input,.m360-field select{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:11px;padding:10px 11px;background:#fff;color:#0f172a;font-weight:700}.m360-note{grid-column:1/-1}.m360-actions{display:flex;flex-wrap:wrap;gap:9px;margin-top:12px}.m360-save,.m360-delete,.m360-confirm,.m360-cancel{border:0;border-radius:11px;padding:10px 14px;font-weight:900;cursor:pointer}.m360-save{background:#2563eb;color:#fff}.m360-delete{background:#fee2e2;color:#b91c1c}.m360-confirm{background:#dc2626;color:#fff}.m360-cancel{background:#e2e8f0;color:#334155}.m360-deletebox{display:none;margin-top:12px;padding:13px;border:1px solid #fecaca;background:#fff1f2;border-radius:14px}.m360-deletebox.open{display:block}.m360-deletebox p{margin:0 0 9px;color:#9f1239;font-size:13px}.m360-deletebox input{width:100%;box-sizing:border-box;margin:5px 0;border:1px solid #fda4af;border-radius:10px;padding:10px}.m360-empty{text-align:center;padding:35px;color:#64748b}
@media(max-width:700px){#m360ManageBtn{right:10px;bottom:10px;padding:12px 14px}.m360-grid{grid-template-columns:1fr}.m360-note{grid-column:auto}.m360-top{display:block}.m360-money{text-align:left;margin-top:8px}#m360Panel{max-height:96vh;border-radius:18px}#m360Head{border-radius:18px 18px 0 0}}
</style>
<button id="m360ManageBtn" type="button">🧾 Buyurtma boshqaruvi</button>
<div id="m360Overlay" aria-hidden="true"><section id="m360Panel"><header id="m360Head"><div><h2>Buyurtmalar va avans shartlari</h2><p>Avans talabini belgilang yoki buyurtmani xavfsiz o‘chiring.</p></div><button id="m360Close" type="button">×</button></header><div id="m360Body"><div class="m360-empty">Yuklanmoqda…</div></div></section></div>
<script id="m360-ext-script">
(()=>{'use strict';
const TOKEN=__M360_TOKEN__;
const overlay=document.getElementById('m360Overlay'),body=document.getElementById('m360Body');
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const money=n=>new Intl.NumberFormat('uz-UZ').format(Number(n||0))+' so‘m';
async function api(url,opt={}){opt.headers=Object.assign({'X-Mebel360-Token':TOKEN,'Content-Type':'application/json'},opt.headers||{});const r=await fetch(url,opt);const j=await r.json().catch(()=>({ok:false,message:'Server javobi o‘qilmadi.'}));if(!r.ok||j.ok===false)throw new Error(j.message||'Xatolik');return j;}
function typeOptions(v){return [['summa','Aniq summa bo‘yicha'],['foiz','Foiz bo‘yicha'],['ishonchli','Doimiy mijoz — avanssiz'],['rahbar','Rahbar ruxsati bilan']].map(x=>`<option value="${x[0]}" ${x[0]===v?'selected':''}>${x[1]}</option>`).join('');}
function card(o){const p=o.policy,s=o.status;return `<article class="m360-order" data-id="${o.id}"><div class="m360-top"><div><div class="m360-code">${esc(o.kod||('ID-'+o.id))}</div><div class="m360-client">${esc(o.mijoz)} · ${esc(o.mahsulot)}</div></div><div class="m360-money"><b>${money(o.umumiy_narx)}</b><br>To‘langan: ${money(o.tolangan)}</div></div><div class="m360-status ${esc(s.badge)}">${esc(s.title)}<div style="font-weight:500;margin-top:4px">${esc(s.message)}</div></div><div class="m360-grid"><div class="m360-field"><label>Avans tartibi</label><select class="m360-kind">${typeOptions(p.turi)}</select></div><div class="m360-field"><label>Talab summa</label><input class="m360-amount" type="number" min="0" value="${Number(p.talab_summa||0)}"></div><div class="m360-field"><label>Talab foiz</label><input class="m360-percent" type="number" min="0" max="100" value="${Number(p.talab_foiz||0)}"></div><div class="m360-field m360-note"><label>Mijozga izoh</label><input class="m360-note-input" value="${esc(p.izoh||'')}" placeholder="Masalan: kelishuv bo‘yicha"></div><div class="m360-field"><label>Rahbar tasdig‘i</label><select class="m360-approved"><option value="0" ${p.rahbar_tasdiq?'':'selected'}>Kutilmoqda</option><option value="1" ${p.rahbar_tasdiq?'selected':''}>Tasdiqlandi</option></select></div></div><div class="m360-actions"><button class="m360-save" type="button">Avans shartini saqlash</button><button class="m360-delete" type="button">Buyurtmani o‘chirish</button></div><div class="m360-deletebox"><p><b>Diqqat:</b> buyurtma ko‘rinishdan o‘chadi, lekin to‘liq zaxira nusxasi bazada saqlanadi. Tasdiqlash uchun kodni aynan yozing.</p><input class="m360-code-confirm" placeholder="${esc(o.kod)}"><input class="m360-reason" placeholder="O‘chirish sababi"><div class="m360-actions"><button class="m360-confirm" type="button">Ha, o‘chirilsin</button><button class="m360-cancel" type="button">Bekor qilish</button></div></div></article>`;}
async function load(){body.innerHTML='<div class="m360-empty">Yuklanmoqda…</div>';try{const j=await api('/mebel360-ext/orders');body.innerHTML=j.orders.length?j.orders.map(card).join(''):'<div class="m360-empty">Buyurtmalar topilmadi.</div>';}catch(e){body.innerHTML='<div class="m360-empty">'+esc(e.message)+'</div>';}}
document.getElementById('m360ManageBtn').onclick=()=>{overlay.classList.add('open');overlay.setAttribute('aria-hidden','false');load();};document.getElementById('m360Close').onclick=()=>{overlay.classList.remove('open');overlay.setAttribute('aria-hidden','true');};overlay.addEventListener('click',e=>{if(e.target===overlay)document.getElementById('m360Close').click();});
body.addEventListener('click',async e=>{const order=e.target.closest('.m360-order');if(!order)return;const id=order.dataset.id;if(e.target.closest('.m360-delete')){order.querySelector('.m360-deletebox').classList.add('open');return;}if(e.target.closest('.m360-cancel')){order.querySelector('.m360-deletebox').classList.remove('open');return;}if(e.target.closest('.m360-save')){const payload={turi:order.querySelector('.m360-kind').value,talab_summa:Number(order.querySelector('.m360-amount').value||0),talab_foiz:Number(order.querySelector('.m360-percent').value||0),izoh:order.querySelector('.m360-note-input').value,rahbar_tasdiq:Number(order.querySelector('.m360-approved').value)};try{await api('/mebel360-ext/policy/'+id,{method:'POST',body:JSON.stringify(payload)});await load();alert('Avans sharti saqlandi.');}catch(err){alert(err.message);}return;}if(e.target.closest('.m360-confirm')){const code=order.querySelector('.m360-code-confirm').value,reason=order.querySelector('.m360-reason').value;try{const j=await api('/mebel360-ext/delete/'+id,{method:'POST',body:JSON.stringify({kod:code,sabab:reason})});alert(j.message);await load();setTimeout(()=>location.reload(),400);}catch(err){alert(err.message);}}});
})();
</script>
'''

_CLIENT_INJECTION = r'''
<style id="m360-client-policy-style">
#m360ClientPolicy{margin:18px auto;width:min(900px,calc(100% - 24px));box-sizing:border-box;border-radius:22px;padding:18px 20px;background:linear-gradient(135deg,#eef2ff,#f5f3ff);border:1px solid #c7d2fe;box-shadow:0 14px 35px rgba(79,70,229,.12);font-family:Arial,sans-serif;color:#172554}
#m360ClientPolicy .mcp-head{display:flex;gap:12px;align-items:center}.mcp-icon{width:46px;height:46px;border-radius:15px;display:grid;place-items:center;background:#4f46e5;color:#fff;font-size:23px}.mcp-title{font-size:18px;font-weight:900}.mcp-sub{margin-top:3px;color:#475569;font-size:13px}.mcp-badge{display:inline-block;margin:13px 0 9px;padding:8px 11px;border-radius:999px;font-weight:900;font-size:12px}.mcp-badge.success{background:#dcfce7;color:#166534}.mcp-badge.waiting{background:#ffedd5;color:#9a3412}.mcp-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:9px;margin-top:8px}.mcp-cell{background:rgba(255,255,255,.8);border:1px solid #e0e7ff;border-radius:13px;padding:11px}.mcp-cell small{display:block;color:#64748b;font-weight:800}.mcp-cell b{display:block;margin-top:4px;font-size:15px}.mcp-note{margin-top:10px;color:#475569;font-size:13px;line-height:1.5}
@media(max-width:650px){.mcp-grid{grid-template-columns:1fr}#m360ClientPolicy{padding:15px}}
</style>
<script id="m360-client-policy-script">
(()=>{'use strict';
const money=n=>new Intl.NumberFormat('uz-UZ').format(Number(n||0))+' so‘m';
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
async function run(){const token=decodeURIComponent(location.pathname.split('/').filter(Boolean).pop()||'');const heading=[...document.querySelectorAll('h1,h2,h3,.order-title,.hero-title')].map(x=>x.textContent.trim()).find(Boolean)||document.body.innerText.slice(0,700);try{const r=await fetch('/mebel360-ext/client-policy?token='+encodeURIComponent(token)+'&q='+encodeURIComponent(heading));if(!r.ok)return;const j=await r.json();if(!j.ok||!j.order)return;const o=j.order,s=o.status;const el=document.createElement('section');el.id='m360ClientPolicy';el.innerHTML=`<div class="mcp-head"><div class="mcp-icon">₸</div><div><div class="mcp-title">Ish boshlash uchun avans holati</div><div class="mcp-sub">${esc(o.kod)} · ${esc(o.mijoz)}</div></div></div><span class="mcp-badge ${esc(s.badge)}">${esc(s.title)}</span><div class="mcp-grid"><div class="mcp-cell"><small>Talab qilinadigan avans</small><b>${money(s.required)}</b></div><div class="mcp-cell"><small>To‘langan</small><b>${money(s.paid)}</b></div><div class="mcp-cell"><small>Ish boshlash uchun qoldi</small><b>${money(s.remaining)}</b></div></div><div class="mcp-note">${esc(s.message)}${s.note?'<br><b>Izoh:</b> '+esc(s.note):''}</div>`;const target=document.querySelector('main')||document.querySelector('.container')||document.body;const first=target.firstElementChild;if(first&&target!==document.body)first.insertAdjacentElement('afterend',el);else target.prepend(el);}catch(e){/* mijoz sahifasini buzmaslik uchun jim qoladi */}}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',run);else run();
})();
</script>
'''


def _inject_html(response: Any, fragment: str) -> Any:
    try:
        content_type = str(response.headers.get("Content-Type", ""))
        if "text/html" not in content_type.lower() or response.status_code != 200:
            return response
        html = response.get_data(as_text=True)
        if not html or fragment[:30] in html:
            return response
        pos = html.lower().rfind("</body>")
        if pos >= 0:
            html = html[:pos] + fragment + html[pos:]
        else:
            html += fragment
        response.set_data(html)
        response.headers["Content-Length"] = str(len(response.get_data()))
    except Exception:
        _log("HTML qo‘shimchasini joylashda xato:\n" + traceback.format_exc())
    return response


def patch_app(module: Any) -> None:
    if getattr(module, _PATCHED_ATTR, False):
        return
    flask_app = getattr(module, "app", None)
    if flask_app is None or not hasattr(flask_app, "route"):
        return

    from flask import jsonify, request, session

    try:
        _ensure_tables(module)
    except Exception:
        _log("Qo‘shimcha jadvallar yaratilmadi:\n" + traceback.format_exc())
        return

    setattr(module, _PATCHED_ATTR, True)

    def _has_rahbar_session() -> bool:
        role = str(session.get("user_role") or session.get("role") or "").strip().lower()
        # Mebel360 ning eski versiyasi logged_in/user, yangi versiyasi esa
        # admin_user_id/user_role kalitlaridan foydalanadi. Ikkalasi ham qo‘llanadi.
        if bool(session.get("logged_in")) and not session.get("staff_role"):
            return True
        if session.get("admin_user_id") and role in {"", "admin", "rahbar", "owner"}:
            return True
        if role in {"admin", "rahbar", "owner"} and (session.get("user") or session.get("admin_username")):
            return True
        return False

    @flask_app.before_request
    def _m360_ext_dashboard_token() -> None:
        if request.path.rstrip("/") == "/dashboard" and _has_rahbar_session():
            session.setdefault("_m360_ext_token", secrets.token_urlsafe(24))

    def _authorized() -> bool:
        sent = request.headers.get("X-Mebel360-Token", "") or request.args.get("_token", "")
        saved = str(session.get("_m360_ext_token", ""))
        return bool(_has_rahbar_session() and sent and saved and secrets.compare_digest(str(sent), saved))

    @flask_app.route("/mebel360-ext/health", methods=["GET"], endpoint="m360_ext_health")
    def _m360_ext_health():
        return jsonify({"ok": True, "version": _EXTENSION_VERSION})

    @flask_app.route("/mebel360-ext/orders", methods=["GET"], endpoint="m360_ext_orders")
    def _m360_ext_orders():
        if not _authorized():
            return jsonify({"ok": False, "message": "Ruxsat yo‘q."}), 403
        try:
            return jsonify({"ok": True, "orders": _all_orders(module)})
        except Exception as exc:
            return jsonify({"ok": False, "message": f"Buyurtmalarni olishda xato: {exc}"}), 500

    @flask_app.route("/mebel360-ext/policy/<int:order_id>", methods=["GET", "POST"], endpoint="m360_ext_policy")
    def _m360_ext_policy(order_id: int):
        if not _authorized():
            return jsonify({"ok": False, "message": "Ruxsat yo‘q."}), 403
        conn = _connect(module)
        try:
            order_row = conn.execute("SELECT * FROM buyurtmalar WHERE id=?", (order_id,)).fetchone()
            if not order_row:
                return jsonify({"ok": False, "message": "Buyurtma topilmadi."}), 404
            if request.method == "POST":
                payload = request.get_json(silent=True) or {}
                kind = str(payload.get("turi") or "summa").strip().lower()
                if kind not in {"summa", "foiz", "ishonchli", "rahbar"}:
                    return jsonify({"ok": False, "message": "Avans tartibi noto‘g‘ri."}), 400
                amount = max(0.0, _num(payload.get("talab_summa")))
                percent = max(0.0, min(100.0, _num(payload.get("talab_foiz"))))
                note = str(payload.get("izoh") or "").strip()[:500]
                approved = 1 if bool(int(_num(payload.get("rahbar_tasdiq")))) else 0
                conn.execute(
                    """
                    INSERT INTO mebel360_avans_siyosati
                        (buyurtma_id, turi, talab_summa, talab_foiz, izoh, rahbar_tasdiq, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(buyurtma_id) DO UPDATE SET
                        turi=excluded.turi,
                        talab_summa=excluded.talab_summa,
                        talab_foiz=excluded.talab_foiz,
                        izoh=excluded.izoh,
                        rahbar_tasdiq=excluded.rahbar_tasdiq,
                        updated_at=CURRENT_TIMESTAMP
                    """,
                    (order_id, kind, amount, percent, note, approved),
                )
                conn.commit()
            policy = _policy_row(conn, order_id)
            order = _public_order(module, conn, _as_dict(order_row), policy)
            return jsonify({"ok": True, "order": order})
        except Exception as exc:
            conn.rollback()
            return jsonify({"ok": False, "message": f"Avans shartini saqlashda xato: {exc}"}), 500
        finally:
            conn.close()

    @flask_app.route("/mebel360-ext/delete/<int:order_id>", methods=["POST"], endpoint="m360_ext_delete")
    def _m360_ext_delete(order_id: int):
        if not _authorized():
            return jsonify({"ok": False, "message": "Ruxsat yo‘q."}), 403
        payload = request.get_json(silent=True) or {}
        ok, message, archive_id = _delete_order(
            module,
            order_id,
            str(payload.get("kod") or ""),
            str(payload.get("sabab") or ""),
        )
        return jsonify({"ok": ok, "message": message, "archive_id": archive_id}), (200 if ok else 400)

    @flask_app.route("/mebel360-ext/client-policy", methods=["GET"], endpoint="m360_ext_client_policy")
    def _m360_ext_client_policy():
        order = _find_order(
            module,
            kod=str(request.args.get("kod") or ""),
            token=str(request.args.get("token") or ""),
            q=str(request.args.get("q") or "")[:1200],
        )
        if not order:
            return jsonify({"ok": False, "message": "Buyurtma topilmadi."}), 404
        conn = _connect(module)
        try:
            public = _public_order(module, conn, order, _policy_row(conn, int(order.get("id") or 0)))
            return jsonify({"ok": True, "order": public})
        finally:
            conn.close()

    @flask_app.after_request
    def _m360_ext_after_request(response: Any) -> Any:
        try:
            path = request.path.rstrip("/")
            if path == "/dashboard":
                token = str(session.get("_m360_ext_token", ""))
                if token:
                    fragment = _DASHBOARD_INJECTION.replace("__M360_TOKEN__", json.dumps(token))
                    response = _inject_html(response, fragment)
            elif request.path.startswith("/kuzatuv/"):
                response = _inject_html(response, _CLIENT_INJECTION)
        except Exception:
            _log("Sahifaga modul qo‘shishda xato:\n" + traceback.format_exc())
        return response

    _log(f"{_EXTENSION_VERSION} versiya muvaffaqiyatli ulandi.")


__all__ = ["patch_app"]
