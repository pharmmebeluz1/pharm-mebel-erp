# -*- coding: utf-8 -*-
"""Mebel360° Excel moduli — avtomatik o‘rnatgich."""
from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path


MARKER_START = "# MEBEL360_EXCEL_MODULE_START"
MARKER_END = "# MEBEL360_EXCEL_MODULE_END"
INJECTION = f"""
{MARKER_START}
try:
    from mebel360_excel_modul import register_excel_module
    register_excel_module(app, get_db)
except Exception as _m360_excel_error:
    print("Mebel360 Excel modul xatosi:", _m360_excel_error)
{MARKER_END}

"""


def main() -> int:
    folder = Path(__file__).resolve().parent
    app_path = folder / "app.py"
    module_path = folder / "mebel360_excel_modul.py"

    print("=" * 62)
    print(" Mebel360° — Excel moliyaviy hisobot modulini o‘rnatish")
    print("=" * 62)

    if not app_path.exists():
        print("XATO: app.py topilmadi.")
        print("Bu fayllarni app.py turgan papkaga qo‘ying.")
        return 1
    if not module_path.exists():
        print("XATO: mebel360_excel_modul.py topilmadi.")
        return 1

    # openpyxl o‘rnatiladi.
    try:
        import openpyxl  # noqa: F401
        print("openpyxl mavjud.")
    except ImportError:
        print("openpyxl o‘rnatilmoqda...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl>=3.1.2"])
        except Exception as exc:
            print("XATO: openpyxl o‘rnatilmadi:", exc)
            print("Internetni tekshirib, yana ishga tushiring.")
            return 1

    text = app_path.read_text(encoding="utf-8-sig")
    if MARKER_START in text:
        print("Excel moduli oldin o‘rnatilgan. Qayta o‘rnatish shart emas.")
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = folder / f"app_excel_oldidan_{stamp}.py"
        shutil.copy2(app_path, backup)
        print("Zaxira nusxa:", backup.name)

        import re
        match = re.search(
            r"(?m)^\s*if\s+__name__\s*==\s*(['\"])__main__\1\s*:",
            text,
        )
        if not match:
            print("XATO: app.py oxiridagi ishga tushirish qismi topilmadi.")
            print("app.py o‘zgartirilmadi.")
            return 1

        text = text[:match.start()] + INJECTION + text[match.start():]
        app_path.write_text(text, encoding="utf-8")
        print("app.py muvaffaqiyatli yangilandi.")

    # Render va boshqa serverlar uchun requirements.txt yangilanadi.
    req_path = folder / "requirements.txt"
    req_text = req_path.read_text(encoding="utf-8") if req_path.exists() else ""
    if "openpyxl" not in req_text.lower():
        if req_text and not req_text.endswith("\n"):
            req_text += "\n"
        req_text += "openpyxl>=3.1.2\n"
        req_path.write_text(req_text, encoding="utf-8")
        print("requirements.txt ga openpyxl qo‘shildi.")

    print()
    print("TAYYOR!")
    print("1) Mebel360° dasturini qayta ishga tushiring.")
    print("2) Rahbar kabinetiga kiring.")
    print("3) Past chapdagi «📊 Excel hisobot» tugmasini bosing.")
    print("4) Sana oralig‘ini tanlab Excelni yuklang.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
