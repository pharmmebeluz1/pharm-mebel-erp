# Mebel360° — telefon va kompyuterga o‘rnatiladigan PWA

Ushbu paket Render/GitHub uchun ishlaydigan Flask + PWA boshlang‘ich qobig‘idir.

## Muhim
Yuborilgan ZIPdagi `app.py` asl ERP web-ilovasi emas, Excel modulini o‘rnatish skripti edi. U yo‘qolmasligi uchun `excel_module_installer.py` nomida saqlandi. Ushbu paketda ishlaydigan web/PWA qobig‘i yangi `app.py` sifatida qo‘shildi.

## Render
- Build: `pip install -r requirements.txt`
- Start: `gunicorn app:app --workers 1 --threads 4 --timeout 120`
- Health: `/api/health`

## Telefonga o‘rnatish
1. Render manzilini Android Chrome’da oching.
2. Menyudan **Ilovani o‘rnatish / Установить приложение** ni bosing.
3. Eski Pharm Mebel yorlig‘i bo‘lsa, avval o‘chirib tashlang.

## Nom va logo
- Ilova nomi: **Mebel360°**
- PWA manifest, brauzer sarlavhasi va ikonkalari yangilangan.
