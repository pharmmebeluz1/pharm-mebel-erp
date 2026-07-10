PHARM MEBEL ERP CLOUD — TAYYOR PAKET

Bu paket internet hostingga joylash uchun tayyorlangan.

Asosiy fayl: app.py
Ishga tushirish: gunicorn --bind 0.0.0.0:$PORT app:app

Muhit sozlamalari:
PHARM_ERP_ADMIN=admin
PHARM_ERP_PASSWORD=mustahkam_parol
PHARM_ERP_SECRET=uzun_maxfiy_kalit
PHARM_ERP_DB=/doimiy_disk/pharm_mebel_erp_pro.db
PORT hosting tomonidan avtomatik beriladi.

MUHIM:
1. SQLite bazasi saqlanishi uchun hostingda doimiy disk/persistent storage kerak.
2. Avvalgi ma'lumotlarni ko'chirish uchun eski pharm_mebel_erp_pro.db faylini data papkasiga qo'ying.
3. Internetda admin/12345 parolini qoldirmang.
4. Hosting ishga tushgach sizga https://... ko'rinishidagi manzil beradi. Shu manzil telefon va kompyuterda ochiladi.
