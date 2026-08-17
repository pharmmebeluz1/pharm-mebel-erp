# Mebel360° — Yulduzchali mijoz baholash

Bu modul telefonga mos Flask + SQLite dastur.

## Asosiy funksiyalar
- 1–5 yulduzcha bilan baholash
- 4–5 yulduz: "Ko‘nglingiz nimadan to‘ldi?"
- 1–3 yulduz: "Biz nimani yaxshilashimiz mumkin?"
- 500 belgigacha izoh
- Har bir buyurtmaga faqat 1 marta baho
- Xodimga avtomatik ball:
  - 1★ = 20
  - 2★ = 40
  - 3★ = 60
  - 4★ = 80
  - 5★ = 100
  - 4–5 yulduzda har bir yaxshi belgi uchun +2 ball
- Rahbar uchun xodimlar reytingi
- Qo‘llab-quvvatlash markazi
- Rahbar uchun qo‘llab-quvvatlash murojaatlari

## Ishga tushirish

1. Python o‘rnating.
2. Terminalda papkaga kiring.
3. Ishga tushiring:

```bash
pip install -r requirements.txt
python app.py
```

4. Brauzerda:
- Mijoz baholashi: http://127.0.0.1:5000/
- Xodimlar reytingi: http://127.0.0.1:5000/rahbar/reytin
- Murojaatlar: http://127.0.0.1:5000/rahbar/murojaatlar

## Mebel360° asosiy dasturiga biriktirish

Hozir bu alohida ishlaydigan tayyor modul.
Asosiy Mebel360° dasturingizdagi:
- `orders`
- `employees`
- mijoz kabineti
- `Yakunlandi` statusi

bilan ulash mumkin.

Muhim:
`ratings.order_id` UNIQUE bo‘lgani uchun mijoz bir buyurtmaga qayta-qayta baho bera olmaydi.

## Call-markaz raqami
`templates/support.html` faylida:
`tel:+998000000000`
joyini haqiqiy telefon raqamiga almashtiring.
