# Karobka: CorelDRAW uchun quti chizmasi (VBA)

Shablonni tanlaysiz va o'lchamlarni kiritasiz. Yoyilma (ko'k — kesish, qizil — bukish) va uning yonida 3D quti darhol sahifada paydo bo'ladi. **Chizish** tugmasi tayyor chizmani guruh qilib qo'yadi.

Talab: CorelDRAW 2020 yoki undan yangi versiya.

## O'rnatish
1. **Code → Download ZIP** orqali yuklab oling va arxivni oching.
2. CorelDRAW'da **Alt+F11** ni bosing. Chapdagi **GlobalMacros** ustida o'ng tugmani bosib, **Import File** ni tanlang va `src` papkadagi 7 ta faylni import qiling:
   `modGeom.bas`, `modTemplates.bas`, `mod3D.bas`, `modDraw.bas`, `modMain.bas`, `clsField.cls`, `frmKarobka.frm`.
   Eski versiya o'rnatilgan bo'lsa, avval eski modullarni o'chiring (o'ng tugma → **Remove**).
3. **Debug → Compile** ni bosing, so'ng **Ctrl+S** bilan saqlang.
4. Ishga tushirish: **Tools → Macros → Karobka**.

> Fayllarni kod oynasiga nusxalab joylamang, "Invalid outside procedure" xatosi chiqadi. Agar Import ishlamasa, `src/joylash/*.txt` dagi toza koddan foydalaning. Uning uchun Insert → Module / Class Module / UserForm qiling, nomini fayl nomi bilan bir xil qo'ying va `.txt` faylning butun matnini joylang.

## 3D
3D quti yoyilmaning o'zidan buklanib hosil bo'ladi, shuning uchun o'lcham o'zgarsa u ham o'zgaradi. Panelning **3D** blokida:

| Boshqaruv | Vazifasi |
|-----------|----------|
| Yig'ish | 0% — tekis karton, 100% — yig'ilgan quti. **Yig'ilishni ko'rsatish** tugmasi animatsiya qiladi |
| Qopqoq | 0% — yopiq, 100% — ochiq. **Ochish / Yopish** tugmasi animatsiya qiladi |
| Aylantirish, Qiyalik, Masshtab | Qutini har tomondan ko'rish |
| Rang | Kraft, Oq, Qora, Qizil, Ko'k, Yashil, Pushti |
| Mockup | Dizayningizni 3D qutiga joylaydi |

**Mockup:**
1. **Chizish** bilan yoyilmani sahifaga qo'ying.
2. Dizaynni yoyilma ustiga **"Dizayn"** nomli qatlamga joylang. Qatlam yo'q bo'lsa, **Mockup** tugmasini birinchi bosganda o'zi yaratiladi.
3. **Mockup** ni bosing: dizayn har bir yuzaga to'g'ri joylanadi. Slayderlarni qimirlatsangiz, oddiy 3D ko'rinishga qaytadi.

Egri shakllar (yostiqcha, silindr, doira quti, sumka-quti, yulduz, tort qutilari, konfet, vstavka, olmos) buklanmaydi, ularda sodda 3D model ko'rsatiladi.

## Shablonlar (339 ta)
Shablonlar 13 guruhga bo'lingan: Gofra · Mailer · Tuck-end · Ruchkali va sumka · Lotok · Lanch-boks · Tort · Ko'p burchak · Piramida · Gulbargli sovg'a · Yostiq va konfet · Kosmetika · Maxsus shakllar.

- Har guruh boshida umumiy shablonlar turadi (66 ta): FEFCO 0201, Mailer, RTE/STE, Lotok, Lanch-boks va boshqalar.
- Keyin `photo.rar` rasmlaridagi chizmalar keladi. Har chizma alohida qator, nomi rasm raqami bilan boshlanadi: `IMG_0311` dagi 1-chizma → **"0311-1 Mailer, osma + qulf"**. `IMG_0311`–`IMG_0466` dagi hamma chizmalar kiritilgan (273 ta).
- Panelda **qidirish** maydoni bor. Rasm raqamini (`0311`) yoki nomdan so'zni (`oyna`, `ruchka`) yozing.

Qaysi rasm qaysi shablonga to'g'ri kelishi [`templates/catalog.md`](templates/catalog.md) da yozilgan.

Qo'shimcha maydonlar 0 bo'lsa, o'sha xususiyat o'chadi: Oyna, Osma, Ruchka, Qulf (tillar soni), Qopqoq quloq, Qiyalik.

## Yangi shablon qo'shish
`src/modTemplates.bas` dagi `InitDefs` ga bitta `AddDef` qatori qo'shing: guruh, nom, quruvchi va standart o'lchamlar. Mavjud quruvchining yangi varianti uchun shu kifoya. Yangi oila bo'lsa, `TBuild` ga yangi quruvchi qo'shing.

Panellar `Rect` / `Poly` bilan chiziladi: ikki panelga umumiy qirra bukish, qolgan qirralar kesish bo'ladi. Qanotlar orasidagi tirqish `Slit` bilan, oyna yoki teshik `CutRect` / `CutPoly` bilan belgilanadi.

3D uchun: `SetRoot x, y, tik` — asosiy panel (yotiq taglik yoki tik devor); `SetFold x1, y1, x2, y2, burchak` — shu chiziqdagi bukish burchagi (standart 90°); `SetLid x1, y1, x2, y2` — qopqoq sharniri. `SetRoot` chaqirilmasa, `Box3` / `Lid3` / `Face` bilan qo'lda yozilgan sodda 3D ishlatiladi.

Koddagi o'zgarishdan keyin:
- `python3 tools/preview.py` — hamma modullarni tekshiradi (lint: e'lon qilinmagan nom, takroriy e'lon, label xatosi) va shablonlarni Corelsiz chizadi: yoyilma, 50% / 100% yig'ilgan va qopqoq ochiq holatlar. VBA kodi `tools/vba2py.py` bilan Python'ga o'giriladi, natija `preview/` papkasida;
- `python3 tools/compare.py RASMLAR CHIQISH 0311 0340` — asl rasmdagi chizma, plagin yoyilmasi va 3D yonma-yon turgan taqqoslash varaqlari;
- `python3 tools/make_paste.py` — `src/joylash` dagi nusxalarni yangilaydi.
