# Karobka: CorelDRAW uchun quti chizmasi (VBA)

Shablonni tanlaysiz va o'lchamlarni kiritasiz. Yoyilma (ko'k — kesish, qizil — bukish) va yig'ilgan qutining sodda 3D ko'rinishi darhol sahifada paydo bo'ladi. **Chizish** tugmasi tayyor chizmani guruh qilib qo'yadi.

Talab: CorelDRAW 2020 yoki undan yangi versiya.

## O'rnatish
1. **Code → Download ZIP** orqali yuklab oling va arxivni oching.
2. CorelDRAW'da **Alt+F11** ni bosing. Chapdagi **GlobalMacros** ustida o'ng tugmani bosib, **Import File** ni tanlang va `src` papkadagi 7 ta faylni import qiling:
   `modGeom.bas`, `modTemplates.bas`, `mod3D.bas`, `modDraw.bas`, `modMain.bas`, `clsField.cls`, `frmKarobka.frm`.
   Eski versiya o'rnatilgan bo'lsa, avval eski modullarni o'chiring (o'ng tugma → **Remove**).
3. **Debug → Compile** ni bosing, so'ng **Ctrl+S** bilan saqlang.
4. Ishga tushirish: **Tools → Macros → Karobka**.

> Fayllarni kod oynasiga nusxalab joylamang, "Invalid outside procedure" xatosi chiqadi. Agar Import ishlamasa, `src/joylash/*.txt` dagi toza koddan foydalaning. Uning uchun Insert → Module / Class Module / UserForm qiling, nomini fayl nomi bilan bir xil qo'ying va `.txt` faylning butun matnini joylang.

## Shablonlar (13 ta)
| Guruh | Shablonlar |
|-------|-----------|
| Gofra | FEFCO 0201, FEFCO 0200 (ochiq) |
| Tuck-end | Teskari qopqoq, To'g'ri qopqoq, Ruchkali quti, Yostiqcha (pillow) |
| Mailer | Mailer (FEFCO 0427), Mailer oynali |
| Lotok | Lotok, Ikki qavat lotok, Taglik + qopqoq, G'ilof, Tortma quti |

## Yangi shablon qo'shish
`src/modTemplates.bas` da quyidagilarni qiling:
1. `TCount` ni bittaga oshiring.
2. `TGroup`, `TName` va `TParams` ga yangi qator qo'shing.
3. `TBuild` da chizing.

Panellar `Rect` / `Poly` bilan chiziladi: ikki panelga umumiy qirra bukish, qolgan qirralar kesish bo'ladi. Qanotlar orasidagi tirqish `Slit` bilan, oyna yoki teshik `CutRect` / `CutPoly` bilan belgilanadi. 3D ko'rinish uchun `Box3`, `Lid3` va `Face` ishlatiladi.

Koddagi o'zgarishdan keyin `python3 tools/make_paste.py` ni ishga tushiring, u `src/joylash` dagi nusxalarni yangilaydi.
