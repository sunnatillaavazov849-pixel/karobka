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

## Shablonlar (66 ta)
| Guruh | Shablonlar |
|-------|-----------|
| Gofra | FEFCO 0201, 0201 ruchkali, 0200, 0203, 0300 (taglik + qopqoq) |
| Mailer | FEFCO 0427: oddiy, quloqsiz, oynali, osma teshikli, qulfli til, ruchkali; Pitsa qutisi; Qopqoqli lotok |
| Tuck-end | RTE, STE, Kub, Oynali, Osma, Qulf tag (1-2-3), Qulf tag oynali, Seal end, Lab bo'yog'i, Uzun quti, Flakon oynali |
| Ruchkali | Ruchkali quti, qulf tagli, yumaloq ruchkali, Vino qutisi, Pirojniy qutisi, Sut qutisi |
| Lotok | Lotok, ruchkali, ikki qavat, ikki qavat ruchkali, bukiladigan burchak, olti/sakkiz burchakli, G'ilof, Tortma quti, Bo'lakli vstavka (2x2, 3x4) |
| Ko'p burchak | Uchburchak prizma, olti/sakkiz burchakli quti, Piramida (3/4/6), Gul quti (6/8), Olmos (4/6) |
| Sovg'a | Gulbargli (yarim doira, yurak, o'tkir uchli), Yostiqcha, Konfet, Lentali piramida, Uycha, Uycha oynali, Silindr |
| Tort | To'rt bargli, To'rt bargli ruchkali, Yumaloq uchli, Lanch-boks (to'rtburchak, olti, sakkiz burchakli) |

Qaysi rasm qaysi shablonga to'g'ri kelishi [`templates/catalog.md`](templates/catalog.md) da yozilgan.

Qo'shimcha maydonlar 0 bo'lsa, o'sha xususiyat o'chadi: Oyna, Osma, Ruchka, Qulf, Qopqoq quloq.

## Yangi shablon qo'shish
`src/modTemplates.bas` dagi `InitDefs` ga bitta `Def` qatori qo'shing: guruh, nom, quruvchi va standart o'lchamlar. Mavjud quruvchining yangi varianti uchun shu kifoya. Yangi oila bo'lsa, `TBuild` ga yangi quruvchi qo'shing.

Panellar `Rect` / `Poly` bilan chiziladi: ikki panelga umumiy qirra bukish, qolgan qirralar kesish bo'ladi. Qanotlar orasidagi tirqish `Slit` bilan, oyna yoki teshik `CutRect` / `CutPoly` bilan belgilanadi. 3D ko'rinish uchun `Box3`, `Lid3` va `Face` ishlatiladi.

Koddagi o'zgarishdan keyin:
- `python3 tools/preview.py` — hamma shablonlarni Corelsiz chizib tekshiradi (VBA kodi `tools/vba2py.py` bilan Python'ga o'giriladi), natija `preview/` papkasida;
- `python3 tools/make_paste.py` — `src/joylash` dagi nusxalarni yangilaydi.
