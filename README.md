# Karobka: CorelDRAW uchun quti chizmasi plagini

Shablonni tanlaysiz, o'lchamlarni kiritasiz va razvyortka darhol sahifada ko'rinadi:
**ko'k chiziq kesish**, **qizil punktir bukish** joyini bildiradi.
"Chizish" tugmasini bossangiz, tayyor chizma bitta guruh bo'lib qoladi.

Talab: CorelDRAW 2020 yoki undan yangi versiya, VBA o'rnatilgan bo'lishi kerak.

## Hozirgi shablonlar
| # | Nomi | O'lchamlar |
|---|------|-----------|
| 1 | Gofra quti (FEFCO 0201) | L, W, H, yelim joyi |
| 2 | Tuck-end, teskari qopqoq | L, W, H, yelim joyi, til |
| 3 | Tuck-end, to'g'ri qopqoq | L, W, H, yelim joyi, til |
| 4 | Lotok (tray) | L, W, H |
| 5 | Taglik + qopqoq (telescope) | L, W, H, qopqoq balandligi, zazor |
| 6 | Mailer (qopqoqli quti) | L, W, H, karton qalinligi |

`photo.rar` dagi boshqa shablonlar keyingi bosqichlarda qo'shiladi.

## O'rnatish (bir marta)
Avval ushbu repodagi kodni yuklab oling: **Code → Download ZIP**. Keyin CorelDRAW'da **Alt+F11** ni bosing, VBA muharriri ochiladi.

> ⚠️ `.bas`, `.cls`, `.frm` fayllarini kod oynasiga **nusxalab joylamang**. Ularning boshidagi `VERSION`, `Begin`, `Attribute` qatorlari kodga tushib qolsa, **"Invalid outside procedure"** xatosi chiqadi. Nusxalab joylash uchun `src/joylash/*.txt` fayllaridan foydalaning.

### A yo'l: Import (oson)
1. Chapdagi **GlobalMacros** ustida o'ng tugmani bosing va **Import File...** ni tanlang.
2. `src` papkasidagi quyidagi fayllarni birma-bir import qiling: `modGeom.bas`, `modTemplates.bas`, `modDraw.bas`, `modMain.bas`, `clsField.cls`, `frmKarobka.frm`.

### B yo'l: qo'lda joylash (import ishlamasa)
1. **GlobalMacros** ni tanlang, so'ng **Insert → Module** ni 4 marta bosing. Pastdagi Properties oynasida modullarning nomini (Name) `modGeom`, `modTemplates`, `modDraw`, `modMain` deb o'zgartiring.
2. **Insert → Class Module** ni bosing va nomini `clsField` qiling.
3. **Insert → UserForm** ni bosing va nomini `frmKarobka` qiling. Keyin formaning ustida ikki marta bosib, uning kod oynasini oching.
4. Har bir oynadagi bor matnni o'chiring. O'rniga `src/joylash` papkasidagi shu nomli `.txt` faylning **butun** matnini joylang.

### Oxirida
- **Tools → References** ichida **Microsoft Forms 2.0 Object Library** belgilangan bo'lsin.
- **Debug → Compile** ni bosing. Xato chiqmasa, **Ctrl+S** bilan saqlang.

## Ishlatish
- **Tools → Macros → Run Macro → Karobka** ni bosing.
- Qulay bo'lishi uchun tugma ham qo'shishingiz mumkin: **Tools → Options → Customization → Commands → Macros** ichidan `Karobka` ni asboblar paneliga sudrab qo'ying.

## Yangi shablon qo'shish
`src/modTemplates.bas` da quyidagilarni qiling:
1. `TCount` ni bittaga oshiring.
2. `TName` va `TParams` ga yangi qator qo'shing.
3. `TBuild` da chizish funksiyasini chaqiring.

Chizishda faqat panellarni `Rect` / `Poly` bilan qo'shish kifoya. Kesish va bukish chiziqlari avtomatik aniqlanadi: ikki panelga umumiy qirra bukish bo'ladi, qolgan qirralar kesish. Qo'shni qanotlar orasidagi tirqishni `Slit` bilan belgilang.

## Tekshirish (Corelsiz)
`python3 tools/preview.py` buyrug'i xuddi shu geometriyani `preview/*.svg` fayllariga chizadi.

`src` dagi kod o'zgartirilgandan keyin `python3 tools/make_paste.py` ni ishga tushiring. U `src/joylash` dagi nusxalarni yangilaydi.
