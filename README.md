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
1. Ushbu repodagi `src` papkasini yuklab oling (**Code → Download ZIP**).
2. CorelDRAW'ni oching va **Alt+F11** ni bosing (VBA muharriri ochiladi).
3. Chapdagi **GlobalMacros** ustida o'ng tugmani bosing va **Import File...** ni tanlang. Keyin quyidagi fayllarni birma-bir import qiling:
   `modGeom.bas`, `modTemplates.bas`, `modDraw.bas`, `modMain.bas`, `clsField.cls`, `frmKarobka.frm`
4. **Tools → References** ichida **Microsoft Forms 2.0 Object Library** belgilangan bo'lsin.
5. Saqlang (**Ctrl+S**).

> Agar `frmKarobka.frm` import bo'lmasa: **Insert → UserForm** qiling, nomini (Name) `frmKarobka` qilib qo'ying, so'ng `frmKarobka.frm` dagi `Option Explicit` dan pastki hamma kodni formaning kod oynasiga joylang.

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
