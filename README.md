# Karobka: CorelDRAW uchun quti chizmasi

Shablonni tanlaysiz va o'lchamlarni kiritasiz. Dastur yoyilmani (kesish va bukish chiziqlari bilan) hamda yig'ilgan qutining 3D ko'rinishini ko'rsatadi. **Chizish** tugmasi chizmani ochiq turgan CorelDRAW hujjatiga chizadi.

## O'rnatish (3 qadam)
1. [`dist/Karobka.exe`](dist/Karobka.exe) faylini yuklab oling.
2. CorelDRAW'ni oching.
3. `Karobka.exe` ni ishga tushiring.

Talab: Windows 10/11 (64-bit), CorelDRAW 2020 yoki undan yangi versiya. .NET yoki boshqa dastur o'rnatish shart emas.

> Birinchi ochishda Windows "Windows protected your PC" oynasini ko'rsatishi mumkin. Unda **More info → Run anyway** ni bosing.

## Shablonlar
| Guruh | Shablonlar |
|-------|-----------|
| Gofra quti | FEFCO 0201, FEFCO 0200 (ochiq) |
| Tuck-end | Teskari qopqoq, To'g'ri qopqoq, Ruchkali quti, Yostiqcha (pillow) |
| Mailer | Mailer (FEFCO 0427), Mailer oynali |
| Lotok | Lotok, Ikki qavat devorli lotok, Taglik + qopqoq, G'ilof (sleeve), Tortma quti |

`photo.rar` dagi qolgan shablonlar guruh-guruh bo'lib qo'shiladi.

## Dasturchilar uchun
- Kod `app/` papkasida (C#, WPF). Oyna ko'rinishi `app/UI/Main.xaml` da, dizayn ranglari va o'lchamlari shu faylning `Window.Resources` qismida joylashgan.
- Yangi shablon qo'shish uchun `app/Templates/Boxes.cs` da `Template` dan meros klass yozing va uni `Catalog.All` ga qo'shing. Panellar `Rect` / `Poly` bilan chiziladi. Ikki panelga umumiy qirra avtomatik bukish bo'ladi, qolgan qirralar kesish bo'ladi. Qo'shni qanotlar orasidagi tirqish `Slit` bilan, oyna yoki teshik `CutRect` / `CutPoly` bilan belgilanadi.
- Geometriyani tekshirish: `dotnet run --project app/tests -- preview`. Buyruq har bir shablonni `preview/*.svg` ga chizadi.
- Yig'ish: `dotnet publish app/Karobka.csproj -c Release -o dist`. Buyruq Linux'da ham ishlaydi.
- Eski VBA versiyasi `src/` papkasida saqlanib qolgan.
