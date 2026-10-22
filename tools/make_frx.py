"""UserForm uchun .frx (ikkilik) faylini yaratadi va .frm sarlavhasini unga moslaydi.

Ishlatish:  python3 tools/make_frx.py            (src/frmKarobka.frm -> src/frmKarobka.frx)
.frx tuzilishi: 24 baytli sarlavha (0x0008424C, CFB hajmi, 0, 0, eni, bo'yi) + Compound File
(\\x01CompObj, f, o oqimlari). Formada boshqaruv elementlari yo'q - hammasi Initialize da kodda yaratiladi,
shuning uchun f oqimida faqat forma xossalari, o oqimi bo'sh.
Yaratilgandan keyin olefile va oletools.oleform bilan tekshiriladi.
"""
import os
import re
import struct
import sys

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src")
FORM_CLSID = "C62A69F0-16DC-11CE-9E98-00AA00574A4F"   # Microsoft Forms 2.0 Form
FONT_CLSID = "0BE35203-8F91-11CE-9DE3-00AA004BB851"   # StdFont
END, FREE, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


def guid(s):
    p = s.split("-")
    return struct.pack("<LHH", int(p[0], 16), int(p[1], 16), int(p[2], 16)) + bytes.fromhex(p[3] + p[4])


def lpstr(s):
    b = s.encode("ascii") + b"\0"
    return struct.pack("<L", len(b)) + b


def lpwstr(s):
    b = s.encode("utf-16-le") + b"\0\0"
    return struct.pack("<L", len(b) // 2) + b


def compobj():
    """[MS-OLEDS] 2.3.8 CompObjStream"""
    d = struct.pack("<LL", 0xFFFE0001, 0x00000A03) + struct.pack("<L", 0xFFFFFFFF) + guid(FORM_CLSID)
    d += lpstr("Microsoft Forms 2.0 Form") + lpstr("Embedded Object") + lpstr("Forms.Form.1")
    d += struct.pack("<L", 0x71B239F4)
    d += lpwstr("Microsoft Forms 2.0 Form") + lpwstr("Embedded Object") + lpwstr("Forms.Form.1")
    return d


def form_stream(w_him, h_him):
    """[MS-OFORMS] 2.2.10.1 FormControl: BackColor, DrawBuffer, DisplayedSize, Font; saytlar yo'q"""
    fBackColor, fDisplayedSize, fFont, fDrawBuffer = 1 << 1, 1 << 10, 1 << 20, 1 << 27
    mask = fBackColor | fDisplayedSize | fFont | fDrawBuffer
    data = struct.pack("<L", 0x8000000F)               # BackColor: tizim rangi (ButtonFace)
    data += struct.pack("<L", 32000)                    # DrawBuffer
    extra = struct.pack("<ll", w_him, h_him)            # DisplayedSize (HIMETRIC)
    body = struct.pack("<L", mask) + data + extra
    s = struct.pack("<BBH", 0, 4, len(body)) + body
    face = b"Tahoma"                                    # StdFont ([MS-OFORMS] 2.4.7, GuidAndFont)
    s += guid(FONT_CLSID) + struct.pack("<BHBHLB", 1, 0, 0, 400, 80000, len(face)) + face
    s += struct.pack("<H", 0)                           # CountOfSiteClassInfo
    s += struct.pack("<LL", 0, 0)                       # CountOfSites, CountOfBytes
    return s


def dir_entry(name, typ, child=NOSTREAM, left=NOSTREAM, right=NOSTREAM, clsid=b"\0" * 16, start=END, size=0):
    n = name.encode("utf-16-le") + b"\0\0"
    e = n.ljust(64, b"\0") + struct.pack("<HBB", len(n), typ, 1)        # 1 - qora
    e += struct.pack("<LLL", left, right, child) + clsid + struct.pack("<L", 0) + b"\0" * 16
    e += struct.pack("<LLL", start, size, 0)
    assert len(e) == 128
    return e


def cfb(streams):
    """Compound File v3 (sektor 512, mini sektor 64): hamma oqimlar mini oqimda.
    streams: [(nom, bayt)] - tartib: \\x01CompObj, f, o"""
    mini, minifat, starts = b"", [], []
    for _, data in streams:
        if not data:
            starts.append(END)
            continue
        n = (len(data) + 63) // 64
        first = len(minifat)
        starts.append(first)
        minifat += [first + k + 1 for k in range(n - 1)] + [END]
        mini += data.ljust(n * 64, b"\0")
    nmini = (len(mini) + 511) // 512
    mini = mini.ljust(nmini * 512, b"\0")
    assert len(minifat) <= 128
    # sektorlar: 0 - FAT, 1 - katalog, 2 - MiniFAT, 3.. - mini oqim
    fat = [FATSECT, END, END] + [3 + k + 1 for k in range(nmini - 1)] + ([END] if nmini else [])
    fat += [FREE] * (128 - len(fat))
    # katalog: 0 root, 1 \x01CompObj, 2 f, 3 o.  Daraxt: root -> o (chapda f, o'ngda \x01CompObj)
    (n1, d1), (n2, d2), (n3, d3) = streams
    dirs = dir_entry("Root Entry", 5, child=3, clsid=guid(FORM_CLSID), start=3 if nmini else END,
                     size=len(minifat) * 64)            # mini oqim hajmi
    dirs += dir_entry(n1, 2, start=starts[0], size=len(d1))
    dirs += dir_entry(n2, 2, start=starts[1], size=len(d2))
    dirs += dir_entry(n3, 2, left=2, right=1, start=starts[2], size=len(d3))
    hdr = bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 16
    hdr += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
    hdr += struct.pack("<LLLLLLLLL", 0, 1, 1, 0, 4096, 2, 1, END, 0)
    hdr += struct.pack("<L", 0) + struct.pack("<L", FREE) * 108
    assert len(hdr) == 512
    mf = minifat + [FREE] * (128 - len(minifat))
    return (hdr + struct.pack("<128L", *fat) + dirs + struct.pack("<128L", *mf) + mini)


def check(frx_path):
    """olefile + oletools.oleform bilan o'qib ko'radi"""
    import olefile
    from oletools import oleform
    raw = open(frx_path, "rb").read()
    sig, size, z1, z2, w, h = struct.unpack("<6L", raw[:24])
    assert sig == 0x0008424C and z1 == 0 and z2 == 0 and size == len(raw) - 24, "sarlavha noto'g'ri"
    ole = olefile.OleFileIO(raw[24:])
    names = sorted("/".join(e) for e in ole.listdir())
    assert names == sorted(["\x01CompObj", "f", "o"]), names
    assert ole.root.clsid.upper() == FORM_CLSID, ole.root.clsid
    cv = ole.openstream("\x01CompObj").read()
    assert b"Forms.Form.1" in cv
    sites = oleform.extract_OleFormVariables(ole, [])
    ole.close()
    return names, len(raw), sites


def main():
    frm = os.path.join(ROOT, "frmKarobka.frm")
    name = os.path.splitext(os.path.basename(frm))[0]
    text = open(frm, "rb").read().decode("ascii")
    cw = int(re.search(r"ClientWidth\s*=\s*(\d+)", text).group(1))
    ch = int(re.search(r"ClientHeight\s*=\s*(\d+)", text).group(1))
    cap = re.search(r'Caption\s*=\s*"([^"]*)"', text).group(1)
    w_him, h_him = round(cw * 2540 / 1440), round(ch * 2540 / 1440)    # twip -> HIMETRIC

    body = cfb([("\x01CompObj", compobj()), ("f", form_stream(w_him, h_him)), ("o", b"")])
    frx = struct.pack("<6L", 0x0008424C, len(body), 0, 0, w_him, h_him) + body
    open(os.path.join(ROOT, name + ".frx"), "wb").write(frx)

    # .frm sarlavhasi (aynan shu ko'rinishda), qolgan kod o'zgarmaydi
    code = text.replace("\r\n", "\n")
    code = code[code.index("Attribute VB_Name"):]
    code = "\n".join(l for l in code.split("\n") if not l.startswith("Attribute VB_")).lstrip("\n")
    head = ("VERSION 5.00\n"
            "Begin {%s} %s\n" % (FORM_CLSID, name) +
            '   Caption         =   "%s"\n' % cap +
            "   ClientHeight    =   %d\n" % ch +
            "   ClientLeft      =   120\n"
            "   ClientTop       =   465\n"
            "   ClientWidth     =   %d\n" % cw +
            '   OleObjectBlob   =   "%s.frx":0000\n' % name +
            "   StartUpPosition =   1  'CenterOwner\n"
            "End\n"
            'Attribute VB_Name = "%s"\n' % name +
            "Attribute VB_GlobalNameSpace = False\n"
            "Attribute VB_Creatable = False\n"
            "Attribute VB_PredeclaredId = True\n"
            "Attribute VB_Exposed = False\n")
    out = (head + code).replace("\n", "\r\n")
    out.encode("ascii")                                 # faqat ASCII
    open(frm, "wb").write(out.encode("ascii"))

    names, size, sites = check(os.path.join(ROOT, name + ".frx"))
    print("%s.frx: %d bayt, oqimlar: %s, saytlar: %d" % (name, size, ", ".join(repr(n) for n in names), len(sites)))


if __name__ == "__main__":
    sys.exit(main())
