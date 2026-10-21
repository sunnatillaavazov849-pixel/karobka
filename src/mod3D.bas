Attribute VB_Name = "mod3D"
' Qutining 3D ko'rinishi.
' 1) Buklash (Fold3D): yoyilma panellari sharnirlar atrofida buriladi - yig'ilish va qopqoq harakati.
' 2) Sodda model: buklab bo'lmaydigan (egri) shablonlar Face/Box3/Lid3 bilan o'zi yuzalar qo'shadi.
' Koordinatalar: x - uzunlik, y - balandlik (yuqoriga), z - chuqurlik.
Option Explicit
Option Private Module          ' Macros ro'yxatida faqat Karobka ko'rinsin

Private Const PI As Double = 3.14159265358979
Private Const LAYER As Double = 0.4  ' qavatlar orasidagi masofa (mm)

Private fStart() As Long, fCnt() As Long, fN As Long
Private px() As Double, py() As Double, pz() As Double, pN As Long
Public fPan() As Long                ' yuza qaysi panel (-1 - sodda model yuzasi)
Public Kinematic As Boolean          ' True - 3D yoyilmadan buklangan
Private pm() As Double               ' panel matritsalari: a11 a12 a13 a21 a22 a23 a31 a32 a33 t1 t2 t3

Public Sub M3Reset()
    fN = 0: pN = 0
    ReDim fStart(0 To 63): ReDim fCnt(0 To 63): ReDim fPan(0 To 63)
    ReDim px(0 To 255): ReDim py(0 To 255): ReDim pz(0 To 255)
End Sub

Private Sub AddFacePt(ByVal x As Double, ByVal y As Double, ByVal z As Double)
    If pN > UBound(px) Then
        ReDim Preserve px(0 To UBound(px) * 2 + 1): ReDim Preserve py(0 To UBound(py) * 2 + 1): ReDim Preserve pz(0 To UBound(pz) * 2 + 1)
    End If
    px(pN) = x: py(pN) = y: pz(pN) = z
    pN = pN + 1
End Sub

Private Sub BeginFace(ByVal pan As Long)
    If fN > UBound(fStart) Then
        ReDim Preserve fStart(0 To UBound(fStart) * 2 + 1): ReDim Preserve fCnt(0 To UBound(fCnt) * 2 + 1)
        ReDim Preserve fPan(0 To UBound(fPan) * 2 + 1)
    End If
    fStart(fN) = pN: fCnt(fN) = 0: fPan(fN) = pan
    fN = fN + 1
End Sub

' ================================================================ buklash

' Yoyilmani buklaydi: tf - yig'ilish (0 - tekis, 1 - yig'ilgan), topen - qopqoq ochiqligi (0 - yopiq, 1 - ochiq).
' Shablon SetRoot chaqirmagan bo'lsa hech narsa qilmaydi (sodda model qoladi)
Public Sub Fold3D(ByVal tf As Double, ByVal topen As Double)
    Dim i As Long, k As Long, h As Long, root As Long, cur As Long, oth As Long
    Dim vis() As Boolean, q() As Long, qh As Long, qt As Long, b As Long, dep() As Long, off As Double
    Kinematic = rootSet And pnN > 0
    If Not Kinematic Then Exit Sub
    ReDim pm(0 To 12 * pnN - 1): ReDim vis(0 To pnN - 1): ReDim q(0 To pnN - 1): ReDim dep(0 To pnN - 1)
    root = FindPanel(rootX, rootY)
    Do While root >= 0
        SetRootMat root
        vis(root) = True: q(qt) = root: qt = qt + 1
        Do While qh < qt
            cur = q(qh): qh = qh + 1
            For h = 0 To hN - 1
                oth = -1
                If Not hBad(h) Then
                    If hA(h) = cur Then oth = hB(h)
                    If hB(h) = cur Then oth = hA(h)
                End If
                If oth >= 0 Then
                    If Not vis(oth) Then
                        ChildMat cur, oth, h, tf, topen
                        vis(oth) = True: q(qt) = oth: qt = qt + 1
                        ' ichki qavat: kichikroq qanot yoki ustiga to'liq buklangan panel; qopqoq - tashqi qavat
                        dep(oth) = dep(cur)
                        If PanelArea(oth) < PanelArea(cur) * 0.98 Or Abs(hAng(h)) >= 150 Then dep(oth) = dep(cur) + 1
                        If hLid(h) Then dep(oth) = hLay(h)
                    End If
                End If
            Next h
        Loop
        root = -1                                ' bog'lanmagan qism qolsa - u ham o'z joyida
        For i = 0 To pnN - 1
            If Not vis(i) Then root = i: Exit For
        Next i
    Loop
    ' qutining markazi (ichki qavatlar shu tomonga suriladi)
    Dim gx As Double, gy As Double, gz As Double, cnt As Long, fx As Double, fy As Double, fz As Double, sg As Double
    For k = 0 To pnN - 1
        b = 12 * k
        For i = pnStart(k) To pnStart(k) + pnCnt(k) - 1
            gx = gx + pm(b) * pvX(i) + pm(b + 1) * pvY(i) + pm(b + 9)
            gy = gy + pm(b + 3) * pvX(i) + pm(b + 4) * pvY(i) + pm(b + 10)
            gz = gz + pm(b + 6) * pvX(i) + pm(b + 7) * pvY(i) + pm(b + 11)
            cnt = cnt + 1
        Next i
    Next k
    gx = gx / cnt: gy = gy / cnt: gz = gz / cnt
    ' yuzalar: har panel uchlari matritsa bilan. Ichki qavatlar o'z tekisligidan markaz tomonga biroz suriladi -
    ' bir tekislikdagi yuzalar ustma-ust ko'rinmasin
    M3Reset
    For k = 0 To pnN - 1
        b = 12 * k
        fx = 0: fy = 0: fz = 0
        For i = pnStart(k) To pnStart(k) + pnCnt(k) - 1
            fx = fx + pm(b) * pvX(i) + pm(b + 1) * pvY(i) + pm(b + 9)
            fy = fy + pm(b + 3) * pvX(i) + pm(b + 4) * pvY(i) + pm(b + 10)
            fz = fz + pm(b + 6) * pvX(i) + pm(b + 7) * pvY(i) + pm(b + 11)
        Next i
        fx = fx / pnCnt(k): fy = fy / pnCnt(k): fz = fz / pnCnt(k)
        sg = pm(b + 2) * (gx - fx) + pm(b + 5) * (gy - fy) + pm(b + 8) * (gz - fz)
        If sg < 0 Then sg = -1 Else sg = 1
        off = LAYER * dep(k) * tf * sg
        BeginFace k
        For i = pnStart(k) To pnStart(k) + pnCnt(k) - 1
            AddFacePt pm(b) * pvX(i) + pm(b + 1) * pvY(i) + pm(b + 2) * off + pm(b + 9), _
                      pm(b + 3) * pvX(i) + pm(b + 4) * pvY(i) + pm(b + 5) * off + pm(b + 10), _
                      pm(b + 6) * pvX(i) + pm(b + 7) * pvY(i) + pm(b + 8) * off + pm(b + 11)
            fCnt(fN - 1) = fCnt(fN - 1) + 1
        Next i
    Next k
End Sub

Private Function PanelArea(ByVal k As Long) As Double
    Dim i As Long, j As Long, a As Double
    For i = 0 To pnCnt(k) - 1
        j = (i + 1) Mod pnCnt(k)
        a = a + pvX(pnStart(k) + i) * pvY(pnStart(k) + j) - pvX(pnStart(k) + j) * pvY(pnStart(k) + i)
    Next i
    PanelArea = Abs(a) / 2
End Function

' (x, y) nuqtasi ichidagi panel; topilmasa eng katta panel
Private Function FindPanel(ByVal x As Double, ByVal y As Double) As Long
    Dim k As Long, i As Long, j As Long, inside As Boolean, best As Long, ar As Double, bestAr As Double
    Dim xi As Double, yi As Double, xj As Double, yj As Double
    best = 0: bestAr = -1
    For k = 0 To pnN - 1
        inside = False: ar = 0
        For i = 0 To pnCnt(k) - 1
            j = (i + 1) Mod pnCnt(k)
            xi = pvX(pnStart(k) + i): yi = pvY(pnStart(k) + i)
            xj = pvX(pnStart(k) + j): yj = pvY(pnStart(k) + j)
            If (yi > y) <> (yj > y) Then
                If x < (xj - xi) * (y - yi) / (yj - yi) + xi Then inside = Not inside
            End If
            ar = ar + xi * yj - xj * yi
        Next i
        If inside Then FindPanel = k: Exit Function
        If Abs(ar) > bestAr Then bestAr = Abs(ar): best = k
    Next k
    FindPanel = best
End Function

' Asosiy panel: yotiq (taglik) yoki tik (old devor). Yoyilmaning +z tomoniga buklanadi
Private Sub SetRootMat(ByVal k As Long)
    Dim b As Long, i As Long
    b = 12 * k
    For i = 0 To 11: pm(b + i) = 0: Next i
    pm(b) = 1
    If rootVert Then
        pm(b + 4) = 1: pm(b + 8) = -1            ' X = x, Y = y, Z = -z
    Else
        pm(b + 5) = 1: pm(b + 7) = -1            ' X = x, Y = z, Z = -y
    End If
End Sub

' Bola panel matritsasi = ota matritsasi x sharnir o'qi atrofida burilish
Private Sub ChildMat(ByVal par As Long, ByVal ch As Long, ByVal h As Long, ByVal tf As Double, ByVal topen As Double)
    Dim ax As Double, ay As Double, ln As Double, cx As Double, cy As Double, i As Long
    Dim ang As Double, c As Double, s As Double, r(0 To 8) As Double, u0 As Double, u1 As Double, u2 As Double
    Dim bp As Long, bc As Long, row As Long, col As Long
    ax = hX2(h) - hX1(h): ay = hY2(h) - hY1(h): ln = Sqr(ax * ax + ay * ay)
    ax = ax / ln: ay = ay / ln
    For i = pnStart(ch) To pnStart(ch) + pnCnt(ch) - 1
        cx = cx + pvX(i): cy = cy + pvY(i)
    Next i
    cx = cx / pnCnt(ch): cy = cy / pnCnt(ch)
    ang = hAng(h) * tf
    If hLid(h) Then ang = (hAng(h) - topen * 110) * tf
    If ax * (cy - hY1(h)) - ay * (cx - hX1(h)) < 0 Then ang = -ang     ' bola panel +z tomonga ko'tarilsin
    ang = ang * PI / 180: c = Cos(ang): s = Sin(ang)
    r(0) = c + (1 - c) * ax * ax: r(1) = (1 - c) * ax * ay: r(2) = s * ay
    r(3) = (1 - c) * ax * ay: r(4) = c + (1 - c) * ay * ay: r(5) = -s * ax
    r(6) = -s * ay: r(7) = s * ax: r(8) = c
    u0 = hX1(h) - (r(0) * hX1(h) + r(1) * hY1(h))
    u1 = hY1(h) - (r(3) * hX1(h) + r(4) * hY1(h))
    u2 = 0 - (r(6) * hX1(h) + r(7) * hY1(h))
    bp = 12 * par: bc = 12 * ch
    For row = 0 To 2
        For col = 0 To 2
            pm(bc + 3 * row + col) = pm(bp + 3 * row) * r(col) + pm(bp + 3 * row + 1) * r(3 + col) + pm(bp + 3 * row + 2) * r(6 + col)
        Next col
        pm(bc + 9 + row) = pm(bp + 3 * row) * u0 + pm(bp + 3 * row + 1) * u1 + pm(bp + 3 * row + 2) * u2 + pm(bp + 9 + row)
    Next row
End Sub

' Face x1, y1, z1, x2, y2, z2, ...
Public Sub Face(ParamArray v() As Variant)
    Dim n As Long, i As Long, m As Long
    n = (UBound(v) + 1) \ 3
    If fN > UBound(fStart) Then
        m = UBound(fStart) * 2 + 1
        ReDim Preserve fStart(0 To m): ReDim Preserve fCnt(0 To m)
    End If
    If pN + n > UBound(px) Then
        m = (UBound(px) + n) * 2
        ReDim Preserve px(0 To m): ReDim Preserve py(0 To m): ReDim Preserve pz(0 To m)
    End If
    If fN > UBound(fPan) Then ReDim Preserve fPan(0 To UBound(fStart))
    fStart(fN) = pN: fCnt(fN) = n: fPan(fN) = -1
    For i = 0 To n - 1
        px(pN) = v(3 * i): py(pN) = v(3 * i + 1): pz(pN) = v(3 * i + 2)
        pN = pN + 1
    Next i
    fN = fN + 1
End Sub

' Quti devorlari. topOn/bottomOn/endsOn: qopqoq, tag, yon devorlar bormi
Public Sub Box3(ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                ByVal L As Double, ByVal H As Double, ByVal W As Double, _
                Optional ByVal topOn As Boolean = True, Optional ByVal bottomOn As Boolean = True, _
                Optional ByVal endsOn As Boolean = True)
    If bottomOn Then Face x, y, z, x + L, y, z, x + L, y, z + W, x, y, z + W
    If topOn Then Face x, y + H, z, x, y + H, z + W, x + L, y + H, z + W, x + L, y + H, z
    Face x, y, z + W, x + L, y, z + W, x + L, y + H, z + W, x, y + H, z + W
    Face x, y, z, x, y + H, z, x + L, y + H, z, x + L, y, z
    If endsOn Then
        Face x, y, z, x, y, z + W, x, y + H, z + W, x, y + H, z
        Face x + L, y, z, x + L, y + H, z, x + L, y + H, z + W, x + L, y, z + W
    End If
End Sub

' Orqa yuqori qirrada ochilgan qopqoq (gradus), front - old til uzunligi
Public Sub Lid3(ByVal L As Double, ByVal H As Double, ByVal W As Double, ByVal angle As Double, Optional ByVal front As Double = 0)
    Dim a As Double, c As Double, s As Double
    a = angle * PI / 180: c = Cos(a): s = Sin(a)
    Face 0, H, 0, 0, H + W * s, W * c, L, H + W * s, W * c, L, H, 0
    If front > 0 Then
        Face 0, H + W * s, W * c, 0, H + W * s - front * c, W * c + front * s, _
             L, H + W * s - front * c, W * c + front * s, L, H + W * s, W * c
    End If
End Sub

' ================================================================ ko'rinish

' Proyeksiya natijasi (ViewProject): nuqtalar ekran koordinatasida (vZ - ko'ruvchiga yaqinligi),
' fOrd - chizish tartibi (uzoqdan yaqinga), fLit - yorug'lik (0..1), fBack - yuzaning ichki tomoni ko'rinadi
Public vX() As Double, vY() As Double, vZ() As Double
Public fOrd() As Long, fLit() As Double, fBack() As Boolean
Public vMinX As Double, vMinY As Double, vMaxX As Double, vMaxY As Double
Private fNx() As Double, fNy() As Double, fNz() As Double, fD() As Double
Private fBx0() As Double, fBx1() As Double, fBy0() As Double, fBy1() As Double

Public Function FaceCount() As Long
    FaceCount = fN
End Function

Public Function FaceStart(ByVal f As Long) As Long
    FaceStart = fStart(f)
End Function

Public Function FaceSize(ByVal f As Long) As Long
    FaceSize = fCnt(f)
End Function

' Yuzalarni yaw (gorizontal burilish) va pitch (qiyalik) burchaklari bilan proyeksiyalaydi va tartiblaydi
Public Sub ViewProject(ByVal yaw As Double, ByVal pitch As Double)
    Dim i As Long, j As Long, k As Long, n As Long, t As Long, pass As Long
    Dim cy As Double, sy As Double, cp As Double, sp As Double, ax As Double, ay As Double, az As Double
    Dim depth() As Double, ln As Double, changed As Boolean
    If fN = 0 Then Exit Sub
    cy = Cos(yaw * PI / 180): sy = Sin(yaw * PI / 180)
    cp = Cos(pitch * PI / 180): sp = Sin(pitch * PI / 180)
    ReDim vX(0 To pN - 1): ReDim vY(0 To pN - 1): ReDim vZ(0 To pN - 1)
    vMinX = 1E+30: vMinY = 1E+30: vMaxX = -1E+30: vMaxY = -1E+30
    For i = 0 To pN - 1
        ax = px(i) * cy + pz(i) * sy: az = -px(i) * sy + pz(i) * cy   ' Y o'qi atrofida
        ay = py(i) * cp - az * sp: az = py(i) * sp + az * cp            ' X o'qi atrofida
        vX(i) = ax: vY(i) = ay: vZ(i) = az
        If ax < vMinX Then vMinX = ax
        If ax > vMaxX Then vMaxX = ax
        If ay < vMinY Then vMinY = ay
        If ay > vMaxY Then vMaxY = ay
    Next i
    ' yuza tekisliklari (Newell usuli), chegaralari, yorug'lik
    ReDim fNx(0 To fN - 1): ReDim fNy(0 To fN - 1): ReDim fNz(0 To fN - 1): ReDim fD(0 To fN - 1)
    ReDim fBx0(0 To fN - 1): ReDim fBx1(0 To fN - 1): ReDim fBy0(0 To fN - 1): ReDim fBy1(0 To fN - 1)
    ReDim fLit(0 To fN - 1): ReDim fBack(0 To fN - 1): ReDim depth(0 To fN - 1): ReDim fOrd(0 To fN - 1)
    For i = 0 To fN - 1
        n = fStart(i)
        fBx0(i) = 1E+30: fBy0(i) = 1E+30: fBx1(i) = -1E+30: fBy1(i) = -1E+30
        For j = 0 To fCnt(i) - 1
            k = n + (j + 1) Mod fCnt(i)
            fNx(i) = fNx(i) + (vY(n + j) - vY(k)) * (vZ(n + j) + vZ(k))
            fNy(i) = fNy(i) + (vZ(n + j) - vZ(k)) * (vX(n + j) + vX(k))
            fNz(i) = fNz(i) + (vX(n + j) - vX(k)) * (vY(n + j) + vY(k))
            depth(i) = depth(i) + vZ(n + j)
            If vX(n + j) < fBx0(i) Then fBx0(i) = vX(n + j)
            If vX(n + j) > fBx1(i) Then fBx1(i) = vX(n + j)
            If vY(n + j) < fBy0(i) Then fBy0(i) = vY(n + j)
            If vY(n + j) > fBy1(i) Then fBy1(i) = vY(n + j)
        Next j
        depth(i) = depth(i) / fCnt(i)
        ln = Sqr(fNx(i) ^ 2 + fNy(i) ^ 2 + fNz(i) ^ 2)
        If ln > 0 Then
            fNx(i) = fNx(i) / ln: fNy(i) = fNy(i) / ln: fNz(i) = fNz(i) / ln
        End If
        fD(i) = fNx(i) * vX(n) + fNy(i) * vY(n) + fNz(i) * vZ(n)
        fBack(i) = fNz(i) < 0
        fLit(i) = 0.55 + 0.45 * Abs(fNx(i) * -0.3 + fNy(i) * 0.8 + fNz(i) * 0.5)
        If fBack(i) Then fLit(i) = fLit(i) * 0.82                     ' ichki tomoni to'qroq
    Next i
    ' 1) markaz chuqurligi bo'yicha (uzoqdan yaqinga)
    For i = 0 To fN - 1: fOrd(i) = i: Next i
    For i = 1 To fN - 1
        t = fOrd(i): j = i - 1
        Do While j >= 0
            If depth(fOrd(j)) <= depth(t) Then Exit Do
            fOrd(j + 1) = fOrd(j): j = j - 1
        Loop
        fOrd(j + 1) = t
    Next i
    ' 2) ustma-ust tushgan yuzalar tekislik bo'yicha tuzatiladi (Newell): orqadagisi oldin
    For pass = 1 To 4
        changed = False
        For i = 0 To fN - 2
            For j = i + 1 To fN - 1
                If Overlap2D(fOrd(i), fOrd(j)) Then
                    If Behind(fOrd(j), fOrd(i)) And Not Behind(fOrd(i), fOrd(j)) Then
                        t = fOrd(j)
                        For k = j To i + 1 Step -1: fOrd(k) = fOrd(k - 1): Next k
                        fOrd(i) = t
                        changed = True
                    End If
                End If
            Next j
        Next i
        If Not changed Then Exit For
    Next pass
End Sub

Private Function Overlap2D(ByVal a As Long, ByVal b As Long) As Boolean
    Overlap2D = fBx0(a) < fBx1(b) - 0.01 And fBx0(b) < fBx1(a) - 0.01 And _
                fBy0(a) < fBy1(b) - 0.01 And fBy0(b) < fBy1(a) - 0.01
End Function

' b yuza a dan orqadami (ko'ruvchiga nisbatan)
Private Function Behind(ByVal b As Long, ByVal a As Long) As Boolean
    Behind = SideAll(b, a, -1) Or SideAll(a, b, 1)
End Function

' f yuzaning hamma uchlari pl tekisligining ko'ruvchi tomonida (want = 1) yoki teskari tomonida (want = -1)
Private Function SideAll(ByVal f As Long, ByVal pl As Long, ByVal want As Long) As Boolean
    Dim j As Long, d As Double, vs As Double, strict As Boolean
    Const E As Double = 0.05
    If Abs(fNz(pl)) < 0.000001 Then Exit Function
    If fNz(pl) > 0 Then vs = 1 Else vs = -1
    For j = fStart(f) To fStart(f) + fCnt(f) - 1
        d = (fNx(pl) * vX(j) + fNy(pl) * vY(j) + fNz(pl) * vZ(j) - fD(pl)) * vs * want
        If d < -E Then Exit Function
        If d > E Then strict = True
    Next j
    SideAll = strict
End Function

' 3D ni sahifaga chizadi. (ox, oy) - chap past burchak, targetH - balandlik (mm), yaw/pitch - ko'rinish burchaklari,
' rgbColor - karton rangi. Shape lar sr ga qo'shiladi
Public Sub Draw3D(sr As ShapeRange, ByVal ox As Double, ByVal oy As Double, ByVal targetH As Double, _
                  Optional ByVal yaw As Double = -35, Optional ByVal pitch As Double = 25, _
                  Optional ByVal rgbColor As Long = &H7EA9CF)
    Dim i As Long, j As Long, k As Long, n As Long, sc As Double, lit As Double
    Dim cr As Long, cg As Long, cb As Long
    Dim crv As Curve, sub_ As SubPath, s As Shape
    If fN = 0 Then Exit Sub
    ViewProject yaw, pitch
    If vMaxY - vMinY < 0.001 Then Exit Sub
    sc = targetH / (vMaxY - vMinY)
    cr = rgbColor And &HFF: cg = (rgbColor \ &H100) And &HFF: cb = (rgbColor \ &H10000) And &HFF
    For k = 0 To fN - 1
        i = fOrd(k): n = fStart(i)
        If fCnt(i) >= 3 Then
            Set crv = Application.CreateCurve(ActiveDocument)
            Set sub_ = crv.CreateSubPath(ox + (vX(n) - vMinX) * sc, oy + (vY(n) - vMinY) * sc)
            For j = 1 To fCnt(i) - 1
                sub_.AppendLineSegment ox + (vX(n + j) - vMinX) * sc, oy + (vY(n + j) - vMinY) * sc
            Next j
            sub_.Closed = True
            Set s = ActiveLayer.CreateCurve(crv)
            lit = fLit(i)
            s.Fill.UniformColor.RGBAssign Int(cr * lit), Int(cg * lit), Int(cb * lit)
            s.Outline.SetProperties Width:=0.2, Color:=CreateRGBColor(Int(cr * 0.4), Int(cg * 0.4), Int(cb * 0.4))
            sr.Add s
        End If
    Next k
End Sub
