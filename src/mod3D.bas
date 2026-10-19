Attribute VB_Name = "mod3D"
' Yig'ilgan qutining sodda 3D ko'rinishi (izometrik).
' Shablon yuzlarni qo'shadi: x = uzunlik, y = balandlik, z = eni.
Option Explicit
Option Private Module          ' Macros ro'yxatida faqat Karobka ko'rinsin

Private Const PI As Double = 3.14159265358979

Private fStart() As Long, fCnt() As Long, fN As Long
Private px() As Double, py() As Double, pz() As Double, pN As Long

Public Sub M3Reset()
    fN = 0: pN = 0
    ReDim fStart(0 To 63): ReDim fCnt(0 To 63)
    ReDim px(0 To 255): ReDim py(0 To 255): ReDim pz(0 To 255)
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
    fStart(fN) = pN: fCnt(fN) = n
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

' 3D ni sahifaga chizadi. (ox, oy) - chap past burchak, targetH - balandlik (mm). Shape lar sr ga qo'shiladi
Public Sub Draw3D(sr As ShapeRange, ByVal ox As Double, ByVal oy As Double, ByVal targetH As Double)
    Dim i As Long, j As Long, k As Long, n As Long
    Dim cy As Double, sy As Double, cp As Double, sp As Double
    Dim rx() As Double, ry() As Double, rz() As Double
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double, sc As Double
    Dim depth() As Double, idx() As Long, t As Long
    Dim crv As Curve, sub_ As SubPath, s As Shape
    Dim ax As Double, ay As Double, az As Double, bx As Double, by As Double, bz As Double
    Dim nx As Double, ny As Double, nz As Double, ln As Double, lit As Double, back As Boolean

    If fN = 0 Then Exit Sub
    cy = Cos(-35 * PI / 180): sy = Sin(-35 * PI / 180)
    cp = Cos(25 * PI / 180): sp = Sin(25 * PI / 180)
    ReDim rx(0 To pN - 1): ReDim ry(0 To pN - 1): ReDim rz(0 To pN - 1)
    minX = 1E+30: minY = 1E+30: maxX = -1E+30: maxY = -1E+30
    For i = 0 To pN - 1
        ax = px(i) * cy + pz(i) * sy: az = -px(i) * sy + pz(i) * cy   ' Y o'qi atrofida
        ay = py(i) * cp - az * sp: az = py(i) * sp + az * cp            ' X o'qi atrofida
        rx(i) = ax: ry(i) = ay: rz(i) = az
        If ax < minX Then minX = ax
        If ax > maxX Then maxX = ax
        If ay < minY Then minY = ay
        If ay > maxY Then maxY = ay
    Next i
    If maxY - minY < 0.001 Then Exit Sub
    sc = targetH / (maxY - minY)

    ' uzoqdagi yuzlar oldin chiziladi
    ReDim depth(0 To fN - 1): ReDim idx(0 To fN - 1)
    For i = 0 To fN - 1
        idx(i) = i
        For j = 0 To fCnt(i) - 1: depth(i) = depth(i) + rz(fStart(i) + j): Next j
        depth(i) = depth(i) / fCnt(i)
    Next i
    For i = 1 To fN - 1
        t = idx(i): j = i - 1
        Do While j >= 0
            If depth(idx(j)) <= depth(t) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = t
    Next i

    For k = 0 To fN - 1
        i = idx(k): n = fStart(i)
        ax = rx(n + 1) - rx(n): ay = ry(n + 1) - ry(n): az = rz(n + 1) - rz(n)
        bx = rx(n + 2) - rx(n): by = ry(n + 2) - ry(n): bz = rz(n + 2) - rz(n)
        nx = ay * bz - az * by: ny = az * bx - ax * bz: nz = ax * by - ay * bx
        ln = Sqr(nx * nx + ny * ny + nz * nz)
        If ln > 0 Then
            back = nz < 0
            If back Then nx = -nx: ny = -ny: nz = -nz
            lit = 0.55 + 0.45 * Abs((nx * -0.3 + ny * 0.8 + nz * 0.5) / ln)
            Set crv = Application.CreateCurve(ActiveDocument)
            Set sub_ = crv.CreateSubPath(ox + (rx(n) - minX) * sc, oy + (ry(n) - minY) * sc)
            For j = 1 To fCnt(i) - 1
                sub_.AppendLineSegment ox + (rx(n + j) - minX) * sc, oy + (ry(n + j) - minY) * sc
            Next j
            sub_.Closed = True
            Set s = ActiveLayer.CreateCurve(crv)
            If back Then lit = lit * 0.82                               ' ichki tomoni to'qroq
            s.Fill.UniformColor.RGBAssign Int(207 * lit), Int(169 * lit), Int(126 * lit)
            s.Outline.SetProperties Width:=0.2, Color:=CreateRGBColor(90, 64, 38)
            sr.Add s
        End If
    Next k
End Sub
