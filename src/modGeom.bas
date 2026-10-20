Attribute VB_Name = "modGeom"
' Razvyortka yadrosi.
' Shablon panellarni (ko'pburchak) qo'shadi. Ikki panelga umumiy qirra -> bukish,
' bitta panelga tegishli qirra -> kesish. Slit bilan ustma-ust qirra -> kesish.
Option Explicit
Option Private Module          ' Macros ro'yxatida faqat Karobka ko'rinsin

Public Const KIND_CUT As Integer = 1
Public Const KIND_FOLD As Integer = 2

Private Const EPS As Double = 0.000001
Private Const TOL As Double = 0.001

Private eX1() As Double, eY1() As Double, eX2() As Double, eY2() As Double, eN As Long
Private sX1() As Double, sY1() As Double, sX2() As Double, sY2() As Double, sN As Long
' erkin chiziqlar: oyna/teshik (CUT) va qo'shimcha bukish (FOLD)
Private fK() As Integer, fX1() As Double, fY1() As Double, fX2() As Double, fY2() As Double, fN As Long

Public segKind() As Integer
Public segX1() As Double, segY1() As Double, segX2() As Double, segY2() As Double
Public segN As Long

Public Sub DieReset()
    eN = 0: sN = 0: segN = 0: fN = 0
    ReDim fK(0 To 63): ReDim fX1(0 To 63): ReDim fY1(0 To 63): ReDim fX2(0 To 63): ReDim fY2(0 To 63)
    ReDim eX1(0 To 255): ReDim eY1(0 To 255): ReDim eX2(0 To 255): ReDim eY2(0 To 255)
    ReDim sX1(0 To 63): ReDim sY1(0 To 63): ReDim sX2(0 To 63): ReDim sY2(0 To 63)
End Sub

Private Function R3(ByVal v As Double) As Double
    R3 = Round(v, 3)
End Function

Private Sub AddEdge(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    If eN > UBound(eX1) Then
        Dim m As Long: m = UBound(eX1) * 2 + 1
        ReDim Preserve eX1(0 To m): ReDim Preserve eY1(0 To m)
        ReDim Preserve eX2(0 To m): ReDim Preserve eY2(0 To m)
    End If
    eX1(eN) = R3(x1): eY1(eN) = R3(y1): eX2(eN) = R3(x2): eY2(eN) = R3(y2)
    eN = eN + 1
End Sub

' Poly x1, y1, x2, y2, x3, y3, ...  (yopiq ko'pburchak)
Public Sub Poly(ParamArray v() As Variant)
    Dim n As Long, i As Long, j As Long
    n = (UBound(v) + 1) \ 2
    For i = 0 To n - 1
        j = (i + 1) Mod n
        AddEdge v(2 * i), v(2 * i + 1), v(2 * j), v(2 * j + 1)
    Next i
End Sub

Public Sub Rect(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    Poly x1, y1, x2, y1, x2, y2, x1, y2
End Sub

' Massivdan ko'pburchak (yoylar uchun)
Public Sub PolyArr(xs() As Double, ys() As Double, ByVal n As Long)
    Dim i As Long, j As Long
    For i = 0 To n - 1
        j = (i + 1) Mod n
        AddEdge xs(i), ys(i), xs(j), ys(j)
    Next i
End Sub

Private Sub AddFree(ByVal k As Integer, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    If fN > UBound(fX1) Then
        Dim m As Long: m = UBound(fX1) * 2 + 1
        ReDim Preserve fK(0 To m): ReDim Preserve fX1(0 To m): ReDim Preserve fY1(0 To m)
        ReDim Preserve fX2(0 To m): ReDim Preserve fY2(0 To m)
    End If
    fK(fN) = k: fX1(fN) = R3(x1): fY1(fN) = R3(y1): fX2(fN) = R3(x2): fY2(fN) = R3(y2)
    fN = fN + 1
End Sub

' Yopiq kesish konturi (oyna, teshik): CutPoly x1, y1, x2, y2, ...
Public Sub CutPoly(ParamArray v() As Variant)
    Dim n As Long, i As Long, j As Long
    n = (UBound(v) + 1) \ 2
    If n = 2 Then AddFree KIND_CUT, v(0), v(1), v(2), v(3): Exit Sub   ' bitta kesish chizig'i
    For i = 0 To n - 1
        j = (i + 1) Mod n
        AddFree KIND_CUT, v(2 * i), v(2 * i + 1), v(2 * j), v(2 * j + 1)
    Next i
End Sub

Public Sub CutRect(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    CutPoly x1, y1, x2, y1, x2, y2, x1, y2
End Sub

' Massivdan yopiq kesish konturi
Public Sub CutPolyArr(xs() As Double, ys() As Double, ByVal n As Long)
    Dim i As Long, j As Long
    For i = 0 To n - 1
        j = (i + 1) Mod n
        AddFree KIND_CUT, xs(i), ys(i), xs(j), ys(j)
    Next i
End Sub

' Qo'shimcha bukish chizig'i
Public Sub Crease(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    AddFree KIND_FOLD, x1, y1, x2, y2
End Sub

' Majburiy kesish (masalan qopqoqlar orasidagi tirqish)
Public Sub Slit(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    If sN > UBound(sX1) Then
        Dim m As Long: m = UBound(sX1) * 2 + 1
        ReDim Preserve sX1(0 To m): ReDim Preserve sY1(0 To m)
        ReDim Preserve sX2(0 To m): ReDim Preserve sY2(0 To m)
    End If
    sX1(sN) = R3(x1): sY1(sN) = R3(y1): sX2(sN) = R3(x2): sY2(sN) = R3(y2)
    sN = sN + 1
End Sub

Private Function OnSeg(ByVal px As Double, ByVal py As Double, _
                       ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    Dim dx As Double, dy As Double, ln2 As Double, t As Double
    dx = x2 - x1: dy = y2 - y1: ln2 = dx * dx + dy * dy
    If ln2 < EPS Then Exit Function
    t = ((px - x1) * dx + (py - y1) * dy) / ln2
    If t < -EPS Or t > 1 + EPS Then Exit Function
    OnSeg = Abs(x1 + t * dx - px) < TOL And Abs(y1 + t * dy - py) < TOL
End Function

Private Sub AddPiece(dict As Object, cnt() As Long, ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double)
    Dim k As String, idx As Long, tx As Double, ty As Double
    ax = R3(ax): ay = R3(ay): bx = R3(bx): by = R3(by)
    If bx < ax Or (bx = ax And by < ay) Then
        tx = ax: ty = ay: ax = bx: ay = by: bx = tx: by = ty
    End If
    k = ax & ";" & ay & ";" & bx & ";" & by
    If dict.Exists(k) Then
        idx = dict(k)
        cnt(idx) = cnt(idx) + 1
    Else
        idx = segN
        If idx > UBound(segX1) Then
            Dim m As Long: m = UBound(segX1) * 2 + 1
            ReDim Preserve segX1(0 To m): ReDim Preserve segY1(0 To m)
            ReDim Preserve segX2(0 To m): ReDim Preserve segY2(0 To m)
            ReDim Preserve segKind(0 To m): ReDim Preserve cnt(0 To m)
        End If
        segX1(idx) = ax: segY1(idx) = ay: segX2(idx) = bx: segY2(idx) = by
        cnt(idx) = 1
        dict.Add k, idx
        segN = segN + 1
    End If
End Sub

' Qirralarni bo'laklarga ajratib, har birini CUT yoki FOLD deb belgilaydi
Public Sub Solve()
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cnt() As Long
    Dim i As Long, j As Long, k As Long, nt As Long
    Dim dx As Double, dy As Double, ln2 As Double, t As Double, tmp As Double
    Dim px As Double, py As Double, mx As Double, my As Double
    Dim ts() As Double

    segN = 0
    ReDim segX1(0 To 255): ReDim segY1(0 To 255): ReDim segX2(0 To 255): ReDim segY2(0 To 255)
    ReDim segKind(0 To 255): ReDim cnt(0 To 255)
    ReDim ts(0 To 2 * eN + 2 * sN + 2)

    For i = 0 To eN - 1
        dx = eX2(i) - eX1(i): dy = eY2(i) - eY1(i): ln2 = dx * dx + dy * dy
        If ln2 > EPS Then
            ts(0) = 0: ts(1) = 1: nt = 2
            For j = 0 To 2 * eN + 2 * sN - 1          ' qirra va tirqish uchlari - bo'luvchi nuqtalar
                If j < 2 * eN Then
                    If j Mod 2 = 0 Then px = eX1(j \ 2): py = eY1(j \ 2) Else px = eX2(j \ 2): py = eY2(j \ 2)
                Else
                    k = j - 2 * eN
                    If k Mod 2 = 0 Then px = sX1(k \ 2): py = sY1(k \ 2) Else px = sX2(k \ 2): py = sY2(k \ 2)
                End If
                t = ((px - eX1(i)) * dx + (py - eY1(i)) * dy) / ln2
                If t > EPS And t < 1 - EPS Then
                    If Abs(eX1(i) + t * dx - px) < TOL And Abs(eY1(i) + t * dy - py) < TOL Then
                        ts(nt) = t: nt = nt + 1
                    End If
                End If
            Next j
            For j = 1 To nt - 1                       ' insertion sort
                tmp = ts(j): k = j - 1
                Do While k >= 0
                    If ts(k) <= tmp Then Exit Do
                    ts(k + 1) = ts(k): k = k - 1
                Loop
                ts(k + 1) = tmp
            Next j
            For j = 0 To nt - 2
                If ts(j + 1) - ts(j) > EPS Then
                    AddPiece dict, cnt, eX1(i) + ts(j) * dx, eY1(i) + ts(j) * dy, _
                                        eX1(i) + ts(j + 1) * dx, eY1(i) + ts(j + 1) * dy
                End If
            Next j
        End If
    Next i

    For i = 0 To segN - 1
        If cnt(i) >= 2 Then segKind(i) = KIND_FOLD Else segKind(i) = KIND_CUT
        If segKind(i) = KIND_FOLD Then
            mx = (segX1(i) + segX2(i)) / 2: my = (segY1(i) + segY2(i)) / 2
            For j = 0 To sN - 1
                If OnSeg(mx, my, sX1(j), sY1(j), sX2(j), sY2(j)) Then segKind(i) = KIND_CUT: Exit For
            Next j
        End If
    Next i

    For i = 0 To fN - 1
        If segN > UBound(segX1) Then
            k = UBound(segX1) * 2 + 1
            ReDim Preserve segX1(0 To k): ReDim Preserve segY1(0 To k)
            ReDim Preserve segX2(0 To k): ReDim Preserve segY2(0 To k): ReDim Preserve segKind(0 To k)
        End If
        segKind(segN) = fK(i)
        segX1(segN) = fX1(i): segY1(segN) = fY1(i): segX2(segN) = fX2(i): segY2(segN) = fY2(i)
        segN = segN + 1
    Next i
End Sub
