Attribute VB_Name = "modTemplates"
' Karobka shablonlari. Yangi shablon qo'shish:
'   1) TCount ni oshiring  2) TName va TParams ga qator qo'shing  3) TBuild da chizing
' TParams formati: "Kalit=Yozuv=Standart|..."   (o'lchamlar mm da)
Option Explicit
Option Private Module          ' Macros ro'yxatida faqat Karobka ko'rinsin

Public Function TCount() As Long
    TCount = 6
End Function

Public Function TName(ByVal i As Long) As String
    Select Case i
        Case 1: TName = "Gofra quti (FEFCO 0201)"
        Case 2: TName = "Tuck-end (teskari qopqoq)"
        Case 3: TName = "Tuck-end (to'g'ri qopqoq)"
        Case 4: TName = "Lotok (tray)"
        Case 5: TName = "Taglik + qopqoq (telescope)"
        Case 6: TName = "Mailer (qopqoqli quti)"
    End Select
End Function

Public Function TParams(ByVal i As Long) As String
    Const LWH As String = "L=Uzunlik=%L|W=Eni=%W|H=Balandlik=%H"
    Select Case i
        Case 1: TParams = Fill(LWH, 300, 200, 150) & "|G=Yelim joyi=35"
        Case 2, 3: TParams = Fill(LWH, 80, 50, 150) & "|G=Yelim joyi=15|T=Til (tuck)=15"
        Case 4: TParams = Fill(LWH, 200, 150, 50)
        Case 5: TParams = Fill(LWH, 200, 150, 50) & "|HL=Qopqoq balandligi=30|K=Qopqoq zazori=2"
        Case 6: TParams = Fill(LWH, 250, 180, 60) & "|T=Karton qalinligi=2"
    End Select
End Function

Private Function Fill(ByVal s As String, ByVal L As Double, ByVal W As Double, ByVal H As Double) As String
    Fill = Replace(Replace(Replace(s, "%L", L), "%W", W), "%H", H)
End Function

' p: Scripting.Dictionary  kalit -> qiymat (mm)
Public Sub TBuild(ByVal i As Long, p As Object)
    DieReset
    Select Case i
        Case 1: Rsc0201 p("L"), p("W"), p("H"), p("G")
        Case 2: TuckEnd p("L"), p("W"), p("H"), p("G"), p("T"), 0
        Case 3: TuckEnd p("L"), p("W"), p("H"), p("G"), p("T"), 2
        Case 4: Tray 0, p("L"), p("W"), p("H")
        Case 5
            Tray 0, p("L"), p("W"), p("H")
            Tray p("L") + 2 * p("H") + 20 + p("HL"), p("L") + 2 * p("K"), p("W") + 2 * p("K"), p("HL")
        Case 6: Mailer p("L"), p("W"), p("H"), p("T")
    End Select
    Solve
End Sub

' ---------------------------------------------------------------- yordamchilar

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

' Yelim | L | W | L | W  yon devorlar. xs(0..4) qaytaradi
Private Sub Body(xs() As Double, ByVal L As Double, ByVal W As Double, ByVal H As Double, ByVal g As Double)
    Dim i As Long, c As Double
    xs(0) = 0: xs(1) = L: xs(2) = L + W: xs(3) = 2 * L + W: xs(4) = 2 * L + 2 * W
    c = Min2(g * 0.6, H / 4)
    Poly 0, 0, -g, c, -g, H - c, 0, H
    For i = 0 To 3
        Rect xs(i), 0, xs(i + 1), H
    Next i
End Sub

' ---------------------------------------------------------------- shablonlar

Private Sub Rsc0201(ByVal L As Double, ByVal W As Double, ByVal H As Double, ByVal g As Double)
    Dim xs(0 To 4) As Double, i As Long, f As Double
    Body xs, L, W, H, g
    f = W / 2
    For i = 0 To 3
        Rect xs(i), H, xs(i + 1), H + f
        Rect xs(i), -f, xs(i + 1), 0
    Next i
    For i = 1 To 3
        Slit xs(i), H, xs(i), H + f
        Slit xs(i), -f, xs(i), 0
    Next i
End Sub

' bottomLid: pastki qopqoq qaysi panelda (0 = teskari, 2 = to'g'ri)
Private Sub TuckEnd(ByVal L As Double, ByVal W As Double, ByVal H As Double, ByVal g As Double, _
                    ByVal tuck As Double, ByVal bottomLid As Long)
    Dim xs(0 To 4) As Double
    Body xs, L, W, H, g
    TuckClosure xs, H, 1, 2, L, W, tuck
    TuckClosure xs, 0, -1, bottomLid, L, W, tuck
End Sub

Private Sub TuckClosure(xs() As Double, ByVal base As Double, ByVal s As Double, ByVal lidIdx As Long, _
                        ByVal L As Double, ByVal W As Double, ByVal tuck As Double)
    Dim i As Long, xa As Double, xb As Double, y1 As Double, y2 As Double
    Dim dh As Double, c As Double, tc As Double
    dh = Min2(W * 0.8, L / 2)
    c = Min2(dh / 3, W / 4)
    For i = 0 To 3
        xa = xs(i): xb = xs(i + 1)
        If i = lidIdx Then
            y1 = base + s * W: y2 = base + s * (W + tuck)
            tc = Min2(tuck * 0.8, (xb - xa) / 4)
            Rect xa, base, xb, y1
            Poly xa, y1, xa + tc, y2, xb - tc, y2, xb, y1
        ElseIf i Mod 2 = 1 Then
            Poly xa, base, xa + c, base + s * dh, xb - c, base + s * dh, xb, base
        End If
    Next i
End Sub

' Burchaklari yelimlanadigan lotok. ox: X bo'yicha siljish
Private Sub Tray(ByVal ox As Double, ByVal L As Double, ByVal W As Double, ByVal H As Double)
    Dim x0 As Double, y0 As Double, sx As Double, sy As Double, c As Double, a As Long, b As Long
    Rect ox, 0, ox + L, W
    Rect ox, -H, ox + L, 0
    Rect ox, W, ox + L, W + H
    Rect ox - H, 0, ox, W
    Rect ox + L, 0, ox + L + H, W
    c = H * 0.25
    For a = 0 To 1
        If a = 0 Then x0 = ox: sx = -1 Else x0 = ox + L: sx = 1
        For b = 0 To 1
            If b = 0 Then y0 = 0: sy = -1 Else y0 = W: sy = 1
            Poly x0, y0, x0, y0 + sy * H, x0 + sx * (H - 2), y0 + sy * (H - c), x0 + sx * (H - 2), y0 + sy * c
            Slit x0, y0, x0 + sx * H, y0
        Next b
    Next a
End Sub

' Mailer (FEFCO 0427 ga o'xshash): ikki qavat devorlar, qopqoq va old til
Private Sub Mailer(ByVal L As Double, ByVal W As Double, ByVal H As Double, ByVal t As Double)
    Dim tc As Double, fc As Double, x0 As Double, sx As Double, y0 As Double, sy As Double
    Dim a As Long, b As Long
    Rect 0, 0, L, W                              ' taglik
    Rect 0, -H, L, 0                             ' old devor
    Rect 0, -2 * H + t, L, -H                    ' old ichki devor
    Rect -H, 0, 0, W                             ' chap devor
    Rect -2 * H + t, 0, -H, W                    ' chap ichki
    Rect L, 0, L + H, W                          ' o'ng devor
    Rect L + H, 0, L + 2 * H - t, W              ' o'ng ichki
    Rect 0, W, L, W + H                          ' orqa devor
    Rect 0, W + H, L, 2 * W + H                  ' qopqoq
    tc = Min2(H * 0.4, L / 6)
    Poly 0, 2 * W + H, tc, 2 * W + 2 * H - t, L - tc, 2 * W + 2 * H - t, L, 2 * W + H   ' old til
    fc = Min2(H * 0.5, W / 4)
    For a = 0 To 1
        If a = 0 Then x0 = 0: sx = -1 Else x0 = L: sx = 1
        ' qopqoq yon qanoti
        Poly x0, W + H + t, x0 + sx * (H - t), W + H + t + fc, _
             x0 + sx * (H - t), 2 * W + H - fc, x0, 2 * W + H - t
        ' quloqlar (old va orqa devorga ulangan)
        For b = 0 To 1
            If b = 0 Then y0 = 0: sy = -1 Else y0 = W: sy = 1
            Poly x0, y0 + sy * 0.5, x0 + sx * (H - 2 * t), y0 + sy * (0.5 + fc * 0.6), _
                 x0 + sx * (H - 2 * t), y0 + sy * (H - 1), x0, y0 + sy * H
        Next b
    Next a
End Sub
