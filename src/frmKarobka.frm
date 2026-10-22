VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKarobka
   Caption         =   "Karobka"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   OleObjectBlob   =   "frmKarobka.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKarobka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Karobka paneli. Barcha elementlar kod orqali yaratiladi, formaga hech narsa qo'yish shart emas.
Option Explicit

' Dizayn: ranglar (BGR), o'lchamlar
Private Const C_BG As Long = &HF4F2F2        ' och kulrang fon
Private Const C_CARD As Long = &HFFFFFF      ' oq blok
Private Const C_FIELD As Long = &HF4F2F2     ' maydon foni
Private Const C_LINE As Long = &HEAE5E5      ' yengil chegara
Private Const C_TEXT As Long = &H1E1C1C      ' asosiy matn
Private Const C_MUTED As Long = &H938E8E     ' kulrang matn
Private Const C_ERR As Long = &H4D48E5       ' xato
Private Const FONT_NAME As String = "Segoe UI"
Private Const ROW_H As Single = 22           ' maydon balandligi
Private Const PAD As Single = 12

Private WithEvents lst As MSForms.ListBox
Private WithEvents cmbGroup As MSForms.ComboBox
Private WithEvents txtFind As MSForms.TextBox
Private ids() As Long                 ' ro'yxat qatori -> shablon raqami
Private curT As Long                  ' maydonlari ko'rsatilgan shablon
Private WithEvents chk3D As MSForms.CheckBox
Private WithEvents chkLive As MSForms.CheckBox
Private WithEvents btnDraw As MSForms.Label
Private fra As MSForms.Frame
Private lblStatus As MSForms.Label
' 3D bloki
Private WithEvents sbFold As MSForms.ScrollBar
Private WithEvents sbLid As MSForms.ScrollBar
Private WithEvents sbYaw As MSForms.ScrollBar
Private WithEvents sbPitch As MSForms.ScrollBar
Private WithEvents sbZoom As MSForms.ScrollBar
Private WithEvents cmbColor As MSForms.ComboBox
Private WithEvents btnFoldAnim As MSForms.Label
Private WithEvents btnLidAnim As MSForms.Label
Private WithEvents btnMockup As MSForms.Label
Private lblFoldV As MSForms.Label, lblLidV As MSForms.Label
Private busy3D As Boolean
Private fields As Collection
Private building As Boolean

' ------------------------------------------------------------ elementlar

Private Function NewLabel(parent As Object, ByVal txt As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single, Optional ByVal bold As Boolean = False, _
                          Optional ByVal color As Long = C_TEXT, Optional ByVal size As Single = 9) As MSForms.Label
    Dim c As MSForms.Label
    Set c = parent.Controls.Add("Forms.Label.1")
    c.Caption = txt: c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.BackStyle = fmBackStyleTransparent
    c.ForeColor = color
    c.Font.Name = FONT_NAME: c.Font.Size = size: c.Font.Bold = bold
    Set NewLabel = c
End Function

' Yassi "tugma" (Label): fon rangli, chegara yo'q, matn markazda
Private Function NewButton(parent As Object, ByVal txt As String, ByVal l As Single, ByVal t As Single, _
                           ByVal w As Single, ByVal h As Single, ByVal back As Long, ByVal fore As Long, _
                           Optional ByVal size As Single = 10) As MSForms.Label
    Dim c As MSForms.Label
    Set c = NewLabel(parent, txt, l, t, w, h, True, fore, size)
    c.BackStyle = fmBackStyleOpaque
    c.BackColor = back
    c.TextAlign = fmTextAlignCenter
    Set NewButton = c
End Function

Private Function NewCheck(ByVal txt As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                          ByVal tip As String) As MSForms.CheckBox
    Dim c As MSForms.CheckBox
    Set c = Me.Controls.Add("Forms.CheckBox.1")
    c.Caption = txt: c.Left = l: c.Top = t: c.Width = w: c.Height = 18
    c.BackStyle = fmBackStyleTransparent
    c.ForeColor = C_TEXT: c.Font.Name = FONT_NAME: c.Font.Size = 9
    c.ControlTipText = tip
    c.Value = True
    Set NewCheck = c
End Function

Private Sub UserForm_Initialize()
    Dim i As Long, cardL As MSForms.Frame, gl() As String
    Me.Caption = "Karobka"
    Me.BackColor = C_BG
    Me.Width = 700: Me.Height = 420
    Me.Font.Name = FONT_NAME

    ' --- Shablon bloki
    Set cardL = Me.Controls.Add("Forms.Frame.1", "cardL")
    With cardL
        .Left = PAD: .Top = PAD: .Width = 200: .Height = 352
        .BackColor = C_CARD: .BorderStyle = fmBorderStyleSingle: .BorderColor = C_LINE
        .SpecialEffect = fmSpecialEffectFlat: .Caption = ""
    End With
    NewLabel cardL, "Shablon", PAD, 8, 150, 16, True, C_TEXT, 10
    Set cmbGroup = cardL.Controls.Add("Forms.ComboBox.1", "cmbGroup")
    With cmbGroup
        .Left = 8: .Top = 30: .Width = 182: .Height = 18
        .Style = fmStyleDropDownList: .SpecialEffect = fmSpecialEffectFlat
        .BackColor = C_FIELD: .ForeColor = C_TEXT
        .Font.Name = FONT_NAME: .Font.Size = 9
        .ControlTipText = "Guruh bo'yicha saralash"
        .AddItem "Hammasi"
    End With
    gl = Split(TGroups, "|")
    For i = 0 To UBound(gl)
        cmbGroup.AddItem gl(i)
    Next i
    For i = 1 To TCount
        If Not HasItem(cmbGroup, TGroup(i)) Then cmbGroup.AddItem TGroup(i)
    Next i
    Set txtFind = cardL.Controls.Add("Forms.TextBox.1", "txtFind")
    With txtFind
        .Left = 8: .Top = 52: .Width = 182: .Height = 18
        .BorderStyle = fmBorderStyleSingle: .BorderColor = C_LINE: .SpecialEffect = fmSpecialEffectFlat
        .BackColor = C_FIELD: .ForeColor = C_TEXT
        .Font.Name = FONT_NAME: .Font.Size = 9
        .ControlTipText = "Qidirish: rasm raqami (masalan 0311) yoki nomi"
    End With
    Set lst = cardL.Controls.Add("Forms.ListBox.1", "lst")
    With lst
        .Left = 6: .Top = 76: .Width = 186: .Height = 268
        .BorderStyle = fmBorderStyleNone: .SpecialEffect = fmSpecialEffectFlat
        .BackColor = C_CARD: .ForeColor = C_TEXT
        .Font.Name = FONT_NAME: .Font.Size = 9
    End With

    ' --- O'lcham bloki
    Set fra = Me.Controls.Add("Forms.Frame.1", "fra")
    With fra
        .Left = PAD + 200 + PAD: .Top = PAD: .Width = 240: .Height = 290
        .BackColor = C_CARD: .BorderStyle = fmBorderStyleSingle: .BorderColor = C_LINE
        .SpecialEffect = fmSpecialEffectFlat: .Caption = ""
        .ScrollBars = fmScrollBarsNone
    End With

    ' --- Sozlamalar va asosiy tugma
    Set chk3D = NewCheck("3D", fra.Left + 2, 308, 60, "Yig'ilgan qutini yoyilma yonida ko'rsatish")
    Set chkLive = NewCheck("Jonli ko'rish", fra.Left + 66, 308, 120, "O'lcham o'zgarganda sahifada darhol ko'rsatish")
    Set btnDraw = NewButton(Me, "Chizish", fra.Left, 330, 240, 24, C_TEXT, C_CARD, 10)
    btnDraw.ControlTipText = "Tayyor chizmani sahifaga qo'yish"
    Set lblStatus = NewLabel(Me, "", PAD, 368, 672, 14, False, C_MUTED, 8)
    Build3DBlock fra.Left + 240 + PAD

    cmbGroup.ListIndex = 0                ' ro'yxatni to'ldiradi
End Sub

' ------------------------------------------------------------ 3D bloki

Private Function NewSlider(parent As Object, ByVal t As Single, ByVal mn As Long, ByVal mx As Long, _
                           ByVal v As Long, ByVal tip As String) As MSForms.ScrollBar
    Dim c As MSForms.ScrollBar
    Set c = parent.Controls.Add("Forms.ScrollBar.1")
    With c
        .Left = PAD: .Top = t: .Width = 184: .Height = 14
        .Orientation = fmOrientationHorizontal
        .Min = mn: .Max = mx: .SmallChange = 5: .LargeChange = 15: .Value = v
        .ControlTipText = tip
    End With
    Set NewSlider = c
End Function

Private Sub Build3DBlock(ByVal l As Single)
    Dim f3 As MSForms.Frame
    Set f3 = Me.Controls.Add("Forms.Frame.1", "fra3")
    With f3
        .Left = l: .Top = PAD: .Width = 208: .Height = 342
        .BackColor = C_CARD: .BorderStyle = fmBorderStyleSingle: .BorderColor = C_LINE
        .SpecialEffect = fmSpecialEffectFlat: .Caption = ""
    End With
    NewLabel f3, "3D", PAD, 8, 100, 16, True, C_TEXT, 10

    NewLabel f3, "Yig'ish", PAD, 34, 100, 14
    Set lblFoldV = NewLabel(f3, "100%", 150, 34, 46, 14, False, C_MUTED, 8)
    lblFoldV.TextAlign = fmTextAlignRight
    Set sbFold = NewSlider(f3, 50, 0, 100, 100, "0% - tekis yoyilma, 100% - yig'ilgan quti")
    Set btnFoldAnim = NewButton(f3, "Yig'ilishni ko'rsatish", PAD, 68, 184, 20, C_FIELD, C_TEXT, 9)

    NewLabel f3, "Qopqoq", PAD, 98, 100, 14
    Set lblLidV = NewLabel(f3, "0%", 150, 98, 46, 14, False, C_MUTED, 8)
    lblLidV.TextAlign = fmTextAlignRight
    Set sbLid = NewSlider(f3, 114, 0, 100, 0, "0% - yopiq, 100% - ochiq")
    Set btnLidAnim = NewButton(f3, "Ochish", PAD, 132, 184, 20, C_FIELD, C_TEXT, 9)

    NewLabel f3, "Aylantirish", PAD, 162, 100, 14
    Set sbYaw = NewSlider(f3, 178, -180, 180, -35, "Qutini chapga-o'ngga aylantirish")
    NewLabel f3, "Qiyalik", PAD, 198, 100, 14
    Set sbPitch = NewSlider(f3, 214, 0, 90, 25, "Yuqoridan qarash burchagi")
    NewLabel f3, "Masshtab", PAD, 234, 100, 14
    Set sbZoom = NewSlider(f3, 250, 40, 200, 100, "3D kattaligi")

    NewLabel f3, "Rang", PAD, 276, 50, 14
    Set cmbColor = f3.Controls.Add("Forms.ComboBox.1", "cmbColor")
    With cmbColor
        .Left = 70: .Top = 273: .Width = 126: .Height = 18
        .Style = fmStyleDropDownList: .SpecialEffect = fmSpecialEffectFlat
        .BackColor = C_FIELD: .ForeColor = C_TEXT
        .Font.Name = FONT_NAME: .Font.Size = 9
        .AddItem "Kraft": .AddItem "Oq": .AddItem "Qora": .AddItem "Qizil"
        .AddItem "Ko'k": .AddItem "Yashil": .AddItem "Pushti"
        .ListIndex = 0
    End With
    Set btnMockup = NewButton(f3, "Mockup", PAD, 306, 184, 24, C_TEXT, C_CARD, 10)
    btnMockup.ControlTipText = "Dizayn qatlamidagi rasmni 3D qutiga joylash"
End Sub

' Mockup: "Dizayn" qatlamidagi dizayn 3D qutining yuzalariga joylanadi
Private Sub btnMockup_Click()
    Dim dsg As Object, created As Boolean
    If Not HasDie Then SetStatus "Avval yoyilmani chizing", True: Exit Sub
    Set dsg = DesignLayer(created)
    If dsg Is Nothing Then Exit Sub
    If created Or dsg.Shapes.Count = 0 Then
        MsgBox "Dizaynni yoyilma ustiga '" & DESIGN_LAYER & "' qatlamiga joylang va Mockup ni qayta bosing." & _
               vbCrLf & "(Qatlam Object Manager oynasida ko'rinadi.)", vbInformation, "Karobka"
        Exit Sub
    End If
    SetStatus "Mockup chizilmoqda...", False
    DoEvents
    Fold3D sbFold.Value / 100, sbLid.Value / 100
    Redraw3D sbYaw.Value, sbPitch.Value, sbZoom.Value / 100, BoxColor, False, dsg
    SetStatus "Mockup tayyor", False
End Sub

Private Function BoxColor() As Long
    Select Case cmbColor.ListIndex
        Case 1: BoxColor = RGB(245, 245, 242)
        Case 2: BoxColor = RGB(70, 70, 72)
        Case 3: BoxColor = RGB(200, 45, 45)
        Case 4: BoxColor = RGB(45, 95, 185)
        Case 5: BoxColor = RGB(60, 150, 80)
        Case 6: BoxColor = RGB(232, 128, 165)
        Case Else: BoxColor = RGB(207, 169, 126)
    End Select
End Function

' 3D ni joriy slayderlar bo'yicha qayta chizish (yoyilma qayta qurilmaydi)
Private Sub Update3D(Optional ByVal final As Boolean = False)
    If busy3D Or sbZoom Is Nothing Then Exit Sub
    lblFoldV.Caption = sbFold.Value & "%": lblLidV.Caption = sbLid.Value & "%"
    If Not chk3D.Value Then Delete3D: Exit Sub
    Fold3D sbFold.Value / 100, sbLid.Value / 100
    Redraw3D sbYaw.Value, sbPitch.Value, sbZoom.Value / 100, BoxColor, final
End Sub

Private Sub Pause(ByVal sec As Single)
    Dim t0 As Single
    t0 = Timer
    Do While Timer < t0 + sec And Timer >= t0
        DoEvents
    Loop
End Sub

' Slayderni a dan b gacha silliq surish (har qadamda 3D qayta chiziladi)
Private Sub Animate(sb As MSForms.ScrollBar, ByVal a As Long, ByVal b As Long)
    Dim k As Long, n As Long
    n = 20
    For k = 0 To n
        busy3D = True
        sb.Value = a + (b - a) * k / n
        busy3D = False
        Update3D
        DoEvents
        Pause 0.02
    Next k
End Sub

Private Sub sbFold_Change()
    Update3D
End Sub

Private Sub sbFold_Scroll()
    Update3D
End Sub

Private Sub sbLid_Change()
    Update3D
End Sub

Private Sub sbLid_Scroll()
    Update3D
End Sub

Private Sub sbYaw_Change()
    Update3D
End Sub

Private Sub sbYaw_Scroll()
    Update3D
End Sub

Private Sub sbPitch_Change()
    Update3D
End Sub

Private Sub sbPitch_Scroll()
    Update3D
End Sub

Private Sub sbZoom_Change()
    Update3D
End Sub

Private Sub sbZoom_Scroll()
    Update3D
End Sub

Private Sub cmbColor_Change()
    Update3D
End Sub


Private Sub btnFoldAnim_Click()
    Animate sbFold, 0, 100
End Sub

Private Sub btnLidAnim_Click()
    If sbLid.Value < 50 Then
        Animate sbLid, sbLid.Value, 100
        btnLidAnim.Caption = "Yopish"
    Else
        Animate sbLid, sbLid.Value, 0
        btnLidAnim.Caption = "Ochish"
    End If
End Sub

Private Function HasItem(c As MSForms.ComboBox, ByVal txt As String) As Boolean
    Dim k As Long
    For k = 0 To c.ListCount - 1
        If c.List(k) = txt Then HasItem = True: Exit Function
    Next k
End Function

Private Sub cmbGroup_Change()
    FillList
End Sub

Private Sub txtFind_Change()
    FillList
End Sub

' Ro'yxatni guruh va qidiruv bo'yicha to'ldirish. Joriy shablon ro'yxatda qolsa, tanlangan bo'lib turadi
Private Sub FillList()
    Dim i As Long, n As Long, g As String, q As String, sel As Long
    g = cmbGroup.Text
    q = LCase(Trim(txtFind.Text))
    lst.Clear
    ReDim ids(0 To TCount)
    sel = -1
    For i = 1 To TCount
        If g = "Hammasi" Or g = TGroup(i) Then
            If q = "" Or InStr(LCase(TGroup(i) & " " & TName(i)), q) > 0 Then
                If g = "Hammasi" Then lst.AddItem TGroup(i) & "  -  " & TName(i) Else lst.AddItem TName(i)
                If i = curT Then sel = n
                ids(n) = i: n = n + 1
            End If
        End If
    Next i
    If sel < 0 And n > 0 Then sel = 0
    If sel >= 0 Then lst.ListIndex = sel
    If n = 0 Then SetStatus "Topilmadi", False Else SetStatus "", False
End Sub

' Maydonlari ko'rsatilgan shablon raqami
Private Function CurId() As Long
    CurId = curT
End Function

Private Sub lst_Change()
    If lst.ListIndex < 0 Then Exit Sub
    If ids(lst.ListIndex) = curT And Not fields Is Nothing Then Exit Sub
    curT = ids(lst.ListIndex)
    BuildFields
    FieldChanged
End Sub

Private Sub BuildFields()
    Dim spec() As String, parts() As String, i As Long, y As Single
    Dim tb As MSForms.TextBox, f As clsField, bg As MSForms.Label, lbl As MSForms.Label

    building = True
    Do While fra.Controls.Count > 0
        fra.Controls.Remove fra.Controls(0).Name
    Loop
    Set fields = New Collection

    NewLabel fra, "O'lcham", PAD, 8, 150, 16, True, C_TEXT, 10
    NewLabel fra, "mm", 200, 10, 28, 14, False, C_MUTED, 8
    spec = Split(TParams(CurId), "|")
    For i = 0 To UBound(spec)
        parts = Split(spec(i), "=")
        y = 34 + i * (ROW_H + 6)
        Set lbl = NewLabel(fra, parts(1), PAD, y + 4, 92, 16)
        If UBound(parts) >= 3 Then lbl.ControlTipText = parts(3)

        ' maydon foni: [-] qiymat [+]
        Set bg = NewButton(fra, "", 108, y, 120, ROW_H, C_FIELD, C_TEXT)
        Set f = New clsField
        Set f.btnMinus = NewButton(fra, "-", 108, y + 2, 24, ROW_H - 4, C_FIELD, C_MUTED, 12)
        Set f.btnPlus = NewButton(fra, "+", 204, y + 2, 24, ROW_H - 4, C_FIELD, C_MUTED, 12)
        Set tb = fra.Controls.Add("Forms.TextBox.1")
        With tb
            .Left = 134: .Top = y + 3: .Width = 68: .Height = ROW_H - 6
            .BorderStyle = fmBorderStyleNone: .SpecialEffect = fmSpecialEffectFlat
            .BackColor = C_FIELD: .ForeColor = C_TEXT
            .TextAlign = fmTextAlignCenter
            .Font.Name = FONT_NAME: .Font.Size = 9: .Font.Bold = True
            .Text = parts(2)
        End With
        Set f.tb = tb: f.Key = parts(0): f.StepV = 1: Set f.Owner = Me
        f.MinV = 1
        If UBound(parts) >= 4 Then f.MinV = Val(parts(4))
        fields.Add f
    Next i
    building = False
End Sub

' Maydonlardagi qiymatlar. Noto'g'ri qiymat bo'lsa Nothing
Private Function ReadParams() As Object
    Dim p As Object, f As clsField, v As Double
    Set p = CreateObject("Scripting.Dictionary")
    For Each f In fields
        v = Val(Replace(f.tb.Text, ",", "."))
        If v < f.MinV Or (f.MinV > 0 And v <= 0) Then
            f.tb.ForeColor = C_ERR
            SetStatus "Qiymat noto'g'ri", True
            Exit Function
        End If
        f.tb.ForeColor = C_TEXT
        p(f.Key) = v
    Next f
    Set ReadParams = p
End Function

Private Function Caption_(p As Object) As String
    Dim k As Variant, dims As String
    For Each k In Array("L", "W", "H", "A", "DI")
        If p.Exists(k) Then
            If dims <> "" Then dims = dims & " x "
            dims = dims & p(k)
        End If
    Next k
    Caption_ = TName(CurId) & "  " & dims & " mm"
End Function

Private Sub SetStatus(ByVal txt As String, ByVal isErr As Boolean)
    lblStatus.Caption = txt
    If isErr Then lblStatus.ForeColor = C_ERR Else lblStatus.ForeColor = C_MUTED
End Sub

Public Sub FieldChanged()
    Dim p As Object
    If building Or fields Is Nothing Then Exit Sub
    Set p = ReadParams()
    If p Is Nothing Then Exit Sub
    SetStatus "", False
    If Not chkLive.Value Then Exit Sub
    TBuild CurId, p
    DrawDie False, Caption_(p)
    Update3D
End Sub

Private Sub chk3D_Click()
    Update3D
End Sub

Private Sub chkLive_Click()
    If chkLive.Value Then
        FieldChanged
    Else
        DeletePreview
        If Not ActiveWindow Is Nothing Then ActiveWindow.Refresh
    End If
End Sub

Private Sub btnDraw_Click()
    Dim p As Object
    If curT = 0 Or fields Is Nothing Then Exit Sub
    Set p = ReadParams()
    If p Is Nothing Then Exit Sub
    TBuild CurId, p
    DrawDie True, Caption_(p)
    Update3D True                  ' 3D ning doimiy nusxasi
    SetStatus "Chizildi", False
    chkLive.Value = False          ' tayyor chizma ustiga preview chizilmasin
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    DeletePreview
    If Not ActiveWindow Is Nothing Then ActiveWindow.Refresh
End Sub
