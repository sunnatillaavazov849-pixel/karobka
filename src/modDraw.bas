Attribute VB_Name = "modDraw"
' Natijani CorelDRAW sahifasiga chizish: yoyilma (DrawDie) va uning yonida 3D (Redraw3D).
' 3D alohida guruh - slayderlar o'zgarganda faqat u qayta chiziladi.
Option Explicit
Option Private Module          ' Macros ro'yxatida faqat Karobka ko'rinsin

Public Const PREVIEW_NAME As String = "KAROBKA_PREVIEW"
Public Const D3_NAME As String = "KAROBKA_3D"

' Oxirgi chizilgan yoyilma chegarasi (3D shu yerning o'ng tomoniga qo'yiladi)
Private dieL As Double, dieR As Double, dieB As Double, dieT As Double, haveDie As Boolean
Private dieGroup As Shape                         ' oxirgi yoyilma guruhi (mockup koordinata boshi uchun)
Public Const DESIGN_LAYER As String = "Dizayn"

Private Sub DeleteNamed(ByVal nm As String)
    Dim s As Shape
    If ActiveDocument Is Nothing Then Exit Sub
    Set s = ActivePage.FindShape(Name:=nm)
    Do While Not s Is Nothing
        s.Delete
        Set s = ActivePage.FindShape(Name:=nm)
    Loop
End Sub

Public Sub DeletePreview()
    DeleteNamed PREVIEW_NAME
    DeleteNamed D3_NAME
End Sub

Public Sub Delete3D()
    DeleteNamed D3_NAME
    If Not ActiveWindow Is Nothing Then ActiveWindow.Refresh
End Sub

Private Function MakeCurve(ByVal kind As Integer) As Shape
    Dim crv As Curve, sp As SubPath, i As Long, any_ As Boolean
    Set crv = Application.CreateCurve(ActiveDocument)
    For i = 0 To segN - 1
        If segKind(i) = kind Then
            Set sp = crv.CreateSubPath(segX1(i), segY1(i))
            sp.AppendLineSegment segX2(i), segY2(i)
            any_ = True
        End If
    Next i
    If any_ Then Set MakeCurve = ActiveLayer.CreateCurve(crv)
End Function

' TBuild dan keyin chaqiriladi. final = False bo'lsa ko'rish (preview) sifatida chiziladi
Public Sub DrawDie(ByVal final As Boolean, ByVal caption As String)
    Dim sr As New ShapeRange, sCut As Shape, sFold As Shape, txt As Shape, g As Shape, org As Shape

    If ActiveDocument Is Nothing Then CreateDocument
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup "Karobka"
    Optimization = True
    On Error GoTo Done

    DeleteNamed PREVIEW_NAME
    If segN = 0 Then GoTo Done

    Set sCut = MakeCurve(KIND_CUT)
    If Not sCut Is Nothing Then
        sCut.Name = "Kesish"
        sCut.Outline.SetProperties Width:=0.25, Color:=CreateRGBColor(0, 90, 255)
        sr.Add sCut
    End If
    Set sFold = MakeCurve(KIND_FOLD)
    If Not sFold Is Nothing Then
        sFold.Name = "Bukish"
        sFold.Outline.SetProperties Width:=0.25, Color:=CreateRGBColor(230, 0, 40)
        On Error Resume Next
        sFold.Outline.Style = OutlineStyles(3)    ' punktir
        On Error GoTo Done
        sr.Add sFold
    End If
    If sr.Count = 0 Then GoTo Done

    ' yoyilma koordinata boshi belgisi (mockup dizaynni shu bo'yicha joylaydi)
    Set org = ActiveLayer.CreateRectangle(0, 0, 0.01, 0.01)
    org.Outline.SetNoOutline
    org.Name = "KAROBKA_ORIGIN"
    sr.Add org

    On Error Resume Next
    Set txt = ActiveLayer.CreateArtisticText(sr.LeftX, sr.BottomY - 8, caption, , , "Segoe UI", 10)
    If Not txt Is Nothing Then sr.Add txt
    On Error GoTo Done

    Set g = sr.Group
    g.CenterX = ActivePage.CenterX
    g.CenterY = ActivePage.CenterY
    If final Then g.Name = "Karobka" Else g.Name = PREVIEW_NAME
    dieL = g.LeftX: dieR = g.RightX: dieB = g.BottomY: dieT = g.TopY: haveDie = True
    Set dieGroup = g

Done:
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
End Sub

' 3D ni yoyilma yonida qayta chizadi (mod3D dagi joriy holat: Fold3D dan keyin).
' yaw/pitch - aylantirish va qiyalik (gradus), zoom - masshtab (1 = yoyilma balandligining 60%),
' rgbColor - karton rangi, final = True bo'lsa doimiy nusxa qoladi
Public Sub Redraw3D(ByVal yaw As Double, ByVal pitch As Double, ByVal zoom As Double, ByVal rgbColor As Long, _
                    Optional ByVal final As Boolean = False, Optional dsg As Object = Nothing)
    Dim sr As New ShapeRange, g As Shape, org As Shape, ox As Double, oy As Double
    If ActiveDocument Is Nothing Or Not haveDie Then Exit Sub
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup "Karobka 3D"
    Optimization = True
    On Error GoTo Done
    DeleteNamed D3_NAME
    If LCase(ActiveLayer.Name) = LCase(DESIGN_LAYER) Then OtherLayer.Activate   ' 3D dizayn qatlamiga tushmasin
    If Not dsg Is Nothing Then
        Set org = dieGroup.Shapes.FindShape(Name:="KAROBKA_ORIGIN")
        ox = org.LeftX: oy = org.BottomY
    End If
    Draw3D sr, dieR + 40, dieB, (dieT - dieB) * 0.6 * zoom, yaw, pitch, rgbColor, dsg, ox, oy
    If sr.Count > 0 Then
        Set g = sr.Group
        If final Then g.Name = "Karobka 3D" Else g.Name = D3_NAME
    End If
Done:
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
End Sub

' Dizayn qatlami (yo'q bo'lsa yaratiladi). created = True - hozir yaratildi
Public Function DesignLayer(created As Boolean) As Object
    Dim ly As Object
    created = False
    If ActiveDocument Is Nothing Then Exit Function
    For Each ly In ActivePage.Layers
        If LCase(ly.Name) = LCase(DESIGN_LAYER) Then Set DesignLayer = ly: Exit Function
    Next ly
    Set DesignLayer = ActivePage.CreateLayer(DESIGN_LAYER)
    created = True
End Function

Private Function OtherLayer() As Object
    Dim ly As Object
    For Each ly In ActivePage.Layers
        If LCase(ly.Name) <> LCase(DESIGN_LAYER) And Not ly.IsSpecialLayer Then Set OtherLayer = ly: Exit Function
    Next ly
    Set OtherLayer = ActivePage.CreateLayer("Karobka")
End Function

' Yoyilma bormi (mockup uchun)
Public Function HasDie() As Boolean
    HasDie = haveDie And Not dieGroup Is Nothing
End Function
