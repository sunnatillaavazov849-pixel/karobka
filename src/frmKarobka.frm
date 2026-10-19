VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKarobka 
   Caption         =   "Karobka"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6450
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKarobka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Karobka paneli. Barcha boshqaruv elementlari kod orqali yaratiladi,
' shuning uchun formaga hech narsa qo'yish shart emas.
Option Explicit

Private WithEvents lst As MSForms.ListBox
Private WithEvents chkLive As MSForms.CheckBox
Private WithEvents btnDraw As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private fra As MSForms.Frame
Private lblInfo As MSForms.Label
Private fields As Collection
Private building As Boolean

Private Sub UserForm_Initialize()
    Dim i As Long
    Me.Caption = "Karobka - quti chizmasi"
    Me.Width = 440: Me.Height = 330

    Set lst = Me.Controls.Add("Forms.ListBox.1", "lst")
    lst.Left = 6: lst.Top = 6: lst.Width = 190: lst.Height = 270
    For i = 1 To TCount
        lst.AddItem TName(i)
    Next i

    Set fra = Me.Controls.Add("Forms.Frame.1", "fra")
    fra.Caption = "O'lchamlar (mm)"
    fra.Left = 204: fra.Top = 6: fra.Width = 224: fra.Height = 196

    Set lblInfo = Me.Controls.Add("Forms.Label.1", "lblInfo")
    lblInfo.Left = 206: lblInfo.Top = 206: lblInfo.Width = 220: lblInfo.Height = 24
    lblInfo.Caption = "Ko'k - kesish, qizil - bukish chizig'i"

    Set chkLive = Me.Controls.Add("Forms.CheckBox.1", "chkLive")
    chkLive.Caption = "Jonli ko'rish (sahifada)"
    chkLive.Left = 206: chkLive.Top = 228: chkLive.Width = 200: chkLive.Value = True

    Set btnDraw = Me.Controls.Add("Forms.CommandButton.1", "btnDraw")
    btnDraw.Caption = "Chizish": btnDraw.Default = True
    btnDraw.Left = 206: btnDraw.Top = 252: btnDraw.Width = 108: btnDraw.Height = 24

    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    btnClose.Caption = "Yopish": btnClose.Cancel = True
    btnClose.Left = 320: btnClose.Top = 252: btnClose.Width = 108: btnClose.Height = 24

    lst.ListIndex = 0
End Sub

Private Sub lst_Change()
    If lst.ListIndex < 0 Then Exit Sub
    BuildFields
    FieldChanged
End Sub

Private Sub BuildFields()
    Dim spec() As String, parts() As String, i As Long
    Dim lbl As MSForms.Label, tb As MSForms.TextBox, f As clsField

    building = True
    Do While fra.Controls.Count > 0
        fra.Controls.Remove fra.Controls(0).Name
    Loop
    Set fields = New Collection

    spec = Split(TParams(lst.ListIndex + 1), "|")
    For i = 0 To UBound(spec)
        parts = Split(spec(i), "=")
        Set lbl = fra.Controls.Add("Forms.Label.1")
        lbl.Caption = parts(1) & " (" & parts(0) & ")"
        lbl.Left = 8: lbl.Top = 10 + i * 26: lbl.Width = 130
        Set tb = fra.Controls.Add("Forms.TextBox.1")
        tb.Left = 140: tb.Top = 6 + i * 26: tb.Width = 70
        tb.Text = parts(2)
        Set f = New clsField
        Set f.tb = tb: f.Key = parts(0): Set f.Owner = Me
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
        If v <= 0 Then Exit Function
        p(f.Key) = v
    Next f
    Set ReadParams = p
End Function

Private Function Caption_(p As Object) As String
    Caption_ = lst.Text & "  " & p("L") & " x " & p("W") & " x " & p("H") & " mm"
End Function

Public Sub FieldChanged()
    Dim p As Object
    If building Or fields Is Nothing Then Exit Sub
    If Not chkLive.Value Then Exit Sub
    Set p = ReadParams()
    If p Is Nothing Then Exit Sub
    TBuild lst.ListIndex + 1, p
    DrawDie False, Caption_(p)
End Sub

Private Sub chkLive_Click()
    If chkLive.Value Then FieldChanged Else DeletePreview: ActiveWindow.Refresh
End Sub

Private Sub btnDraw_Click()
    Dim p As Object
    Set p = ReadParams()
    If p Is Nothing Then
        MsgBox "O'lchamlar musbat son bo'lishi kerak.", vbExclamation, "Karobka"
        Exit Sub
    End If
    TBuild lst.ListIndex + 1, p
    DrawDie True, Caption_(p)
    chkLive.Value = False          ' tayyor chizma ustiga preview chizilmasin
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    DeletePreview
    If Not ActiveWindow Is Nothing Then ActiveWindow.Refresh
End Sub
