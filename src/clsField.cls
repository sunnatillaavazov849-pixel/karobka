VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Bitta o'lcham maydoni: [-] qiymat [+]. O'zgarganda formaga xabar beradi
Option Explicit

Public WithEvents tb As MSForms.TextBox
Public WithEvents btnMinus As MSForms.Label
Public WithEvents btnPlus As MSForms.Label
Public Key As String
Public StepV As Double
Public Owner As Object

Private Sub tb_Change()
    Owner.FieldChanged
End Sub

Private Sub Bump(ByVal d As Double)
    Dim v As Double
    v = Val(Replace(tb.Text, ",", ".")) + d
    If v < StepV Then v = StepV
    tb.Text = Replace(CStr(v), ",", ".")
End Sub

Private Sub btnMinus_Click()
    Bump 0 - StepV
End Sub

Private Sub btnPlus_Click()
    Bump StepV
End Sub
