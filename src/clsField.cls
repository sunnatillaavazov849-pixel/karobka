VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Bitta o'lcham maydoni: o'zgarganda formaga xabar beradi
Option Explicit

Public WithEvents tb As MSForms.TextBox
Public Key As String
Public Owner As Object

Private Sub tb_Change()
    Owner.FieldChanged
End Sub
