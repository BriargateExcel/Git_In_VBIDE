Attribute VB_Name = "modTestCode"
Option Explicit

Public Sub Procedure_One()
    MsgBox "Procedure One"
End Sub

Public Sub Procedure_Two()
    MsgBox "Procedure Two"
End Sub

Public Sub Auto_Open()
    AddNewVBEControls
End Sub

Public Sub Auto_Close()
    DeleteMenuItems
End Sub
