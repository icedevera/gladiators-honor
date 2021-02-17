VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} blacksmithform 
   Caption         =   "UserForm1"
   ClientHeight    =   6810
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9560
   OleObjectBlob   =   "blacksmithform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "blacksmithform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub smithexit_Click()

    blacksmithform.Hide
    mapform.Show
    
End Sub

Private Sub upgradesw_Click()

    slvl = Cells(11, 8)
    
If Cells(4, 8).Value >= slvl * 15 Then
    slvl = slvl + 1
    Range("H11") = slvl
    actctr = actctr + 1
    MsgBox "You have upgraded your sword to level " & slvl & "!"
    Call Module1.dayactchecker
Else
    MsgBox "Train first. You are not worthy of a better sword."
End If

End Sub

