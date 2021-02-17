VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} aresfightwndow 
   Caption         =   "Fight "
   ClientHeight    =   6570
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   8170
   OleObjectBlob   =   "aresfightwndow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "aresfightwndow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
If TextBox4.Value = 0 Then
    Exit Sub
Else
    Powers.Show
End If

End Sub

Private Sub CommandButton2_Click()

Call enemyfightatk

End Sub

Private Sub CommandButton3_Click()

Instructions.Show


End Sub

Private Sub Defend_Click()
block = Cells(10, 2).Value

Range("b10") = 1

Call enemyfightsp

Range("B10") = 0

End Sub

Private Sub Userform_Initialize()

    Call MainFight.Enemyname
    enemytype 'refreshes stats of enemy
    
End Sub
