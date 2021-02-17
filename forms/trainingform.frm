VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} trainingform 
   Caption         =   "UserForm4"
   ClientHeight    =   5930
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9220
   OleObjectBlob   =   "trainingform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "trainingform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Atk1_Click()

  Cells(4, 8).Value = Cells(4, 8).Value + 3 'increases your attack by 3
    actctr = actctr + 1
    MsgBox "Your Attack increased by 3!"
    Call Module1.dayactchecker

End Sub


Private Sub Def1_Click()
    Cells(3, 8).Value = Cells(3, 8).Value + 3 'increases your defense by 3
        actctr = actctr + 1
        MsgBox "Your Defense increased by 3!"
        Call Module1.dayactchecker
End Sub

Private Sub Fns1_Click()
    Cells(12, 8).Value = Cells(12, 8).Value + 1 'increases your finesse by 1
        actctr = actctr + 1
        MsgBox "Your Finesse increased by 1!"
        Call Module1.dayactchecker
End Sub


Private Sub smithexit_Click()

    trainingform.Hide
    mapform.Show
    
End Sub

