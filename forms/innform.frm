VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} innform 
   Caption         =   "UserForm2"
   ClientHeight    =   6080
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7650
   OleObjectBlob   =   "innform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "innform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub innexit_Click()

    innform.Hide
    mapform.Show
    
End Sub


Private Sub restbut_Click()

    resetplayer
    arenaform.s1.Visible = True
    arenaform.s2.Visible = True
    dayctr = dayctr + 1
    actctr = 0
    Call Module1.dayactchecker
    MsgBox "You have rested. 1 day has passed."
    
End Sub
