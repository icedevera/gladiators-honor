VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainmenu 
   Caption         =   "Gladiator's Honor"
   ClientHeight    =   9040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   12580
   OleObjectBlob   =   "mainmenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CommandButton1_Click()

    playerstats
    resetplayer
    Me.Hide
    Cinematic1.Show
    
End Sub
Sub CommandButton2_Click()

    End
    
End Sub

Private Sub Image1_Click()
mainmenu.Hide
creditsform.Show

End Sub
