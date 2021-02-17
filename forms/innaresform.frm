VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} innaresform 
   Caption         =   "UserForm1"
   ClientHeight    =   5900
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7670
   OleObjectBlob   =   "innaresform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "innaresform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub restbutares_Click()

    resetplayer
    dayctr = dayctr + 1
    actctr = 0
    Call Module1.dayactchecker
    MsgBox "You have rested. 1 day has passed."
    innaresform.Hide
    Cinematic5.Show
    
End Sub

    
