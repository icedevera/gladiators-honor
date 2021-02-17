VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} arenaform 
   Caption         =   "Colloseum"
   ClientHeight    =   9050
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   12580
   OleObjectBlob   =   "arenaform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "arenaform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Userform_Initialize()

    enemytype 'makes sure even first enemy stats is appropriate to battlectr
    
End Sub
Private Sub CommandButton5_Click()

    'pause and resume. can also be resumed by moving the character.
If pausetimer = 0 Then
pausetimer = 1
CommandButton5.Caption = "Resume"
Else
CommandButton5.Caption = "Pause"
pausetimer = 0
Module1.moving
End If
    
End Sub

Private Sub vDown_Click()
CommandButton5.Caption = "Pause"
vx = 0
vy = 1
pausetimer = 0
Module1.moving
End Sub

Private Sub vLeft_Click()
CommandButton5.Caption = "Pause"
vx = -1
vy = 0
pausetimer = 0
Module1.moving
End Sub

Private Sub vRight_Click()
vx = 1
vy = 0
CommandButton5.Caption = "Pause"
pausetimer = 0
Module1.moving
End Sub


Private Sub vUp_Click()
CommandButton5.Caption = "Pause"
vy = -1
vx = 0

pausetimer = 0
Module1.moving
End Sub

