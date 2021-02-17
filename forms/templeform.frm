VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} templeform 
   Caption         =   "UserForm3"
   ClientHeight    =   6600
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9100
   OleObjectBlob   =   "templeform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "templeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim praychance As Integer
Private Sub PraytoNeptune_Click()
Randomize
praychance = Int(Rnd * 20) + 1

If praychance = 1 Then
    permatk = Cells(4, 8).Value
    Range("H4") = permatk + 20
    MsgBox "Neptune has listened to your grovelling!"
    
Else
    MsgBox ("Nothing seemed to happen.")
End If
    

    actctr = actctr + 1
    Call Module1.dayactchecker

End Sub

Private Sub PraytoJupiter_Click()
    
Dim permhealth As Integer

Randomize
praychance = Int(Rnd * 20) + 1

If praychance = 1 Then
    permhealth = Cells(2, 8).Value
    Range("H2") = permhealth + 50
    MsgBox "Jupiter has listened to your grovelling!"

Else
    MsgBox ("Nothing seemed to happen.")
End If
    
    actctr = actctr + 1
    Call Module1.dayactchecker


End Sub

Private Sub PraytoApollo_Click()

Randomize
praychance = Int(Rnd * 20) + 1

If praychance = 1 Then
    permdef = Cells(3, 8).Value
    Range("H3") = permdef + 20
    MsgBox "Apollo has listened to your grovelling!"

Else
    MsgBox ("Nothing seemed to happen.")
End If
    
    actctr = actctr + 1
    Call Module1.dayactchecker

End Sub

Private Sub smithexit_Click()

    templeform.Hide
    mapform.Show
    
End Sub

