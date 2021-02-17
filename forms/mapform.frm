VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mapform 
   Caption         =   "UserForm1"
   ClientHeight    =   7110
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9850
   OleObjectBlob   =   "mapform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mapform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub arenabut_Click()

    actctr = actctr + 5
    Call Module1.dayactchecker
    mapform.Hide
    
    If battlectr >= 2 Then
    arenaform.s2.Visible = True
    ElseIf battlectr >= 10 Then
    arenaform.s2.Visible = False
    Else
    arenaform.s2.Visible = False
    End If
    
    If battlectr >= 3 Then
    arenaform.s3.Visible = True
    ElseIf battlectr >= 10 Then
    arenaform.s3.Visible = False
    Else
    arenaform.s3.Visible = False
    End If
    
    If battlectr >= 4 Then
    arenaform.s4.Visible = True
    ElseIf battlectr >= 10 Then
    arenaform.s4.Visible = False
    Else
    arenaform.s4.Visible = False
    End If
    
    arenaform.mp.Left = 300
    arenaform.mp.Top = 198
    arenaform.s1.Left = 606
    arenaform.s1.Top = 204
    arenaform.s2.Left = 300
    arenaform.s2.Top = 436
    arenaform.s3.Left = 288
    arenaform.s3.Top = 12
    arenaform.s4.Left = 6
    arenaform.s4.Top = 204
    
    enemytype 'resets enemy
    arenaform.Show
    
End Sub

Private Sub blacksmithbut_Click()
    
    mapform.Hide
    blacksmithform.Show
    
End Sub

Private Sub guardtbut_Click()
    
    mapform.Hide
    templeform.Show
    
End Sub

Private Sub innbut_Click()

    mapform.Hide
    innform.Show
    
End Sub

Private Sub trainingbut_Click()

    mapform.Hide
    trainingform.Show
    
End Sub
