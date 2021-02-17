Attribute VB_Name = "Module1"
Public pausetimer As Integer
Public vy As Integer 'determines up and down direction movement
Public vx As Integer 'deteermines left and right direction movement
Public dayctr As Integer 'counts days
Public actctr As Integer 'counts actions per day and finishes day after 5 actions
Public battlectr As Integer 'counts # of battles encountered
Public slvl As Integer 'counts sword level
Public Fns As Single
Public health As Long
Public Thunderbolt As Integer 'relationship with gods
Public Trident As Integer
Public Bow As Integer
Public Sub playerstats() 'when start game is clicked, resets all stats

'starting stats
Range("H11") = 1    'slvl (changes values in cells to startin stats)
Range("H3") = 10    'dfn
Range("H4") = 20    'atk
Range("H12") = 5    'fns
Range("H2") = 100 'health


dayctr = 1
battlectr = 1

'calculated stats
permhealth = Cells(3, 8) * 10 'health calculated thru dfn
permdfn = Cells(3, 8) 'dfn based on starting stats
permatk = Cells(4, 8) 'atk based on...
Fns = Cells(12, 8) 'fns based on...
slvl = Cells(11, 8) 'slvl based on...

'relationship with gods
Thunderbolt = 1
Trident = 1
Bow = 1

End Sub

'MOVEMENT CODES

'movement codes taken from ETV and modified for group project use
'we did not make the codes and we acknowledge the creators of ETV as the actual creators of the codes

Sub mover()

e1
e2
e3
e4

'responsible for player movement
lefryt
topbot

'THE TIMER--> most important for the seemingly infinite looping
Dim begin As Double

'begin is the current time and will not change
begin = Timer 'timer is a function that shows realtime in Integer form
    Sheet1.Cells(1, 16) = begin
Do While Timer < (begin + 0.1) '0.1 is the delay between the loop's repeat
DoEvents 'it will do events while the timer is .1 second less than the initial
Loop 'overall, it will repeat all events every 0.1 seconds

End Sub

Sub moving()

'will keep doing the mover sub as long as the game is unpaused

Do While pausetimer = 0
Call mover
Loop

End Sub

Sub e1()

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
Randomize
a = Rnd() * 8
b = Rnd() * 8
c = Rnd()
d = Rnd()
If c = 0 Then
e = -1
Else
e = 1
End If
If d = 0 Then
f = -1
Else
f = 1
End If


If arenaform.mp.Left > (arenaform.s1.Left - 12) And arenaform.mp.Left < (arenaform.s1.Left + 24) And arenaform.mp.Top > (arenaform.s1.Top - 12) And arenaform.mp.Top < (arenaform.s1.Top + 24) Then
'stops if in range
ps = 0
Else
ps = 1
End If

If arenaform.s1.Visible = True Then

    'stops the player when they reach the boundaries set
    If arenaform.s1.Top < 90 Then
    arenaform.s1.Top = 90
    Else
    End If
    If arenaform.s1.Top > 342 Then
    arenaform.s1.Top = 342
    Else
    End If
    If arenaform.s1.Left > 516 Then
    arenaform.s1.Left = 516
    Else
    End If
    If arenaform.s1.Left < 90 Then
    arenaform.s1.Left = 90
    Else
    End If

    'moving the position of the character
    
    arenaform.s1.Left = arenaform.s1.Left + a * e * ps
    arenaform.s1.Top = arenaform.s1.Top + b * f * ps

    If ps = 0 Then 'proximity condition is met then..?
    
    AMainFightWindow.Show 'if u win
    arenaform.s1.Visible = False 'enemy becomes invisible
        
    Else
    End If

Else

    arenaform.s1.Left = 606
    arenaform.s1.Top = 204
    arenaform.s1.Visible = False
    If arenaform.s1.Visible = False And arenaform.s2.Visible = False And arenaform.s3.Visible = False And arenaform.s4.Visible = False Then
        
        arenaform.Hide
        MsgBox "You win."
        battlectr = battlectr + 1
        
        If battlectr = 10 Then
            innaresform.Show
        End If
        
        mapform.Show
        
    End If

End If

End Sub
Sub e2()

'checks whether 2nd...nth enemy should be present depending on how many battles the mc went through already

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
Randomize
a = Rnd() * 8
b = Rnd() * 8
c = Rnd()
d = Rnd()
If c = 0 Then
e = -1
Else
e = 1
End If
If d = 0 Then
f = -1
Else
f = 1
End If

If arenaform.s2.Visible = True Then

    If arenaform.mp.Left > (arenaform.s2.Left - 12) And arenaform.mp.Left < (arenaform.s2.Left + 24) And arenaform.mp.Top > (arenaform.s2.Top - 12) And arenaform.mp.Top < (arenaform.s2.Top + 24) Then
    'stops if in range
    ps = 0
    Else
    ps = 1
    End If
    
    'stops the player when they reach the boundaries set
    If arenaform.s2.Top < 90 Then
    arenaform.s2.Top = 90
    Else
    End If
    If arenaform.s2.Top > 342 Then
    arenaform.s2.Top = 342
    Else
    End If
    If arenaform.s2.Left > 516 Then
    arenaform.s2.Left = 516
    Else
    End If
    If arenaform.s2.Left < 90 Then
    arenaform.s2.Left = 90
    Else
    End If
    
    'moving the position of the character
    
    arenaform.s2.Left = arenaform.s2.Left + a * e * ps
    arenaform.s2.Top = arenaform.s2.Top + b * f * ps

    If ps = 0 Then 'proximity condition is met then..?
    
    AMainFightWindow.Show 'you battle
    arenaform.s2.Visible = False 'if it does exit, enemy should die
    arenaform.s2.Left = 300
    arenaform.s2.Top = 436
        
    End If
    
Else

    arenaform.s2.Left = 300
    arenaform.s2.Top = 436
    arenaform.s2.Visible = False
    If arenaform.s1.Visible = False And arenaform.s2.Visible = False And arenaform.s3.Visible = False And arenaform.s4.Visible = False Then
        
        arenaform.Hide
        MsgBox "You win."
        battlectr = battlectr + 1
        
        If battlectr = 10 Then
            innaresform.Show
        End If
        
        mapform.Show
        
    End If
    
End If

End Sub
Sub e3()

'checks whether 2nd...nth enemy should be present depending on how many battles the mc went through already

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
Randomize
a = Rnd() * 8
b = Rnd() * 8
c = Rnd()
d = Rnd()
If c = 0 Then
e = -1
Else
e = 1
End If
If d = 0 Then
f = -1
Else
f = 1
End If

If arenaform.s3.Visible = True Then

    If arenaform.mp.Left > (arenaform.s3.Left - 12) And arenaform.mp.Left < (arenaform.s3.Left + 24) And arenaform.mp.Top > (arenaform.s3.Top - 12) And arenaform.mp.Top < (arenaform.s3.Top + 24) Then
    'stops if in range
    ps = 0
    Else
    ps = 1
    End If
    
    'stops the player when they reach the boundaries set
    If arenaform.s3.Top < 90 Then
    arenaform.s3.Top = 90
    Else
    End If
    If arenaform.s3.Top > 342 Then
    arenaform.s3.Top = 342
    Else
    End If
    If arenaform.s3.Left > 516 Then
    arenaform.s3.Left = 516
    Else
    End If
    If arenaform.s3.Left < 90 Then
    arenaform.s3.Left = 90
    Else
    End If
    
    'moving the position of the character
    
    arenaform.s3.Left = arenaform.s3.Left + a * e * ps
    arenaform.s3.Top = arenaform.s3.Top + b * f * ps

    If ps = 0 Then 'proximity condition is met then..?
    
    AMainFightWindow.Show
    arenaform.s3.Visible = False
    
    End If
        
Else

    arenaform.s3.Left = 288
    arenaform.s3.Top = 12
    arenaform.s3.Visible = False
    If arenaform.s1.Visible = False And arenaform.s2.Visible = False And arenaform.s3.Visible = False And arenaform.s4.Visible = False Then
        
        arenaform.Hide
        MsgBox "You win."
        battlectr = battlectr + 1
        
        If battlectr = 10 Then
            innaresform.Show
        End If
        
        mapform.Show
        
    End If
    
End If

End Sub
Sub e4()

'checks whether 2nd...nth enemy should be present depending on how many battles the mc went through already

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
Randomize
a = Rnd() * 8
b = Rnd() * 8
c = Rnd()
d = Rnd()
If c = 0 Then
e = -1
Else
e = 1
End If
If d = 0 Then
f = -1
Else
f = 1
End If

If arenaform.s4.Visible = True Then

    If arenaform.mp.Left > (arenaform.s4.Left - 12) And arenaform.mp.Left < (arenaform.s4.Left + 24) And arenaform.mp.Top > (arenaform.s4.Top - 12) And arenaform.mp.Top < (arenaform.s4.Top + 24) Then
    'stops if in range
    ps = 0
    Else
    ps = 1
    End If
    
    'stops the player when they reach the boundaries set
    If arenaform.s4.Top < 90 Then
    arenaform.s4.Top = 90
    Else
    End If
    If arenaform.s4.Top > 342 Then
    arenaform.s4.Top = 342
    Else
    End If
    If arenaform.s4.Left > 516 Then
    arenaform.s4.Left = 516
    Else
    End If
    If arenaform.s4.Left < 90 Then
    arenaform.s4.Left = 90
    Else
    End If
    
    'moving the position of the character
    
    arenaform.s4.Left = arenaform.s4.Left + a * e * ps
    arenaform.s4.Top = arenaform.s4.Top + b * f * ps

    If ps = 0 Then 'proximity condition is met then..?
    
    AMainFightWindow.Show
    arenaform.s4.Visible = False
        
    End If
    
Else

    arenaform.s4.Left = 6
    arenaform.s4.Top = 204
    arenaform.s4.Visible = False
    If arenaform.s1.Visible = False And arenaform.s2.Visible = False And arenaform.s3.Visible = False And arenaform.s4.Visible = False Then
        
        arenaform.Hide
        MsgBox "You win."
        battlectr = battlectr + 1
        
        If battlectr = 10 Then
            innaresform.Show
        End If
        
        mapform.Show
        
    End If
    
End If

End Sub

'PLAYERMOVEMENT CODE

'sub lefryt is for left right motion
'vx is for the motion along x axis

Sub lefryt()

'boundaries for left right motion

If arenaform.mp.Left < 90 Then
arenaform.mp.Left = 90
Else
End If
If arenaform.mp.Left > 516 Then
arenaform.mp.Left = 516
Else
End If

'actual movement along x-axis

arenaform.mp.Left = arenaform.mp.Left + vx * Fns

End Sub

'sub topbot is for up and down motion
'vy is for the motion along y axis

Sub topbot()

'boundaries for top bottom motion

If arenaform.mp.Top < 90 Then
arenaform.mp.Top = 90
arenaform.mp.Enabled = True
Else
End If
If arenaform.mp.Top > 342 Then
arenaform.mp.Top = 342
arenaform.mp.Enabled = True
Else
End If

'actual movement

arenaform.mp.Top = arenaform.mp.Top + vy * Fns

End Sub

'GAME TIME CODES

Sub dayactchecker() 'disables buttons when day is done and forces player to move on to next day
    
    If actctr >= 5 Then
        
        'Maps
        mapform.blacksmithbut.Enabled = False
        mapform.arenabut.Enabled = False
        mapform.guardtbut.Enabled = False
        mapform.trainingbut.Enabled = False
        mapform.innbut.Enabled = True
        
        'Blacksmith
        blacksmithform.upgradesw.Enabled = False
        
        'Temple
        templeform.PraytoNeptune.Enabled = False
        templeform.PraytoJupiter.Enabled = False
        templeform.PraytoApollo.Enabled = False
        
        'Training
        trainingform.Atk1.Enabled = False
        trainingform.Def1.Enabled = False
        trainingform.Fns1.Enabled = False
        
    ElseIf actctr = 0 Then
        
        mapform.blacksmithbut.Enabled = True
        mapform.arenabut.Enabled = True
        mapform.guardtbut.Enabled = True
        mapform.trainingbut.Enabled = True
        mapform.innbut.Enabled = True
        blacksmithform.upgradesw.Enabled = True
        
         'Blacksmith
        blacksmithform.upgradesw.Enabled = True
        
        'Temple
        templeform.PraytoNeptune.Enabled = True
        templeform.PraytoJupiter.Enabled = True
        templeform.PraytoApollo.Enabled = True
        
        'Training
        trainingform.Atk1.Enabled = True
        trainingform.Def1.Enabled = True
        trainingform.Fns1.Enabled = True
        
    ElseIf actctr >= 1 Then
    
        mapform.blacksmithbut.Enabled = True
        mapform.arenabut.Enabled = False
        mapform.guardtbut.Enabled = True
        mapform.trainingbut.Enabled = True
        mapform.innbut.Enabled = True
        blacksmithform.upgradesw.Enabled = True
        
         'Blacksmith
        blacksmithform.upgradesw.Enabled = True
        
        'Temple
        templeform.PraytoNeptune.Enabled = True
        templeform.PraytoJupiter.Enabled = True
        templeform.PraytoApollo.Enabled = True
        
        'Training
        trainingform.Atk1.Enabled = True
        trainingform.Def1.Enabled = True
        trainingform.Fns1.Enabled = True
    
    End If
    
End Sub
