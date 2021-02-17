Attribute VB_Name = "MainFight"
Sub enemyfightatk()

Dim moves As Integer
Dim health As Long
Dim ehealth As Long
Dim edfn As Long
Dim eatk As Long
Dim atk As Long
Dim dfn As Long
Dim first As Double 'determines first turn chance

Randomize
moves = Int(Rnd * 10) + 1
first = Int(Rnd * 10) + 1

'these will reset when resetplayer macros is triggered
health = Cells(2, 2).Value
atk = Cells(4, 2).Value
dfn = Cells(3, 2).Value
atkbuff = Cells(6, 2).Value
dfnbuff = Cells(7, 2).Value
swordatk = Cells(4, 8) + (Cells(4, 8) * (Cells(11, 8) / 4))
Fns = Cells(12, 8).Value

edfn = Cells(3, 4).Value
eatk = Cells(4, 4).Value
ehealth = Cells(2, 4).Value
blck = Cells(10, 2).Value

If first > 5 + (Fns / 50) Then


    If moves <= 6 Then              'enemy attacks
    
            If eatk > (dfn + dfnbuff) Then         'health after being attacked by enemy
                Range("b2").Value = health - (eatk - (dfn + dfnbuff))
                MsgBox ("Enemy Attacks You!")
                
                If health - (eatk - (dfn + dfnbuff)) <= 0 Then      'checks if dead
                    MsgBox ("YOU DIED")
                    If battlectr >= 10 Then
                            aresfightwndow.Hide
                            mainmenu.Show
                        Else
                            AMainFightWindow.Hide
                            arenaform.Hide
                            mainmenu.Show
                        End If
                    Exit Sub
                End If
      
            
            Else
                MsgBox ("Your Defense is Too Superior for the enemy!")
            End If
            
        If swordatk + atkbuff > edfn Then
            Range("d2") = ehealth - ((swordatk + atkbuff) - edfn)
                    If ehealth - ((swordatk + atkbuff) - edfn) <= 0 Then
                        MsgBox "You Have Slain Your Opponent!"
                        If battlectr >= 10 Then
                            aresfightwndow.Hide
                            Cinematic6.Show
                        Else
                            AMainFightWindow.Hide
                            enemytype
                        End If
                        Exit Sub
                    End If
            MsgBox ("You Attacked the Enemy")
            
        Else
            MsgBox ("Your Enemy's defense is too high!")
        
        End If
        
    ElseIf moves = 7 Or moves = 8 Then 'enemy defends
        Range("d3").Value = edfn + 10
        MsgBox ("Enemy Charges his Defense!")
        
        If swordatk + atkubff > edfn Then
            Range("d2") = ehealth - ((swordatk + atkbuff) - edfn)
            
                If ehealth - ((swordatk + atkbuff) - edfn) <= 0 Then
                    MsgBox "You Have Slain Your Opponent!"
                    If battlectr >= 10 Then
                            aresfightwndow.Hide
                            Cinematic6.Show
                        Else
                            AMainFightWindow.Hide
                            enemytype
                        End If
                    Exit Sub
                End If
                MsgBox ("You Attacked the Enemy")
            
        Else
            MsgBox ("Your Enemy's defense is too high!")
                
    End If
        
    Else 'enemy charges attack
            Range("d4").Value = eatk + 10
            MsgBox ("Enemy Charges His Attack!")
        
            If swordatk + atkbuff > edfn Then
                Range("d2") = ehealth - ((swordatk + atkbuff) - edfn)
            
                    If ehealth - ((swordatk + atkbuff) - edfn) <= 0 Then
                        MsgBox "You Have Slain Your Opponent!"
                        If battlectr >= 10 Then
                            aresfightwndow.Hide
                            Cinematic6.Show
                        Else
                            AMainFightWindow.Hide
                            enemytype
                        End If
                        Exit Sub
                    End If
                MsgBox ("You Attacked the Enemy")
            
            Else
            
            MsgBox ("Your Enemy's defense is too high!")
            
            End If
        
    End If

Else     'player moves first

    If moves <= 6 Then              'enemy attacks
    
            
            If swordatk + atkbuff > edfn Then
                Range("d2") = ehealth - ((swordatk + atkbuff) - edfn)
                        If ehealth - ((swordatk + atkbuff) - edfn) <= 0 Then
                            MsgBox "You Have Slain Your Opponent!"
                            If battlectr >= 10 Then
                            aresfightwndow.Hide
                            Cinematic6.Show
                        Else
                            AMainFightWindow.Hide
                            enemytype
                        End If
                            Exit Sub
                        End If
                MsgBox ("You Attacked the Enemy")
            
            Else
                MsgBox ("Your Enemy's defense is too high!")
        
            End If
            
            If eatk > (dfn + dfnbuff) Then
                Range("b2").Value = health - (eatk - (dfn + dfnbuff))
                MsgBox ("Enemy Attacks You!")
                
                If health - (eatk - (dfn + dfnbuff)) <= 0 Then
                    MsgBox ("YOU DIED")
                    If battlectr >= 10 Then
                            aresfightwndow.Hide
                            mainmenu.Show
                        Else
                            AMainFightWindow.Hide
                            arenaform.Hide
                            mainmenu.Show
                        End If
                    Exit Sub
                End If
      
            
            Else
                MsgBox ("Your Defense is Too Superior for the enemy!")
            End If
        
    ElseIf moves = 7 Or moves = 8 Then 'enemy charges defense
    
        
        If swordatk + atkubff > edfn Then
            Range("d2") = ehealth - ((atk + atkbuff) - edfn)
            
                If ehealth - ((swordatk + atkbuff) - edfn) <= 0 Then
                    MsgBox "You Have Slain Your Opponent!"
                    If battlectr >= 10 Then
                            aresfightwndow.Hide
                            Cinematic6.Show
                        Else
                            AMainFightWindow.Hide
                            enemytype
                        End If
                    Exit Sub
                End If
                MsgBox ("You Attacked the Enemy")
            
        Else
            MsgBox ("Your Enemy's defense is too high!")
                
        End If
        
        Range("d3").Value = edfn + 10
        MsgBox ("Enemy Charges his Defense!")
        
    Else 'enemy charges attack

            If swordatk + atkbuff > edfn Then
                Range("d2") = ehealth - ((swordatk + atkbuff) - edfn)
            
                    If ehealth - ((swordatk + atkbuff) - edfn) <= 0 Then
                        MsgBox "You Have Slain Your Opponent!"
                        If battlectr >= 10 Then
                            aresfightwndow.Hide
                            Cinematic6.Show
                        Else
                            AMainFightWindow.Hide
                            enemytype
                        End If
                        Exit Sub
                    End If
                MsgBox ("You Attacked the Enemy")
            
            Else
            
            MsgBox ("Your Enemy's defense is too high!")
            
            End If
            
            Range("d4").Value = eatk + 10
            MsgBox ("Enemy Charges His Attack!")
        
        
    End If
End If
    
End Sub

Sub enemyfightsp()
Dim moves As Integer
Dim health As Long
Dim ehealth As Long
Dim edfn As Long
Dim eatk As Long
Dim atk As Long
Dim dfn As Long
Dim first As Double

Randomize
moves = Int(Rnd * 4) + 1
first = Int(Rnd * 10) + 1
health = Cells(2, 2).Value
atk = Cells(4, 2).Value
dfn = Cells(3, 2).Value
blck = Cells(10, 2).Value

edfn = Cells(3, 4).Value
eatk = Cells(4, 4).Value
ehealth = Cells(2, 4).Value

If first > 5 + (Fns / 50) Then 'enemy goes first
    If moves <= 8 Then 'enemy attacks
        If blck = 0 Then
                If eatk > (dfn + dfnbuff) Then
                    Range("b2").Value = health - (eatk - (dfn + dfnbuff))
                    MsgBox ("Enemy Attacks You!")
                    
                    If health - (eatk - (dfn + dfnbuff)) <= 0 Then
                        MsgBox ("YOU DIED")
                        If battlectr >= 10 Then
                            aresfightwndow.Hide
                            mainmenu.Show
                        Else
                            AMainFightWindow.Hide
                            arenaform.Hide
                            mainmenu.Show
                        End If
                        Exit Sub
                    End If
                Else
                
                MsgBox ("Your Defense is Too Superior for the enemy!")
                
                End If
        Else
            MsgBox ("Your Enemy's speed proved your block USELESS!")
            
        End If
         
    ElseIf moves = 7 Or moves = 8 Then 'enemy defends
        Range("d3").Value = edfn + 10
        MsgBox ("Enemy Charges his Defense!")
        
        
    Else 'enemy charges attack
        Range("d4").Value = eatk + 10
        MsgBox ("Enemy Charges His Attack!")
        
        
    End If

Else 'player moves first

    If moves <= 6 Then 'enemy attacks
    If blck = 0 Then
            If eatk > (dfn + dfnbuff) Then
                Range("b2").Value = health - (eatk - (dfn + dfnbuff))
                MsgBox ("Enemy Attacks You!")
                
                If health - (eatk - (dfn + dfnbuff)) <= 0 Then
                    MsgBox ("YOU DIED")
                   If battlectr >= 10 Then
                            aresfightwndow.Hide
                            mainmenu.Show
                        Else
                            AMainFightWindow.Hide
                            arenaform.Hide
                            mainmenu.Show
                        End If
                    Exit Sub
                End If
            Else
            
            MsgBox ("Your Defense is Too Superior for the enemy!")
            
            End If
            
        Else
        
            MsgBox ("You Have Blocked Your Opponent's Attack")
    
        End If
            
        
    ElseIf moves = 7 Or moves = 8 Then 'enemy defends
        Range("d3").Value = edfn + 10
        MsgBox ("Enemy Charges his Defense!")
        
        
    Else 'enemy charges attack
        Range("d4").Value = eatk + 10
        MsgBox ("Enemy Charges His Attack!")
        
        
    End If
    
End If

End Sub

Public Sub Enemyname()

Dim firstname(1 To 12) As String
Dim lastname(1 To 7) As String

For i = 1 To 12
    firstname(i) = Cells(28 + i, 1)

Next i

For i = 1 To 7
    lastname(i) = Cells(28 + i, 2)
    
Next i

Randomize
fn = firstname(Int(Rnd * 12) + 1)
Ln = lastname(Int(Rnd * 7) + 1)

Cells(28, 4) = fn & " " & Ln

End Sub
