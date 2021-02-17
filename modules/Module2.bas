Attribute VB_Name = "Module2"
Public Sub resetenemy1()
Attribute resetenemy1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' resetenemy Macro
'

'
    Range("J1:J5").Select
    Selection.Copy
    Range("D1:D5").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("D8").Select
End Sub


Public Sub enemytype()

If battlectr = 1 Then
    resetenemy1
    
ElseIf battlectr = 2 Then
    resetenemy2
    
ElseIf battlectr = 3 Then
    resetenemy3
    
ElseIf battlectr = 4 Then
    resetenemy4
    
ElseIf battlectr = 5 Then
    resetenemy5
    
ElseIf battlectr = 6 Then
    resetenemy6

ElseIf battlectr = 7 Then
    resetenemy7
    
ElseIf battlectr = 8 Then
    resetenemy8

ElseIf battlectr = 9 Then
    resetenemy9
    
ElseIf battlectr >= 10 Then
    aresthegod
    
End If

End Sub
Sub resetenemy2()
Attribute resetenemy2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' resetenemy2 Macro
'

'
    Range("J7:J11").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub resetenemy3()
Attribute resetenemy3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' resetenemy3 Macro
'

'
    Range("J13:J17").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub resetenemy4()
Attribute resetenemy4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' resetenemy4 Macro
'

'
    Range("J19:J23").Select
    Selection.Copy
    Range("H18").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub resetplayer()
'
' resetplayer Macro
'

'
    Range("H2:H13").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub

Sub aresthegod()
'
' aresthegod Macro
'

'
    Range("H15:H19").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub resetenemy5()
'
' resetenemy5 Macro
'

'
    Range("L1:L5").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub resetenemy6()
'
' resetenemy6 Macro
'

'
    Range("L7:L11").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub resetenemy7()
'
' resetenemy7 Macro
'

'
    Range("L13:L17").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub resetenemy8()
'
' resetenemy8 Macro
'

'
    Range("L19:L23").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
End Sub
Sub resetenemy9()
'
' resetenemy9 Macro
'

'
    Range("N1:N5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub


