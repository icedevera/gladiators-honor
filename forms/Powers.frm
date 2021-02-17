VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Powers 
   Caption         =   "UserForm2"
   ClientHeight    =   3390
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "Powers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Powers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim ehealth As Long
Dim sp As Integer


MsgBox ("JUPITER HAS HEEDED YOUR CALL AND CRIPPLED YOUR FOE!")

ehealth = Cells(2, 4).Value
sp = Cells(5, 2).Value

Range("D2").Value = ehealth - Int(0.5 * ehealth)

Call enemyfightsp
Range("B5").Value = sp - 1
Powers.Hide

End Sub

Private Sub CommandButton2_Click()
Dim sp As Integer
Dim atk As Integer

MsgBox ("YOU HAVE GAINED NEPTUNE'S POWER. UNLEASH HIS WRATH!")
sp = Cells(5, 2)
atk = Cells(8, 2)
atkbuff = Cells(6, 2).Value

Range("B6").Value = Int(0.5 * (atk + atkbuff)) + atkbuff
Range("B8").Value = atk + atkbuff

Call enemyfightsp
Range("B5").Value = sp - 1
Powers.Hide

End Sub

Private Sub CommandButton3_Click()
Dim sp As Integer
Dim dfn As Integer

MsgBox ("YOU HAVE GAINED APOLLO'S BLESSING. UNLEASH HIS WRATH!")

sp = Cells(5, 2)
dfn = Cells(9, 2)
dfnbuff = Cells(7, 2).Value

Range("B7").Value = Int(0.5 * (dfn + dfnbuff)) + dfnbuff
Range("B9").Value = dfn + dfnbuff


Call enemyfightsp
Range("B5").Value = sp - 1
Powers.Hide

End Sub
