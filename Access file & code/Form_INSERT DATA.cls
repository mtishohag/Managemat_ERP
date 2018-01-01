VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_INSERT DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command111_Click()


' Engineering Drawing

If EN_D_PC >= 20 And EN_D_PF >= 20 Then
EN_D_TOTAL = EN_D_PC + EN_D_PF
Else: EN_D_TOTAL = 0
End If

If EN_D_PC < 20 Then
EN_D_PC.BackColor = vbRed
Else
EN_D_PC.BackColor = vbWhite
End If
If EN_D_PF < 20 Then
EN_D_PF.BackColor = vbRed
Else
EN_D_PF.BackColor = vbWhite
End If


If EN_D_TOTAL = 0 Then
EN_D_TOTAL.BackColor = vbRed
Else
EN_D_TOTAL.BackColor = 12632256
End If


EN_D = EN_D_TOTAL
If EN_D >= 80 Then
EN_D_LG = "A+"
EN_D_GP = 4
ElseIf EN_D >= 75 Then
EN_D_LG = "A"
EN_D_GP = 3.75
ElseIf EN_D >= 70 Then
EN_D_LG = "A-"
EN_D_GP = 3.5
ElseIf EN_D >= 65 Then
EN_D_LG = "B+"
EN_D_GP = 3.25
ElseIf EN_D >= 60 Then
EN_D_LG = "B"
EN_D_GP = 3
ElseIf EN_D >= 55 Then
EN_D_LG = "B-"
EN_D_GP = 2.75
ElseIf EN_D >= 20 Then
EN_D_LG = "C+"
EN_D_GP = 2.5
ElseIf EN_D >= 45 Then
EN_D_LG = "C"
EN_D_GP = 2.25
ElseIf EN_D >= 40 Then
EN_D_LG = "D"
EN_D_GP = 2
Else: EN_D_LG = "F"
EN_D_GP = 0
End If

' Engineering Materials

EN_M_T_TOTAL = EN_M_TC + EN_M_TF
If EN_M_T_TOTAL >= 40 And EN_M_PC >= 10 And EN_M_PF >= 10 Then
EN_M_TOTAL = EN_M_T_TOTAL + EN_M_PC + EN_M_PF
Else
EN_M_TOTAL = 0
End If

If EN_M_T_TOTAL < 40 Then
EN_M_T_TOTAL.BackColor = vbRed
Else
EN_M_T_TOTAL.BackColor = 12632256
End If
If EN_M_PC < 10 Then
EN_M_PC.BackColor = vbRed
Else
EN_M_PC.BackColor = vbWhite
End If
If EN_M_PF < 10 Then
EN_M_PF.BackColor = vbRed
Else
EN_M_PF.BackColor = vbWhite
End If
If EN_M_TOTAL = 0 Then
EN_M_TOTAL.BackColor = vbRed
Else
EN_M_TOTAL.BackColor = 12632256
End If

EN_M = (EN_M_TOTAL / 3) * 2
If EN_M >= 80 Then
EN_M_LG = "A+"
EN_M_GP = 4
ElseIf EN_M >= 75 Then
EN_M_LG = "A"
EN_M_GP = 3.75
ElseIf EN_M >= 70 Then
EN_M_LG = "A-"
EN_M_GP = 3.5
ElseIf EN_M >= 65 Then
EN_M_LG = "B+"
EN_M_GP = 3.25
ElseIf EN_M >= 60 Then
EN_M_LG = "B"
EN_M_GP = 3
ElseIf EN_M >= 55 Then
EN_M_LG = "B-"
EN_M_GP = 2.75
ElseIf EN_M >= 50 Then
EN_M_LG = "C+"
EN_M_GP = 2.5
ElseIf EN_M >= 45 Then
EN_M_LG = "C"
EN_M_GP = 2.25
ElseIf EN_M >= 40 Then
EN_M_LG = "D"
EN_M_GP = 2
Else: EN_M_LG = "F"
EN_M_GP = 0
End If


' English-1

ENG_TOTAL = ENG_TC + ENG_TF
'If ENG_TOTAL < 40 Then
'ENG_TOTAL = 0
'Else
'End If

If ENG_TOTAL < 40 Then
ENG_TOTAL.BackColor = vbRed
Else
ENG_TOTAL.BackColor = 12632256
End If

'If ENG_TOTAL = 0 Then
'ENG_TOTAL.BackColor = vbRed
'Else
'ENG_TOTAL.BackColor = 12632256
'End If

ENG = ENG_TOTAL
If ENG >= 80 Then
ENG_LG = "A+"
ENG_GP = 4
ElseIf ENG >= 75 Then
ENG_LG = "A"
ENG_GP = 3.75
ElseIf ENG >= 70 Then
ENG_LG = "A-"
ENG_GP = 3.5
ElseIf ENG >= 65 Then
ENG_LG = "B+"
ENG_GP = 3.25
ElseIf ENG >= 60 Then
ENG_LG = "B"
ENG_GP = 3
ElseIf ENG >= 55 Then
ENG_LG = "B-"
ENG_GP = 2.75
ElseIf ENG >= 50 Then
ENG_LG = "C+"
ENG_GP = 2.5
ElseIf ENG >= 45 Then
ENG_LG = "C"
ENG_GP = 2.25
ElseIf ENG >= 40 Then
ENG_LG = "D"
ENG_GP = 2
Else: ENG_LG = "F"
ENG_GP = 0
End If


' Physics

PHY_T_TOTAL = PHY_TC + PHY_TF
If PHY_T_TOTAL >= 60 And PHY_PC >= 10 And PHY_PF >= 10 Then
PHY_TOTAL = PHY_T_TOTAL + PHY_PC + PHY_PF
Else
PHY_TOTAL = 0
End If

If PHY_T_TOTAL < 60 Then
PHY_T_TOTAL.BackColor = vbRed
Else
PHY_T_TOTAL.BackColor = 12632256
End If
If PHY_PC < 10 Then
PHY_PC.BackColor = vbRed
Else
PHY_PC.BackColor = vbWhite
End If
If PHY_PF < 10 Then
PHY_PF.BackColor = vbRed
Else
PHY_PF.BackColor = vbWhite
End If
If PHY_TOTAL = 0 Then
PHY_TOTAL.BackColor = vbRed
Else
PHY_TOTAL.BackColor = 12632256
End If

PHY = PHY_TOTAL / 2
If PHY >= 80 Then
PHY_LG = "A+"
PHY_GP = 4
ElseIf PHY >= 75 Then
PHY_LG = "A"
PHY_GP = 3.75
ElseIf PHY >= 70 Then
PHY_LG = "A-"
PHY_GP = 3.5
ElseIf PHY >= 65 Then
PHY_LG = "B+"
PHY_GP = 3.25
ElseIf PHY >= 60 Then
PHY_LG = "B"
PHY_GP = 3
ElseIf PHY >= 55 Then
PHY_LG = "B-"
PHY_GP = 2.75
ElseIf PHY >= 50 Then
PHY_LG = "C+"
PHY_GP = 2.5
ElseIf PHY >= 45 Then
PHY_LG = "C"
PHY_GP = 2.25
ElseIf PHY >= 40 Then
PHY_LG = "D"
PHY_GP = 2
Else: PHY_LG = "F"
PHY_GP = 0
End If

' Basic Workshop Practice

If BWP_PC >= 20 And BWP_PF >= 20 Then
BWP_TOTAL = BWP_PC + BWP_PF
Else: BWP_TOTAL = 0
End If

If BWP_PC < 20 Then
BWP_PC.BackColor = vbRed
Else
BWP_PC.BackColor = vbWhite
End If
If BWP_PF < 20 Then
BWP_PF.BackColor = vbRed
Else
BWP_PF.BackColor = vbWhite
End If


If BWP_TOTAL = 0 Then
BWP_TOTAL.BackColor = vbRed
Else
BWP_TOTAL.BackColor = 12632256
End If


BWP = BWP_TOTAL
If BWP >= 80 Then
BWP_LG = "A+"
BWP_GP = 4
ElseIf BWP >= 75 Then
EN_D_LG = "A"
EN_D_GP = 3.75
ElseIf BWP >= 70 Then
BWP_LG = "A-"
BWP_GP = 3.5
ElseIf BWP >= 65 Then
BWP_LG = "B+"
BWP_GP = 3.25
ElseIf BWP >= 60 Then
BWP_LG = "B"
BWP_GP = 3
ElseIf BWP >= 55 Then
BWP_LG = "B-"
BWP_GP = 2.75
ElseIf BWP >= 20 Then
BWP_LG = "C+"
BWP_GP = 2.5
ElseIf BWP >= 45 Then
BWP_LG = "C"
BWP_GP = 2.25
ElseIf BWP >= 40 Then
BWP_LG = "D"
BWP_GP = 2
Else: BWP_LG = "F"
BWP_GP = 0
End If

' MACHINE SHOP  PRACTICE

MSP_T_TOTAL = MSP_TC + MSP_TF
If MSP_T_TOTAL >= 60 And MSP_PC >= 10 And MSP_PF >= 10 Then
MSP_TOTAL = MSP_T_TOTAL + MSP_PC + MSP_PF
Else
MSP_TOTAL = 0
End If

If MSP_T_TOTAL < 60 Then
MSP_T_TOTAL.BackColor = vbRed
Else
MSP_T_TOTAL.BackColor = 12632256
End If
If MSP_PC < 10 Then
MSP_PC.BackColor = vbRed
Else
MSP_PC.BackColor = vbWhite
End If
If MSP_PF < 10 Then
MSP_PF.BackColor = vbRed
Else
MSP_PF.BackColor = vbWhite
End If
If MSP_TOTAL = 0 Then
MSP_TOTAL.BackColor = vbRed
Else
MSP_TOTAL.BackColor = 12632256
End If

MSP = MSP_TOTAL / 2
If MSP >= 80 Then
MSP_LG = "A+"
MSP_GP = 4
ElseIf MSP >= 75 Then
MSP_LG = "A"
MSP_GP = 3.75
ElseIf MSP >= 70 Then
MSP_LG = "A-"
MSP_GP = 3.5
ElseIf MSP >= 65 Then
MSP_LG = "B+"
MSP_GP = 3.25
ElseIf MSP >= 60 Then
MSP_LG = "B"
MSP_GP = 3
ElseIf MSP >= 55 Then
MSP_LG = "B-"
MSP_GP = 2.75
ElseIf MSP >= 50 Then
MSP_LG = "C+"
MSP_GP = 2.5
ElseIf MSP >= 45 Then
MSP_LG = "C"
MSP_GP = 2.25
ElseIf MSP >= 40 Then
MSP_LG = "D"
MSP_GP = 2
Else: MSP_LG = "F"
MSP_GP = 0
End If


' MATH HEMATH ICS

MATH_T_TOTAL = MATH_TC + MATH_TF
If MATH_T_TOTAL >= 60 And MATH_PC >= 20 Then
MATH_TOTAL = MATH_T_TOTAL + MATH_PC
Else
MATH_TOTAL = 0
End If

If MATH_T_TOTAL < 60 Then
MATH_T_TOTAL.BackColor = vbRed
Else
MATH_T_TOTAL.BackColor = 12632256
End If
If MATH_PC < 20 Then
MATH_PC.BackColor = vbRed
Else
MATH_PC.BackColor = vbWhite
End If
If MATH_TOTAL = 0 Then
MATH_TOTAL.BackColor = vbRed
Else
MATH_TOTAL.BackColor = 12632256
End If

 MA = (MATH_TOTAL / 4) * 2
If MA >= 80 Then
MATH_LG = "A+"
MATH_GP = 4
ElseIf MA >= 75 Then
MATH_LG = "A"
MATH_GP = 3.75
ElseIf MA >= 70 Then
MATH_LG = "A-"
MATH_GP = 3.5
ElseIf MA >= 65 Then
MATH_LG = "B+"
MATH_GP = 3.25
ElseIf MA >= 60 Then
MATH_LG = "B"
MATH_GP = 3
ElseIf MA >= 55 Then
MATH_LG = "B-"
MATH_GP = 2.75
ElseIf MA >= 50 Then
MATH_LG = "C+"
MATH_GP = 2.5
ElseIf MA >= 45 Then
MATH_LG = "C"
MATH_GP = 2.25
ElseIf MA >= 40 Then
MATH_LG = "D"
MATH_GP = 2
Else: MATH_LG = "F"
MATH_GP = 0
End If


' GPA Calculation

If EN_D_GP <> 0 And EN_M_GP <> 0 And ENG_GP <> 0 And PHY_GP <> 0 And MATH_GP <> 0 And BWP_GP <> 0 And MSP_GP <> 0 Then

GPA = (EN_D_GP * 2 + EN_M_GP * 3 + ENG_GP * 2 + PHY_GP * 4 + MATH_GP * 4 + BWS_GP * 2 + MSP_GP * 4) / 21
Else
GPA = 0
End If

If GPA = 0 Then
GPA.BackColor = vbRed
Else
GPA.BackColor = vbWhite
End If

If GPA = 4 Then
Letter_Grade = "A+"
ElseIf GPA >= 3.75 And GPA < 4 Then
Letter_Grade = "A"
ElseIf GPA >= 3.5 And GPA < 3.75 Then
Letter_Grade = "A-"
ElseIf GPA >= 3.25 And GPA < 3.5 Then
Letter_Grade = "B+"
ElseIf GPA >= 3 And GPA < 3.25 Then
Letter_Grade = "B"
ElseIf GPA >= 2.75 And GPA < 3 Then
Letter_Grade = "B-"
ElseIf GPA >= 2.5 And GPA < 2.75 Then
Letter_Grade = "C+"
ElseIf GPA >= 2.25 And GPA < 2.5 Then
Letter_Grade = "C"
ElseIf GPA >= 2 And GPA < 2.25 Then
Letter_Grade = "D"
Else
Letter_Grade = "F"
End If



If EN_D_GP <> 0 And EN_M_GP <> 0 And ENG_GP <> 0 And PHY_GP <> 0 And MATH_GP <> 0 And BWP_GP <> 0 And MSP_GP <> 0 Then
Remarks = "Passed"
Else
Remarks = "Referred"
End If

If EN_D_GP <> 0 And EN_M_GP <> 0 And ENG_GP <> 0 And PHY_GP <> 0 And MATH_GP <> 0 And BWP_GP <> 0 And MSP_GP <> 0 Then
Remarks.BackColor = vbWhite
Else
Remarks.BackColor = vbRed
End If

End Sub

