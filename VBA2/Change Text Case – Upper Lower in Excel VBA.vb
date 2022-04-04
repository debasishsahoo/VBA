Sub sbChangeCASE()
      'Upper Case
      Range("A3") = UCase(Range("A3"))
      
      'Lower Case
      Range("A4") = LCase(Range("A4"))
End Sub




Sub sbCompareColumns_1()
iCntr = 1
Do While Cells(iCntr, 1) <> ""
If Cells(iCntr, 1) = Cells(iCntr, 2) Then
    Cells(iCntr, 3) = "Matched"
Else
    Cells(iCntr, 3) = "Not Matched"
End If
iCntr = iCntr + 1
Loop
End Sub




Sub sbCompareColumns_2()
iCntr = 1
Do While Cells(iCntr, 1) <> ""
If UCase(Cells(iCntr, 1)) = UCase(Cells(iCntr, 2)) Then
    Cells(iCntr, 3) = "Matched"
Else
    Cells(iCntr, 3) = "Not Matched"
End If
iCntr = iCntr + 1
Loop
End Sub