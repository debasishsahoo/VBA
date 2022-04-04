Sub sbRangeFillColorExample1()
'Using Cell Object
Cells(3, 2).Interior.ColorIndex = 5 ' 5 indicates Blue Color
End Sub



Sub sbRangeFillColorExample2()
'Using Range Object
Range("B3").Interior.ColorIndex = 5
End Sub


Sub sbRangeFillColorExample3()
'Using Cell Object
Cells(3, 2).Interior.Color = RGB(0, 0, 250)
'Using Range Object
Range("B3").Interior.Color = RGB(0, 0, 250)
End Sub


Sub sbPrintColorIndexColors()
Dim iCntr
For iCntr = 1 To 56
    Cells(iCntr, 1).Interior.ColorIndex = iCntr
    Cells(iCntr, 1) = iCntr
Next iCntr
End Sub