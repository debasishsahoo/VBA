Sub addsAlphabets1()
Dim i As Integer
For i = 65 To 90
ActiveCell.Value = Chr(i)
ActiveCell.Offset(1, 0).Select
Next i
End Sub

Sub addsAlphabets2()
Dim i As Integer
For i = 97 To 122
ActiveCell.Value = Chr(i)
ActiveCell.Offset(1, 0).Select
Next i
End Sub