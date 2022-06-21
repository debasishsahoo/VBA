
Sub If_Loop()
Dim Cell as Range
  For Each Cell In Range("A2:A6")
    If Cell.Value > 0 Then
      Cell.Offset(0, 1).Value = "Positive"
    ElseIf Cell.Value < 0 Then
      Cell.Offset(0, 1).Value = "Negative"
    Else
      Cell.Offset(0, 1).Value = "Zero"
     End If
  Next Cell
End Sub