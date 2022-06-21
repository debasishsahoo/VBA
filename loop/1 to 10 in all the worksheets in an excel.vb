Sub ForLoopPrintNumbers()
Dim loop_ctr As Integer
Dim sheet As Integer

For sheet = 1 To Worksheets.Count
For loop_ctr = 1 To 10
Worksheets(sheet).Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
Next loop_ctr
Next sheet

MsgBox "For Loop Completed!"
End Sub