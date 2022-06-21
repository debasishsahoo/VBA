Sub ForLoopPrintNumbers()
Dim loop_ctr As Integer
For loop_ctr = 1 To 10
ActiveSheet.Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
Next loop_ctr
MsgBox "For Loop Completed!"
End Sub