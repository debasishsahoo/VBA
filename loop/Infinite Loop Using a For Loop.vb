'Do not run this code
Sub InfiniteForLoop()
Dim loop_ctr As Integer
Dim cell As Integer

For loop_ctr = 1 To 10
ActiveSheet.Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
loop_ctr = loop_ctr - 1
Next loop_ctr

MsgBox "For Loop Completed!"
End Sub