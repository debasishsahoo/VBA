Sub ReverseForLoop()
Dim loop_ctr As Integer
Dim cell As Integer
cell = 1

For loop_ctr = 10 To 1 Step -1
ActiveSheet.Range("A1").Offset(cell - 1, 0).Value = loop_ctr
cell = cell + 1
Next loop_ctr

MsgBox "For Loop Completed!"
End Sub