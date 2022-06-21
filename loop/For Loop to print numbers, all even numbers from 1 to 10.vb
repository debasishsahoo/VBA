Sub ForLoopToPrintEvenNumbers()
Dim loop_ctr As Integer
Dim cell As Integer
cell = 1

For loop_ctr = 1 To 10
If loop_ctr Mod 2 = 0 Then
ActiveSheet.Range("A1").Offset(cell - 1, 0).Value = loop_ctr
cell = cell + 1
End If
Next loop_ctr

MsgBox "For Loop Completed!"
End Sub

Sub ForLoopToPrintEvenNumbers()
Dim loop_ctr As Integer
Dim cell As Integer
cell = 1

For loop_ctr = 2 To 10 Step 2
ActiveSheet.Range("A1").Offset(cell - 1, 0).Value = loop_ctr
cell = cell + 1
Next loop_ctr

MsgBox "For Loop Completed!"
End Sub