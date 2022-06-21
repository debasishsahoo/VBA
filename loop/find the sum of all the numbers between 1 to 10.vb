Sub ForLoopSumNumbers()
Dim loop_ctr As Integer
Dim result As Integer
result = 0
For loop_ctr = 1 To 10
result = result + loop_ctr
Next loop_ctr
MsgBox "Sum of numbers from 1-10 is : " & result
End Sub