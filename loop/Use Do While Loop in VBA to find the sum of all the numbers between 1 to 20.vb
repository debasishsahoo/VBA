Sub WhileLoopSumNumbers()
Dim loop_ctr As Integer
Dim result As Integer
loop_ctr = 1
result = 0

Do While loop_ctr <= 20
result = result + loop_ctr
loop_ctr = loop_ctr + 1
Loop

MsgBox "Sum of numbers from 1-20 is : " & result
End Sub