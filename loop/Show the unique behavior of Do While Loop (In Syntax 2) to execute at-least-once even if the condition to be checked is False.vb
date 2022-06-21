Sub DoWhileLoopTest()
Dim loop_ctr As Integer
loop_ctr = 100

Do
MsgBox "Loop Counter : " & loop_ctr
loop_ctr = loop_ctr + 1
Loop While loop_ctr <= 10

End Sub

Sub DoWhileLoopTest()
Dim loop_ctr As Integer
loop_ctr = 100

Do While loop_ctr <= 10
MsgBox "Loop Counter : " & loop_ctr
loop_ctr = loop_ctr + 1
Loop

End Sub