Sub DoUntilLoopPrintNumbers()
Dim loop_ctr As Integer
loop_ctr = 1

Do Until loop_ctr < 10
ActiveSheet.Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
loop_ctr = loop_ctr + 1
Loop

MsgBox ("Loop Ends")
End Sub
'--------------------------------------------
Sub DoUntilLoopSumNumbers()
Dim loop_ctr As Integer
Dim result As Integer
loop_ctr = 1
result = 0

Do Until loop_ctr > 20
result = result + loop_ctr
loop_ctr = loop_ctr + 1
Loop

MsgBox "Sum of numbers from 1-20 is : " & result
End Sub

'----------------------------------------------
Sub DoUntilLoopTest()
Dim loop_ctr As Integer
loop_ctr = 100

Do
MsgBox "Loop Counter : " & loop_ctr
loop_ctr = loop_ctr + 1
Loop Until loop_ctr > 10
End Sub
'-----------------------------------------------
Sub DoUntilLoopTest()
Dim loop_ctr As Integer
loop_ctr = 100

Do Until loop_ctr > 10
MsgBox "Loop Counter : " & loop_ctr
loop_ctr = loop_ctr + 1
Loop
End Sub
'-----------------------------------------------
Sub NestedDoUntilLoop()
Dim loop_ctr As Integer
Dim sheet As Integer
sheet = 1

Do Until sheet > Worksheets.Count
loop_ctr = 1
Do Until loop_ctr > 5
Worksheets(sheet).Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
loop_ctr = loop_ctr + 1
Loop
sheet = sheet + 1
Loop

MsgBox "Nested Do Until Loop Completed!"
End Sub

'-------------------------------------
'Do not run this code
Sub InfiniteDoUntilLoop()
Dim loop_ctr As Integer
loop_ctr = 1

Do Until loop_ctr > 10
ActiveSheet.Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
Loop

MsgBox ("Loop Ends")
End Sub
