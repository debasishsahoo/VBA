'Do not run this code
Sub InfiniteDoWhileLoop()
Dim loop_ctr As Integer
loop_ctr = 1

Do While loop_ctr <= 10
ActiveSheet.Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
Loop

MsgBox ("Loop Ends")
End Sub