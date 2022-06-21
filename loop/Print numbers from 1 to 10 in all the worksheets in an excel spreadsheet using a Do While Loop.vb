Sub NestedDoWhileLoop()
Dim loop_ctr As Integer
Dim sheet As Integer
sheet = 1

Do While sheet <= Worksheets.Count
loop_ctr = 1
Do While loop_ctr <= 10
Worksheets(sheet).Range("A1").Offset(loop_ctr - 1, 0).Value = loop_ctr
loop_ctr = loop_ctr + 1
Loop
sheet = sheet + 1
Loop

MsgBox "Nested While Loop Completed!"
End Sub