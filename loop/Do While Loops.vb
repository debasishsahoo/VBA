' The VBA Do While and Do Until (see next section) are very similar. They will repeat a loop while (or until) a condition is met.

' The Do While Loop will repeat a loop while a condition is met.

' Here is the Do While Syntax:

' Do While Condition
' [Do Something]
' Loop
' Where:

' Condition – The condition to test
' [Do Something] – The code block to repeat
' You can also set up a Do While loop with the Condition at the end of the loop:

' Do
' [Do Something]
' Loop While Condition

'DoWhile

Sub DoWhileLoop()
    Dim n As Integer
    n = 1
    Do While n < 11
        MsgBox n
        n = n + 1
    Loop
End Sub

'Loop While
Sub DoLoopWhile()
    Dim n As Integer
    n = 1
    Do
        MsgBox n
        n = n + 1
    Loop While n < 11
End Sub