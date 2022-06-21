' VBA Do Until Loop
' Do Until Loops will repeat a loop until a certain condition is met. The syntax is essentially the same as the Do While loops:

' Do Until Condition
' [Do Something]
' Loop
' and similarly the condition can go at the start or the end of the loop:

' Do
' [Do Something]
' Loop Until Condition

'Do Until
Sub DoUntilLoop()
    Dim n As Integer
    n = 1
    Do Until n > 10
        MsgBox n
        n = n + 1
    Loop
End Sub

'Loop Until

Sub DoLoopUntil()
    Dim n As Integer
    n = 1
    Do
        MsgBox n
        n = n + 1
    Loop Until n > 10
End Sub