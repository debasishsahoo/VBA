' The For Next Loop allows you to repeat a block of code a specified number of times. The syntax is:

' [Dim Counter as Integer]
 
' For Counter = Start to End [Step Value]
'     [Do Something]
' Next [Counter]
' Where the items in brackets are optional.

' [Dim Counter as Long] – Declares the counter variable. Required if Option Explicit is declared at the top of your module.
' Counter – An integer variable used to count
' Start – The start value (Ex. 1)
' End – The end value (Ex. 10)
' [Step Value] – Allows you to count every n integers instead of every 1 integer. You can also go in reverse with a negative value (ex. Step -1)
' [Do Something] – The code that will repeat
' Next [Counter] – Closing statement to the For Next Loop. You can include the Counter or not. However, I strongly recommend including the counter as it makes your code easier to read.
' If that’s confusing, don’t worry. We will review some examples:

' Count to 10
' This code will count to 10 using a For-Next Loop:

Sub ForEach_CountTo10()
Dim n As Integer
For n = 1 To 10
    MsgBox n
Next n
End Sub

'Count to 10 – Only Even Numbers
Sub ForEach_CountTo10_Even()
Dim n As Integer
For n = 2 To 10 Step 2
    MsgBox n
Next n
End Sub

Sub ForLoop()
    Dim i As Integer
    For i = 1 To 10
        MsgBox i
    Next i
End Sub

'For Loop Step – Inverse
Sub ForEach_Countdown_Inverse()
 
Dim n As Integer
For n = 10 To 1 Step -1
    MsgBox n
Next n
MsgBox "Lift Off"
 
End Sub

'Delete Rows if Cell is Blank
Sub ForEach_DeleteRows_BlankCells()
 
Dim n As Integer
For n = 10 To 1 Step -1
    If Range("a" & n).Value = "" Then
        Range("a" & n).EntireRow.Delete
    End If
Next n
 
End Sub
'Nested For Loop
Sub Nested_ForEach_MultiplicationTable()
 
Dim row As Integer, col As Integer
 
For row = 1 To 9
    For col = 1 To 9
        Cells(row + 1, col + 1).Value = row * col
    Next col
Next row
 
End Sub
'Exit For
Sub ExitFor_Loop()
 
Dim i As Integer
For i = 1 To 1000
    If Range("A" & i).Value = "error" Then
        Range("A" & i).Select
        MsgBox "Error Found"
        Exit For
    End If
Next i
 
End Sub