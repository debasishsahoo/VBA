' Exit Do Loop
' Similar to using Exit For to exit a For Loop, you use the Exit Do command to exit a Do Loop immediately

' Exit Do

Sub ExitDo_Loop()
 
Dim i As Integer
i = 1
 
Do Until i > 1000
    If Range("A" & i).Value = "error" Then
        Range("A" & i).Select
        MsgBox "Error Found"
        Exit Do
    End If
    i = i + 1
Loop
 
End Sub