Public Sub LoopExample()
    Dim Check As Boolean, Counter As Long, Total As Long
    Check = True: Counter = 0: Total = 0 ' Initialize variables.
    Do ' Outer loop.
        Do While Counter < 20 ' Inner Loop
            Counter = Counter + 1 ' Increment Counter.
            If Counter Mod 10 = 0 Then ' Check in with the user on every multiple of 10.
                Check = (MsgBox("Keep going?", vbYesNo) = vbYes) ' Stop when user click's on No
                If Not Check Then Exit Do ' Exit inner loop.
            End If
        Loop
        Total = Total + Counter ' Exit Do Lands here.
        Counter = 0
    Loop Until Check = False ' Exit outer loop immediately.
    MsgBox "Counted to: " & Total
End Sub