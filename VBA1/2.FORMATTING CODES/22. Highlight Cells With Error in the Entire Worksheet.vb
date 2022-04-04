Sub highlightErrors()
Dim rng As Range
Dim i As Integer
For Each rng In ActiveSheet.UsedRange
If WorksheetFunction.IsError(rng) Then
i = i + 1
rng.Style = "bad"
End If
Next rng
MsgBox _
"There are total " & i _
& " error(s) in this worksheet."
End Sub