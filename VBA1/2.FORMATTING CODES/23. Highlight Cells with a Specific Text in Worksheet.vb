Sub highlightSpecificValues()
Dim rng As range
Dim i As Integer
Dim c As Variant
c = InputBox("Enter Value To Highlight")
For Each rng In ActiveSheet.UsedRange
If rng = c Then
rng.Style = "Note"
i = i + 1
End If
Next rng
MsgBox "There are total " & i & " " & c & " in this worksheet."
End Sub