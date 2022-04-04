Sub HighlightMisspelledCells()
Dim rng As Range
For Each rng In ActiveSheet.UsedRange
If Not Application.CheckSpelling(word:=rng.Text) Then
rng.Style = "Bad"
End If
Next rng
End Sub