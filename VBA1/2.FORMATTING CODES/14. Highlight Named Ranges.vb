Sub HighlightRanges()
Dim RangeName As Name
Dim HighlightRange As Range
On Error Resume Next
For Each RangeName In ActiveWorkbook.Names
Set HighlightRange = RangeName.RefersToRange
HighlightRange.Interior.ColorIndex = 36
Next RangeName
End Sub