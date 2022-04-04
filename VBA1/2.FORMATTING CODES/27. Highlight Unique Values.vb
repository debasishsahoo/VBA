Sub highlightUniqueValues()
Dim rng As Range
Set rng = Selection
rng.FormatConditions.Delete
Dim uv As UniqueValues
Set uv = rng.FormatConditions.AddUniqueValues
uv.DupeUnique = xlUnique
uv.Interior.Color = vbGreen
End Sub