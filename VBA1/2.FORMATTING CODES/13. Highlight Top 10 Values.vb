Sub TopTen()
Selection.FormatConditions.AddTop10
Selection.FormatConditions(Selection.FormatConditions.Count).S
tFirstPriority
With Selection.FormatConditions(1)
.TopBottom = xlTop10Top
.Rank = 10
.Percent = False
End With
With Selection.FormatConditions(1).Font
.Color = -16752384
.TintAndShade = 0
End With
With Selection.FormatConditions(1).Interior
.PatternColorIndex = xlAutomatic
.Color = 13561798
.TintAndShade = 0
End With
Selection.FormatConditions(1).StopIfTrue = False
End Sub