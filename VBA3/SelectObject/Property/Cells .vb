If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Shading.BackgroundPatternColorIndex = wdRed 
Else 
 MsgBox "The insertion point is not in a table." 
End If