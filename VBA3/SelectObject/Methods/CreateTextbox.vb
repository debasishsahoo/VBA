If Selection.Type = wdSelectionNormal Then 
 Selection.CreateTextbox 
 Selection.ShapeRange(1).Line.DashStyle =msoLineDashDot 
End If