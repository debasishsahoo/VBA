If Selection.MoveLeft = 1 Then MsgBox "Move was successful"


ActiveDocument.ActiveWindow.View.FieldShading = _ 
 wdFieldShadingWhenSelected 
With Selection 
 .Fields.Add Range:=Selection.Range, Type:=wdFieldDate 
 .MoveLeft Unit:=wdWord, Count:=1 
End With

If Selection.Information(wdWithInTable) = True Then 
 Selection.MoveLeft Unit:=wdCell, Count:=1, Extend:=wdMove 
End If