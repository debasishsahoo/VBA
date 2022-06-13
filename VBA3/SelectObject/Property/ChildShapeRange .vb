Sub ChildShapes() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 'Create a new document with a drawing canvas and shapes 
 Set docNew = Documents.Add 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=100, Top:=100, Width:=200, Height:=200) 
 shpCanvas.CanvasItems.AddShape msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=100, Height:=100 
 shpCanvas.CanvasItems.AddShape msoShapeOval, _ 
 Left:=0, Top:=50, Width:=100, Height:=100 
 shpCanvas.CanvasItems.AddShape msoShapeDiamond, _ 
 Left:=0, Top:=100, Width:=100, Height:=100 
 
 'Select all shapes in the canvas 
 shpCanvas.CanvasItems.SelectAll 
 
 'Fill canvas child shapes with a pattern 
 If Selection.HasChildShapeRange = True Then 
 Selection.ChildShapeRange.Fill.Patterned msoPatternDivot 
 Else 
 MsgBox "This is not a range of child shapes." 
 End If 
 
End Sub