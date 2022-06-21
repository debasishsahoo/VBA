' VBA For Each Loop
' The VBA For Each Loop will loop through all objects in a collection:

' All cells in a range
' All worksheets in a workbook
' All shapes in a worksheet
' All open workbooks
' You can also use Nested For Each Loops to:

' All cells in a range on all worksheets
' All shapes on all worksheets
' All sheets in all open workbooks
' and so on…

' For Each Object in Collection
' [Do Something]
' Next [Object]

' For Each Cell in Range
Sub ForEachCell_inRange()
Dim cell As Range
For Each cell In Range("a1:a10")
    cell.Value = cell.Offset(0,1).Value
Next cell
End Sub

' For Each Worksheet in Workbook
Sub ForEachSheet_inWorkbook()
DiM ws As Worksheet
For Each ws In Worksheets
    ws.Unprotect "password"
Next ws
End Sub

' For Each Open Workbook
'This code will save and close all open workbooks:

Sub ForEachWB_inWorkbooks()
Dim wb As Workbook
For Each wb In Workbooks
    wb.Close SaveChanges:=True
Next wb
End Sub

'For Each Shape in Worksheet

Sub ForEachShape()
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    shp.Delete
Next shp
End Sub

' For Each Shape in Each Worksheet in Workbook
Sub ForEachShape_inAllWorksheets()
Dim shp As Shape, ws As Worksheet
For Each ws In Worksheets
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
Next ws
End Sub

' For Each – IF Loop
Sub ForEachCell_inRange()
Dim cell As Range
For Each cell In Range("a1:a10")
    If cell.Value = "" Then _
       cell.EntireRow.Hidden = True
Next cell
End Sub