Sub InsertPivotTable()
'Macro By debasishsahoo

'Declare Variables
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long

'Insert a New Blank Worksheet
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTable").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTable")
Set DSheet = Worksheets("Data")

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="SalesPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="SalesPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Year")
.Orientation = xlRowField
.Position = 1
End With
With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Month")
.Orientation = xlRowField
.Position = 2
End With

'Insert Column Fields
With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Zone")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Data Field
With ActiveSheet.PivotTables("SalesPivotTable").PivotFields ("Amount")
.Orientation = xlDataField
.Function = xlSum
.NumberFormat = "#,##0"
.Name = "Revenue "
End With

'Format Pivot Table
ActiveSheet.PivotTables("SalesPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("SalesPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub