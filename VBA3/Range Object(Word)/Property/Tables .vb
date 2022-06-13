Selection.Collapse Direction:=wdCollapseStart 
Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
NumRows:=5, NumColumns:=5) 
myTable.AutoFormat Format:=wdTableFormatClassic2

num = 90 
For Each acell In ActiveDocument.Tables(1).Columns(1).Cells 
 acell.Range.Text = num & " Sales" 
 num = num + 1 
Next acell