Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Selection.Range, 3, 3) 
i = 1 
For Each c In myTable.Range.Cells 
 c.Range.InsertAfter "Cell " & i 
 i = i + 1 
Next c