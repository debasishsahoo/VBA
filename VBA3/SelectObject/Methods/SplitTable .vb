If Selection.Information(wdWithInTable) = True Then 
 Selection.SplitTable 
End If

ActiveDocument.Tables(1).Rows(2).Select 
Selection.SplitTable