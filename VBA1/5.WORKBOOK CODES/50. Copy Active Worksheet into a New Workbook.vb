Sub CopyWorksheetToNewWorkbook()
ThisWorkbook.ActiveSheet.Copy _
Before:=Workbooks.Add.Worksheets(1)
End Sub