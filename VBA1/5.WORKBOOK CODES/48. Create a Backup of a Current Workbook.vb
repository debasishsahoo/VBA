Sub FileBackUp()
ThisWorkbook.SaveCopyAs Filename:=ThisWorkbook.Path & _
"" & Format(Date, "mm-dd-yy") & " " & _
ThisWorkbook.name
End Sub