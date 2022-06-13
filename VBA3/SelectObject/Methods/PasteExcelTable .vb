Sub PasteExcelFormatted() 
 Selection.PasteExcelTable _ 
 LinkedToExcel:=True, _ 
 WordFormatting:=False, _ 
 RTF:=True 
End Sub