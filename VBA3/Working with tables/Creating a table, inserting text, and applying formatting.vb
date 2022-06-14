Sub CreateNewTable() 
 Dim docActive As Document 
 Dim tblNew As Table 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 Set docActive = ActiveDocument 
 Set tblNew = docActive.Tables.Add( _ 
 Range:=docActive.Range(Start:=0, End:=0), NumRows:=3, _ 
 NumColumns:=4) 
 intCount = 1 
 
 For Each celTable In tblNew.Range.Cells 
  celTable.Range.InsertAfter "Cell " & intCount 
  intCount = intCount + 1 
 Next celTable 
 
 tblNew.AutoFormat Format:=wdTableFormatColorful2, _ 
 ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True 
End Sub