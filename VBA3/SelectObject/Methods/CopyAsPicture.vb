Sub CopyPasteAsPicture() 
 ActiveDocument.Content.Select 
 With Selection 
 .CopyAsPicture 
 .Collapse Direction:=wdCollapseEnd 
 .PasteSpecial DataType:=wdPasteMetafilePicture 
 End With 
End Sub