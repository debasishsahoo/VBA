If ActiveDocument.Endnotes.Count >= 5 Then 
 ActiveDocument.Range.GoTo What:=wdGoToEndnote, _ 
 Which:=wdGoToAbsolute, Count:=5 
End If

If ActiveDocument.Footnotes.Count >= 1 Then 
 Set R1 = ActiveDocument.Range.GoTo(What:=wdGoToFootnote, _ 
 Which:=wdGoToFirst) 
 R1.Expand Unit:=wdCharacter 
End If
