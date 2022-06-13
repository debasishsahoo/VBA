Dim objLink As Hyperlink 
Dim objRange As Range 
 
Set objRange = ActiveDocument.Range( _ 
 Paragraphs(1).Range.Start, _ 
 Paragraphs(10).Range.End) 
 
For Each objLink In objRange.Hyperlinks 
 If InStr(LCase(objLink.Address), "microsoft") <> 0 Then 
 MsgBox objLink.Name 
 End If 
Next objLink