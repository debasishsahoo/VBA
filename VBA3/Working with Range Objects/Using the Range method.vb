Sub SetNewRange() 
 Dim rngDoc As Range 
 Set rngDoc = ActiveDocument.Range(Start:=0, End:=10) 
End Sub

Sub SetBoldRange() 
 Dim rngDoc As Range 
 Set rngDoc = ActiveDocument.Range(Start:=0, End:=10) 
 rngDoc.Bold = True 
End Sub

Sub BoldRange() 
 ActiveDocument.Range(Start:=0, End:=10).Bold = True 
End Sub

Sub InsertTextBeforeRange() 
 Dim rngDoc As Range 
 Set rngDoc = ActiveDocument.Range(Start:=0, End:=0) 
 rngDoc.InsertBefore "Hello " 
End Sub

Sub NewRange() 
 Dim doc As Document 
 Dim rngDoc As Range 
 
 Set doc = ActiveDocument 
 Set rngDoc = doc.Range(Start:=doc.Paragraphs(2).Range.Start, _ 
 End:=doc.Paragraphs(3).Range.End) 
End Sub