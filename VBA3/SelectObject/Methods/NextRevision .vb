Dim rngTemp as Range 
Dim revTemp as Revision 
 
If ActiveDocument.Paragraphs.Count >= 5 Then 
 Set rngTemp = ActiveDocument.Paragraphs(5).Range 
 rngTemp.Select 
 Set revTemp = Selection.NextRevision(Wrap:=False) 
 If Not (revTemp Is Nothing) Then revTemp.Reject 
End If
'------------------------------------------------'

Dim revTemp as Revision 
 
Set revTemp = Selection.NextRevision(Wrap:=True) 
If Not (revTemp Is Nothing) Then 
 If revTemp.Type = wdRevisionInsert Then revTemp.Accept 
End If
'------------------------------------------------'
Dim revTemp as Revision 
Dim strAuthor as String 
 
strAuthor = ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor) 
 
Do While True 
 Set revTemp = Selection.NextRevision(Wrap:=False) 
 If Not (revTemp Is Nothing) Then 
 If revTemp.Author = strAuthor Then 
 MsgBox Prompt:="Another revision by " & strAuthor & "!" 
 Exit Do 
 End If 
 Else 
 MsgBox Prompt:="No more revisions!" 
 Exit Do 
 End If 
Loop