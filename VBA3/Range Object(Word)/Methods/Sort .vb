Sub NewParagraphSort() 
 Dim newDoc As Document 
 Set newDoc = Documents.Add 
 newDoc.Content.InsertAfter "pear" & Chr(13) _ 
 & "zucchini" & Chr(13) & "apple" & Chr(13) 
 newDoc.Content.Sort SortOrder:=wdSortOrderAscending 
End Sub