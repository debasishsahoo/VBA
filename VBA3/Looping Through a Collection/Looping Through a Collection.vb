Sub LoopThroughOpenDocuments() 
 Dim docOpen As Document 
 For Each docOpen In Documents 
 MsgBox docOpen.Name 
 Next docOpen 
End Sub

Sub LoopThroughBookmarks() 
 Dim bkMark As Bookmark 
 Dim strMarks() As String 
 Dim intCount As Integer 
 
 If ActiveDocument.Bookmarks.Count > 0 Then 
 ReDim strMarks(ActiveDocument.Bookmarks.Count - 1) 
 intCount = 0 
 For Each bkMark In ActiveDocument.Bookmarks 
 strMarks(intCount) = bkMark.Name 
 intCount = intCount + 1 
 Next bkMark 
 End If 
End Sub

Sub UpdateDateFields() 
 Dim fldDate As Field 
 
 For Each fldDate In ActiveDocument.Fields 
 If InStr(1, fldDate.Code, "Date", 1) Then fldDate.Update 
 Next fldDate 
End Sub

Sub FindAutoTextEntry() 
 Dim atxtEntry As AutoTextEntry 
 
 For Each atxtEntry In ActiveDocument.AttachedTemplate.AutoTextEntries 
 If atxtEntry.Name = "Filename" Then _ 
 MsgBox "The Filename AutoText entry exists." 
 Next atxtEntry 
End Sub