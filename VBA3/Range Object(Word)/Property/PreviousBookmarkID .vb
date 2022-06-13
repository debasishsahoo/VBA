num = ActiveDocument.Paragraphs(2).Range.PreviousBookmarkID 
If num <> 0 Then MsgBox ActiveDocument.Content.Bookmarks(num).Name