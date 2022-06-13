Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseStart 
If myRange.BookmarkID = 0 Then 
 ActiveDocument.Bookmarks.Add Name:="temp", Range:=myRange 
End If