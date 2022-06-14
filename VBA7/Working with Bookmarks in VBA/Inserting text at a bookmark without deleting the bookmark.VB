Dim BMRange As Range
'Identify current Bookmark range and insert text
Set BMRange = ActiveDocument.Bookmarks("MyBookmark").Range
BMRange.Text = "Hello world"
'Re-insert the bookmark
ActiveDocument.Bookmarks.Add "MyBookmark", BMRange

'If you have a macro which updates many bookmarks (for example, a userform  macro), then repeating the above code over and over again would be very laborious, so the best plan is to call a subroutine with arguments as in the following example:

Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = TextToUse
    ActiveDocument.Bookmarks.Add BookmarkToUpdate, BMRange
End Sub

You could call it like this:

UpdateBookmark "NameOfBookMark", "String you want to insert"