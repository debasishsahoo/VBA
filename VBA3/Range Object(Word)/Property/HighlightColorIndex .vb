Selection.Range.HighlightColorIndex = wdNoHighlight

For Each abookmark In ActiveDocument.Bookmarks 
 abookmark.Range.HighlightColorIndex = wdYellow 
Next abookmark