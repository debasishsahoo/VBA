Dim bmRange As Range

Set bmRange = ActiveDocument.Bookmarks("myBookmark").Range

bmRange.Text = "Inserted Text"

ActiveDocument.Bookmarks.Add _
   Name:="myBookmark", _
   Range:=bmRange