Set objWord = CreateObject("Word.Application")
objWord.Caption = "Test Caption"
objWord.Visible = True

Set objDoc = objWord.Documents.Open("c:\scripts\word\bookmarkdoc.doc")

Set objRange = objDoc.Bookmarks("NameBookmark").Range
objRange.Text = "Bob"

Set objRange = objDoc.Bookmarks("AddressBookmark").Range
objRange.Text = "999"