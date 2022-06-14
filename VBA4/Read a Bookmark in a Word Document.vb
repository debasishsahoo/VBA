Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open("c:\scripts\word\bookmarkdoc.doc")
Set objRange = objDoc.Bookmarks("NameBookmark").Range

Wscript.Echo objRange.Text 

Set objRange = objDoc.Bookmarks("AddressBookmark").Range
Wscript.Echo objRange.Text 

objWord.Quit