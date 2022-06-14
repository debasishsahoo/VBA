Sub QuickMarkGoTo()
  If ActiveDocument.Bookmarks.Exists("QuickMark") = True Then
    Selection.GoTo What:=wdGoToBookmark, Name:="QuickMark"
  End If
End Sub