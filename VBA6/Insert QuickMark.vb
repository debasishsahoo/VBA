Sub QuickMarkInsert()
  If ActiveDocument.Bookmarks.Exists("QuickMark") = True Then
    ActiveDocument.Bookmarks("QuickMark").Delete
  End If
  With ActiveDocument.Bookmarks
    .Add Range:=Selection.Range, Name:="QuickMark"
  End With
End Sub