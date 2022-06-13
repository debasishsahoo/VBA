Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
Selection.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=1

Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=4

Selection.GoTo What:=wdGoToLine, Which:=wdGoToPrevious, Count:=2

Selection.GoTo What:=wdGoToField, Name:="Date"

Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext

If ActiveDocument.Endnotes.Count >= 5 Then
 Selection.GoTo What:=wdGoToEndnote, _
 Which:=wdGoToAbsolute, Count:=5
End If

Selection.GoTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=4

Selection.GoTo What:=wdGoToPage, Which:=wdGoToPrevious, Count:=2