Sub HideDeletedText()
Dim r As Range

Set r = ActiveDocument.Range
r.ShowAll = False
Debug.Print r.Text

End Sub