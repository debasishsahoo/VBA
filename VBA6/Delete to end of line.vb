Sub DeleteToEndOfLine()
  Selection.MoveEnd Unit:=wdLine
  If Selection.Characters.Last = vbCr Then
    Selection.MoveEnd Unit:=wdCharacter, Count:=-1
  End If
  Selection.Text = ""
End Sub