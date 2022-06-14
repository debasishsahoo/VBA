Sub DeleteCurrentLine()
  Selection.HomeKey Unit:=wdLine
  Selection.MoveEnd Unit:=wdLine
  Selection.Text = ""
End Sub