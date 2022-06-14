Sub SelectToStyleBoundariesAndDelete()
  Dim StyleName As Variant
  Application.ScreenUpdating = False
  Set StyleName = Selection.Style
  While Selection.Characters.First.Previous.Style = StyleName
    Selection.MoveStart Unit:=wdCharacter, Count:=-1
  Wend
  While Selection.Characters.Last.Next.Style = StyleName
    Selection.MoveEnd Unit:=wdCharacter, Count:=1
  Wend
  Application.ScreenUpdating = True
  Selection.Delete
End Sub