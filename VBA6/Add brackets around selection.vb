Sub AddBracketsAroundSelection()
  ' Simply inserts adjacent brackets if nothing is selected
  If Selection.Type = wdSelectionIP Then
    Selection.InsertBefore "["
    Selection.InsertAfter "]"
    Selection.MoveStart Unit:=wdCharacter, Count:=1
    Selection.Collapse Direction:=wdCollapseStart
  Else
    ' Shrinks selection to exclude leading or trailing spaces
    ' Also excludes trailing paragraph break
    While Selection.Characters.First = " "
      Selection.MoveStart Unit:=wdCharacter, Count:=1
    Wend
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
      Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    Selection.InsertBefore "["
    Selection.InsertAfter "]"
    ' Brackets shed any character style acquired from selection
    Selection.Characters.First.Font.Reset
    Selection.Characters.Last.Font.Reset
  End If
End Sub