Public Sub LoopThroughRows()
Dim cell As Range
For Each cell In Range("A:A")
    If cell.value <> "" Then MsgBox cell.address & ": " & cell.Value
Next cell
End Sub