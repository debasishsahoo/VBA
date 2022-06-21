Public Sub LoopThroughColumns()
 
Dim cell As Range
 
For Each cell In Range("1:1")
    If cell.Value <> "" Then MsgBox cell.Address & ": " & cell.Value
Next cell
 
End Sub