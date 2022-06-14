Sub CheckTableCells()

Dim oCell As Cell
Dim oRow As Row

For Each oRow In Selection.Tables(1).Rows
    For Each oCell In oRow.Cells
        If oCell.Range.Text = Chr(13) & Chr(7) Then
            MsgBox oCell.RowIndex & " " & oCell.ColumnIndex & " is empty."
        End If
    Next oCell
Next oRow

End Sub


Sub CheckTableCells()

Dim oCell As Cell
Dim oRow As Row
Dim MyRange As Range

For Each oRow In Selection.Tables(1).Rows
    For Each oCell In oRow.Cells
        Set MyRange = oCell.Range
        MyRange.End = MyRange.End - 1
        If Len(MyRange.Text) = 0 Then
            MsgBox oCell.RowIndex & " " & oCell.ColumnIndex & " is empty."
        End If
    Next oCell
Next oRow

End Sub


Sub CheckTableCells()

Dim oCell As Cell
Dim oRow As Row
Dim MyRange As Range

For Each oRow In Selection.Tables(1).Rows
    For Each oCell In oRow.Cells
        If Selection.Text = Chr(13) & Chr(7) Then
            oCell.Select
            MsgBox oCell.RowIndex & " " & oCell.ColumnIndex & " is empty."
        End If
    Next oCell
Next oRow

End Sub