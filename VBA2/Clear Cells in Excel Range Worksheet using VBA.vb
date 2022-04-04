'Clearing a Cells/Range using Clear Method
Sub sbClearCells()
Range("A1:C10").Clear
End Sub

'Clearing Only Data of a Range using ClearContents Method
Sub sbClearCellsOnlyData()
Range("A1:C10").ClearContents
End Sub

'Clearing Entire Worksheet using Clear Method
Sub sbClearEntireSheet()
Sheets("SheetName").Cells.Clear
End Sub

'Clearing Only Data from Worksheet using ClearContents Method
Sub sbClearEntireSheetOnlyData()
Sheets("SheetName").Cells.ClearContents
End Sub