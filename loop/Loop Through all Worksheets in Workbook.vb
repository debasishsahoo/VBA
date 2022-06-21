Sub LoopThroughSheets()
Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Visible = True
    Next
End Sub