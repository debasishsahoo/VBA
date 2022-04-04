Sub ProtectAllWorskeets()
Dim ws As Worksheet
Dim ps As String
ps = InputBox("Enter a Password.", vbOKCancel)
For Each ws In ActiveWorkbook.Worksheets
ws.Protect Password:=ps
Next ws
End Sub