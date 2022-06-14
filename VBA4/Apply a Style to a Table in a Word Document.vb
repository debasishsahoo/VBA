Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

Set objRange = objDoc.Range()
objDoc.Tables.Add objRange,1,3
Set objTable = objDoc.Tables(1)
objTable.Range.Font.Size = 10
objTable.Range.Style = "Table Contemporary"

x=2

objTable.Cell(x, 1).Range.Text = "Service Name"
objTable.Cell(x, 2).Range.text = "Display Name"
objTable.Cell(x, 3).Range.text = "State"

strComputer = "."
Set objWMIService = _
    GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objItem in colItems
    If x > 1 Then
        objTable.Rows.Add()
    End If
    objTable.Cell(x, 1).Range.Text = objItem.Name
    objTable.Cell(x, 2).Range.text = objItem.DisplayName
    objTable.Cell(x, 3).Range.text = objItem.State
    x = x + 1
Next