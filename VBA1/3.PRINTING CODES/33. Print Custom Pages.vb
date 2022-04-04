Sub printCustomSelection()
Dim startpage As Integer
Dim endpage As Integer
startpage = _
InputBox("Please Enter Start Page number.", "Enter Value")
If Not WorksheetFunction.IsNumber(startpage) Then
MsgBox _
"Invalid Start Page number. Please try again.", "Error"
Exit Sub
End If
endpage = _
InputBox("Please Enter End Page number.", "Enter Value")
If Not WorksheetFunction.IsNumber(endpage) Then
MsgBox _
"Invalid End Page number. Please try again.", "Error"
Exit Sub
End If
Selection.PrintOut From:=startpage, _
To:=endpage, Copies:=1, Collate:=True
End Sub