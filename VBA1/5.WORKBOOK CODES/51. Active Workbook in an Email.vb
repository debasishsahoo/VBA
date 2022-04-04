Sub Send_Mail()
Dim OutApp As Object
Dim OutMail As Object
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
With OutMail
.to = "Sales@FrontLinePaper.com"
.Subject = "Growth Report"
.Body = "Hello Team, Please find attached Growth Report."
.Attachments.Add ActiveWorkbook.FullName
.display
End With
Set OutMail = Nothing
Set OutApp = Nothing
End Sub