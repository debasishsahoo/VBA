Sub ShowOpenDialog() 
 Dialogs(wdDialogFileOpen).Show 
End Sub

Sub ShowPrintDialog() 
 Dialogs(wdDialogFilePrint).Show 
End Sub

Sub ShowBorderDialog() 
 With Dialogs(wdDialogFormatBordersAndShading) 
 .DefaultTab = wdDialogFormatBordersAndShadingTabPageBorder 
 .Show 
 End With 
End Sub

Sub DisplayUserInfo() 
 MsgBox Application.UserName 
End Sub

Sub ShowAndSetUserInfoDialogBox() 
 With Dialogs(wdDialogToolsOptionsUserInfo) 
 .Display 
 If .Name <> "" Then .Execute 
 End With 
End Sub

Sub SetUserName() 
 Application.UserName = "Jeff Smith" 
 Dialogs(wdDialogToolsOptionsUserInfo).Display 
End Sub