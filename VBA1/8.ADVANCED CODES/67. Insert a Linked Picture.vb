Sub LinkedPicture()
Selection.Copy
ActiveSheet.Pictures.Paste(Link:=True).Select
End Sub