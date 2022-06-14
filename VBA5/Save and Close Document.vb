Sub Save_Close_Document()
  Documents.Close SaveChanges:=wdSaveChanges
  Application.Quit SaveChanges:=wdSaveChanges
End Sub