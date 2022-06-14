Sub DoNotSave_Close_Document()
  Documents.Close SaveChanges:=wdDoNotSaveChanges
  Application.Quit SaveChanges:=wdDoNotSaveChanges
End Sub