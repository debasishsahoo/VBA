Sub Macro2()

  Documents.Open FileName:= "Lorem.doc" , _

      ConfirmConversions:= False , _

      ReadOnly:= False , AddToRecentFiles:= False , _

      PasswordDocument:= "" , PasswordTemplate:= "" , _

      Revert:= False , WritePasswordDocument:= "" , _

      WritePasswordTemplate:= "" , _

      Format:=wdOpenFormatAuto, XMLTransform:= ""

  ' more code, for example ...

  ActiveWindow.View = wdPrintView

End Sub