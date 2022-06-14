Sub HeaderFooterProperty()
   Dim MyText As String
   MyText = "<Replace this with your text>"
   ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
   Selection.HeaderFooter.Range.Text = "MyText"
   ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub