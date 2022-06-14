Option Explicit

Sub CreateBookMarkMenu()

Dim oBookmark As Bookmark
Dim oBar As CommandBar
Dim oPopup As CommandBarPopup
Dim oButton As CommandBarButton
Dim ShowHiddenStatus As Boolean

'Find out whether hidden bookmarks set as "visible" or not,
'storing this setting in a variable so it can be returned to at the end.
'Then make the hidden bookmarks invisible
'(we don't want cross-refs etc to appear in our menu)
ShowHiddenStatus = ActiveDocument.Bookmarks.ShowHidden
ActiveDocument.Bookmarks.ShowHidden = False

CustomizationContext = ActiveDocument
Set oBar = CommandBars.ActiveMenuBar

'First delete Bookmark menu if it already exists
Set oPopup = CommandBars.FindControl(Tag:="Recreate")
If Not oPopup Is Nothing Then
    oPopup.Delete
End If


If ActiveDocument.Bookmarks.Count > 0 Then

      Set oPopup = oBar.Controls.Add(Type:=msoControlPopup, _
               Before:=oBar.Controls.Count + 1)

      With oPopup
          .Caption = "Bookmarks"
          .Tag = "Recreate"
      End With

      For Each oBookmark In ActiveDocument.Bookmarks
          Set oButton = oPopup.Controls.Add(Type:=msoControlButton)

          With oButton
              .Caption = oBookmark.Name
              .Style = msoButtonCaption
              .OnAction = "BookMarkSelect"
          End With

      Next

      'Add a Refresh button at the bottom
      Set oButton = oPopup.Controls.Add(Type:=msoControlButton)

      With oButton
          .Caption = "Refresh list"
          .Style = msoButtonCaption
          .OnAction = "CreateBookMarkMenu"
          .BeginGroup = True
      End With

End If

ActiveDocument.Bookmarks.ShowHidden = ShowHiddenStatus

Set oButton = Nothing
Set oPopup = Nothing
Set oBar = Nothing
Set oBookmark = Nothing

End Sub

Private Sub BookMarkSelect()
    If ActiveDocument.Bookmarks.Exists(CommandBars.ActionControl.Caption) Then
        ActiveDocument.Bookmarks(CommandBars.ActionControl.Caption). _
                Range.Select
    End If
End Sub

Sub AutoOpen()
    'Make sure the document's menu is visible when the document opens
    'If the "customisation context" has been changed since it was last opened,
    'the document-specific menus won't be visible!
    CustomizationContext = ActiveDocument
End Sub

