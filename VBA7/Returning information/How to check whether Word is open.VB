Sub AutoExec()

Dim Hits As Long, oTask As Task

  For Each oTask In Tasks
      If Left(oTask.Name, 14) = "Microsoft Word" Then Hits = Hits + 1
      If Hits = 2 Then
          oTask.WindowState = wdWindowStateMaximize
          Application.Quit
      End If
  Next oTask

End Sub