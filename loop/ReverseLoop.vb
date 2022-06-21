Sub ReverseLoop()
'PURPOSE: Loop Through A Range of Cells in Reverse Order


Dim x As Long

'Loop Through 0-100 in Reverse Order
  For x = 100 To 0 Step -1
    ActiveSheet.Cells(x, 2).Value = x * 100
  Next x

End Sub