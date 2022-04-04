'In this Example I am changing the Range B4 Font to Bold
Sub sbChangeFontToBold()

'Using Cell Object
Cells(4, 2).Font.Bold = True

'Using Range Object
Range("B4").Font.Bold = True

End Sub