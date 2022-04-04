'In this Example I am changing the Range B4 Font Color
Sub sbChangeFontColor()

'Using Cell Object
Cells(4, 2).Font.ColorIndex = 3 ' 3 indicates Red Color

'Using Range Object
Range("B4").Font.ColorIndex = 3

'--- You can use use RGB, instead of ColorIndex -->
'Using Cell Object
Cells(4, 2).Font.Color = RGB(255, 0, 0)

'Using Range Object
Range("B4").Font.Color = RGB(255, 0, 0)

End Sub