Set Range1 = ActiveDocument.Words(1) 
Set Range2 = ActiveDocument.Words(2)

'Setting a Range object variable equal to another Range object variable

Set Range2 = Range1
Set Range2 = Range1.Duplicate