Dim varArrayList As Variant
Dim strSelected As String
varArrayList = Array("value1", "value2", "value3")
strSelected = SelectionBoxSingle(List:=varArrayList)
If Len(strSelected) > 0 Then
    Debug.Print strSelected
End If
'SelectionBoxMulti to allow the user to select multiple values


Dim varArrayList As Variant
Dim varArraySelected As Variant
varArrayList = Array("value1", "value2", "value3")
varArraySelected = SelectionBoxMulti(List:=varArrayList, Prompt:="Select one or more values", _
                                    SelectionType:=fmMultiSelectMulti, Title:="Select multiple")
If Not IsEmpty(varArraySelected) Then 'not cancelled
    Debug.Print varArraySelected(0)
End If