Sub LoopThroughFiles ()
 
Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
 
Set oFSO = CreateObject("Scripting.FileSystemObject")
 
Set oFolder = oFSO.GetFolder("C:\Demo)
 
i = 2
 
For Each oFile In oFolder.Files
    Range("A" & i).value = oFile.Name
    i = i + 1
Next oFile
 
End Sub