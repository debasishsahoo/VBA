Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()

objWord.FileSearch.FileName = "*.mp3"
objWord.FileSearch.LookIn = "C:\"
objWord.FileSearch.SearchSubfolders = True
objWord.FileSearch.Execute

For Each objFile in objWord.FileSearch.FoundFiles
    Wscript.Echo objFile
Next

objWord.Quit