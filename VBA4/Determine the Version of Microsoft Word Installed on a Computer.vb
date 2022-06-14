Set objWord = CreateObject("Word.Application")
Wscript.Echo "Version: " & objWord.Version
Wscript.Echo "Build: " & objWord.Build
Wscript.Echo "Product Code: " & _
    objWord.ProductCode()
objWord.Quit