Sub SearchWindow32()
Dim chromePath As String
Dim search_string As String
Dim query As String
query = InputBox("Enter here your search here", "Google Search")
search_string = query
search_string = Replace(search_string, " ", "+")
'Uncomment the following line for Windows 64 versions and comment out Windows 32 versions'
'chromePath = "C:Program FilesGoogleChromeApplicationchrome.exe"
'Uncomment the following line for Windows 32 versions and comment out Windows 64 versions
'chromePath = "C:Program Files (x86)GoogleChromeApplicationchrome.exe"
Shell (chromePath & " -url http://google.com/#q=" & search_string)
End Sub