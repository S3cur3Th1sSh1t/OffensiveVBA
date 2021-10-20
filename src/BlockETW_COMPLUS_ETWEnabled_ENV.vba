Sub AutoOpen()

Dim objUserEnvVars As Object
Set objUserEnvVars = CreateObject("WScript.Shell").Environment("User")
objUserEnvVars.Item("COMPlus_ETWEnabled") = "0"

End Sub
