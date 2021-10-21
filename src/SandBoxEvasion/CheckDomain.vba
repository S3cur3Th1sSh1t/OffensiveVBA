Sub CheckDomain()
On Error Resume Next
Set objRootDSE = GetObject("LDAP://RootDSE")
If Err.Number <> 0 Then
wscript.Quit
Else
MsgBox ("Evil")
End If
On Error GoTo 0
End Sub
