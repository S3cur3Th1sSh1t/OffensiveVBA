Function CheckAvailableStorage()
  Set VRfvQEHCs=GetObject("winmgmts:\\.\root\cimv2")
  Set uxSpQuH=VRfvQEHCs.ExecQuery("Select * from Win32_LogicalDisk")
  For Each XqDiZDh In uxSpQuH
    zhlOwdMZ=zhlOwdMZ+Int(XqDiZDh.Size / Clng("1073741824"))
  Next
  If zhlOwdMZ < Cint("60") Then
    wscript.Quit
End If
End Function
