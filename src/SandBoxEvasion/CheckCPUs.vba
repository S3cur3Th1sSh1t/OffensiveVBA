Function CheckAvailableCPUs() 
  TnHWRZVH=Cint("0")
  Set VRfvQEHCs=GetObject("winmgmts:\\.\root\cimv2")
  Set uxSpQuH=VRfvQEHCs.ExecQuery("Select * from Win32_Processor", , Cint("48"))
  For Each XqDiZDh In uxSpQuH
    If XqDiZDh.NumberOfCores < Cint("2") Then
      TnHWRZVH=True
    End If
  Next
  If TnHWRZVH Then
    wscript.Quit
  Else
  End If
End Function
