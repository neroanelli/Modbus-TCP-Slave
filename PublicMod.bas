Attribute VB_Name = "PublicMod"
Public OpenFlag As Boolean
Public ConnectFlag As Boolean
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Function Hexn(ByVal number As Long, ByVal n As Integer) As String
Dim str As String

    str = String(n, "0") + Hex(number)
    str = Right(str, n)
    Hexn = str
End Function


Function Crc_16(ByVal Str1 As String) As Long
Dim i As Integer
Dim j As Integer
Dim CVal As Long
Dim Temp1 As Integer
Dim Const1 As Long
    CVal = 65535        '&HFFFF
    Const1 = 40961      '&HA001
    For i = 1 To LenB(Str1)
        Temp1 = AscB(MidB(Str1, i, 1))
        CVal = Temp1 Xor CVal
        For j = 0 To 7
            If (CVal Mod 2) = 0 Then
                CVal = CVal \ 2
            Else
                CVal = CVal \ 2
                CVal = CVal Xor Const1
            End If
        Next j
    Next i
    Crc_16 = CVal And &HFFFF&
End Function
