Attribute VB_Name = "Module1DLL"




Public Function a2hex(alpha, length)
    hextemplate = "0123456789ABCDEF"
    alpha = Right("00000000" & UCase(alpha), length)
    If Len(alpha) = 2 Then
        ahindex = InStr(1, hextemplate, Left(alpha, 1)) - 1
        alindex = InStr(1, hextemplate, Right(alpha, 1)) - 1
        a2hex = (ahindex * (16 ^ 1)) + (alindex * (16 ^ 0))
    ElseIf Len(alpha) = 4 Then
        lb = Left(alpha, 2)
        hb = Right(alpha, 2)
        hx = a2hex(hb, 2) * (16 ^ 2) + a2hex(lb, 2)
        a2hex = hx
    ElseIf Len(alpha) = 8 Then
        aH = a2hex(Right(alpha, 4), 4)
        aL = a2hex(Left(alpha, 4), 4)
        a2hex = (aH * (16 ^ 4)) + aL
    End If
End Function

Function bytes2hexStr(din)
    For i = 1 To UBound(din)
            c = c & n2h(din(i), 1)
    Next
    bytes2hexStr = c
End Function

Function DecodePacket(value)
    Dim mask As Byte
        mask = &HAD
        DecodePacket = value Xor mask
End Function
Public Function DecodePacketString(ByVal strRawPacket As String) As String
    Dim i As Long
    Dim hstr As String
        hstr = Replace(strRawPacket, " ", "")
        
    Debug.Print hstr
    
    Dim pHex As String
    pHex = ""
    'ReDim Preserve pHex((Len(hstr) / 2 - 1))
    'ReDim Preserve pHex(Len(Hstr) - 1)
    
    Dim vb As ScriptControl
    Set vb = New ScriptControl
    
    vb.Language = "vbscript"
    
    For i = 0 To Len(hstr) / 2 - 1
        pHex = pHex & " " & Right("00" & Hex(DecodePacket(vb.Eval("&H" & Mid(hstr, (i * 2) + 1, 2)))), 2)
      
    Next
    'Size = CInt(Len(strPacket) / 2)
    'ReDim pHex(size)
    DecodePacketString = pHex
    
End Function

Public Function MakePacket(ByVal strPacket As String) As Byte()
    Dim i As Long
    Dim hstr As String
        hstr = Replace(strPacket, " ", "")
     Dim pHex() As Byte
   ReDim Preserve pHex((Len(hstr) / 2 - 1))
    'ReDim Preserve pHex(Len(Hstr) - 1)
    For i = 0 To Len(hstr) / 2 - 1
      pHex(i) = DecodePacket(a2hex(Mid(hstr, (i * 2) + 1, 2), 2))
    Next
    Size = CInt(Len(strPacket) / 2)
   ' ReDim pHex(size)
    MakePacket = pHex
End Function

Public Function getNumeric(data, pos, nbytes)
    n = nbytes * 2
    getNumeric = a2hex(Mid(data, pos, n), n)
End Function


Public Function cInt1(ByRef d() As Byte, pos) As Long
    cInt1 = d(pos)
End Function
Public Function cInt2(ByRef d() As Byte, pos)
    hx = d(pos + 1)
    lx = d(pos)
    cInt2 = hx * (16 ^ 2) + lx
End Function
Public Function cInt4(ByRef d() As Byte, pos)
    hx = d(pos + 3)
    lx = d(pos + 2)
    v1 = hx * (16 ^ 2) + lx
    hx2 = d(pos + 1)
    lx2 = d(pos)
    v2 = hx2 * (16 ^ 2) + lx2
    cInt4 = v1 * (16 ^ 4) + v2
End Function
Public Function cStrN(ByRef d() As Byte, pos, n)
    buf = ""
    For i = pos To pos + n
    
        buf = buf & Chr(d(i))
    Next
    cStrN = buf
End Function

Function String2Hex(str)
    cx = ""
    For i = 0 To Len(str) - 1
        cx = cx & Hex(Asc(Mid(str, i + 1, 1)))
    Next
    String2Hex = cx
End Function
Function n2h(number, nb)
    For i = 1 To (nb * 2)
        zerotem = zerotem & "0"
    Next
    rn = Right(zerotem & Hex(number), (nb * 2))
    out = ""
    For i = 1 To Len(rn) Step 2
       out = Mid(rn, i, 2) & out
    Next
    n2h = out
End Function
Function a2hstr(str)
hstr = ""
        For i = 1 To Len(str)
            ch = Hex(Asc(Mid(str, i, 1)))
            hstr = hstr & ch
        Next
a2hstr = hstr
End Function
