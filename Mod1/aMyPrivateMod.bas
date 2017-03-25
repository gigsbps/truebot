Attribute VB_Name = "aMyPrivateMod"
Option Explicit

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
    
Public tempMoveSlot As Integer
Public tempMoveNum As Integer
Public RequestToWork As Boolean

Public LicenseId() As Long
Public LastText As String
Public spcMode As Boolean
Public logID As String
Public logPass As String

Public Type MemberDetail
    id As Long
    Name As String
    Lvl As Integer
    ele As Integer
    Con As Long
    Reborn As Boolean
End Type

Public Type ArmyX
    AName As String
    ASubCount As Integer
    AMemCount As Integer
    ALeader As MemberDetail
    ASubLeader(1 To 3) As MemberDetail
    AMember(1 To 95) As MemberDetail
End Type
    
Public myArmy As ArmyX
Public memCount As Integer

Public Sub addLicId(id As Long, ByRef cLast As Integer)
    ReDim Preserve LicenseId(cLast + 1)
    LicenseId(cLast + 1) = id
        cLast = cLast + 1
End Sub

Public Function checkLicId(id As Long) As Boolean
Dim kkk As Integer
    If (UBound(LicenseId) > 0) And (id > 0) Then
        For kkk = 1 To UBound(LicenseId)
            If id = LicenseId(kkk) Then
                checkLicId = True
                Exit For
            Else
                checkLicId = False
            End If
        Next kkk
    End If
End Function

Public Sub initLicenseList()
Dim currentLast As Integer
    spcMode = False
    currentLast = 0
    addLicId 158168, currentLast
    addLicId 284931, currentLast
    addLicId 661976, currentLast
End Sub

Function Hex2Bin(ByVal strHex As String) As String
    Dim i           As Integer
    Dim j           As Integer
    Dim dec         As Integer
    Dim tmpDec      As Integer
    Dim strbin      As String
    Dim HexChars    As String
    
    HexChars = "0123456789ABCDEF"
    strbin = ""
    For i = 1 To Len(strHex)
        dec = InStr(1, HexChars, Mid(strHex, i, 1)) - 1
        tmpDec = 0
        For j = 3 To 0 Step -1
            If tmpDec + (2 ^ j) <= dec Then
                strbin = strbin & "1"
                tmpDec = tmpDec + (2 ^ j)
            Else
                strbin = strbin & "0"
            End If
        Next
        
    Next
    Hex2Bin = strbin
End Function

Function Hex2Double(ByVal strHex As String) As Double
    Dim i           As Integer
    Dim j           As Integer
    Dim sign        As Integer
    Dim expo        As Integer
    Dim dec         As Double
    Dim strbin      As String
    Dim strtmp      As String
    
    strtmp = ""
    For i = (Len(strHex) - 1) To 1 Step -2
        strtmp = strtmp & Mid(strHex, i, 2)
    Next
    
    strbin = Hex2Bin(strtmp)
    ' bit 63 Sign Bit
    sign = 1
    If Mid(strbin, 1, 1) = "1" Then
        sign = -1
    End If
    
    ' Bits 62 - 52 Exponent Field
    expo = 0
    For i = 2 To 12
        If Mid(strbin, i, 1) = "1" Then
            expo = expo + (2 ^ (12 - i))
        End If
    Next
    ' Bits 51 - 0 Significand
    dec = 1
    For i = 13 To 64
        If Mid(strbin, i, 1) = "1" Then
            dec = dec + (2 ^ (12 - i))
        End If
    Next
    Hex2Double = sign * (2 ^ (expo - 1023)) * dec
End Function


Public Function ele2text(ByVal ele As Integer) As String
    Select Case ele
        Case 1
            ele2text = "ดิน"
        Case 2
            ele2text = "น้ำ"
        Case 3
            ele2text = "ไฟ"
        Case 4
            ele2text = "ลม"
        Case 5
            ele2text = "จิต"
        Case 0
            ele2text = "ไร้ธาตุ"
        Case Else
            ele2text = ele
    End Select
End Function

Public Function PlusOrMin(ByVal ssign As Integer) As String
    If ssign = 1 Then
        PlusOrMin = "-"
    ElseIf ssign = 0 Then
        PlusOrMin = "+"
    Else
        PlusOrMin = ""
    End If
End Function

Public Function StatusText(ByVal ssign As Integer) As String
    If ssign = &H19 Then
        StatusText = " Hp "
    ElseIf ssign = &H1A Then
        StatusText = " Sp "
    ElseIf ssign >= 220 And ssign <= 223 Then
        StatusText = " สำเร็จ "
    ElseIf ssign = 0 Then
        StatusText = " Miss "
    Else
        StatusText = " *" & ssign & " "
    End If
End Function

