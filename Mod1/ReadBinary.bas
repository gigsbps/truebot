Attribute VB_Name = "ReadBinary"
Option Explicit

Private Const BlockSize = 32768 * 1024


Function ReadFile(sFileName As String, ByRef fsize As Long) As Variant
 On Error Resume Next
  
    Dim i As Integer
    Dim FileLength As Long, LeftOver As Long
    Dim RetVal As Variant

    ' Open the source file.
    Dim hFile As Integer
    hFile = FreeFile
    Open sFileName For Binary Access Read As hFile

    ' Get the length of the file.
    Dim nFileSize As Long
    nFileSize = LOF(hFile)
    fsize = nFileSize
    If nFileSize = 0 Then
        ReadFile = Empty
        Exit Function
    End If

    ' Read file
    Dim arrData() As Byte
    ReDim arrData(nFileSize)
    Get hFile, , arrData
    Close hFile
    
    ReadFile = arrData

End Function
