Attribute VB_Name = "IniMod"
' The following module uses API functions available in kernel32
' to interact with specified files

Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
    As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetPrivateProfileSection Lib "kernel32" _
    Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    ' Get the section from the INI file
    ' lpAppName refers to the section in the INI file
    ' lpReturnedString refers to the retrieved string
    ' nSize refers to the size of retrieved string
    ' lpFileName refers to the INI file

Public Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
   As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    ' Gets the settings from INI file
    ' lpApplicationName refers to the section in the INI file
    ' lpKeyName refers to the key we wish to change
    ' lpDefault refers to the value the key takes if no value is entered for it
    ' lpReturnedString refers to the retrieved string
    ' nSize refers to the size of retrieved string
    ' lpFileName refers to the INI file

Public Declare Function WritePrivateProfileSection Lib "kernel32" _
    Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
    ByVal lpString As String, ByVal lpFileName As String) As Long
    ' Writes the section to the INI file
    ' lpAppName refers to the section in the INI file
    ' lpString refers to the string being written
    ' lpFileName refers to the INI file
    
Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    ' Writes the settings to the INI file
    ' lpApplicationName refers to the section in the INI file
    ' lpKeyName refers to the Key being written to
    ' lpString refers to the value being written to the specified Key
    ' lpFileName refers to the INI file
    
    'Public gstrKeyValue As String * 256


'Public Declare Function fnTrueDLL Lib "TrueDLL" () As Integer

