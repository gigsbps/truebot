Attribute VB_Name = "ModSystray"
Option Explicit
' This code was adapted from original system tray module published by Ben Baird.
' Tray Icon add/remove functions implemented within this module.
' Created by E.Spencer (elliot@spnc.demon.co.uk) - This code is public domain.
' Added call back to handle mouse events correctly

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_ACTIVATEAPP = &H1C
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64
Public Const GWL_WNDPROC = (-4)

Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * MAX_TOOLTIP
End Type
Public nfIconData As NOTIFYICONDATA
Private FHandle As Long     ' Storage for form handle
Private WndProc As Long     ' Address of our handler
Private Hooking As Boolean  ' Hooking indicator

' Add your application to the system tray.
' Param 1 = Handle of form (which deals with sys tray events)
' Param 2 = Icon (form icon - any icon)
' Param 3 = Handle of icon (form icon - any icon)
' Param 4 = Tip for sys tray icon.
'
' Example - AddIconToTray Me.Hwnd, Me.Icon, Me.Icon.Handle, "This is a test tip"
'
Public Sub AddIconToTray(MeHwnd As Long, MeIcon As Long, MeIconHandle As Long, Tip As String)
With nfIconData
   .hwnd = MeHwnd
   .uID = MeIcon
   .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
   .uCallbackMessage = WM_RBUTTONUP
   .hIcon = MeIconHandle
   .szTip = Tip & Chr$(0)
   .cbSize = Len(nfIconData)
End With
Shell_NotifyIcon NIM_ADD, nfIconData
End Sub

' Remove your application from the system tray.
' Call when you quit your application.
'
Public Sub RemoveIconFromTray()
Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub

' Call this routine to ensure my app gets notified of all events
' Example - Hook Me.hWnd
'
Public Sub Hook(Lwnd As Long)
If Hooking = False Then
   FHandle = Lwnd
   WndProc = SetWindowLong(Lwnd, GWL_WNDPROC, AddressOf WindowProc)
   Hooking = True
End If
End Sub

' Call this routine to transfer event notification back to standard handler
' Example - Unhook
'
Public Sub Unhook()
If Hooking = True Then
   SetWindowLong FHandle, GWL_WNDPROC, WndProc
   Hooking = False
End If
End Sub

' Detect a right click event on our system tray icon - pass control to a handler routine
' in the main form (change as required)
Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Ensure that its our app thats affected and that its the right event
If Hooking = True Then
   If uMsg = WM_RBUTTONUP And lParam = WM_RBUTTONDOWN Then
      Form1.SysTrayMouseEventHandler  ' Pass the event back to the form handler
      WindowProc = True               ' Let windows know we handled it
      Exit Function
   End If
   WindowProc = CallWindowProc(WndProc, hw, uMsg, wParam, lParam) ' Pass it along
End If
End Function


