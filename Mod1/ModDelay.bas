Attribute VB_Name = "ModDelay"
Option Explicit
'********************************************
'*    (c) 1999-2000 Sergey Merzlikin        *
'********************************************

Public Const STATUS_TIMEOUT = &H102&
Public Const INFINITE = -1& ' Infinite interval
Public Const QS_KEY = &H1&
Public Const QS_MOUSEMOVE = &H2&
Public Const QS_MOUSEBUTTON = &H4&
Public Const QS_POSTMESSAGE = &H8&
Public Const QS_TIMER = &H10&
Public Const QS_PAINT = &H20&
Public Const QS_SENDMESSAGE = &H40&
Public Const QS_HOTKEY = &H80&
Public Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT _
        Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON _
        Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Public Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

' The MsgWaitObj function replaces Sleep,
' WaitForSingleObject, WaitForMultipleObjects functions.
' Unlike these functions, it
' doesn't block thread messages processing.
' Using instead Sleep:
'     MsgWaitObj dwMilliseconds
' Using instead WaitForSingleObject:
'     retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
' Using instead WaitForMultipleObjects:
'     retval = MsgWaitObj(dwMilliseconds, hObj(0&), n),
'     where n - wait objects quantity,
'     hObj() - their handles array.

Public Function MsgWaitObj(Interval As Long, _
            Optional hObj As Long = 0&, _
            Optional nObj As Long = 0&) As Long
Dim T As Long, T1 As Long
If Interval <> INFINITE Then
    T = GetTickCount()
    On Error Resume Next
    T = T + Interval
    ' Overflow prevention
    If Err <> 0& Then
        If T > 0& Then
            T = ((T + &H80000000) _
            + Interval) + &H80000000
        Else
            T = ((T - &H80000000) _
            + Interval) - &H80000000
        End If
    End If
    On Error GoTo 0
    ' T contains now absolute time of the end of interval
Else
    T1 = INFINITE
End If
Do
    If Interval <> INFINITE Then
        T1 = GetTickCount()
        On Error Resume Next
     T1 = T - T1
        ' Overflow prevention
        If Err <> 0& Then
            If T > 0& Then
                T1 = ((T + &H80000000) _
                - (T1 - &H80000000))
            Else
                T1 = ((T - &H80000000) _
                - (T1 + &H80000000))
            End If
        End If
        On Error GoTo 0
        ' T1 contains now the remaining interval part
        If IIf((T1 Xor Interval) > 0&, _
            T1 > Interval, T1 < 0&) Then
            ' Interval expired
            ' during DoEvents
            MsgWaitObj = STATUS_TIMEOUT
            Exit Function
        End If
    End If
    ' Wait for event, interval expiration
    ' or message appearance in thread queue
    MsgWaitObj = MsgWaitForMultipleObjects(nObj, _
            hObj, 0&, T1, QS_ALLINPUT)
    ' Let's message be processed
    DoEvents
    If MsgWaitObj <> nObj Then Exit Function
    ' It was message - continue to wait
Loop
End Function
