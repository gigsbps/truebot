Attribute VB_Name = "mLVSort"
'Make sure the module name is
'mLVSort.  Set this in
'your properties window

Option Explicit
Public objFind As LV_FINDINFO
Public objItem As LV_ITEM
  
'variable to hold the sort order (ascending or descending)
Public sOrder As Boolean
'variable to hold sort column
Public sColumn As Long
Public sTag

Public Type POINT
  x As Long
  y As Long
End Type

Public Type LV_FINDINFO
  flags As Long
  psz As String
  lParam As Long
  pt As POINT
  vkDirection As Long
End Type

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
    iTag As Variant
End Type
 
'Constants
Public Const LVFI_PARAM = 1
Public Const LVIF_TEXT = &H1

Public Const LVM_FIRST = &H1000
Public Const LVM_FINDITEM = LVM_FIRST + 13
Public Const LVM_GETITEMTEXT = LVM_FIRST + 45
Public Const LVM_SORTITEMS = LVM_FIRST + 48
     
'API declarations
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long
Public Function CompareDates(ByVal lParam1 As Long, _
                             ByVal lParam2 As Long, _
                             ByVal hWnd As Long) As Long
     
'CompareDates: This is the sorting routine that gets passed to the
'ListView control to provide the comparison test for date values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than

Dim dDate1 As Date, dDate2 As Date, dE As Boolean, d2E As Boolean
On Error GoTo CDERR

  'Obtain the item names and dates corresponding to the
  'input parameters
   dDate1 = ListView_GetItemDate(hWnd, lParam1)
   dDate2 = ListView_GetItemDate(hWnd, lParam2)
     
  'based on the Public variable sOrder set in the
  'columnheader click sub, sort the dates appropriately:
   Select Case sOrder
      Case True:    'sort descending
            
            If dDate1 < dDate2 Then
                  CompareDates = 0
            ElseIf dDate1 = dDate2 Then
                  CompareDates = 1
            Else
                CompareDates = 2
            End If
      
      Case Else: 'sort ascending
   
            If dDate1 > dDate2 Then
                  CompareDates = 0
            ElseIf dDate1 = dDate2 Then
                  CompareDates = 1
            Else
                CompareDates = 2
            End If
   
   End Select
   Exit Function
CDERR:
    CompareDates = 1
End Function


Public Function CompareValues(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hWnd As Long) As Long
     
'CompareValues: This is the sorting routine that gets passed to the
'ListView control to provide the comparison test for numeric values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
  
Dim val1 As Long, val2 As Long
On Error GoTo CDERR
    'Obtain the item names and values corresponding
    'to the input parameters
    val1 = ListView_GetItemValueStr(hWnd, lParam1)
    val2 = ListView_GetItemValueStr(hWnd, lParam2)
     
    'based on the Public variable sOrder set in the
    'columnheader click sub, sort the values appropriately:
    Select Case sOrder
        Case True:    'sort descending
            
            If val1 < val2 Then
                CompareValues = 0
            ElseIf val1 = val2 Then
                CompareValues = 1
            Else
                CompareValues = 2
            End If
      
        Case Else: 'sort ascending
   
            If val1 > val2 Then
                CompareValues = 0
            ElseIf val1 = val2 Then
                CompareValues = 1
            Else
                CompareValues = 2
            End If
   
    End Select
    Exit Function
CDERR:
    CompareValues = 1
End Function

Public Function CompareCurrency(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hWnd As Long) As Long
     
'CompareValues: This is the sorting routine that gets passed to the
'ListView control to provide the comparison test for numeric values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
  
Dim val1 As Currency, val2 As Currency
On Error GoTo CDERR
    'Obtain the item names and values corresponding
    'to the input parameters
    val1 = ListView_GetItemCurrency(hWnd, lParam1)
    val2 = ListView_GetItemCurrency(hWnd, lParam2)
     
    'based on the Public variable sOrder set in the
    'columnheader click sub, sort the values appropriately:
    Select Case sOrder
        Case True:    'sort descending
            
            If val1 < val2 Then
                CompareCurrency = 0
            ElseIf val1 = val2 Then
                CompareCurrency = 1
            Else
                CompareCurrency = 2
            End If
      
        Case Else: 'sort ascending
   
            If val1 > val2 Then
                CompareCurrency = 0
            ElseIf val1 = val2 Then
                CompareCurrency = 1
            Else
                CompareCurrency = 2
            End If
   
    End Select
    Exit Function
CDERR:
    CompareCurrency = 1
End Function

Public Function ComparePercent(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hWnd As Long) As Long
     
'CompareValues: This is the sorting routine that gets passed to the
'ListView control to provide the comparison test for numeric values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
  
Dim val1 As Single, val2 As Single
On Error GoTo CDERR
    'Obtain the item names and values corresponding
    'to the input parameters
    val1 = ListView_GetItemPercent(hWnd, lParam1)
    val2 = ListView_GetItemPercent(hWnd, lParam2)
     
    'based on the Public variable sOrder set in the
    'columnheader click sub, sort the values appropriately:
    Select Case sOrder
        Case True:    'sort descending
            
            If val1 < val2 Then
                ComparePercent = 0
            ElseIf val1 = val2 Then
                ComparePercent = 1
            Else
                ComparePercent = 2
            End If
      
        Case Else: 'sort ascending
   
            If val1 > val2 Then
                ComparePercent = 0
            ElseIf val1 = val2 Then
                ComparePercent = 1
            Else
                ComparePercent = 2
            End If
   
    End Select
    Exit Function
CDERR:
    ComparePercent = 1
End Function

Private Function ListView_GetItemDate(hWnd As Long, lParam As Long) As Date
Dim r As Long, hIndex As Long
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = sColumn
    objItem.pszText = Space$(32)
    objItem.cchTextMax = Len(objItem.pszText)
     
    'get the string at subitem 1
    r = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
     
    'and convert it into a date and exit
    If r > 0 Then
        If IsDate(Left$(objItem.pszText, r)) Then
            ListView_GetItemDate = CDate(Left$(objItem.pszText, r))
        Else
            ListView_GetItemDate = DateSerial(4501, 1, 1)
        End If
    End If
End Function


Public Function ListView_GetItemValueStr(hWnd As Long, lParam As Long) As Long
Dim r As Long, hIndex As Long
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = sColumn
    objItem.pszText = Space$(32)
    objItem.cchTextMax = Len(objItem.pszText)
    
    'get the string at subitem 2
    r = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
     
    'and convert it into a long
    If r > 0 Then
        ListView_GetItemValueStr = CLng(Left$(objItem.pszText, r))
    End If
End Function

Public Function ListView_GetItemCurrency(hWnd As Long, lParam As Long) As Long
Dim r As Long, hIndex As Long
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = sColumn
    objItem.pszText = Space$(32)
    objItem.cchTextMax = Len(objItem.pszText)
     
    'get the string at subitem 2
    r = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
     
    'and convert it into a long
    If r > 0 Then
        ListView_GetItemCurrency = CCur(Left$(objItem.pszText, r))
    End If
End Function

Public Function ListView_GetItemPercent(hWnd As Long, lParam As Long) As Long
Dim r As Long, hIndex As Long, temp As String
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = sColumn
    objItem.pszText = Space$(32)
    objItem.cchTextMax = Len(objItem.pszText)
     
    'get the string at subitem 2
    r = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
     
    'and convert it into a long
    If r > 0 Then
        temp = Left$(objItem.pszText, r)
        If Right$(temp, 1) = "%" Then
            temp = Left$(temp, Len(temp) - 1)
        End If
        ListView_GetItemPercent = CSng(temp)
    End If
End Function

Public Sub SortLvwOnDate(lvw As ListView, ColIndex As Long)
    lvw.Sorted = False
    If lvw.SortKey = ColIndex - 1 Then
        If lvw.SortOrder = lvwAscending Then
            lvw.SortOrder = lvwDescending
        Else
            lvw.SortOrder = lvwAscending
        End If
    Else
        lvw.SortKey = ColIndex - 1
        lvw.SortOrder = lvwAscending
    End If
    'mLVSort.sTag = lvw.Tag
    mLVSort.sColumn = ColIndex - 1
    mLVSort.sOrder = (lvw.SortOrder = lvwAscending)
    SendMessageLong lvw.hWnd, LVM_SORTITEMS, lvw.hWnd, AddressOf CompareDates
End Sub

Public Sub SortLvwOnLong(lvw As ListView, ColIndex As Long)
    lvw.Sorted = False
    If lvw.SortKey = ColIndex - 1 Then
        If lvw.SortOrder = lvwAscending Then
            lvw.SortOrder = lvwDescending
        Else
            lvw.SortOrder = lvwAscending
        End If
    Else
        lvw.SortKey = ColIndex - 1
        lvw.SortOrder = lvwAscending
    End If
'    mLVSort.sTag
    mLVSort.sColumn = ColIndex - 1
    mLVSort.sOrder = (lvw.SortOrder = lvwAscending)
    SendMessageLong lvw.hWnd, LVM_SORTITEMS, lvw.hWnd, AddressOf CompareValues
End Sub

Public Sub SortLvwOnCurrency(lvw As ListView, ColIndex As Long)
    lvw.Sorted = False
    If lvw.SortKey = ColIndex - 1 Then
        If lvw.SortOrder = lvwAscending Then
            lvw.SortOrder = lvwDescending
        Else
            lvw.SortOrder = lvwAscending
        End If
    Else
        lvw.SortKey = ColIndex - 1
        lvw.SortOrder = lvwAscending
    End If
    mLVSort.sColumn = ColIndex - 1
    mLVSort.sOrder = (lvw.SortOrder = lvwAscending)
    SendMessageLong lvw.hWnd, LVM_SORTITEMS, lvw.hWnd, AddressOf CompareCurrency
End Sub

Public Sub SortLvwOnPercent(lvw As ListView, ColIndex As Long)
    lvw.Sorted = False
    If lvw.SortKey = ColIndex - 1 Then
        If lvw.SortOrder = lvwAscending Then
            lvw.SortOrder = lvwDescending
        Else
            lvw.SortOrder = lvwAscending
        End If
    Else
        lvw.SortKey = ColIndex - 1
        lvw.SortOrder = lvwAscending
    End If
    mLVSort.sColumn = ColIndex - 1
    mLVSort.sOrder = (lvw.SortOrder = lvwAscending)
    SendMessageLong lvw.hWnd, LVM_SORTITEMS, lvw.hWnd, AddressOf ComparePercent
End Sub
