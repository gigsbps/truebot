VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPCSCombat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public uID As Long
'                        unknown = cInt2(din, begin + 6)
Public ownerid As Long
Public Row As Integer
Public Col As Integer
Public MAXHP As Long
Public MAXSP As Long
'Public HP As Long
'Public SP As Long
Public lv As Integer
Public elem As Integer


'local variable(s) to hold property value(s)
Private mvarHP As Long 'local copy
Private mvarSP As Long 'local copy
Public Property Let SP(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SP = 5
    mvarSP = vData
End Property


Public Property Get SP() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SP
    SP = mvarSP
End Property



Public Property Let HP(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HP = 5
    mvarHP = vData
End Property


Public Property Get HP() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HP
    HP = mvarHP
End Property



