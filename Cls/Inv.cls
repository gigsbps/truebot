VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarslot As Integer 'local copy
Private mvaritemid As Long 'local copy
Private mvarnum As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarfine As Integer 'local copy
Public Property Let fine(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fine = 5
    mvarfine = vData
End Property


Public Property Get fine() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fine
    fine = mvarfine
End Property



Public Property Let num(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.num = 5
    mvarnum = vData
End Property


Public Property Get num() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.num
    num = mvarnum
End Property



Public Property Let itemid(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.itemid = 5
    mvaritemid = vData
End Property


Public Property Get itemid() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.itemid
    itemid = mvaritemid
End Property



Public Property Let slot(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.slot = 5
    mvarslot = vData
End Property


Public Property Get slot() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.slot
    slot = mvarslot
End Property



