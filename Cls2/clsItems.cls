VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public itemid As Long
Public itemname As String
Public itemtype As String
Public itemvalue As Long
'local variable(s) to hold property value(s)
Private mvaritemDesc As String 'local copy
Private mvaritemvalue2 As Long 'local copy
Private mvaritemtype2 As String 'local copy
'local variable(s) to hold property value(s)
Private mvareqposition As String 'local copy
Public Property Let eqposition(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.eqposition = 5
    mvareqposition = vData
End Property


Public Property Get eqposition() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.eqposition
    eqposition = mvareqposition
End Property



Public Property Let itemtype2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.itemtype2 = 5
    mvaritemtype2 = vData
End Property


Public Property Get itemtype2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.itemtype2
    itemtype2 = mvaritemtype2
End Property



Public Property Let itemvalue2(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.itemvalue2 = 5
    mvaritemvalue2 = vData
End Property


Public Property Get itemvalue2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.itemvalue2
    itemvalue2 = mvaritemvalue2
End Property



Public Property Let itemdesc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.itemDesc = 5
    mvaritemDesc = vData
End Property


Public Property Get itemdesc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.itemDesc
    itemdesc = mvaritemDesc
End Property




Public Function isHPItem()
    isHPItem = IsHP(itemid)
End Function
Public Function isSPItem()
    isSPItem = IsSP(itemid)
End Function

Public Function getName()
    getName = getItemName(itemid)
End Function
