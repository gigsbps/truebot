VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CharStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarINTX As Long 'local copy
Private mvarATKX As Long 'local copy
Private mvarDEFX As Long 'local copy
Private mvarAGIX As Long 'local copy
Private mvarHPX As Long 'local copy
Private mvarSPX As Long 'local copy
Public Property Let SPX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SPX = 5
    mvarSPX = vData
End Property


Public Property Get SPX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SPX
    SPX = mvarSPX
End Property



Public Property Let HPX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HPX = 5
    mvarHPX = vData
End Property


Public Property Get HPX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HPX
    HPX = mvarHPX
End Property



Public Property Let AGIX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AGIX = 5
    mvarAGIX = vData
End Property


Public Property Get AGIX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AGIX
    AGIX = mvarAGIX
End Property



Public Property Let DEFX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DEFX = 5
    mvarDEFX = vData
End Property


Public Property Get DEFX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DEFX
    DEFX = mvarDEFX
End Property



Public Property Let ATKX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ATKX = 5
    mvarATKX = vData
End Property


Public Property Get ATKX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ATKX
    ATKX = mvarATKX
End Property



Public Property Let INTX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.INTX = 5
    mvarINTX = vData
End Property


Public Property Get INTX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.INTX
    INTX = mvarINTX
End Property




