VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarlevel As Integer 'local copy
Private mvarmaxexp As Long 'local copy
Private mvartexp As Long 'local copy
Public Property Let Texp(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.texp = 5
    mvartexp = vData
End Property


Public Property Get Texp() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.texp
    Texp = mvartexp
End Property



Public Property Let maxexp(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.maxexp = 5
    mvarmaxexp = vData
End Property


Public Property Get maxexp() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.maxexp
    maxexp = mvarmaxexp
End Property



Public Property Let level(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.level = 5
    mvarlevel = vData
End Property


Public Property Get level() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.level
    level = mvarlevel
End Property



