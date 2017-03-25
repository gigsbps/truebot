VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DamageTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarRow As Integer 'local copy
Private mvarCol As Integer 'local copy
Private mvarDamagePoint As Long 'local copy
Private mvarMiss As Boolean 'local copy

Private mvarStatusCount As Integer
Private mvarDamagePoint2 As Long
Private mvarStatus1 As Integer
Private mvarStatus2 As Integer
Private mvarSign1 As Integer
Private mvarSign2 As Integer

Public Property Let StatusCount(ByVal vData As Integer)
    mvarStatusCount = vData
End Property
Public Property Get StatusCount() As Integer
    StatusCount = mvarStatusCount
End Property

Public Property Let Status1(ByVal vData As Integer)
    mvarStatus1 = vData
End Property
Public Property Get Status1() As Integer
    Status1 = mvarStatus1
End Property

Public Property Let Status2(ByVal vData As Integer)
    mvarStatus2 = vData
End Property
Public Property Get Status2() As Integer
    Status2 = mvarStatus2
End Property

Public Property Let Sign2(ByVal vData As Integer)
    mvarSign2 = vData
End Property
Public Property Get Sign2() As Integer
    Sign2 = mvarSign2
End Property

Public Property Let Sign1(ByVal vData As Integer)
    mvarSign1 = vData
End Property
Public Property Get Sign1() As Integer
    Sign1 = mvarSign1
End Property

Public Property Let DamagePoint2(ByVal vData As Long)
    mvarDamagePoint2 = vData
End Property
Public Property Get DamagePoint2() As Long
    DamagePoint2 = mvarDamagePoint2
End Property


Public Property Let Miss(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Miss = 5
    mvarMiss = vData
End Property


Public Property Get Miss() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Miss
    Miss = mvarMiss
End Property



Public Property Let DamagePoint(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DamagePoint = 5
    mvarDamagePoint = vData
End Property


Public Property Get DamagePoint() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DamagePoint
    DamagePoint = mvarDamagePoint
End Property



Public Property Let Col(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Col = 5
    mvarCol = vData
End Property


Public Property Get Col() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Col
    Col = mvarCol
End Property



Public Property Let Row(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Row = 5
    mvarRow = vData
End Property


Public Property Get Row() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Row
    Row = mvarRow
End Property



