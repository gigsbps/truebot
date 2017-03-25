VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DamageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarAttkFromRow As Integer 'local copy
Private mvarAttkFromCol As Integer 'local copy
Private mvarAttkToRow As Integer 'local copy
Private mvarAttkToCol As Integer 'local copy
Private mvarAttkSkill As Long 'local copy
Private mvarDamagePoint As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarnEffect As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarDmgTarget As Scripting.Dictionary 'local copy
Public Property Set DmgTarget(ByVal vData As Scripting.Dictionary)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DmgTarget = Form1
    Set mvarDmgTarget = vData
End Property


Public Property Get DmgTarget() As Scripting.Dictionary
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DmgTarget
    Set DmgTarget = mvarDmgTarget
End Property



Public Property Let nEffect(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.nEffect = 5
    mvarnEffect = vData
End Property


Public Property Get nEffect() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.nEffect
    nEffect = mvarnEffect
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



Public Property Let AttkSkill(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttkSkill = 5
    mvarAttkSkill = vData
End Property


Public Property Get AttkSkill() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttkSkill
    AttkSkill = mvarAttkSkill
End Property



Public Property Let AttkToCol(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttkToCol = 5
    mvarAttkToCol = vData
End Property


Public Property Get AttkToCol() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttkToCol
    AttkToCol = mvarAttkToCol
End Property



Public Property Let AttkToRow(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttkToRow = 5
    mvarAttkToRow = vData
End Property


Public Property Get AttkToRow() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttkToRow
    AttkToRow = mvarAttkToRow
End Property



Public Property Let AttkFromCol(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttkFromCol = 5
    mvarAttkFromCol = vData
End Property


Public Property Get AttkFromCol() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttkFromCol
    AttkFromCol = mvarAttkFromCol
End Property



Public Property Let AttkFromRow(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttkFromRow = 5
    mvarAttkFromRow = vData
End Property


Public Property Get AttkFromRow() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttkFromRow
    AttkFromRow = mvarAttkFromRow
End Property



Private Sub Class_Initialize()
    Set mvarDmgTarget = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
  Set mvarDmgTarget = Nothing
End Sub
