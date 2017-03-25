VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CharStatus"
Attribute VB_Ext_KEY = "Member1" ,"CharStatus"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarisLicenseId As Boolean
Private mvaruid As Variant 'local copy
Private mvarCharName As String 'local copy
Private mvarlevel As Integer 'local copy
Private mvarElement As Integer 'local copy
Private mvarCharStatus As CharStatus
Private mvarRow As Integer 'local copy
Private mvarCol As Integer 'local copy
Private mvartexp As Long 'local copy
Private mvarOrder As Integer 'local copy
Private mvarfai As Integer 'local copy
Private mvarIsSleep As Boolean 'local copy
Private mvarHP As Long 'local copy
Private mvarSP As Long 'local copy
Private mvarMAXHP As Long 'local copy
Private mvarMAXSP As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarHor As Long 'local copy
Private mvararmy_rate As Scripting.Dictionary  'local copy
'local variable(s) to hold property value(s)
Private mvarSkills As Scripting.Dictionary 'local copy
'local variable(s) to hold property value(s)
Private mvarX As Long 'local copy
Private mvarY As Long 'local copy
Private mvarmapid As Long 'local copy
'local variable(s) to hold property value(s)
Private mvargold As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarNewBorn As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarGhost As Variant 'local copy
Private mvarGodOfLuck As Variant 'local copy
Public PvPranking
Public Property Let GodOfLuck(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GodOfLuck = 5
    mvarGodOfLuck = vData
End Property


Public Property Set GodOfLuck(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.GodOfLuck = Form1
    Set mvarGodOfLuck = vData
End Property


Public Property Get GodOfLuck() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GodOfLuck
    If IsObject(mvarGodOfLuck) Then
        Set GodOfLuck = mvarGodOfLuck
    Else
        GodOfLuck = mvarGodOfLuck
    End If
End Property



Public Property Let Ghost(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ghost = 5
    mvarGhost = vData
End Property


Public Property Set Ghost(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Ghost = Form1
    Set mvarGhost = vData
End Property


Public Property Get Ghost() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ghost
    If IsObject(mvarGhost) Then
        Set Ghost = mvarGhost
    Else
        Ghost = mvarGhost
    End If
End Property



Public Property Let NewBorn(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewBorn = 5
    mvarNewBorn = vData
End Property


Public Property Get NewBorn() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewBorn
    NewBorn = mvarNewBorn
End Property



Public Property Let gold(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gold = 5
    mvargold = vData
End Property


Public Property Get gold() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gold
    gold = mvargold
End Property



Public Property Let mapid(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mapid = 5
    mvarmapid = vData
End Property


Public Property Get mapid() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mapid
    mapid = mvarmapid
End Property



Public Property Let y(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
End Property


Public Property Get y() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    y = mvarY
End Property



Public Property Let x(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
End Property


Public Property Get x() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    x = mvarX
End Property



Public Property Set Skills(ByVal vData As Scripting.Dictionary)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Skills = Form1
    Set mvarSkills = vData
End Property


Public Property Get Skills() As Scripting.Dictionary
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Skills
    Set Skills = mvarSkills
End Property



Public Property Set army_rate(ByVal vData As Scripting.Dictionary)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.army_rate = Form1
    Set mvararmy_rate = vData
End Property


Public Property Get army_rate() As Scripting.Dictionary
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.army_rate
    Set army_rate = mvararmy_rate
End Property



Public Property Let Hor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hor = 5
    mvarHor = vData
End Property


Public Property Get Hor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hor
    Hor = mvarHor
End Property




Public Property Let MAXSP(ByVal vData As Long)
    mvarMAXSP = vData
End Property
Public Property Get MAXSP() As Long
    MAXSP = mvarMAXSP
End Property
Public Property Let MAXHP(ByVal vData As Long)
    mvarMAXHP = vData
End Property
Public Property Get MAXHP() As Long
    MAXHP = mvarMAXHP
End Property
Public Property Let SP(ByVal vData As Long)
    mvarSP = vData
End Property
Public Property Get SP() As Long
    SP = mvarSP
End Property
Public Property Let HP(ByVal vData As Long)
    mvarHP = vData
End Property
Public Property Get HP() As Long
    HP = mvarHP
End Property
Public Property Let IsSleep(ByVal vData As Boolean)
    mvarIsSleep = vData
End Property
Public Property Get IsSleep() As Boolean
    IsSleep = mvarIsSleep
End Property
Public Property Let fai(ByVal vData As Integer)
    mvarfai = vData
End Property
Public Property Get fai() As Integer
    fai = mvarfai
End Property
Public Property Let Order(ByVal vData As Integer)
    mvarOrder = vData
End Property
Public Property Get Order() As Integer
    Order = mvarOrder
End Property
Public Property Let Texp(ByVal vData As Long)
    mvartexp = vData
End Property
Public Property Get Texp() As Long
    Texp = mvartexp
End Property

Public Property Let Col(ByVal vData As Integer)
    mvarCol = vData
End Property
Public Property Get Col() As Integer
    Col = mvarCol
End Property



Public Property Let Row(ByVal vData As Integer)
    mvarRow = vData
End Property


Public Property Get Row() As Integer
    Row = mvarRow
End Property





Public Property Get CharStatus() As CharStatus
    Set CharStatus = mvarCharStatus
End Property


Public Property Set CharStatus(vData As CharStatus)
    Set mvarCharStatus = vData
End Property

Public Property Let Element(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Element = 5
    mvarElement = vData
End Property


Public Property Get Element() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Element
    Element = mvarElement
End Property



Public Property Let level(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Level = 5
    mvarlevel = vData
End Property


Public Property Get level() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Level
    level = mvarlevel
End Property



Public Property Let charname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CharName = 5
    mvarCharName = vData
End Property


Public Property Get charname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CharName
    charname = mvarCharName
End Property





'Public Property Get CharStatus() As npcStatus
'    Set CharStatus = mvarCharStatus
'End Property
'
'
'Public Property Set CharStatus(vData As npcStatus)
'    Set mvarCharStatus = vData
'End Property



'Public Property Get MySkills() As MySkills
'    If mvarMySkills Is Nothing Then
'        Set mvarMySkills = New MySkills
'    End If
'
'
'    Set MySkills = mvarMySkills
'End Property
'
'
'Public Property Set MySkills(vData As MySkills)
'    Set mvarMySkills = vData
'End Property



'Public Property Get Equipments() As Equipments
'    If mvarEquipments Is Nothing Then
'        Set mvarEquipments = New Equipments
'    End If
'
'
'    Set Equipments = mvarEquipments
'End Property
'
'
'Public Property Set Equipments(vData As Equipments)
'    Set mvarEquipments = vData
'End Property



Private Sub Class_Initialize()
  'create the mCharStatus object when the Character class is created
 ' Set mvarCharStatus = New npcStatus
  'create the mCharStatus object when the Character class is created
  Set mvarCharStatus = New CharStatus
  Set mvararmy_rate = New Scripting.Dictionary
  Set mvarSkills = New Scripting.Dictionary
End Sub


Private Sub Class_Terminate()
  Set mvarCharStatus = Nothing
  'Set mvarCharStatus = Nothing
'  Set mvarMySkills = Nothing
'  Set mvarEquipments = Nothing
End Sub



Public Property Let uID(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.uid = 5
    mvaruid = vData
End Property


Public Property Set uID(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.uid = Form1
    Set mvaruid = vData
End Property


Public Property Get uID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.uid
    If IsObject(mvaruid) Then
        Set uID = mvaruid
    Else
        uID = mvaruid
    End If
End Property

Public Property Let isLicenseId(ByVal vData As Boolean)
    mvarisLicenseId = vData
End Property
Public Property Get isLicenseId() As Boolean
    isLicenseId = mvarisLicenseId
End Property



