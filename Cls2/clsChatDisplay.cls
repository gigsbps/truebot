VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChatDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarbgcolor As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarobj As RichTextBox   'local copy
Public Property Set obj(ByVal vData As RichTextBox)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.obj = Form1
    Set mvarobj = vData
End Property


Public Property Get obj() As RichTextBox
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.obj
    Set obj = mvarobj
End Property



Public Property Let bgcolor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bgcolor = 5
    mvarbgcolor = vData
End Property


Public Property Get bgcolor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bgcolor
    bgcolor = mvarbgcolor
End Property



