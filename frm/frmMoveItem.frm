VERSION 5.00
Begin VB.Form frmMoveItem 
   Caption         =   "ระบบเคลื่อนย้าย item"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3720
   LinkTopic       =   "Form2"
   ScaleHeight     =   1035
   ScaleWidth      =   3720
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ยกเลิก"
      Height          =   315
      Left            =   2220
      TabIndex        =   5
      Top             =   540
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
      Default         =   -1  'True
      Height          =   315
      Left            =   2220
      TabIndex        =   4
      Top             =   180
      Width           =   1275
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "50"
      Top             =   540
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "25"
      Top             =   180
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "จำนวน"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ย้ายไปยัง Slot"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmMoveItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    tempMoveSlot = Combo1.Text
    tempMoveNum = Combo2.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    tempMoveSlot = 0
    Unload Me
End Sub

Private Sub Form_Load()
Dim k As Integer
    For k = 1 To 25
        Combo1.AddItem k
    Next k
    For k = 1 To 50
        Combo2.AddItem k
    Next k
    
End Sub
