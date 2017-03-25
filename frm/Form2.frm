VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   255
   ClientWidth     =   9795
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "บริจาค"
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ทิ้ง"
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Equipment(test)"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ใช้ Item (ขุนพล)"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ใช้ Item (ตัวเอง)"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtNum 
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSlotNo 
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ListView ListItems1 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Num"
         Object.Width           =   706
      EndProperty
   End
   Begin VB.CommandButton cmdSendItem 
      Caption         =   "ส่งของ"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtPlayerName 
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Menu mnuCmdAutoItem 
      Caption         =   "Command"
      Visible         =   0   'False
      Begin VB.Menu mnuUseItem 
         Caption         =   "Useitem"
         Begin VB.Menu mnuUseItemPlayer 
            Caption         =   "Player"
         End
         Begin VB.Menu mnuUseItemPartner 
            Caption         =   "Partner"
         End
      End
      Begin VB.Menu mnuSendItem 
         Caption         =   "SendItem"
      End
      Begin VB.Menu mnuDrop 
         Caption         =   "Drop"
      End
      Begin VB.Menu mnuContribute 
         Caption         =   "Contribute"
      End
      Begin VB.Menu mnuNone5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveTo 
         Caption         =   "MoveTO"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdSendItem_Click()
On Error Resume Next
    uID = Form1.getPlayerId(txtPlayerName.Text)
    If uID <> 0 Then
        Form1.ts.SendItemTo uID, ListItems1.SelectedItem.Tag, ListItems1.SelectedItem.ListSubItems(2).Text
    End If
End Sub

Private Sub Form_Activate()
    Me.Caption = Form1.ts.Character.charname
End Sub

Private Sub Form_Load()
On Error Resume Next
   Form1.updateinv
   
   LastSelect = 1
End Sub


Private Sub ListItems1_DblClick()
On Error Resume Next
    Form1.ts.UseItem ListItems1.SelectedItem.Index, 1, 0
End Sub

Private Sub ListItems1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    LastSelectItem = Item.Index
    
End Sub

Private Sub ListItems1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim itm As ListItem
    Set itm = ListItems1.HitTest(x, y)
         If itm Is Nothing Then Exit Sub
    
    If Button = vbRightButton Then
            itm.Selected = True
            mnuUseItemPlayer.Caption = Form1.ts.Character.charname
            mnuUseItemPartner.Caption = Form1.ts.CurrentPartner.charname
            PopupMenu mnuCmdAutoItem
    Else
        ListItems1.ToolTipText = itm.ToolTipText
    End If
End Sub

Private Sub mnuContribute_Click()
    On Error Resume Next
    Form1.ts.Contribute 0, ListItems1.SelectedItem.Tag
End Sub

Private Sub mnuDrop_Click()
    On Error Resume Next
    Form1.ts.DropItem ListItems1.SelectedItem.Tag, ListItems1.SelectedItem.ListSubItems(2).Text
End Sub

Private Sub mnuMoveup_Click()
     
End Sub

Private Sub mnuMoveTo_Click()
On Error Resume Next
'     Form1.ts.MoveItem ListItems1.SelectedItem

End Sub

Private Sub mnuSendItem_Click()
On Error Resume Next
    
     Form1.ts.SendItemTo Form1.getPlayerId(Form1.Text3.Text), ListItems1.SelectedItem.Tag, ListItems1.SelectedItem.ListSubItems(2).Text
End Sub

Private Sub mnuUseItemPartner_Click()
On Error Resume Next
'MsgBox ListItems1.SelectedItem.Index
    Form1.ts.UseItem ListItems1.SelectedItem.Index, 1, 1
End Sub

Private Sub mnuUseItemPlayer_Click()
On Error Resume Next
    Form1.ts.UseItem ListItems1.SelectedItem.Index, 1, 0
End Sub
