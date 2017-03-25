VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrade 
      Caption         =   "แลกเปลี่ยน"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtMoney2 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   3600
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtMoney1 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
   End
   Begin MSComctlLib.ListView RecvTrade 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   10081
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
         Object.Width           =   758
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Iventory Items"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Num"
         Object.Width           =   900
      EndProperty
   End
   Begin MSComctlLib.ListView SendTrade 
      Height          =   5715
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   10081
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
         Object.Width           =   758
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Iventory Items"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Num"
         Object.Width           =   900
      EndProperty
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
