VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "<<DreaM ProjecT>>"
   ClientHeight    =   9165
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   14760
   ForeColor       =   &H00C0FFC0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerDelay 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   7800
   End
   Begin VB.CommandButton cmdShowSena 
      Caption         =   "แสดงเสนา"
      Height          =   315
      Left            =   6300
      TabIndex        =   73
      Top             =   3780
      Width           =   915
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   480
      Top             =   7320
   End
   Begin VB.Frame mainFrame 
      Caption         =   "Frame2"
      Height          =   4095
      Index           =   4
      Left            =   9000
      TabIndex        =   58
      Top             =   6960
      Width           =   3675
      Begin VB.Image Image2 
         Height          =   2910
         Left            =   180
         Picture         =   "Form1.frx":078A
         Top             =   300
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageCombo cmbMix1 
      Height          =   330
      Left            =   5640
      TabIndex        =   68
      Top             =   3060
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "0"
   End
   Begin VB.CommandButton cmdMixIt 
      Caption         =   "หลอมรวม"
      Height          =   315
      Left            =   6300
      TabIndex        =   67
      Top             =   3420
      Width           =   915
   End
   Begin VB.Frame mainFrame 
      Caption         =   "Frame2"
      Height          =   555
      Index           =   3
      Left            =   5880
      TabIndex        =   57
      Top             =   6720
      Width           =   615
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   0
         TabIndex        =   64
         Top             =   3240
         Width           =   3735
         Begin VB.Label Label4 
            Caption         =   "จำนวนผู้เล่นที่ Online อยู่ในขณะนี้ ="
            Height          =   255
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   2595
         End
         Begin VB.Label Label5 
            Caption         =   "0"
            Height          =   255
            Left            =   2580
            TabIndex        =   65
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   3840
         TabIndex        =   63
         Top             =   3180
         Width           =   1575
      End
      Begin MSComctlLib.ListView listOnline 
         Height          =   315
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชื่อ"
            Object.Width           =   4357
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลเวล"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ธาตุ"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "วีรบุรุษ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ประลอง"
            Object.Width           =   1288
         EndProperty
      End
   End
   Begin VB.Frame mainFrame 
      Caption         =   "Frame2"
      Height          =   555
      Index           =   2
      Left            =   5160
      TabIndex        =   56
      Top             =   6720
      Width           =   615
      Begin MSComctlLib.ListView listArmy 
         Height          =   315
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   617
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อ"
            Object.Width           =   5062
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "เลเวล"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ธาตุ"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "บริจาค"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame mainFrame 
      Caption         =   "Frame2"
      Height          =   555
      Index           =   1
      Left            =   4440
      TabIndex        =   55
      Top             =   6720
      Width           =   615
      Begin MSComctlLib.ListView listFriend 
         Height          =   315
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame mainFrame 
      Caption         =   "Frame2"
      Height          =   555
      Index           =   0
      Left            =   3720
      TabIndex        =   53
      Top             =   6720
      Width           =   615
      Begin RichTextLib.RichTextBox txtDisplay 
         Height          =   315
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":9F5E
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   315
         Left            =   360
         TabIndex        =   59
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393217
         BackColor       =   15790320
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":9FF0
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   7500
      TabIndex        =   39
      Top             =   6780
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageCombo text2 
      Height          =   330
      Left            =   1140
      TabIndex        =   38
      Top             =   5880
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo Text3 
      Height          =   330
      Left            =   5640
      TabIndex        =   37
      Top             =   5880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Setting"
      Height          =   315
      Left            =   5640
      TabIndex        =   35
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackPack 
      Caption         =   "Open BackPack"
      Height          =   315
      Left            =   5640
      TabIndex        =   34
      Top             =   2340
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Request Party"
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   5460
      Width           =   1575
   End
   Begin VB.CommandButton cmdHorse 
      Caption         =   "Riding horse"
      Height          =   315
      Left            =   5640
      TabIndex        =   25
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtpversion 
      Height          =   285
      Left            =   3060
      TabIndex        =   32
      Text            =   "185"
      Top             =   6840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00F96844&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -120
      ScaleHeight     =   225
      ScaleWidth      =   10005
      TabIndex        =   28
      Top             =   6240
      Width           =   10035
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ( Philipine Edition )"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   0
         Width           =   1380
      End
      Begin VB.Label txtCurrentLoc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mapid = mapid (x,y)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5925
         TabIndex        =   29
         Top             =   0
         Width           =   1350
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sit [1]"
      Height          =   315
      Left            =   5640
      TabIndex        =   31
      Top             =   4860
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartTimer 
      Caption         =   "Start Timer"
      Height          =   315
      Left            =   5640
      TabIndex        =   26
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inventory >>"
      Height          =   315
      Left            =   5640
      TabIndex        =   27
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Timer ReConnectTimer 
      Left            =   2520
      Top             =   6720
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   6720
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   6720
   End
   Begin MSComctlLib.ImageCombo icbPartnerList 
      Height          =   330
      Left            =   1620
      TabIndex        =   22
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin VB.TextBox txtPasswd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "•"
      TabIndex        =   3
      ToolTipText     =   "PASSWORD"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtAccount 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "TS ID"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtServerIP 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "64.27.19.234"
      ToolTipText     =   "IP SERVER"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer ScriptTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   6720
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   6720
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   483
      TabIndex        =   5
      Top             =   560
      Width           =   7275
      Begin VB.PictureBox expscale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   5520
         Picture         =   "Form1.frx":A082
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   103
         TabIndex        =   24
         Top             =   60
         Width           =   1575
         Begin VB.Image imgexpscale 
            Height          =   120
            Index           =   1
            Left            =   0
            Picture         =   "Form1.frx":A44F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox expscale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   1800
         Picture         =   "Form1.frx":A488
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   103
         TabIndex        =   23
         Top             =   60
         Width           =   1575
         Begin VB.Image imgexpscale 
            Height          =   120
            Index           =   0
            Left            =   0
            Picture         =   "Form1.frx":A855
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox pscale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":A88E
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         Begin VB.Image imgscale 
            Height          =   120
            Index           =   0
            Left            =   0
            Picture         =   "Form1.frx":AC5B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox pscale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   0
         Picture         =   "Form1.frx":ACA8
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   8
         Top             =   600
         Width           =   1695
         Begin VB.Image imgscale 
            Height          =   120
            Index           =   2
            Left            =   0
            Picture         =   "Form1.frx":B075
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox pscale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   3
         Left            =   3720
         Picture         =   "Form1.frx":B0C2
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   7
         Top             =   600
         Width           =   1695
         Begin VB.Image imgscale 
            Height          =   120
            Index           =   3
            Left            =   0
            Picture         =   "Form1.frx":B48F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox pscale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   3720
         Picture         =   "Form1.frx":B4DC
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         Begin VB.Image imgscale 
            Height          =   120
            Index           =   1
            Left            =   0
            Picture         =   "Form1.frx":B8A9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label txtExp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   195
         Index           =   2
         Left            =   2220
         TabIndex        =   48
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label txtExp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   195
         Index           =   5
         Left            =   5940
         TabIndex        =   51
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label txtExp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   195
         Index           =   3
         Left            =   5880
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label txtExp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label txtExp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.Label txtExp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Index           =   4
         Left            =   6300
         TabIndex        =   50
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Next : "
         Height          =   195
         Index           =   5
         Left            =   5520
         TabIndex        =   45
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exp : "
         Height          =   195
         Index           =   3
         Left            =   5520
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp/Min :"
         Height          =   195
         Index           =   1
         Left            =   5520
         TabIndex        =   43
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Next : "
         Height          =   195
         Index           =   4
         Left            =   1800
         TabIndex        =   42
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp/Min :"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   870
      End
      Begin VB.Line Line1 
         X1              =   232
         X2              =   232
         Y1              =   0
         Y2              =   64
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exp : "
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Partner Name"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hp/maxhp"
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   13
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hp/maxhp"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "sp/maxsp"
         Height          =   255
         Index           =   2
         Left            =   540
         TabIndex        =   11
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "sp/maxsp"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   70
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListItems1 
      Height          =   5715
      Left            =   7320
      TabIndex        =   33
      Top             =   540
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
   Begin MSComctlLib.ListView ListItems2 
      Height          =   5715
      Left            =   9780
      TabIndex        =   36
      Top             =   540
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
         Text            =   "BackPack Items"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Num"
         Object.Width           =   900
      EndProperty
   End
   Begin MSComctlLib.ImageCombo imgChatType 
      Height          =   330
      Left            =   60
      TabIndex        =   41
      Top             =   5880
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "กระซิบ"
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3915
      Left            =   0
      TabIndex        =   40
      Top             =   1920
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6906
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "System"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Chat Box"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Army"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Online Player"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show System"
      Height          =   315
      Left            =   5640
      TabIndex        =   52
      Top             =   4260
      Width           =   1575
   End
   Begin MSComctlLib.ImageCombo cmbMix2 
      Height          =   330
      Left            =   5640
      TabIndex        =   69
      Top             =   3420
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "0"
   End
   Begin MSComctlLib.ImageCombo cmbMix3 
      Height          =   330
      Left            =   5640
      TabIndex        =   70
      Top             =   3780
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "0"
   End
   Begin VB.Label lblFreeDebug 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4740
      TabIndex        =   72
      Top             =   1620
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Free Debug ::"
      Height          =   255
      Left            =   3660
      TabIndex        =   71
      Top             =   1620
      Width           =   1275
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   6480
      TabIndex        =   21
      Top             =   180
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Online times"
      Height          =   195
      Left            =   5520
      TabIndex        =   20
      Top             =   180
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   -600
      Picture         =   "Form1.frx":B8F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13860
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gold"
      Height          =   195
      Left            =   75
      TabIndex        =   18
      Top             =   1635
      Width           =   330
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadScript 
         Caption         =   "ReLoad Script"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuLoadWayPoint 
         Caption         =   "Load waypoint"
         Enabled         =   0   'False
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfirmExit 
         Caption         =   "Confirm Exit"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuEnableReconnect 
         Caption         =   "Enable Auto Reconnect"
      End
      Begin VB.Menu mnuAutoEat 
         Caption         =   "Enable HP&SP Auto Eating"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Enable Systray when minimize"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAvoid9am 
         Caption         =   "Auto avoid server maintenance (9.00)"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGMOnline 
         Caption         =   "Auto disconnect when GM online"
      End
      Begin VB.Menu mnuGMInmap 
         Caption         =   "Auto disconnect when GM in map"
      End
   End
   Begin VB.Menu mnuCommand 
      Caption         =   "Command"
      Begin VB.Menu mnuOpenInventory 
         Caption         =   "Inventories"
      End
      Begin VB.Menu mnunone4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTestClickNPC 
         Caption         =   "Test Click NPC"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Call Start()"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Call Stop()"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMakeShop 
         Caption         =   "Open Shop"
      End
      Begin VB.Menu mnuCloseShop 
         Caption         =   "Close Shop"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckTime 
         Caption         =   "Check Air Time"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContent 
         Caption         =   "Contents..."
         Enabled         =   0   'False
      End
      Begin VB.Menu none2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu RCPopup 
      Caption         =   "RCPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
   End
   Begin VB.Menu mnuCmdAutoItem 
      Caption         =   "ItemCommand"
      Visible         =   0   'False
      Begin VB.Menu mnuUseItem 
         Caption         =   "Use Item"
         Begin VB.Menu mnuUseItemPlayer 
            Caption         =   "Player"
         End
         Begin VB.Menu mnuUseItemPartner 
            Caption         =   "Partner"
         End
      End
      Begin VB.Menu mnuSendItem 
         Caption         =   "Send Item"
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
         Caption         =   "Move To Slot"
      End
      Begin VB.Menu mnuSendToBackp 
         Caption         =   "Send To BackPack"
      End
   End
   Begin VB.Menu mnuCmdAutoBackPack 
      Caption         =   "BackPackCommand"
      Visible         =   0   'False
      Begin VB.Menu mnuSendToInven 
         Caption         =   "Send To Inventory"
      End
   End
   Begin VB.Menu mnuUtils 
      Caption         =   "Utils"
      Begin VB.Menu mnuMaskConvert 
         Caption         =   "MaskConvert"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ts As tspacket
Attribute ts.VB_VarHelpID = -1
Public WithEvents VBscript As MSScriptControl.ScriptControl
Attribute VBscript.VB_VarHelpID = -1
Public fso As Scripting.FileSystemObject
Public f As Scripting.file
Public tso As Scripting.TextStream
Public sv As clsServer

Dim ScriptFileName As String

Public ChatDisplay As New clsChatDisplay
Public ChatLength As Integer
Public DataLength As Integer

Public colBack As Long
Public colAll As Long
Public colPublic As Long
Public colWhisper As Long
Public colTeam As Long
Public colGuild As Long
Public colGFriend As Long
Public colDrop As Long
Public colFight As Long

Public chkAll As Boolean
Public chkPublic As Boolean
Public chkWhisper As Boolean
Public chkTeam As Boolean
Public chkGuild As Boolean
Public chkGFriend As Boolean
Public chkDrop As Boolean
Public chkFight As Boolean

Public FightingFlag As Boolean
Public LastChatType As Integer
Public AvoidFlag As Boolean

Dim skill As clsSkill
Dim initscr As Boolean

Dim charname
Dim setanswer As Integer
Dim fuckgod As Boolean
 

Dim StartExp(2) As Long
Dim StartTime(2) As Date
Dim LastExp(2) As Long
Dim LastTime(2) As Date

Public GameFolder As String

Public UserCommand As Scripting.Dictionary


Public RestartTimer As Boolean


Dim sittype

Private Sub cmdOption_Click()
    frmOption.Show
End Sub

Public Function cdelay2(sec)

Dim pauseTime
Dim start
Dim finish
Dim totaltime

    pauseTime = sec
    start = Timer
    Do While Timer < start + pauseTime
        'cdelay2 (1)
        DoEvents
    Loop
    finish = Timer
    
End Function
Public Function cdelay(sec)
    MsgWaitObj sec * 1000
'Dim pauseTime
'Dim start
'Dim finish
'Dim totaltime
'
'    pauseTime = sec
'    start = Timer
'    Do While Timer < start + pauseTime
'        DoEvents
'    Loop
'    finish = Timer
    
End Function


Public Function writeINIdata(inisec, inikey, inival)
    Dim lbAppName As String ' To carry the name of the section - [Set]
    Dim lpFileName As String ' Carries INI file name
    Dim sHsKey As String ' Carries Key name - for UserName
    Dim sHsValue As String 'Carries Key value - for UserName
    Dim file As String
        file = App.Path & "\profile.ini" 'name of file
        lpFileName = file
        lpAppName = inisec 'Section name
        sHsKey = inikey 'Key name
        sHsValue = inival 'Key value
        writeINIdata = WritePrivateProfileString(lpAppName, sHsKey, sHsValue, lpFileName)
End Function

Public Sub SaveConfig()
    StatusUser = writeINIdata("Server", "ServerIP", txtServerIP.Text)
    StatusUser = writeINIdata("Server", "ID", txtAccount.Text)
    StatusUser = writeINIdata("Server", "PASSWORD", txtPasswd.Text)
    StatusUser = writeINIdata("Server", "PVERSION", txtpversion.Text)
    StatusUser = writeINIdata("Server", "GameFolder", GameFolder)
    
    StatusUser = writeINIdata("Color", "Background", colBack)
    StatusUser = writeINIdata("Color", "ChatAll", colAll)
    StatusUser = writeINIdata("Color", "ChatPubLic", colPublic)
    StatusUser = writeINIdata("Color", "ChatWhisper", colWhisper)
    StatusUser = writeINIdata("Color", "ChatParty", colTeam)
    StatusUser = writeINIdata("Color", "ChatGuild", colGuild)
    StatusUser = writeINIdata("Color", "ChatGFriend", colGFriend)
    StatusUser = writeINIdata("Color", "DataDroping", colDrop)
    StatusUser = writeINIdata("Color", "DataFighting", colFight)
    
    StatusUser = writeINIdata("RefreshLength", "DataBox", DataLength)
    StatusUser = writeINIdata("RefreshLength", "ChatBox", ChatLength)
    
    StatusUser = writeINIdata("OnOff", "ChatAll", chkAll)
    StatusUser = writeINIdata("OnOff", "ChatPubLic", chkPublic)
    StatusUser = writeINIdata("OnOff", "ChatWhisper", chkWhisper)
    StatusUser = writeINIdata("OnOff", "ChatParty", chkTeam)
    StatusUser = writeINIdata("OnOff", "ChatGuild", chkGuild)
    StatusUser = writeINIdata("OnOff", "ChatGFriend", chkGFriend)
    StatusUser = writeINIdata("OnOff", "DataDroping", chkDrop)
    StatusUser = writeINIdata("OnOff", "DataFighting", chkFight)
    
    StatusUser = writeINIdata("Menu", "AutoRecon", mnuEnableReconnect.Checked)
    StatusUser = writeINIdata("Menu", "AutoEat", mnuAutoEat.Checked)
    StatusUser = writeINIdata("Menu", "AutoSystray", mnuSystray.Checked)
    StatusUser = writeINIdata("Menu", "AutoAvoid9am", mnuAvoid9am.Checked)
    StatusUser = writeINIdata("Menu", "AutoGMOnline", mnuGMOnline.Checked)
    StatusUser = writeINIdata("Menu", "AutoGMInmap", mnuGMInmap.Checked)
End Sub

Private Sub cmdBackPack_Click()
    If cmdBackPack.Caption = "Open BackPack" Then
        cmdBackPack.Caption = "Open Invent"
        ListItems1.Visible = False
        ListItems2.Visible = True
    Else
        cmdBackPack.Caption = "Open BackPack"
        ListItems2.Visible = False
        ListItems1.Visible = True
    End If
End Sub

Private Sub cmdHorse_Click()
    If ts.IsHorse = False Then
        ts.Horse
        cmdHorse.Caption = "Take off horse"
    Else
        ts.UnHorse
        cmdHorse.Caption = "Riding horse"
    End If
End Sub

Private Sub cmdLogin_Click()
    DoLogin
End Sub

Sub DoLogin()
On Error Resume Next
    If cmdLogin.Caption = "Login" Then
        Set VBscript = New ScriptControl
        Set ts = New tspacket
           initscr = False
           ts.pversion = txtpversion.Text
           ts.Disconect
           ts.ConnectServer txtServerIP.Text, 6414
  '         ts.ConnectServer "127.0.0.1", 1000
           SaveConfig
    ElseIf cmdLogin.Caption = "Logout" Then
        Set VBscript = Nothing
            initscr = True
            ts.Disconect
       ' Set ts = Nothing
    End If
End Sub

Private Sub cmdMixIt_Click()
Dim YesNoReturn As Integer
Dim strTemp1 As String
Dim strTemp2 As String
Dim strTemp3 As String
Dim slotTemp1 As Integer
Dim slotTemp2 As Integer
Dim slotTemp3 As Integer

strTemp1 = ""
strTemp2 = ""
strTemp3 = ""
slotTemp1 = 0
slotTemp2 = 0
slotTemp3 = 0

If CInt(cmbMix1.Text) > 0 Then
    slotTemp1 = ts.MyItems(CInt(cmbMix1.Text)).num
    If ts.MyItems(CInt(cmbMix1.Text)).itemid > 0 Then
        strTemp1 = getItemName(ts.MyItems(CInt(cmbMix1.Text)).itemid) & _
        " #" & ts.MyItems(CInt(cmbMix1.Text)).num & " ชิ้น"
    Else
        strTemp1 = "(ไม่มี item ในช่องนี้)"
    End If
Else
    strTemp1 = "(ไม่มี item ในช่องนี้)"
End If
If CInt(cmbMix2.Text) > 0 Then
    slotTemp2 = ts.MyItems(CInt(cmbMix2.Text)).num
    If ts.MyItems(CInt(cmbMix2.Text)).itemid > 0 Then
        strTemp2 = getItemName(ts.MyItems(CInt(cmbMix2.Text)).itemid) & _
        " #" & ts.MyItems(CInt(cmbMix2.Text)).num & " ชิ้น"
    Else
        strTemp2 = "(ไม่มี item ในช่องนี้)"
    End If
Else
    strTemp2 = "(ไม่มี item ในช่องนี้)"
End If
If CInt(cmbMix3.Text) > 0 Then
    slotTemp3 = ts.MyItems(CInt(cmbMix3.Text)).num
    If ts.MyItems(CInt(cmbMix3.Text)).itemid > 0 Then
        strTemp3 = getItemName(ts.MyItems(CInt(cmbMix3.Text)).itemid) & _
        " #" & ts.MyItems(CInt(cmbMix3.Text)).num & " ชิ้น"
    Else
        strTemp3 = "(ไม่มี item ในช่องนี้)"
    End If
Else
    strTemp3 = "(ไม่มี item ในช่องนี้)"
End If

    YesNoReturn = MsgBox("คุณต้องการหลอม items ต่อไปนี้ใช่หรือไม่" & vbCrLf & _
    "ช่องที่ " & cmbMix1.Text & " " & strTemp1 & vbCrLf & _
    "ช่องที่ " & cmbMix2.Text & " " & strTemp2 & vbCrLf & _
    "ช่องที่ " & cmbMix3.Text & " " & strTemp3 _
    , vbQuestion + vbYesNo, "หลอมรวม")
    
    If YesNoReturn = vbYes Then
        If CInt(cmbMix1.Text) > 0 Then
            If CInt(cmbMix2.Text) > 0 Then
                If CInt(cmbMix3.Text) > 0 Then
                    If Not ((CInt(cmbMix1.Text) = CInt(cmbMix2.Text)) Or (CInt(cmbMix1.Text) = CInt(cmbMix3.Text)) Or (CInt(cmbMix2.Text) = CInt(cmbMix3.Text))) Then
                        ts.MixItem CInt(cmbMix1.Text), slotTemp1, CInt(cmbMix2.Text), slotTemp2, CInt(cmbMix3.Text), slotTemp3
                    End If
                Else
                    If Not (CInt(cmbMix1.Text) = CInt(cmbMix2.Text)) Then
                        ts.MixItem CInt(cmbMix1.Text), slotTemp1, CInt(cmbMix2.Text), slotTemp2, CInt(cmbMix3.Text), slotTemp3
                    End If
                End If
            Else
                If Not (CInt(cmbMix1.Text) = CInt(cmbMix3.Text)) Then
                    ts.MixItem CInt(cmbMix1.Text), slotTemp1, CInt(cmbMix2.Text), slotTemp2, CInt(cmbMix3.Text), slotTemp3
                End If
            End If
        Else
            If CInt(cmbMix2.Text) > 0 Then
                If CInt(cmbMix3.Text) > 0 Then
                    ts.MixItem CInt(cmbMix1.Text), slotTemp1, CInt(cmbMix2.Text), slotTemp2, CInt(cmbMix3.Text), slotTemp3
                End If
            End If
        End If
        
    End If
End Sub

Private Sub cmdRefresh_Click()
    CreatePlayerList
End Sub

Private Sub cmdShowSena_Click()
    If cmdShowSena.Caption = "แสดงเสนา" Then
        cmdShowSena.Caption = "ซ่อนเสนา"
    Else
        cmdShowSena.Caption = "แสดงเสนา"
    End If
End Sub

Private Sub cmdStartTimer_Click()
    If ScriptTimer.Enabled = False Then
        cmdStartTimer.Caption = "Stop Timer"
        ScriptTimer.Enabled = True
    Else
        ScriptTimer.Enabled = False
        cmdStartTimer.Caption = "Start Timer"
    End If
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Inventory >>" Then
        Command1.Caption = "Inventory <<"
        Me.Width = 9810
    Else
        Command1.Caption = "Inventory >>"
        Me.Width = 7410
    End If
End Sub

Private Sub Command3_Click()
    frmTrade.Show
End Sub

Private Sub Command4_Click()
    If TabStrip1.SelectedItem.Index = 2 Then
        If Command4.Caption = "Show System" Then
            Command4.Caption = "Hide System"
            txtDisplay.Visible = True
            txtChat.Visible = True
            
            txtDisplay.Height = 1730
            txtDisplay.Width = 5415
            txtDisplay.Left = 0
            txtDisplay.Top = 0
            
            txtChat.Height = 1730
            txtChat.Width = 5415
            txtChat.Left = 0
            txtChat.Top = 1760
        Else
            Command4.Caption = "Show System"
            txtDisplay.Visible = False
            txtChat.Visible = True
            InitSizeAllBox
        End If
    End If
End Sub

Private Sub Command5_Click()
    frmChatSetting.Show
End Sub

Private Sub Command6_Click()
On Error Resume Next
    
    
    If Command6.Caption = "Sit [1]" Then
        ts.SendAction (46)
        Command6.Caption = "Sit [2]"
    ElseIf Command6.Caption = "Sit [2]" Then
        ts.SendAction (47)
        Command6.Caption = "Sit [3]"
    ElseIf Command6.Caption = "Sit [3]" Then
        ts.SendAction (48)
        Command6.Caption = "Sit [4]"
    ElseIf Command6.Caption = "Sit [4]" Then
        ts.SendAction (49)
        Command6.Caption = "Sit [1]"
    End If
    
    
 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Form2.Move Form1.Left + Form1.Width, Form1.Top, Form2.Width, Form2.Height

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Form2.Move Form1.Left + Form1.Width, Form1.Top, Form2.Width, Form2.Height
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call mnuExit_Click
' Cancel = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = 1 And mnuSystray.Checked = True Then
        Hook Me.hwnd   ' Set up our handler
        AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, Label1(0).Caption
        Me.Hide
    End If
End Sub
Public Sub SysTrayMouseEventHandler()
    SetForegroundWindow Me.hwnd
    PopupMenu RCPopup, vbPopupMenuRightButton
End Sub

Private Sub icbPartnerList_click()
On Error Resume Next
    If icbPartnerList.SelectedItem.Tag = 0 Then
        Call ts.UnSelectPartner
    End If
    

    If icbPartnerList.SelectedItem.Tag <> ts.CurrentPartner.uID Then
        Call ts.SelectPartner(icbPartnerList.SelectedItem.Tag)
    End If
End Sub

Private Sub imgChatType_Click()
    text2.SetFocus
End Sub


Private Sub listArmy_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
Dim i As Integer

    If ColumnHeader = "เลเวล" Then
        For i = 1 To listArmy.ListItems.Count
            While Len(listArmy.ListItems(i).ListSubItems.Item(2).Text) < 3
                listArmy.ListItems(i).ListSubItems.Item(2).Text = 0 & listArmy.ListItems(i).ListSubItems.Item(2).Text
            Wend
        Next i
    ElseIf ColumnHeader = "บริจาค" Then
        For i = 1 To listArmy.ListItems.Count
            While Len(listArmy.ListItems(i).ListSubItems.Item(4).Text) < 10
                listArmy.ListItems(i).ListSubItems.Item(4).Text = 0 & listArmy.ListItems(i).ListSubItems.Item(4).Text
            Wend
        Next i
    End If

    If listArmy.SortOrder = lvwAscending Then
        listArmy.SortOrder = lvwDescending
    Else
        listArmy.SortOrder = lvwAscending
    End If
    listArmy.Sorted = True
    listArmy.SortKey = ColumnHeader.Index - 1
    
    If ColumnHeader = "เลเวล" Then
        For i = 1 To listArmy.ListItems.Count
            While Left(listArmy.ListItems(i).ListSubItems.Item(2).Text, 1) = "0"
                listArmy.ListItems(i).ListSubItems.Item(2).Text = Mid(listArmy.ListItems(i).ListSubItems.Item(2).Text, 2, Len(listArmy.ListItems(i).ListSubItems.Item(1).Text) - 1)
            Wend
        Next i
    ElseIf ColumnHeader = "บริจาค" Then
        For i = 1 To listArmy.ListItems.Count
            While (Left(listArmy.ListItems(i).ListSubItems.Item(4).Text, 1) = "0") And (Len(listArmy.ListItems(i).ListSubItems.Item(4).Text) > 1)
                listArmy.ListItems(i).ListSubItems.Item(4).Text = Mid(listArmy.ListItems(i).ListSubItems.Item(4).Text, 2, Len(listArmy.ListItems(i).ListSubItems.Item(4).Text) - 1)
            Wend
        Next i
    End If
End Sub

Private Sub listArmy_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Form1.Text3.Text = Item.SubItems(1)
End Sub

Private Sub listOnline_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim i As Integer
    If ColumnHeader = "เลเวล" Then
        For i = 1 To listOnline.ListItems.Count
            While Len(listOnline.ListItems(i).ListSubItems.Item(1).Text) < 3
                listOnline.ListItems(i).ListSubItems.Item(1).Text = 0 & listOnline.ListItems(i).ListSubItems.Item(1).Text
            Wend
        Next i
    End If
    
    If listOnline.SortOrder = lvwAscending Then
        listOnline.SortOrder = lvwDescending
    Else
        listOnline.SortOrder = lvwAscending
    End If
    listOnline.Sorted = True
    listOnline.SortKey = ColumnHeader.Index - 1
    
    If ColumnHeader = "เลเวล" Then
        For i = 1 To listOnline.ListItems.Count
            While Left(listOnline.ListItems(i).ListSubItems.Item(1).Text, 1) = "0"
                listOnline.ListItems(i).ListSubItems.Item(1).Text = Mid(listOnline.ListItems(i).ListSubItems.Item(1).Text, 2, Len(listOnline.ListItems(i).ListSubItems.Item(1).Text) - 1)
            Wend
        Next i
    End If
End Sub

Private Sub listOnline_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Form1.Text3.Text = Item.Text
End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAutoEat_Click()
    mnuAutoEat.Checked = IIf(mnuAutoEat.Checked = True, False, True)
End Sub

Private Sub mnuAvoid9am_Click()
    mnuAvoid9am.Checked = IIf(mnuAvoid9am.Checked = True, False, True)
End Sub

Private Sub mnuCheckTime_Click()
    ts.CheckAirTime
End Sub

Private Sub mnuCloseShop_Click()
  On Error Resume Next
  ts.CloseShop
  VBscript.ExecuteStatement "onCloseShop()"
End Sub

Private Sub mnuConfirmExit_Click()
    mnuConfirmExit.Checked = IIf(mnuConfirmExit.Checked = True, False, True)
End Sub

Private Sub mnuEnableReconnect_Click()
    mnuEnableReconnect.Checked = IIf(mnuEnableReconnect.Checked = True, False, True)
End Sub

Private Sub mnuExit_Click()
    Dim ret As Integer
    If mnuConfirmExit.Checked = True Then
        ret = MsgBox("Do you want to Exit ?", vbCritical + vbOKCancel)
        If ret = vbOK Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub


Private Sub mnuGMInmap_Click()
    mnuGMInmap.Checked = IIf(mnuGMInmap.Checked = True, False, True)
End Sub

Private Sub mnuGMOnline_Click()
    mnuGMOnline.Checked = IIf(mnuGMOnline.Checked = True, False, True)
End Sub

Private Sub mnuLoadScript_Click()
On Error Resume Next
        initscript
        InitScript1
End Sub

Private Sub mnuMakeShop_Click()
  On Error Resume Next
  VBscript.ExecuteStatement "MakeShop()"
End Sub

Private Sub mnuMaskConvert_Click()
    Dialog.Show
    
End Sub

Private Sub mnuOpenInventory_Click()
On Error Resume Next
    Command1_Click
End Sub

Private Sub mnuRestore_Click()
    Unhook    ' Return event control to windows
    Me.WindowState = 0
    Me.Show
    RemoveIconFromTray
    
End Sub

Private Sub mnuStart_Click()
On Error Resume Next
    VBscript.ExecuteStatement "Start()"
End Sub

Private Sub mnuStop_Click()
On Error Resume Next
    VBscript.ExecuteStatement "Stop()"
End Sub

Private Sub mnuSystray_Click()
    mnuSystray.Checked = IIf(mnuSystray.Checked = True, False, True)
End Sub

Private Sub mnuTestClickNPC_Click()
On Error Resume Next
    Err.Clear
    npcid = InputBox("Insert NPCID For Simulate single click")
    If Err.number = 0 Then
        ts.ClickOnNPC CInt(npcid)
    End If
End Sub

Private Sub ReConnectTimer_Timer()
    AvoidFlag = False
    Call cmdLogin_Click
End Sub

Public Sub CreateTop10List()
On Error Resume Next
Dim oTs As Dictionary
Dim uID
Dim itm As ListSubItem
Dim strTemp As String
    
    Set oTs = Form1.ts.ol
    Label5.Caption = oTs.Count
    listOnline.ListItems.Clear
    For Each uID In oTs.Keys
        With listOnline.ListItems.Add
            .Tag = uID
            .Text = oTs.Item(uID).charname
            Set itm = .ListSubItems.Add(, , oTs.Item(uID).level)
            Select Case oTs.Item(uID).Element
                Case 1
                    strTemp = "ดิน"
                Case 2
                    strTemp = "น้ำ"
                Case 3
                    strTemp = "ไฟ"
                Case 4
                    strTemp = "ลม"
            End Select
            Set itm = .ListSubItems.Add(, , strTemp)
            Set itm = .ListSubItems.Add(, , 0)
                
            If oTs.Item(uID).NewBorn = 1 Then
                .Bold = True
            Else
                
            End If
        End With
    Next
End Sub

Public Sub CreatePlayerList()
On Error Resume Next
Dim oTs As Dictionary
Dim uID
Dim itm As ListSubItem
Dim strTemp As String
    
    Set oTs = Form1.ts.ol
    Label5.Caption = oTs.Count
    listOnline.ListItems.Clear
    For Each uID In oTs.Keys
        With listOnline.ListItems.Add
            .Tag = uID
            .Text = oTs.Item(uID).charname
            Set itm = .ListSubItems.Add(, , oTs.Item(uID).level)
            Select Case oTs.Item(uID).Element
                Case 1
                    strTemp = "ดิน"
                Case 2
                    strTemp = "น้ำ"
                Case 3
                    strTemp = "ไฟ"
                Case 4
                    strTemp = "ลม"
            End Select
            Set itm = .ListSubItems.Add(, , strTemp)
            Select Case oTs.Item(uID).NewBorn
                Case 3
                    strTemp = "จอมยุทธ"
                Case 4
                    strTemp = "จอมทัพ"
                Case 5
                    strTemp = "กุนซือ"
                Case 6
                    strTemp = "เซียน"
                Case Else
                    strTemp = ""
            End Select
            
            Set itm = .ListSubItems.Add(, , strTemp)
                
            If oTs.Item(uID).NewBorn = 1 Then
                .Bold = True
            ElseIf oTs.Item(uID).NewBorn > 1 Then
                .Bold = True
                .ForeColor = vbRed
            End If
            
            Set itm = .ListSubItems.Add(, , oTs.Item(uID).PvPranking)
        
        End With
    Next
End Sub

Private Sub TabStrip1_Click()
    'InitSizeAllBox
    Select Case TabStrip1.SelectedItem.Index
        Case 1
            mainFrame(0).Visible = True
            mainFrame(1).Visible = False
            mainFrame(2).Visible = False
            mainFrame(3).Visible = False
            mainFrame(4).Visible = False
            txtDisplay.Visible = True
            txtChat.Visible = False
            InitSizeAllBox
        Case 2
            mainFrame(0).Visible = True
            mainFrame(1).Visible = False
            mainFrame(2).Visible = False
            mainFrame(3).Visible = False
            mainFrame(4).Visible = False
            txtDisplay.Visible = False
            txtChat.Visible = True
            InitSizeAllBox
'        Case 3
'            mainFrame(0).Visible = False
'            mainFrame(1).Visible = False
'            mainFrame(2).Visible = False
'            mainFrame(3).Visible = False
'            mainFrame(4).Visible = True
        Case 3
            mainFrame(0).Visible = False
            mainFrame(1).Visible = False
            mainFrame(2).Visible = True
            mainFrame(3).Visible = False
            mainFrame(4).Visible = False
        Case 4
            mainFrame(0).Visible = False
            mainFrame(1).Visible = False
            mainFrame(2).Visible = False
            mainFrame(3).Visible = True
            mainFrame(4).Visible = False
            CreatePlayerList
    End Select
End Sub

Private Sub Text3_Click()
    text2.SetFocus
End Sub

Private Sub Timer2_Timer()

'    If FindWindow(vbNullString, "WPE PRO - " & App.EXEName & ".EXE") > 0 Then
'            Unload Me
'            End
'    ElseIf FindWindow(vbNullString, "WPE PRO - " & App.EXEName & ".EXE" & " - [WPEPRO1]") > 0 Then
'            Unload Me
'            End
'    ElseIf FindWindow(vbNullString, "WPE PRO - " & App.EXEName & ".EXE" & " - [WPEPRO2]") > 0 Then
'            Unload Me
'            End
'    ElseIf FindWindow(vbNullString, "WPE PRO - " & App.EXEName & ".EXE" & " - [WPEPRO3]") > 0 Then
'            Unload Me
'            End
'    ElseIf FindWindow(vbNullString, "WPE PRO - " & App.EXEName & ".EXE" & " - [WPEPRO4]") > 0 Then
'            Unload Me
'            End
'    ElseIf FindWindow(vbNullString, "WPE PRO - " & App.EXEName & ".EXE" & " - [WPEPRO5]") > 0 Then
'            Unload Me
'            End
'    End If
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
   sec = DateDiff("s", StartTime(0), Now)
    NewDate = DateAdd("s", sec, "00:00:00")


    Label11.Caption = Format(NewDate, "h:m:s")
End Sub



Private Sub ts_AppearCurrentOnlinePlayers(ByVal objPlayerCharacter As Character)
On Error Resume Next
'   With dlgPlayerOnline.ListView1.ListItems.Add
'            .Text = oTs.ol.Item(uid).CharName
'            .ListSubItems.Add , , oTs.ol.Item(uid).level
'            If oTs.ol.Item(uid).NewBorn = True Then
'                .Bold = True
'            End If
'            .Selected = True
'        End With
    'dlgPlayerOnline.ListView1.Refresh
    If VBscript.Language = "Javascript" Then
        VBscript.ExecuteStatement "PlayerCurrentOnline(" & objPlayerCharacter.uID & ")"
    End If
    
  
     
End Sub

Private Sub ts_AppearOnlinePlayers(ByVal objPlayerCharacter As Character)
On Error Resume Next
    VBscript.ExecuteStatement "PlayerOnline(" & objPlayerCharacter.uID & ")"
End Sub

Private Sub ts_Closed()
On Error Resume Next
    
    'VBscript.ExecuteStatement "Closed()"
    'Set VBscript = Nothing
    cmdLogin.Caption = "Login"
    AppendDisplay "Connection Closed.", vbRed
    
    If mnuEnableReconnect.Checked = True Or (mnuAvoid9am.Checked = True And AvoidFlag) Then
        ReConnectTimer.Interval = 5000
        If mnuAvoid9am.Checked = True Then ReConnectTimer.Interval = 30000
        ReConnectTimer.Enabled = True
    End If
    Timer4.Enabled = False
    ScriptTimer.Enabled = False
    
    
End Sub


Public Sub AppendDisplay(Msg, cColor)
On Error Resume Next

    linea = txtDisplay.GetLineFromChar(Len(txtDisplay.Text))
    If linea > DataLength Then
        txtDisplay.Text = ""
    End If
    
    txtDisplay.SelStart = Len(txtDisplay.Text)
    txtDisplay.SelText = Msg & vbNewLine
    
    txtDisplay.SelStart = Len(txtDisplay.Text) - Len(Msg) - 2
    txtDisplay.SelLength = Len(Msg)
    txtDisplay.SelColor = cColor
    txtDisplay.SelFontName = "MS Sans Serif"
End Sub

Sub AppendChat(Msg, Optional ByVal cColor As VBRUN.ColorConstants)
On Error Resume Next
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelText = Msg & vbNewLine
    
    txtChat.SelStart = Len(txtChat.Text) - Len(Msg) - 2
    txtChat.SelLength = Len(Msg)
    txtChat.SelColor = cColor
    txtChat.SelFontName = "MS Sans Serif"
End Sub
Sub SetPlayerMeter(Index, obj As Character)
On Error Resume Next
    'If obj.MAXHP > 0 Then
    percentofhp = ((obj.HP * 100) / obj.MAXHP)
    imgscale(Index).ToolTipText = obj.HP & "/" & obj.MAXHP
    imgscale(Index).Width = percentofhp * pscale(Index).Width / 100
    Label9(Index).Caption = obj.HP & "/" & obj.MAXHP
    
    percentofsp = ((obj.SP * 100) / obj.MAXSP)
    imgscale(Index + 2).ToolTipText = obj.SP & "/" & obj.MAXSP
    imgscale(Index + 2).Width = percentofsp * pscale(Index + 2).Width / 100
    Label9(Index + 2).Caption = obj.SP & "/" & obj.MAXSP
    
    Label1(0).Caption = ts.Character.charname & "(" & ts.Character.level & ")"
    Form1.Caption = "TrueBot - [" & ts.Character.charname & "]" & " Lv." & ts.Character.level
    Label1(1).Caption = ts.CurrentPartner.charname & "(" & ts.CurrentPartner.level & ")"
    'End If
End Sub


Public Sub alert(Msg)
On Error Resume Next
    MsgBox Msg
End Sub

Sub AppendFreeDebug(ByVal strTemp As String)
On Error Resume Next
    lblFreeDebug.Caption = strTemp
End Sub

Function GetTextFreeDebug() As String
On Error Resume Next
    GetTextFreeDebug = lblFreeDebug.Caption
End Function

Sub initscript()
On Error Resume Next
    
    Set VBscript = New ScriptControl
        VBscript.Language = "Javascript"
        VBscript.AllowUI = True
        VBscript.Timeout = 30000
        VBscript.AddObject "Timer", ScriptTimer
        VBscript.AddObject "frm", Form1
        VBscript.AddObject "MenuReConnect", Form1.mnuEnableReconnect
        VBscript.AddObject "Server", sv
        VBscript.AddObject "Display", txtDisplay
        VBscript.AddObject "Chat", ChatDisplay
        VBscript.AddObject "SKILL", skill
        VBscript.AddObject "NPC", dnpcs
        VBscript.AddObject "MAP", dmaps
        VBscript.AddObject "ITEMS", ditems
End Sub

Public Function AddText3(strTemp As String)
Dim i As Integer
Dim notFound As Boolean
    notFound = True
    For i = 1 To Text3.ComboItems.Count
        If Text3.ComboItems(i).Text = strTemp Then notFound = False
    Next i
    If notFound Then Text3.ComboItems.Add.Text = strTemp
End Function

Public Function AddText2(strTemp As String)
Dim i As Integer
Dim notFound As Boolean
    notFound = True
    For i = 1 To text2.ComboItems.Count
        If text2.ComboItems(i).Text = strTemp Then notFound = False
    Next i
    If notFound Then text2.ComboItems.Add.Text = strTemp
End Function

Private Sub Command2_Click()
On Error Resume Next
    ts.RequestParty getPlayerId(Text3.Text)
    AddText3 (Text3.Text)
    
End Sub
Public Function getPlayerName(playerid)
On Error Resume Next
    If ts.ol.Exists(playerid) Then
        getPlayerName = ts.ol.Item(playerid).charname
    End If
End Function

Public Function getItemName2(itemid)
On Error Resume Next
        getItemName2 = getItemName(itemid)
End Function

Public Function getPlayerId(playerName) As Long
On Error Resume Next
    For Each uID In ts.ol.Keys
        If ts.ol.Item(uID).charname = playerName Then
            getPlayerId = ts.ol.Item(uID).uID
            Exit Function
        End If
    Next
    getPlayerId = -1
End Function

Private Function LoadINIdata(inisec, inikey) As String
Dim GetSetting As Long 'Get user on form load
Dim temp1 As String * 50 ' stores retreived value
Dim sHsUser As String
    file = App.Path & "\profile.ini" ' file name
    lpAppName = inisec 'Section name
    sHsUser = inikey 'Key name
    lpDefault = Empty ' Default for any of the declared Keys
    lpFileName = file
    GetSetting = GetPrivateProfileString(lpAppName, sHsUser, lpDefault, temp1, Len(temp1), lpFileName)
    LoadINIdata = temp1
End Function


Private Sub InitSizeAllBox()
Dim i As Integer
    For i = 0 To 4
        With mainFrame(i)
            .Height = 3495
            .Width = 5415
            .Left = 60
            .Top = 1980
            .Visible = False
            .BorderStyle = 0
        End With
    Next i
    mainFrame(0).Visible = True
    
    txtDisplay.Height = 3495
    txtDisplay.Width = 5415
    txtDisplay.Left = 0
    txtDisplay.Top = 0
    
    txtChat.Height = 3495
    txtChat.Width = 5415
    txtChat.Left = 0
    txtChat.Top = 0
    
    listFriend.Height = 3495
    listFriend.Width = 5415
    listFriend.Left = 0
    listFriend.Top = 0
    
    listArmy.Height = 3495
    listArmy.Width = 5415
    listArmy.Left = 0
    listArmy.Top = 0
    
    listOnline.Height = 3135
    listOnline.Width = 5415
    listOnline.Left = 0
    listOnline.Top = 0
    
End Sub

Private Sub Form_Load()
Dim tempk As Integer
On Error Resume Next

    Timer2.Enabled = True
    
    initLicenseList
    LastExp(0) = 0
    LastExp(1) = 0
    
    Me.Height = 7245
    Me.Width = 7410
    ListItems2.Left = ListItems1.Left
    ListItems2.Top = ListItems1.Top
    ListItems2.Visible = False
    'Assign Default Value
    colBack = vbWhite
    colAll = &H969696
    colPublic = vbBlack
    colWhisper = vbRed
    colTeam = &H18080
    colGuild = &HCC0101
    colGFriend = &H770101
    colDrop = vbBlue
    colFight = vbRed
    ChatLength = 300
    DataLength = 500
    chkBack = True
    chkAll = True
    chkPublic = True
    chkWhisper = True
    chkTeam = True
    chkGuild = True
    chkGFriend = True
    chkDrop = True
    chkFight = True
    
    Set sv = New clsServer
    Set fso = New Scripting.FileSystemObject
    Set ts = New tspacket
    Set skill = New clsSkill
        initscript
        initscr = False
        fuckgod = False
        
        StartExp(0) = 0
        StartExp(1) = 0
        
    imgChatType.ComboItems.Add.Text = "กระซิบ"
    imgChatType.ComboItems.Add.Text = "กลุ่ม"
    imgChatType.ComboItems.Add.Text = "กองทัพ"
    imgChatType.ComboItems.Add.Text = "พันธมิตร"
    imgChatType.ComboItems.Add.Text = "ส่วนตัว"
        
    For tempk = 0 To 25
        cmbMix1.ComboItems.Add.Text = tempk
        cmbMix2.ComboItems.Add.Text = tempk
        cmbMix3.ComboItems.Add.Text = tempk
    Next tempk
    
    Set ChatDisplay = New clsChatDisplay
    Set ChatDisplay.obj = txtChat
    
    InitSizeAllBox
    
    txtServerIP.Text = LoadINIdata("Server", "ServerIP")
    txtAccount.Text = LoadINIdata("Server", "ID")
    txtPasswd.Text = LoadINIdata("Server", "PASSWORD")
    txtpversion.Text = LoadINIdata("Server", "PVERSION")
    GameFolder = LoadINIdata("Server", "GameFolder")
    GameFolder = Trim(GameFolder)
    While Asc(Right(GameFolder, 1)) = 0
        GameFolder = Mid(GameFolder, 1, Len(GameFolder) - 1)
    Wend
    
    If Right(GameFolder, 1) = "\" Then
        GameFolder = Mid(GameFolder, 1, Len(GameFolder) - 1)
    End If
    
    Call LoadItemData
    Call LoadNPCData
    Call LoadSkillData
    
    colBack = LoadINIdata("Color", "Background")
    colAll = LoadINIdata("Color", "ChatAll")
    colPublic = LoadINIdata("Color", "ChatPubLic")
    colWhisper = LoadINIdata("Color", "ChatWhisper")
    colTeam = LoadINIdata("Color", "ChatParty")
    colGuild = LoadINIdata("Color", "ChatGuild")
    colGFriend = LoadINIdata("Color", "ChatGFriend")
    colDrop = LoadINIdata("Color", "DataDroping")
    colFight = LoadINIdata("Color", "DataFighting")

    ChatLength = LoadINIdata("RefreshLength", "ChatBox")
    DataLength = LoadINIdata("RefreshLength", "DataBox")
    
    chkBack = LoadINIdata("OnOff", "Background")
    chkAll = LoadINIdata("OnOff", "ChatAll")
    chkPublic = LoadINIdata("OnOff", "ChatPubLic")
    chkWhisper = LoadINIdata("OnOff", "ChatWhisper")
    chkTeam = LoadINIdata("OnOff", "ChatParty")
    chkGuild = LoadINIdata("OnOff", "ChatGuild")
    chkGFriend = LoadINIdata("OnOff", "ChatGFriend")
    chkDrop = LoadINIdata("OnOff", "DataDroping")
    chkFight = LoadINIdata("OnOff", "DataFighting")
    
    mnuEnableReconnect.Checked = LoadINIdata("Menu", "AutoRecon")
    mnuAutoEat.Checked = LoadINIdata("Menu", "AutoEat")
    mnuSystray.Checked = LoadINIdata("Menu", "AutoSystray")
    mnuAvoid9am.Checked = LoadINIdata("Menu", "AutoAvoid9am")
    mnuGMOnline.Checked = LoadINIdata("Menu", "AutoGMOnline")
    mnuGMInmap.Checked = LoadINIdata("Menu", "AutoGMInmap")
    
    FightingFlag = False
    AvoidFlag = False
    
    txtChat.BackColor = colBack

'    Label12.Caption = GetVersion() & Label12.Caption
    Label12.Caption = "Dream Project" & Label12.Caption
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'    End
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        AddText2 (text2.Text)
        KeyAscii = 0
        flag = Mid(text2.Text, 1, 6)
        If LCase(flag) = "bot://" Then
            statement = Trim(Mid(text2.Text, 7))
            VBscript.ExecuteStatement statement
             text2.Text = ""
            
            Exit Sub
        End If
        
        ctype = 2
        Select Case imgChatType.Text
            Case "กระซิบ"
                ctype = 2
            Case "กลุ่ม"
                ctype = 5
            Case "กองทัพ"
                ctype = 6
            Case "พันธมิตร"
                ctype = 7
            Case "ส่วนตัว"
                ctype = 3
        End Select
        
        If text2.Text <> "" Then
            If ctype = 2 Then
                ts.Chat ctype, text2.Text
                AppendChat TimeText & " กระซิบ [" & ts.Character.charname & "] " & ":" & text2.Text, colPublic
            ElseIf ctype = 3 Then
                ts.Chat ctype, text2.Text, getPlayerId(Text3.Text)
                AddText3 (Text3.Text)
            ElseIf ctype = 5 Then
                ts.Chat ctype, text2.Text
            ElseIf ctype = 6 Then
                ts.Chat ctype, text2.Text
            ElseIf ctype = 7 Then
                ts.Chat ctype, text2.Text
            End If
            text2.Text = ""
        End If
    
    End If
    
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    'HpRecover
End Sub

Private Sub ScriptTimer_Timer()
On Error Resume Next
    cmdStartTimer.Caption = Now

    VBscript.ExecuteStatement "OnTimer()"
End Sub

Private Sub ts_AppearAnotherCombat(ByVal playerid As Long)
On Error Resume Next
    If VBscript.Language = "Javascript" Then
        VBscript.ExecuteStatement "FoundCombat(" & playerid & ")"
    End If
    
    
End Sub


Sub InitScript1()
         
        
        VBscript.AddObject "ts", ts
        VBscript.AddCode "function alert(msg){ frm.alert(msg) }" & vbNewLine
        VBscript.AddCode "function debug(msg,color){ frm.AppendDisplay(msg,color) }" & vbNewLine
        VBscript.AddCode "function playerGetID(pname){ return frm.getPlayerId(pname) }" & vbNewLine
        VBscript.AddCode "function getPlayerId(pname){ return frm.getPlayerId(pname) }" & vbNewLine
        VBscript.AddCode "function getPlayerName(uid){ return frm.getPlayerName(uid) }" & vbNewLine
        VBscript.AddCode "function getItemName(itemid){ return frm.getItemName2(itemid) }" & vbNewLine
        VBscript.AddCode "function include(fname){ return frm.Include(fname)}" & vbNewLine
        VBscript.AddCode "function cdelay (sec){ return frm.cdelay (sec) }" & vbNewLine
        VBscript.AddCode "function FreeDebug(strTemp){ return frm.AppendFreeDebug(strTemp) }" & vbNewLine
        VBscript.AddCode "function GetFreeDebug(){ return frm.GetTextFreeDebug() }" & vbNewLine
        
        'Include ("common.js")
        'Include ("QA.js")
        'Include ("Wrong.js")
        Include ("main.js")
        
'        fname = fso.BuildPath(App.Path, "\common.js")
'        If fso.FileExists(fname) = True Then
'            Set f = fso.GetFile(fname)
'            If f.Size > 0 Then
'                Set tso = f.OpenAsTextStream(ForReading, TristateUseDefault)
'                VBscript.AddCode tso.ReadAll
'            End If
'        End If
'
'        fname = fso.BuildPath(App.Path, "\QA.js")
'        If fso.FileExists(fname) = True Then
'            Set f = fso.GetFile(fname)
'            If f.Size > 0 Then
'                Set tso = f.OpenAsTextStream(ForReading, TristateUseDefault)
'                VBscript.AddCode tso.ReadAll
'            End If
'        End If
'       '  VBscript.AddCode txtScript.Text
'        fname = fso.BuildPath(App.Path, "\script.js")
'        If fso.FileExists(fname) = True Then
'            Set f = fso.GetFile(fname)
'            If f.Size > 0 Then
'                Set tso = f.OpenAsTextStream(ForReading, TristateUseDefault)
'                VBscript.AddCode tso.ReadAll
'            End If
'        End If
'        For i = 1 To VBscript.Procedures.Count
'           AppendDisplay "Load " & VBscript.Procedures(i).Name, vbBlue
'        Next
        AppendDisplay "Loaded " & VBscript.Procedures.Count & " functions", vbBlue

End Sub

Public Sub Include(filename)
    fname = fso.BuildPath(App.Path, filename)
        If fso.FileExists(fname) = True Then
            Set f = fso.GetFile(fname)
            'MsgBox fname
            If f.Size > 0 Then
                Set tso = f.OpenAsTextStream(ForReading, TristateUseDefault)
                VBscript.AddCode tso.ReadAll
            End If
        End If
End Sub

Public Function TimeText() As String
Dim xHour As String
Dim xMin As String
Dim xSec As String

    xHour = Hour(Time)
    xMin = Minute(Time)
    xSec = Second(Time)
    If Len(xHour) < 2 Then
        xHour = "0" & xHour
    End If
    If Len(xMin) < 2 Then
        xMin = "0" & xMin
    End If
    If Len(xSec) < 2 Then
        xSec = "0" & xSec
    End If
    TimeText = "[" & xHour & ":" & xMin & ":" & xSec & "]"
    
End Function


Private Sub ts_ChatMessage(typeid As Variant, Msg As Variant, sender As Variant)
On Error Resume Next
LastChatType = typeid

Dim typetext As String
    If typeid = 11 And initscr = False Then
        initscr = True

        initscript
        InitScript1
        'HpRecover
    End If
    
    linea = txtChat.GetLineFromChar(Len(txtChat.Text))
    If linea > ChatLength Then
        txtChat.Text = ""
    End If
    
    Select Case typeid
        Case 1
                If chkAll Then
                    typetext = TimeText & " ทั้งหมด "
                    AppendChat typetext & "[" & getPlayerName(sender) & "] :" & Msg, colAll
                    VBscript.ExecuteStatement "OnPublicMsg('" & getPlayerName(sender) & "','" & Msg & "')"
                End If
                Exit Sub
        Case 2
                If chkPublic Then
                    typetext = TimeText & " กระซิบ "
                    AppendChat typetext & "[" & getPlayerName(sender) & "] :" & Msg, colPublic
                    VBscript.ExecuteStatement "OnWhisperMsg('" & getPlayerName(sender) & "','" & Msg & "')"
                End If
            Exit Sub
        Case 3
                If chkWhisper Then
                    typetext = TimeText & " ส่วนตัว "
                    If sender = ts.Character.uID Then
                        AppendChat typetext & "[" & getPlayerName(sender) & "] - [" & Text3.Text & "] " & ":" & Msg, colWhisper
                        AddText3 (Text3.Text)
                    Else
                        AppendChat typetext & "[" & getPlayerName(sender) & "] - [" & ts.Character.charname & "] " & ":" & Msg, colWhisper
                        VBscript.ExecuteStatement "OnPrivateMsg('" & getPlayerName(sender) & "','" & Msg & "')"
                        AddText3 (getPlayerName(sender))
                    End If
                End If
            Exit Sub
        Case 4
            typetext = TimeText & " เทพสวรรค์ "
            AppendChat typetext & "[" & getPlayerName(sender) & "] :" & Msg, vbCyan
            VBscript.ExecuteStatement "OnGodMsg('" & getPlayerName(sender) & "','" & Msg & "')"
            Exit Sub
        Case 5
                If chkTeam Then
                    typetext = TimeText & " กลุ่ม "
                    AppendChat typetext & "[" & getPlayerName(sender) & "] :" & Msg, colTeam
                    VBscript.ExecuteStatement "OnTeamMsg('" & getPlayerName(sender) & "','" & Msg & "')"
                End If
            Exit Sub
        Case 6
                If chkGuild Then
                    typetext = TimeText & " กองทัพ "
                    AppendChat typetext & "[" & getPlayerName(sender) & "] :" & Msg, colGuild
                    VBscript.ExecuteStatement "OnGuildMsg('" & getPlayerName(sender) & "','" & Msg & "')"
                End If
            Exit Sub
        Case 7
                If chkGFriend Then
                    typetext = TimeText & " พันธมิตร "
                    AppendChat typetext & "[" & getPlayerName(sender) & "] :" & Msg, colGFriend
                    VBscript.ExecuteStatement "OnAllyMsg('" & getPlayerName(sender) & "','" & Msg & "')"
                End If
            Exit Sub
        Case 11
            If Right(Msg, 4) = "#end" Then
                Msg = Mid(Msg, 1, Len(Msg) - 4)
            End If
            AppendChat "ระบบ : " & Msg, vbRed
            Exit Sub
        Case 0
            If Right(Msg, 4) = "#end" Then
                Msg = Mid(Msg, 1, Len(Msg) - 4)
            End If
            AppendChat "ระบบ : " & Msg, vbRed
            Exit Sub
        Case 8
            If LastChatType = 0 Then
                AppendChat "Server Maintenance DETECTED !!", vbRed
                VBscript.ExecuteStatement "onMaintenance()"
                If mnuAvoid9am.Checked = False Then
                    AppendChat "Avoid 9am. option = Off , do nothing.", vbRed
                Else
                    AppendChat "Avoid 9am. option = On , Disconnect now.", vbRed
                    AvoidFlag = True
                    ts.Disconect
                End If
            End If
            Exit Sub
        Case Else
            AppendChat "Type : " & typeid & " Msg = " & Msg, &H4080&
            Exit Sub
    End Select
End Sub
Public Function GetVersion()
On Error Resume Next
   GetVersion = "TrueBot " & App.Major & "." & App.Minor & "." & App.Revision
End Function


Private Sub ts_Connected()
On Error Resume Next
    AppendDisplay "Connected for " & txtServerIP.Text & ":" & "6414", vbBlack

    ts.RequestLogin
End Sub

Private Sub ts_Connecting()
On Error Resume Next
    AppendDisplay "Connecting Server.......", vbBlack
End Sub



Private Sub ts_doAcceptParty(ByVal partyid As Long)
On Error Resume Next
    AppendDisplay "[Party System] :: ผู้เล่น " & getPlayerName(partyid) & " ได้เข้าร่วม Party แล้ว", vbRed
End Sub


Private Sub ts_doNotEnoughSlot(ByVal itemid As Long, ByVal n As Integer)
On Error Resume Next
    AppendDisplay "ช่องเต็มคุณไม่ได้ " & getItemName(itemid) & " จำนวน " & n & " อัน", vbRed
End Sub

Private Sub ts_doNotEnoughBackPackSlot(ByVal itemid As Long, ByVal n As Integer)
On Error Resume Next
    AppendDisplay "ไม่สามารถเก็บ " & getItemName(itemid) & " จำนวน " & n & " อัน เข้าเป้หลังได้", vbRed
End Sub

Private Sub ts_DoSelectPartner(ByVal partnerid As Long)
On Error Resume Next
    Label1(1).Caption = ts.CurrentPartner.charname
    StartExp(1) = ts.CurrentPartner.Texp
    StartTime(1) = Now

   For i = 1 To icbPartnerList.ComboItems.Count
        If icbPartnerList.ComboItems(i).Tag = ts.CurrentPartner.uID Then
            icbPartnerList.ComboItems(i).Selected = True
            DoEvents
            Exit For
        End If
   Next
End Sub

Private Sub ts_DuplicateLogin()
On Error Resume Next
    AppendDisplay "Login ซ้ำซ้อน กรุณาลองเชื่อมต่อใหม่", vbRed
    'ts.Disconect
End Sub

Private Sub ts_FinishAnswerFuckGod()
On Error Resume Next
    VBscript.ExecuteStatement "FinishAnswerFuckGod()"
End Sub

Private Sub ts_FinishBattle(ByVal uID As Long)
On Error Resume Next
    If VBscript.Language = "Javascript" Then
        VBscript.ExecuteStatement "FinishBattle(" & uID & ")"
    End If
End Sub

Private Sub ts_InitInventoryList()
On Error Resume Next
updateinv
End Sub

Private Sub ts_InitBackPackList()
On Error Resume Next
updatebackp
End Sub

Private Sub ts_InitPlayerStatus()
On Error Resume Next
    StartExp(0) = ts.Character.Texp
    StartTime(0) = Now
    
    DoEvents
End Sub

Private Sub ts_InvalidLicence()
On Error Resume Next
    AppendDisplay "Invalid Licence.", vbBlack
End Sub

Private Sub ts_InventoryChange()
On Error Resume Next
    updateinv
End Sub

Private Sub ts_BackPackChange()
On Error Resume Next
    updatebackp
End Sub


Public Sub updateinv()
On Error Resume Next
    Dim oitem As Inv
    Form1.ListItems1.ListItems.Clear
    For i = 1 To 25
        Set oitem = ts.MyItems(i)
        With Form1.ListItems1.ListItems.Add
            .Tag = i
            .Text = i
            .ToolTipText = ditems(oitem.itemid).itemname & _
                " " & ditems(oitem.itemid).itemtype & _
                " " & ditems(oitem.itemid).itemvalue & _
                " " & ditems(oitem.itemid).itemtype2 & _
                " " & ditems(oitem.itemid).itemvalue2 & _
                " (" & ditems(oitem.itemid).itemdesc & ")"
            If oitem.itemid = 99999 Then
                AppendDisplay "DANGER !! Unknown-item found, please update your ITEM.DAT", vbRed
                'AppendDisplay "อันตราย !! พบ item ที่ไม่รู้จัก, กรุณาเปลี่ยนไฟล์ ITEM.DAT ใหม่", vbRed
                ts.Disconect
            End If
            With .ListSubItems.Add
                .Tag = oitem.itemid
                .Text = getItemName(oitem.itemid)
            End With
            With .ListSubItems.Add
                .Text = oitem.num
                
            End With
            If oitem.num = 50 Then
                .ForeColor = vbRed
                .ListSubItems.Item(1).ForeColor = vbRed
                .ListSubItems.Item(2).ForeColor = vbRed
            Else
                .ForeColor = vbBlack
                .ListSubItems.Item(1).ForeColor = vbBlack
                .ListSubItems.Item(2).ForeColor = vbBlack
            End If
        End With
    Next
    Form1.ListItems1.ListItems(LastSelectItem).Selected = True
    
End Sub

Public Sub updatebackp()
On Error Resume Next
    Dim ooitem As Inv
    Form1.ListItems2.ListItems.Clear
    For i = 1 To 25
        'AppendChat ditems(ts.BackPack(i).itemid).itemname
        Set ooitem = ts.BackPack(i)
        With Form1.ListItems2.ListItems.Add
            .Tag = i
            .Text = i
            .ToolTipText = ditems(ooitem.itemid).itemname & _
                " " & ditems(ooitem.itemid).itemtype & _
                " " & ditems(ooitem.itemid).itemvalue & _
                " " & ditems(ooitem.itemid).itemtype2 & _
                " " & ditems(ooitem.itemid).itemvalue2 & _
                " (" & ditems(ooitem.itemid).itemdesc & ")"
            If ooitem.itemid = 99999 Then
                AppendDisplay "DANGER !! Unknown-item found, please update your ITEM.DAT", vbRed
                'AppendDisplay "อันตราย !! พบ item ที่ไม่รู้จัก, กรุณาเปลี่ยนไฟล์ ITEM.DAT ใหม่", vbRed
            End If

            With .ListSubItems.Add
                .Tag = ooitem.itemid
                .Text = getItemName(ooitem.itemid)
            End With
            With .ListSubItems.Add
                .Text = ooitem.num
                
            End With
            If ooitem.num = 50 Then
                .ForeColor = vbRed
                .ListSubItems.Item(1).ForeColor = vbRed
                .ListSubItems.Item(2).ForeColor = vbRed
            Else
                .ForeColor = vbBlack
                .ListSubItems.Item(1).ForeColor = vbBlack
                .ListSubItems.Item(2).ForeColor = vbBlack
            End If
        End With
    Next
    Form1.ListItems2.ListItems(LastSelectBPItem).Selected = True
    
End Sub

Private Sub ts_LoginFail()
    AppendDisplay "Unable to login, please check your server ip or account,password again.", vbRed
    ts.Disconect
     cmdLogin.Caption = "Login"
     If mnuEnableReconnect.Checked = True Then
         ReConnectTimer.Enabled = True
     End If
     
End Sub

Private Sub ts_Loginok()
On Error Resume Next
    cmdLogin.Caption = "Logout"
    
    With icbPartnerList.ComboItems.Add
        .Tag = 0
        .Text = "sleep"
    End With
    
    VBscript.ExecuteStatement "Logon()"
    
    
    DisplayLocation
    
    Timer4.Enabled = True
    ReConnectTimer.Enabled = False
    'HpRecover
    
End Sub

Private Sub ts_MyAttack()
On Error Resume Next
    VBscript.ExecuteStatement "MyAttack()"
End Sub

Private Sub ts_NpcDialog(ByVal DialogId As Long)
On Error Resume Next
    VBscript.ExecuteStatement "NpcDialog(" & DialogId & ")"
End Sub

Private Sub ts_NpcDialogMenu(ByVal DialogId As Long)
On Error Resume Next
    VBscript.ExecuteStatement "NpcDialogMenu(" & DialogId & ")"
End Sub

Private Sub ts_NpcWalkThenDialog(ByVal DialogId As Long)
On Error Resume Next
    VBscript.ExecuteStatement "NpcWalkThenDialog(" & DialogId & ")"
End Sub

Private Sub ts_CombatSceneDialog(ByVal DialogId As Long)
On Error Resume Next
    VBscript.ExecuteStatement "CombatSceneDialog(" & DialogId & ")"
End Sub

Private Sub ts_odEattem(ByVal slot As Integer, ByVal n As Integer, ByVal pos As Integer)
On Error Resume Next
    If pos = 0 Then
        VBscript.ExecuteStatement "onMyEatItem(" & slot & "," & n & ")"
    Else
        VBscript.ExecuteStatement "onPartnerEatItem(" & slot & "," & n & "," & pos & ")"
    End If
End Sub

Private Sub ts_on140B()
On Error Resume Next
    VBscript.ExecuteStatement "NpcHiddenDialog()"
End Sub

Private Sub ts_onAnswerRight(ByVal Question As String, ByVal answer As String)
On Error Resume Next
    VBscript.ExecuteStatement "onAnswerRight('" & Question & "','" & answer & "')"
End Sub


Private Sub ts_onAnswerWrong(ByVal Question As String, ByVal answer As String)
On Error Resume Next
    VBscript.ExecuteStatement "onAnswerWrong('" & Question & "','" & answer & "')"

End Sub

Private Sub ts_onBattleStarted()
On Error Resume Next
    FightingFlag = True
    
    If ScriptTimer.Enabled = True Then
        ScriptTimer.Enabled = False
        RestartTimer = True
        
    End If
    
    
    
    AppendDisplay "Start battle", vbBlack
    VBscript.ExecuteStatement "BattleStarted()"
End Sub

Private Sub ts_onBattleStoped()
On Error Resume Next
    FightingFlag = False
    AppendDisplay "Battle Stopped", vbBlack
    UpdateExpPerMin
    VBscript.ExecuteStatement "BattleStoped()"
    
    If RestartTimer = True Then
        ScriptTimer.Enabled = True
    End If
    
End Sub

Private Sub UpdateExpPerMin()
Dim onlinesec As Long
Dim CharExpPer As Double, PartExpPer As Double
Dim CurExp As Double, CurPartExp As Double
Dim ExpPerMin1 As Double, ExpPerMin2 As Double
Dim ExpLeft1 As Double, ExpLeft2 As Double
Dim LeftDay As Integer, LeftHour As Integer, LeftMin As Integer, LeftSec As Integer
Dim strLeftTime As String

    onlinesec = DateDiff("s", StartTime(0), Now)
    If Not FightingFlag Then
    
    If ts.Character.NewBorn = 0 Then
        CurExp = Getexp(ts.Character.level, ts.Character.Texp)
        CharExpPer = ((CurExp * 100) / dicExp1.Item(ts.Character.level).maxexp)
        ExpLeft1 = dicExp1.Item(ts.Character.level).maxexp - CurExp
    ElseIf ts.Character.NewBorn = 1 Then
        CurExp = Getexp2(ts.Character.level, ts.Character.Texp)
        CharExpPer = ((CurExp * 100) / dicExp2.Item(ts.Character.level).maxexp)
        ExpLeft1 = dicExp2.Item(ts.Character.level).maxexp - CurExp
    Else
        CurExp = Getexp3(ts.Character.level, ts.Character.Texp)
        CharExpPer = ((CurExp * 100) / dicExp3.Item(ts.Character.level).maxexp)
        ExpLeft1 = dicExp3.Item(ts.Character.level).maxexp - CurExp
    
    End If
    
    If ((ts.CurrentPartner.uID >= 45000) And (ts.CurrentPartner.uID < 46000)) Then
        CurPartExp = Getexp2(ts.CurrentPartner.level, ts.CurrentPartner.Texp)
        PartExpPer = ((CurPartExp * 100) / dicExp2.Item(ts.CurrentPartner.level).maxexp)
        ExpLeft2 = dicExp2.Item(ts.CurrentPartner.level).maxexp - CurPartExp
    Else
        CurPartExp = Getexp(ts.CurrentPartner.level, ts.CurrentPartner.Texp)
        PartExpPer = ((CurPartExp * 100) / dicExp1.Item(ts.CurrentPartner.level).maxexp)
        ExpLeft2 = dicExp1.Item(ts.CurrentPartner.level).maxexp - CurPartExp
    End If

        txtExp(0).Caption = Format(ts.Character.Texp - StartExp(0), "##,###0") & "(" & Format(CharExpPer, "###0.00") & "%)"
        'Label5.Caption = Format((ts.Character.Texp - StartExp(0)) / onlineminute, "######0.00")
        ExpPerMin1 = ((ts.Character.Texp - StartExp(0)) / onlinesec) * 60
        txtExp(1).Caption = Format(ExpPerMin1, "###,###0.00")
        
        LeftDay = 0
        LeftHour = 0
        LeftMin = 0
        LeftSec = 0
        
        If ExpPerMin1 <= 0 Then
            txtExp(2).Caption = " Infinity"
        Else
            LeftMin = Int(ExpLeft1 / ExpPerMin1)
            LeftSec = Int(((ExpLeft1 / ExpPerMin1) - LeftMin) * 60)
            If LeftMin >= 1440 Then
               LeftDay = Int(LeftMin / 1440)
               LeftMin = LeftMin - (LeftDay * 1440)
            End If
            If LeftMin >= 60 Then
               LeftHour = Int(LeftMin / 60)
               LeftMin = LeftMin - (LeftHour * 60)
            End If
            
            If LeftDay > 0 Then
                strLeftTime = LeftDay & "d:"
            Else
                strLeftTime = ""
            End If
            strLeftTime = strLeftTime & LeftHour & "h:"
            strLeftTime = strLeftTime & LeftMin & "m:"
            strLeftTime = strLeftTime & LeftSec & "s"
            
            txtExp(2).Caption = strLeftTime
        End If
        
        txtExp(3).Caption = Format(ts.CurrentPartner.Texp - StartExp(1), "##,###0") & "(" & Format(PartExpPer, "###0.00") & "%)"
        'Label7.Caption = Format((ts.CurrentPartner.Texp - StartExp(1)) / onlineminute, "######0.00")
        ExpPerMin2 = ((ts.CurrentPartner.Texp - StartExp(1)) / onlinesec) * 60
        txtExp(4).Caption = Format(ExpPerMin2, "###,###0.00")
        
        LeftDay = 0
        LeftHour = 0
        LeftMin = 0
        LeftSec = 0
        
        If ExpPerMin2 <= 0 Then
            txtExp(5).Caption = " Infinity"
        Else
            LeftMin = Int(ExpLeft2 / ExpPerMin2)
            LeftSec = Int(((ExpLeft2 / ExpPerMin2) - LeftMin) * 60)
            If LeftMin >= 1440 Then
               LeftDay = Int(LeftMin / 1440)
               LeftMin = LeftMin - (LeftDay * 1440)
            End If
            If LeftMin >= 60 Then
               LeftHour = Int(LeftMin / 60)
               LeftMin = LeftMin - (LeftHour * 60)
            End If
            
            If LeftDay > 0 Then
                strLeftTime = LeftDay & "d:"
            Else
                strLeftTime = ""
            End If
            strLeftTime = strLeftTime & LeftHour & "h:"
            strLeftTime = strLeftTime & LeftMin & "m:"
            strLeftTime = strLeftTime & LeftSec & "s"

            txtExp(5).Caption = strLeftTime
        End If
    End If
End Sub

Private Sub ts_onChangeStatus()
On Error Resume Next
Dim onlineminute As Long
Dim onlinesec As Long
    HpRecover
     
    ' ts.Character.NewBorn = 0
    If ts.Character.NewBorn = 0 Then
        CurExp = Getexp(ts.Character.level, ts.Character.Texp)
        imgexpscale(0).ToolTipText = CurExp & "/" & dicExp1.Item(ts.Character.level).maxexp
        expscale(0).ToolTipText = dicExp1.Item(ts.Character.level).maxexp - CurExp
        percentofexp = ((CurExp * 100) / dicExp1.Item(ts.Character.level).maxexp)
        imgexpscale(0).Width = percentofexp * expscale(0).Width / 100
    ElseIf ts.Character.NewBorn = 1 Then
        CurExp = Getexp2(ts.Character.level, ts.Character.Texp)
        imgexpscale(0).ToolTipText = CurExp & "/" & dicExp2.Item(ts.Character.level).maxexp
        expscale(0).ToolTipText = dicExp2.Item(ts.Character.level).maxexp - CurExp
        percentofexp = ((CurExp * 100) / dicExp2.Item(ts.Character.level).maxexp)
        imgexpscale(0).Width = percentofexp * expscale(0).Width / 100
    Else
        CurExp = Getexp3(ts.Character.level, ts.Character.Texp)
        imgexpscale(0).ToolTipText = CurExp & "/" & dicExp3.Item(ts.Character.level).maxexp
        expscale(0).ToolTipText = dicExp3.Item(ts.Character.level).maxexp - CurExp
        percentofexp = ((CurExp * 100) / dicExp3.Item(ts.Character.level).maxexp)
        imgexpscale(0).Width = percentofexp * expscale(0).Width / 100
    End If
    
    
    If ((ts.CurrentPartner.uID >= 45000) And (ts.CurrentPartner.uID < 46000)) Then
        CurExp = Getexp2(ts.CurrentPartner.level, ts.CurrentPartner.Texp)
        imgexpscale(1).ToolTipText = CurExp & "/" & dicExp2.Item(ts.CurrentPartner.level).maxexp
        expscale(1).ToolTipText = dicExp2.Item(ts.CurrentPartner.level).maxexp - CurExp
        percentofexp = ((CurExp * 100) / dicExp2.Item(ts.CurrentPartner.level).maxexp)
        imgexpscale(1).Width = percentofexp * expscale(1).Width / 100
    ElseIf ts.CurrentPartner.uID >= 0 Then
        CurExp = Getexp(ts.CurrentPartner.level, ts.CurrentPartner.Texp)
        imgexpscale(1).ToolTipText = CurExp & "/" & dicExp1.Item(ts.CurrentPartner.level).maxexp
        expscale(1).ToolTipText = dicExp1.Item(ts.CurrentPartner.level).maxexp - CurExp
        percentofexp = ((CurExp * 100) / dicExp1.Item(ts.CurrentPartner.level).maxexp)
        imgexpscale(1).Width = percentofexp * expscale(1).Width / 100
    End If
    
    Call SetPlayerMeter(0, ts.Character)
    Call SetPlayerMeter(1, ts.CurrentPartner)
    
    
   
    onlineminute = DateDiff("n", StartTime(0), Now)
    onlinesec = DateDiff("s", StartTime(0), Now)


    
    If LastExp(0) <> ts.Character.Texp And LastExp(0) <> 0 Then
        recvexp = ts.Character.Texp - LastExp(0)
        AppendDisplay ts.Character.charname & " ได้รับ Exp +" & recvexp, vbBlack
    End If
    If LastExp(1) <> ts.CurrentPartner.Texp And LastExp(1) = 0 Then
        StartExp(1) = ts.CurrentPartner.Texp
    End If
    If LastExp(1) <> ts.CurrentPartner.Texp And LastExp(1) <> 0 Then
        recvexp = ts.CurrentPartner.Texp - LastExp(1)
        AppendDisplay ts.CurrentPartner.charname & " ได้รับ Exp +" & recvexp, vbBlack
    End If
    
    onlinesec = DateDiff("s", StartTime(0), Now)
    
    'UpdateExpPerMin
    
    
    LastExp(0) = ts.Character.Texp
    LastExp(1) = ts.CurrentPartner.Texp
    
    
    
      'MsgBox ts.Character.Element
'    If ts.Character.Element = 3 Then
'        Picture1.BackColor = &HFF&
'    Else
'        Picture1.BackColor = &HFF00&
'    End If
    
End Sub

Sub HpRecover()
On Error Resume Next
    If mnuAutoEat.Checked = False Then
        If VBscript.Language = "Javascript" Then
            VBscript.ExecuteStatement "HpRecover()"
        End If
        Exit Sub
    End If


    Dim oitem  As Inv
    Dim itm As clsItems
    For i = 1 To 25
        Set oitem = ts.MyItems(i)
        If Not oitem Is Nothing Then
            If IsHP(oitem.itemid) = True Then
            
                If ts.Character.HP < ts.Character.MAXHP Then
                    Set itm = getItem(oitem.itemid)
                    loss = ts.Character.MAXHP - ts.Character.HP
                    nCount = oitem.num
                    addon = nCount * itm.itemvalue
                    If addon <= loss Then
                        ts.EatItemForAuto i, nCount, 0
                        Exit Sub
                    Else
                        nCount = Round(loss / itm.itemvalue)
                        If nCount > 0 Then
                            ts.EatItemForAuto i, nCount, 0
                            Exit Sub
                        End If
                    End If
                End If
            
            
                If ts.CurrentPartner.HP < ts.CurrentPartner.MAXHP Then
                    Set itm = getItem(oitem.itemid)
                    loss = ts.CurrentPartner.MAXHP - ts.CurrentPartner.HP
                    nCount = oitem.num
                    addon = nCount * itm.itemvalue
                    If addon <= loss Then
                        ts.EatItemForAuto i, nCount, ts.CurrentPartner.Order
                        Exit Sub
                    Else
                        nCount = Round(loss / itm.itemvalue)
                        If nCount > 0 Then
                            ts.EatItemForAuto i, nCount, ts.CurrentPartner.Order
                            Exit Sub
                        End If
                    End If
                End If
            
            End If
        
        
            If IsSP(oitem.itemid) = True Then
            
                If ts.Character.SP < ts.Character.MAXSP Then
                    Set itm = getItem(oitem.itemid)
                    loss = ts.Character.MAXSP - ts.Character.SP
                    nCount = oitem.num
                    addon = nCount * itm.itemvalue
                    If addon <= loss Then
                        ts.EatItemForAuto i, nCount, 0
                        Exit Sub
                    Else
                        nCount = Round(loss / itm.itemvalue)
                        If nCount > 0 Then
                            ts.EatItemForAuto i, nCount, 0
                            Exit Sub
                        End If
                    End If
                End If
            
            
                If ts.CurrentPartner.SP < ts.CurrentPartner.MAXSP Then
                    Set itm = getItem(oitem.itemid)
                    loss = ts.CurrentPartner.MAXSP - ts.CurrentPartner.SP
                    nCount = oitem.num
                    addon = nCount * itm.itemvalue
                    If addon <= loss Then
                        ts.EatItemForAuto i, nCount, ts.CurrentPartner.Order
                        Exit Sub
                    Else
                        nCount = Round(loss / itm.itemvalue)
                        If nCount > 0 Then
                            ts.EatItemForAuto i, nCount, ts.CurrentPartner.Order
                            Exit Sub
                        End If
                                            
                    End If
                End If
            
            End If
        
        
        End If
    Next
End Sub


Private Sub ts_onGotGhost()
On Error Resume Next
    VBscript.ExecuteStatement "onEvilGod()"
End Sub

Private Sub ts_onGotLuckyGod()
On Error Resume Next
    VBscript.ExecuteStatement "onLuckyGod()"
End Sub

Private Sub ts_onNPCAppear(ByVal npcmapid As Integer, ByVal x As Long, ByVal y As Long)
On Error Resume Next
    VBscript.ExecuteStatement "onNPCAppear(" & npcmapid & "," & x & ", " & y & ")"
End Sub

Private Sub ts_onOpenCombat()
On Error Resume Next
    '    txtDisplay.Text = "Combat request." & vbNewLine & txtDisplay.Text
'        FightingFlag = True
        AppendDisplay "Combat request.", vbBlue
End Sub

Private Sub ts_onRequestSleep(ByVal price As Long)
On Error Resume Next
    AppendDisplay "sleep " & price & " gold", vbBlue
    ts.doSleep
End Sub

Private Sub ts_onSales(itemid, num, money)
On Error Resume Next
    AppendDisplay "ขาย Item " & getItemName(itemid) & " จำนวน " & num & " ชิ้น ได้รับ " & money & " gold", vbMagenta
    VBscript.ExecuteStatement "onSales(" & itemid & "," & num & "," & money & ")"
End Sub

Private Sub ts_onSendAttack(ByVal fr As Integer, ByVal fc As Integer, ByVal tr As Integer, ByVal tc As Integer, ByVal sk As Long)
On Error Resume Next

End Sub

Private Sub ts_onSetsena(ByVal uID As Long)
On Error Resume Next
    If Form1.cmdShowSena.Caption = "แสดงเสนา" Then
       AppendDisplay "[Party System] :: ผู้เล่น " & getPlayerName(uID) & " ได้รับการแต่งตั้งเป็นเสนาฯ ของ Party", vbRed
    End If
End Sub

Public Sub DisplayLocation()
    txtCurrentLoc.Caption = "mapid = " & ts.Character.mapid & " (" & ts.Character.x & "," & ts.Character.y & ")"
End Sub


Private Sub ts_onWalk(x As Variant, y As Variant)
On Error Resume Next
    DisplayLocation
    VBscript.ExecuteStatement "onWalk(" & x & "," & y & ")"

End Sub

Private Sub ts_onWarp(ByVal uID As Long, ByVal mapid As Long, ByVal warpid As Integer)
On Error Resume Next
Dim kk As Long
    If uID = ts.CurrentParty Then
        ts.LastWarpId = warpid
    End If
    'AppendDisplay "Current map id = " & mapid & ", Last map id = " & ts.lastMapID & ", warp id = " & warpid, vbBlue
    AppendDisplay "Current map id = " & mapid & ", Last warp id = " & warpid, vbBlue
    
    If warpid < 10 Then
        kk = ts.lastMapID
        kk = (kk * 10) + warpid
    Else
        kk = ts.lastMapID
        kk = (kk * 100) + warpid
    End If
    
    DisplayLocation
    AppendDisplay "Current map name = " & GetMapName(kk), vbBlue
 End Sub

Private Sub ts_PartnerAttack()
On Error Resume Next
    VBscript.ExecuteStatement "MyPartnerAttack()"
End Sub

Private Sub ts_PartyStop(ByVal playerid As Long)
On Error Resume Next
    
    If playerid = ts.CurrentParty Then
        Form1.AppendDisplay getPlayerName(playerid) & "สลาย Party แล้ว", vbRed
    End If
    
    VBscript.ExecuteStatement "PartyStop(" & playerid & ")"
    
    
'    MsgBox "กลุ่มสลาย " & getPlayerName(playerid)
End Sub

Private Sub ts_PatchIncorrect()
On Error Resume Next
    AppendDisplay "หมายเลข Patch ไม่ถูกต้องกรุณาตรวจสอบใหม่", vbRed
End Sub

Private Sub ts_PlayerAppearInMap(ByVal playerid As Long, x As Variant, y As Variant)
On Error Resume Next

    VBscript.ExecuteStatement "PlayerAppearInMap(" & playerid & "," & x & "," & y & ")"

End Sub

Private Sub ts_PlayerLeaveMap(ByVal playerid As Long, mapid As Variant, x As Variant, y As Variant)
On Error Resume Next

    VBscript.ExecuteStatement "PlayerLeaveMap(" & mapid & "," & playerid & "," & x & "," & y & ")"

End Sub

Private Sub ts_PlayerOffline(ByVal playerid As Long)
On Error Resume Next

    VBscript.ExecuteStatement "PlayerOffline(" & playerid & ")"
End Sub

Private Sub ts_PlayerOnline(ByVal onlineDatetime As Date)
On Error Resume Next
    Label1(0).Caption = ts.Character.charname
    If ts.Character.NewBorn = 1 Then
        Label1(0).FontBold = True
    End If
    
    'Dim f As Scripting.file
    'Set fso = New Scripting.FileSystemObject
    'Set f = fso.GetFile(ScriptFileName)
    Form1.Caption = "TrueBot - [" & ts.Character.charname & "]" & " Lv." & ts.Character.level
    
    AppendDisplay "Current map id " & ts.Character.mapid & " at (" & ts.Character.x & "," & ts.Character.y & ")", vbBlack
End Sub

Private Sub ts_PlayerWalk(ByVal playerid As Long, ByVal direction As Integer, ByVal x As Long, ByVal y As Long)
On Error Resume Next
    VBscript.ExecuteStatement "onPlayerWalk(" & playerid & ", " & x & "," & y & ")"
    If playerid = ts.CurrentParty Then
        ts.Character.x = x
        ts.Character.y = y
        DisplayLocation
    End If
End Sub

Private Sub ts_ReadyForLogin()
On Error Resume Next
    ts.Login txtAccount.Text, txtPasswd.Text
    cmdLogin.Caption = "Logout"
    icbPartnerList.ComboItems.Clear

End Sub

Private Sub ts_RecvDropItems(ByVal itemid As Long, ByVal num As Integer)
On Error Resume Next
    If chkDrop Then
        AppendDisplay "ได้รับ " & getItemName(itemid) & " จำนวน " & num & " ชิ้น", colDrop
    End If
    'AppendChat "Drop " & getItemName(itemid) & " " & num & " ea.", vbBlue
    VBscript.ExecuteStatement "RecvDropItems(" & itemid & "," & num & ")"
End Sub

Private Sub ts_RecvBackPackItems(ByVal itemid As Long, ByVal num As Integer)
On Error Resume Next
    If chkDrop Then
        AppendDisplay "เก็บของ " & getItemName(itemid) & " จำนวน " & num & " ชิ้น เข้าเป้หลัง", colDrop
    End If
    'AppendChat "Drop " & getItemName(itemid) & " " & num & " ea.", vbBlue
    VBscript.ExecuteStatement "RecvBackPackItems(" & getItemName(itemid) & "," & num & ")"
End Sub


Private Sub ts_RecvItemFrom(ByVal uID As Long, ByVal itemid As Long, ByVal n As Integer)
On Error Resume Next
    AppendDisplay "ได้รับ Item " & getItemName(itemid) & " จาก " & getPlayerName(uID) & " จำนวน " & n & " ชิ้น", vbRed
    VBscript.ExecuteStatement "RecvItemFrom(" & uID & "," & itemid & "," & n & ")"
End Sub

Private Sub ts_RecvMoney(ByVal money As Long)
On Error Resume Next
    Label10.Caption = Format(ts.Character.gold, "###,###,###") & " gold"
End Sub

Private Sub ts_RecvNPCCombat()
On Error Resume Next
Dim npc As NPCSCombat
    ImageCombo2.ComboItems.Clear
    For i = 0 To ts.oNPCCombat.Count - 1
        Set npc = ts.oNPCCombat(i)
        If npc.HP > 0 Then
            AppendDisplay GetNPCName(npc.uID) & "(" & npc.Row & "," & npc.Col & ") Level=" & npc.lv & " HP=" & npc.HP & "/" & npc.MAXHP & " ธาตุ=" & ele2text(npc.elem), vbBlue
            'With ImageCombo2.ComboItems.Add
            '    .Text = "(" & npc.Row & "," & npc.Col & ")" & GetNPCName(npc.uID)
            'End With
        End If
    Next
End Sub
Function GetNPCFormPos(r, c) As NPCSCombat
On Error Resume Next
Dim npc As NPCSCombat
    For i = 0 To ts.oNPCCombat.Count - 1
        Set npc = ts.oNPCCombat.Item(i)
        If npc.Row = r And npc.Col = c Then
            Set GetNPCFormPos = npc
            Exit Function
        End If
    Next
    Set GetNPCFormPos = Nothing
    Exit Function
End Function

Private Sub ts_onDamage()
On Error Resume Next
Dim npcdmg As DamageInfo
Dim NPCFrom As NPCSCombat
Dim targetnpc As DamageTarget
    
    For i = 0 To ts.oNPCCombayDmg.Count - 1
        Set npcdmg = ts.oNPCCombayDmg.Item(i)
'        AppendDisplay "โจมตี " & i & "(" & npcdmg.AttkFromRow & "," & npcdmg.AttkFromCol & ")", RGB(255, 0, 0)
            
            If npcdmg.AttkFromRow = ts.Character.Row And npcdmg.AttkFromCol = ts.Character.Col Then
                msgtext = ts.Character.charname & " " & getSkillName(npcdmg.AttkSkill)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    Set NPCFrom = GetNPCFormPos(targetnpc.Row, targetnpc.Col)
                    
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") "
                    If targetnpc.Row <= 1 Then
                        msgtext1 = msgtext1 & GetNPCName(NPCFrom.uID)
                    End If
                    msgtext1 = msgtext1 & " for " & PlusOrMin(targetnpc.Sign1) & targetnpc.DamagePoint & StatusText(targetnpc.Status1)
                    If (targetnpc.StatusCount = 2) Then
                        If targetnpc.Status2 = &H19 Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = &H1A Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf (targetnpc.Status2 >= &HDC) And (targetnpc.Status2 <= &HDF) Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = 0 Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        End If
                    End If
                    
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
'                Set npcfrom = GetNPCFormPos(npcdmg.AttkFromRow, npcdmg.AttkFromCol)
'                    If Not npcform Is Nothing Then
'                        AppendDisplay GetNPCName(npcform.uid) & " โจมตี", vbBlack
'                    End If
            ElseIf npcdmg.AttkFromRow = ts.CurrentPartner.Row And npcdmg.AttkFromCol = ts.CurrentPartner.Col Then
                msgtext = ts.CurrentPartner.charname & " " & getSkillName(npcdmg.AttkSkill)
                'Set NPCFrom = GetNPCFormPos(npcdmg.AttkToRow, npcdmg.AttkToCol)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    Set NPCFrom = GetNPCFormPos(targetnpc.Row, targetnpc.Col)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") "
                    If targetnpc.Row <= 1 Then
                        msgtext1 = msgtext1 & GetNPCName(NPCFrom.uID)
                    End If
                    msgtext1 = msgtext1 & " for " & PlusOrMin(targetnpc.Sign1) & targetnpc.DamagePoint & StatusText(targetnpc.Status1)
                    If (targetnpc.StatusCount = 2) Then
                        If targetnpc.Status2 = &H19 Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = &H1A Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf (targetnpc.Status2 >= &HDC) And (targetnpc.Status2 <= &HDF) Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = 0 Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        End If
                    End If
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
            ElseIf npcdmg.AttkFromRow = ts.Character.Row Then
                msgtext = "Team" & " " & getSkillName(npcdmg.AttkSkill)
                
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    Set NPCFrom = GetNPCFormPos(targetnpc.Row, targetnpc.Col)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") "
                    If targetnpc.Row <= 1 Then
                        msgtext1 = msgtext1 & GetNPCName(NPCFrom.uID)
                    End If
                    msgtext1 = msgtext1 & " for " & PlusOrMin(targetnpc.Sign1) & targetnpc.DamagePoint & StatusText(targetnpc.Status1)
                    If (targetnpc.StatusCount = 2) Then
                        If targetnpc.Status2 = &H19 Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = &H1A Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf (targetnpc.Status2 >= &HDC) And (targetnpc.Status2 <= &HDF) Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = 0 Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        End If
                    End If
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
            ElseIf npcdmg.AttkFromRow = 2 Then 'ts.CurrentPartner.Row
                msgtext = "Team Partner" & " " & getSkillName(npcdmg.AttkSkill)
                
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    Set NPCFrom = GetNPCFormPos(targetnpc.Row, targetnpc.Col)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") "
                    If targetnpc.Row <= 1 Then
                        msgtext1 = msgtext1 & GetNPCName(NPCFrom.uID)
                    End If
                    msgtext1 = msgtext1 & " for " & PlusOrMin(targetnpc.Sign1) & targetnpc.DamagePoint & StatusText(targetnpc.Status1)
                    If (targetnpc.StatusCount = 2) Then
                        If targetnpc.Status2 = &H19 Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = &H1A Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf (targetnpc.Status2 >= &HDC) And (targetnpc.Status2 <= &HDF) Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = 0 Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        End If
                    End If
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
            Else
                Set NPCFrom = GetNPCFormPos(npcdmg.AttkFromRow, npcdmg.AttkFromCol)
                msgtext = GetNPCName(NPCFrom.uID) & "(" & npcdmg.AttkFromRow & "," & npcdmg.AttkFromCol & ")" & " " & getSkillName(npcdmg.AttkSkill)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") "
                    msgtext1 = msgtext1 & " for " & PlusOrMin(targetnpc.Sign1) & targetnpc.DamagePoint & StatusText(targetnpc.Status1)
                    If (targetnpc.StatusCount = 2) Then
                        If targetnpc.Status2 = &H19 Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = &H1A Then
                            msgtext1 = msgtext1 & PlusOrMin(targetnpc.Sign2) & targetnpc.DamagePoint2 & StatusText(targetnpc.Status2)
                        ElseIf (targetnpc.Status2 >= &HDC) And (targetnpc.Status2 <= &HDF) Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        ElseIf targetnpc.Status2 = 0 Then
                            msgtext1 = msgtext1 & StatusText(targetnpc.Status2)
                        End If
                    End If
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
            End If
    Next
    If ts.oNPCCombayDmg.Count > 1 Then
        If chkFight Then
            AppendDisplay "COMBO", vbGreen
        End If
        'VBscript.ExecuteStatement "Combo()"
    End If
End Sub



Private Sub ts_RecvPartnerLists(ByVal p As Character)
On Error Resume Next
    With icbPartnerList.ComboItems.Add
        .Tag = p.uID
        .Text = p.charname
    End With
End Sub

Private Sub ts_RecvQuestion()
On Error Resume Next
Dim Ques

    Ques = ts.LastQuestion
    Msg = "Question. " & Ques
    AppendDisplay Msg, vbBlue
    
    For Each c In ts.LastAnswers.Keys
        AppendDisplay "Choice is " & ts.LastAnswers(c) & ":" & c, vbBlack
    Next
    VBscript.ExecuteStatement "doRecvQuestion()"

'Winsock1.Connect "www.truedev.net", 80
End Sub
Private Sub ts_ResponseAnswer()
On Error Resume Next
    VBscript.ExecuteStatement "ResponseAnswer()"
    
    fuckgod = True
   ' AppendDisplay "Auto answer choice (" & setanswer & ")", vbBlack
End Sub



Private Sub ts_RequestPartyAcceptFrom(ByVal uID As Long)
On Error Resume Next
    ts.CurrentParty = uID
    Form1.AppendDisplay "เข้าร่วมกลุ่มของ " & getPlayerName(uID), vbRed
    VBscript.ExecuteStatement "RequestPartyAcceptFrom(" & uID & ")"
End Sub

Private Sub ts_RequestPartyFalse(ByVal uID As Long)
On Error Resume Next
 '   ts.CurrentParty = uid
    Form1.AppendDisplay "เข้าร่วมกลุ่มของ " & getPlayerName(uID) & " ล้มเหลว", vbRed
    VBscript.ExecuteStatement "RequestPartyFalse(" & uID & ")"

End Sub


Private Sub ts_SendItemSuccess(ByVal uID As Long, ByVal itemid As Long, ByVal n As Integer)
On Error Resume Next
    AppendDisplay "ส่ง Item " & getItemName(itemid) & " จำนวน " & n & " ชิ้น ไปยัง [" & getPlayerName(uID) & "] สำเร็จ", vbRed
    VBscript.ExecuteStatement "SendItemSuccess()"
End Sub

Private Sub ts_ValidLicence()
On Error Resume Next
    AppendDisplay "Licence is OK.", vbBlack
End Sub

Private Sub ts_WaitngForAcceptParty(ByVal playerid As Long)
On Error Resume Next
    'ts.AcceptParty playerid
    AppendDisplay "[Party System] :: ผู้เล่น " & getPlayerName(playerid) & " ต้องการเข้าร่วม Party", vbRed
    VBscript.ExecuteStatement "AcceptedParty( " & playerid & " )"
    
    
End Sub

Private Sub ts_warpFinish()
On Error Resume Next
    VBscript.ExecuteStatement "warpFinish()"
End Sub

Private Sub VBscript_Error()
On Error Resume Next
   ' MsgBox "Execute script error."
   'AppendChat "Script error....", vbRed
End Sub

Private Sub ListItems2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    LastSelectBPItem = Item.Index
End Sub

Private Sub ListItems2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim itm As ListItem
    Set itm = ListItems2.HitTest(x, y)
         If itm Is Nothing Then Exit Sub
    
    If Button = vbRightButton Then
            itm.Selected = True
            LastSelectBPItem = itm.Index
            PopupMenu mnuCmdAutoBackPack
    Else
        ListItems2.ToolTipText = itm.ToolTipText
    End If
End Sub

Private Sub mnuSendToInven_click()
    ts.BackPackToInven (LastSelectBPItem)
End Sub

Private Sub mnuSendToBackp_click()
    ts.InvenToBackPack (LastSelectItem)
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
            LastSelectItem = itm.Index
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

Private Sub mnuMoveTo_Click()
Dim tempso As Integer
Dim tempnum As Integer

On Error Resume Next
'     Form1.ts.MoveItem ListItems1.SelectedItem
    'tempso = CInt(InputBox("Move To Slot >", "Moving", Default))
    frmMoveItem.Show vbModal, Me
    If tempMoveSlot > 0 Then
        If tempMoveNum >= ts.MyItems(ListItems1.SelectedItem.Index).num Then
            tempMoveNum = ts.MyItems(ListItems1.SelectedItem.Index).num
        End If
        
        If (ts.MyItems(tempMoveSlot).itemid = 0) Then
            If tempMoveNum >= ts.MyItems(ListItems1.SelectedItem.Index).num Then
                tempMoveNum = ts.MyItems(ListItems1.SelectedItem.Index).num
            End If
            ts.MoveItem ListItems1.SelectedItem.Index, tempMoveSlot, tempMoveNum
        ElseIf (ts.MyItems(tempMoveSlot).itemid = ts.MyItems(ListItems1.SelectedItem.Index).itemid) Then
            If tempMoveNum + ts.MyItems(tempMoveSlot).num > 50 Then
                tempMoveNum = 50 - ts.MyItems(tempMoveSlot).num
            End If
            ts.MoveItem ListItems1.SelectedItem.Index, tempMoveSlot, tempMoveNum
        End If
    End If
End Sub

Private Sub mnuSendItem_Click()
On Error Resume Next
    
     Form1.ts.SendItemTo Form1.getPlayerId(Form1.Text3.Text), ListItems1.SelectedItem.Tag, ListItems1.SelectedItem.ListSubItems(2).Text
     AddText3 (Text3.Text)
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

Private Sub ts_ArmyListOnline(ByVal playerid As Long)
Dim i As Integer
    isArmy = False
    If playerid = 558403 Then
        i = 1
    End If
    For i = 1 To listArmy.ListItems.Count
        If listArmy.ListItems.Item(i).Tag = playerid Then
            listArmy.ListItems.Item(i).Bold = True
            listArmy.ListItems.Item(i).Text = "On"
            listArmy.ListItems.Item(i).ForeColor = &H11EE11
            memCount = memCount + 1
            TabStrip1.Tabs.Item(3).Caption = "Army - " & myArmy.AName & "(" & memCount & "/" & 1 + myArmy.AMemCount + myArmy.ASubCount & ")"
            Exit For
        End If
    Next i
End Sub

Private Sub ts_ArmyListOffline(ByVal playerid As Long)
Dim i As Integer
    isArmy = False
    For i = 1 To listArmy.ListItems.Count
        If listArmy.ListItems.Item(i).Tag = playerid Then
            listArmy.ListItems.Item(i).Text = ""
            memCount = memCount - 1
            TabStrip1.Tabs.Item(3).Caption = "Army - " & myArmy.AName & "(" & memCount & "/" & 1 + myArmy.AMemCount + myArmy.ASubCount & ")"
            Exit For
        End If
    Next i
End Sub

Private Sub ts_InitArmyList()
Dim i As Integer
    
    memCount = 0
        listArmy.ListItems.Clear
        With listArmy.ListItems.Add
            .Tag = myArmy.ALeader.id
            If getPlayerId(myArmy.ALeader.Name) = myArmy.ALeader.id Then
                    .Text = "On"
                    .ForeColor = &H11EE11
                    .Bold = True
                    memCount = memCount + 1
            Else
                .Text = ""
            End If
            With .ListSubItems.Add
                .Tag = myArmy.ALeader.id
                .Text = myArmy.ALeader.Name
                .ForeColor = &H2222CC
                If myArmy.ALeader.Reborn Then
                    .Bold = True
                End If
            End With
            With .ListSubItems.Add
                .Tag = myArmy.ALeader.id
                .Text = myArmy.ALeader.Lvl
                .ForeColor = &H2222CC
            End With
            With .ListSubItems.Add
                .Tag = myArmy.ALeader.id
                Select Case myArmy.ALeader.ele
                    Case 1
                        .Text = "ดิน"
                    Case 2
                        .Text = "น้ำ"
                    Case 3
                        .Text = "ไฟ"
                    Case 4
                        .Text = "ลม"
                End Select
                .ForeColor = &H2222CC
            End With
            With .ListSubItems.Add
                .Tag = myArmy.ALeader.id
                .Text = myArmy.ALeader.Con
                .ForeColor = &H2222CC
            End With
        End With
        
        For i = 1 To myArmy.ASubCount
            With listArmy.ListItems.Add
                .Tag = myArmy.ASubLeader(i).id
                If getPlayerId(myArmy.ASubLeader(i).Name) = myArmy.ASubLeader(i).id Then
                    .Text = "On"
                    .ForeColor = &H11EE11
                    .Bold = True
                    memCount = memCount + 1
                Else
                    .Text = ""
                End If
                With .ListSubItems.Add
                    .Tag = myArmy.ASubLeader(i).id
                    .Text = myArmy.ASubLeader(i).Name
                    .ForeColor = &HCC2222
                    If myArmy.ASubLeader(i).Reborn Then
                        .Bold = True
                    End If
                End With
                With .ListSubItems.Add
                    .Tag = myArmy.ASubLeader(i).id
                    .Text = myArmy.ASubLeader(i).Lvl
                    .ForeColor = &HCC2222
                End With
                With .ListSubItems.Add
                    .Tag = myArmy.ASubLeader(i).id
                    Select Case myArmy.ASubLeader(i).ele
                        Case 1
                            .Text = "ดิน"
                        Case 2
                            .Text = "น้ำ"
                        Case 3
                            .Text = "ไฟ"
                        Case 4
                            .Text = "ลม"
                    End Select
                    .ForeColor = &HCC2222
                End With
                With .ListSubItems.Add
                    .Tag = myArmy.ASubLeader(i).id
                    .Text = myArmy.ASubLeader(i).Con
                    .ForeColor = &HCC2222
                End With
            End With
        Next i
        For i = 1 To myArmy.AMemCount
            With listArmy.ListItems.Add
                .Tag = myArmy.AMember(i).id
                If getPlayerId(myArmy.AMember(i).Name) = myArmy.AMember(i).id Then
                    .Text = "On"
                    .ForeColor = &H11EE11
                    .Bold = True
                    memCount = memCount + 1
                Else
                    .Text = ""
                End If
                With .ListSubItems.Add
                    .Tag = myArmy.AMember(i).id
                    .Text = myArmy.AMember(i).Name
                    If myArmy.AMember(i).Reborn Then
                        .Bold = True
                    End If
                End With
                With .ListSubItems.Add
                    .Tag = myArmy.AMember(i).id
                    .Text = myArmy.AMember(i).Lvl
                End With
                With .ListSubItems.Add
                    .Tag = myArmy.AMember(i).id
                    Select Case myArmy.AMember(i).ele
                        Case 1
                            .Text = "ดิน"
                        Case 2
                            .Text = "น้ำ"
                        Case 3
                            .Text = "ไฟ"
                        Case 4
                            .Text = "ลม"
                    End Select
                End With
                With .ListSubItems.Add
                    .Tag = myArmy.AMember(i).id
                    .Text = myArmy.AMember(i).Con
                End With
            End With
        Next i
        listArmy.ListItems(listArmy.ListItems.Count).Selected = True
    
    TabStrip1.Tabs.Item(3).Caption = "Army - " & myArmy.AName & "(" & memCount & "/" & 1 + myArmy.AMemCount + myArmy.ASubCount & ")"
End Sub


