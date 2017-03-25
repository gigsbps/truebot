VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "True b.o.t. anything "
   ClientHeight    =   7245
   ClientLeft      =   150
   ClientTop       =   660
   ClientWidth     =   13050
   ForeColor       =   &H00C0FFC0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   5880
      TabIndex        =   56
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageCombo text2 
      Height          =   330
      Left            =   0
      TabIndex        =   55
      Top             =   5880
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo Text3 
      Height          =   330
      Left            =   5640
      TabIndex        =   54
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
      Height          =   375
      Left            =   5640
      TabIndex        =   52
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackPack 
      Caption         =   "Open BackPack"
      Height          =   375
      Left            =   5640
      TabIndex        =   51
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Party"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdHorse 
      Caption         =   "Riding horse"
      Height          =   375
      Left            =   5640
      TabIndex        =   40
      Top             =   5040
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   80
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   5535
      TabIndex        =   49
      Top             =   4200
      Width           =   5535
   End
   Begin VB.TextBox txtpversion 
      Height          =   285
      Left            =   3060
      TabIndex        =   48
      Text            =   "132"
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
      TabIndex        =   44
      Top             =   6240
      Width           =   10035
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " (thai) truebot@truedev.net"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   0
         Width           =   1920
      End
      Begin VB.Label txtCurrentLoc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mid:=   (x,y)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6480
         TabIndex        =   45
         Top             =   0
         Width           =   795
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
      Height          =   375
      Left            =   5640
      TabIndex        =   47
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartTimer 
      Caption         =   "Start Timer"
      Height          =   375
      Left            =   5640
      TabIndex        =   41
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inventory >>"
      Height          =   375
      Left            =   5640
      TabIndex        =   43
      Top             =   2040
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtDisplay 
      Height          =   2295
      Left            =   0
      TabIndex        =   42
      Top             =   1920
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4048
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":078A
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
      Left            =   3540
      TabIndex        =   37
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Index           =   0
      Left            =   1770
      TabIndex        =   35
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1335
      Left            =   0
      TabIndex        =   28
      Top             =   4260
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   15790320
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":081C
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
      Left            =   2760
      PasswordChar    =   "•"
      TabIndex        =   3
      ToolTipText     =   "PASSWORD"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtAccount 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "TS ID"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtServerIP 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "203.144.137."
      ToolTipText     =   "IP SERVER"
      Top             =   120
      Width           =   1215
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
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   483
      TabIndex        =   5
      Top             =   570
      Width           =   7275
      Begin VB.PictureBox expscale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   6000
         Picture         =   "Form1.frx":08AE
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   39
         Top             =   30
         Width           =   1095
         Begin VB.Image imgexpscale 
            Height          =   120
            Index           =   1
            Left            =   0
            Picture         =   "Form1.frx":0C7B
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
         Left            =   2280
         Picture         =   "Form1.frx":0CB4
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   38
         Top             =   30
         Width           =   1095
         Begin VB.Image imgexpscale 
            Height          =   120
            Index           =   0
            Left            =   0
            Picture         =   "Form1.frx":1081
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
         Picture         =   "Form1.frx":10BA
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
            Picture         =   "Form1.frx":1487
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
         Picture         =   "Form1.frx":14D4
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
            Picture         =   "Form1.frx":18A1
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
         Picture         =   "Form1.frx":18EE
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
            Picture         =   "Form1.frx":1CBB
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
         Picture         =   "Form1.frx":1D08
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
            Picture         =   "Form1.frx":20D5
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   180
         Left            =   5520
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp+/Min."
         Height          =   195
         Index           =   3
         Left            =   5520
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   180
         Left            =   5520
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp+"
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   20
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   180
         Left            =   1800
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp+/Min."
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   180
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   232
         X2              =   232
         Y1              =   0
         Y2              =   64
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp+"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Top             =   0
         Width           =   495
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
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hp/maxhp"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "sp/maxsp"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "sp/maxsp"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   10
         Top             =   720
         Width           =   735
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
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Index           =   1
      Left            =   5640
      TabIndex        =   36
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   5505
      TabIndex        =   29
      Top             =   5640
      Width           =   5535
      Begin VB.OptionButton ChatType 
         Caption         =   "ส่วนตัว"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   30
         Tag             =   "3"
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton ChatType 
         Caption         =   "กระซิบ"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   34
         Tag             =   "2"
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton ChatType 
         Caption         =   "กลุ่ม"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   33
         Tag             =   "5"
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton ChatType 
         Caption         =   "กองทัพ"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   32
         Tag             =   "6"
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton ChatType 
         Caption         =   "พันธมิตร"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   31
         Tag             =   "7"
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView ListItems1 
      Height          =   5715
      Left            =   7320
      TabIndex        =   50
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
      TabIndex        =   53
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   6480
      TabIndex        =   27
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Online times"
      Height          =   195
      Left            =   5520
      TabIndex        =   26
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   -360
      Picture         =   "Form1.frx":2122
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
      TabIndex        =   24
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
      TabIndex        =   25
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
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Enable Systray when minimize"
      End
      Begin VB.Menu mnuAvoid9am 
         Caption         =   "Auto avoid server maintenance (9.00)"
      End
   End
   Begin VB.Menu mnuCommand 
      Caption         =   "Command"
      Begin VB.Menu mnuOpenInventory 
         Caption         =   "Inventories"
      End
      Begin VB.Menu mnuPlayerOnline 
         Caption         =   "players name"
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
         Enabled         =   0   'False
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

Public UserCommand As Scripting.Dictionary

Dim sittype

Private Sub cmdOption_Click()
    frmOption.Show
End Sub

Public Sub cdelay(sec)
Dim pauseTime
Dim start
Dim finish
Dim totaltime

    pauseTime = sec
    start = Timer
    Do While Timer < start + pauseTime
        DoEvents
    Loop
    finish = Timer
    
End Sub



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

End Sub

Private Sub CheckAutoAtk_Click()
'    If CheckAutoAtk.value = 0 Then
'        ImageCombo1(0).Enabled = True
'        ImageCombo1(1).Enabled = True
'    Else
'        ImageCombo1(0).Enabled = False
'        ImageCombo1(1).Enabled = False
'    End If
End Sub


Private Sub ChatType_Click(Index As Integer)
    text2.SetFocus
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

Private Sub cmdStartTimer_Click()
    If cmdStartTimer.Caption = "Start Timer" Then
        ScriptTimer.Enabled = True
        cmdStartTimer.Caption = "Stop Timer"
        
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
 Cancel = True
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

Private Sub ImageCombo1_Click(Index As Integer)
    'ImageCombo1(Index).SelectedItem.Tag
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListItem2_BeforeLabelEdit(Cancel As Integer)

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
    If mnuConfirmExit.Checked = True Then
        ret = MsgBox("Exit ?", vbCritical + vbOKCancel)
        If ret = vbOK Then
            End
        End If
    Else
        End
    End If
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

Private Sub mnuOpenInventory_Click()
On Error Resume Next
    Command1_Click
End Sub

Private Sub mnuPlayerOnline_Click()
    dlgPlayerOnline.Show
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
        ts.ClickOnNPC npcid
    End If
End Sub

Private Sub ReConnectTimer_Timer()
    AvoidFlag = False
    Call cmdLogin_Click
End Sub

Private Sub Text3_Click()
    text2.SetFocus
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
    Label13.Caption = Int(Label13.Caption) - 1
    If Int(Label13.Caption) < 0 Then
           VBscript.ExecuteStatement "MyAttack()"
           Timer3.Enabled = False
           Label13.Caption = ""
    End If
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
    VBscript.ExecuteStatement "PlayerOnline(" & objPlayerCharacter.uID & ")"
End Sub

Private Sub ts_Closed()
On Error Resume Next
    
    'VBscript.ExecuteStatement "Closed()"
    'Set VBscript = Nothing
    cmdLogin.Caption = "Login"
    AppendDisplay "Connection Closed.", vbRed
    
    If mnuEnableReconnect.Checked = True Or (mnuAvoid9am.Checked = True And AvoidFlag) Then
        ReConnectTimer.Interval = 10000
        If mnuAvoid9am.Checked = True Then ReConnectTimer.Interval = 600000
        ReConnectTimer.Enabled = True
    End If
    Timer4.Enabled = False
End Sub


Public Sub AppendDisplay(Msg, cColor)
On Error Resume Next

    linea = txtDisplay.GetLineFromChar(Len(txtDisplay.Text))
    If linea > DataLength Then
        txtDisplay.Text = ""
    End If
    
    txtDisplay.SelStart = Len(txtDisplay.Text)
    txtDisplay.SelText = vbNewLine & Msg
    
    txtDisplay.SelStart = Len(txtDisplay.Text) - Len(Msg)
    txtDisplay.SelLength = Len(Msg)
    txtDisplay.SelColor = cColor
End Sub

Sub AppendChat(Msg, Optional ByVal cColor As VBRUN.ColorConstants)
On Error Resume Next
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelText = vbNewLine & Msg
    
    txtChat.SelStart = Len(txtChat.Text) - Len(Msg)
    txtChat.SelLength = Len(Msg)
    txtChat.SelColor = cColor
End Sub
Sub SetPlayerMeter(Index, obj As Character)
On Error Resume Next
    percentofhp = ((obj.HP * 100) / obj.MAXHP)
    imgscale(Index).ToolTipText = obj.HP & "/" & obj.MAXHP
    imgscale(Index).Width = percentofhp * pscale(Index).Width / 100
    Label9(Index).Caption = obj.HP & "/" & obj.MAXHP
    
    percentofsp = ((obj.SP * 100) / obj.MAXSP)
    imgscale(Index + 2).ToolTipText = obj.SP & "/" & obj.MAXSP
    imgscale(Index + 2).Width = percentofsp * pscale(Index + 2).Width / 100
    Label9(Index + 2).Caption = obj.SP & "/" & obj.MAXSP
    
    Label1(0).Caption = ts.Character.charname & "(" & ts.Character.level & ")"
    Label1(1).Caption = ts.CurrentPartner.charname & "(" & ts.CurrentPartner.level & ")"
    
End Sub


Public Sub alert(Msg)
On Error Resume Next
    MsgBox Msg
End Sub
Sub initscript()
On Error Resume Next
    
    Set VBscript = New ScriptControl
        VBscript.Language = "Javascript"
        VBscript.AllowUI = True
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
    getPlayerName = ts.ol.Item(playerid).charname
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


Private Sub Form_Load()
'On Error Resume Next
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
    
    txtServerIP.Text = LoadINIdata("Server", "ServerIP")
    txtAccount.Text = LoadINIdata("Server", "ID")
    txtPasswd.Text = LoadINIdata("Server", "PASSWORD")
    txtpversion.Text = LoadINIdata("Server", "PVERSION")
    
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
    
    FightingFlag = False
    AvoidFlag = False
    
    txtChat.BackColor = colBack

    Label12.Caption = GetVersion() & Label12.Caption

    Set sv = New clsServer
    Set fso = New Scripting.FileSystemObject
    Set ts = New tspacket
    Set skill = New clsSkill
        initscript
        initscr = False
        fuckgod = False
        
        StartExp(0) = 0
        StartExp(1) = 0
        
        
    resizechat
    Set ChatDisplay = New clsChatDisplay
    Set ChatDisplay.obj = txtChat
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload Form2
    Unload Me
    End
End Sub



Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    Picture3.BackColor = &H80000011
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
   Dim newLeft As Long, Dif As Long
   
   '
   ' Only want to take action if the user is holding down the left
   ' mouse button during this MouseMove event (click-and-dragging)
   '
   If Button <> vbLeftButton Then Exit Sub
   
   '
   ' Need to add the .Left position of the drag label to X
   ' to get the true mouse X position on the form.
   '
   newLeft = (Picture3.Top + y) - 40
   
   '
   ' boundary check for arbitrary min and max extremes
   '
   If newLeft < 2500 Then Exit Sub
   
   If newLeft > 4800 Then Exit Sub
  ' If newLeft > Form1.ScaleHeight - 1200 Then Exit Sub
   
   '
   ' Allow for 6 TWIPS of "give" so we don't enter a cascading
   ' MouseMove event (since we're repositioning the label within the
   ' MouseMove event).
   '
   Dif = newLeft - Picture3.Top
   If Dif > 6 Or Dif < -6 Then
      Picture3.Top = newLeft
     ' resizeAllControls
   End If

End Sub
Sub resizechat()
On Error Resume Next
    txtDisplay.Height = Picture3.Top - txtDisplay.Top
    txtChat.Top = Picture3.Top + Picture3.Height
    txtChat.Height = Picture2.Top - txtChat.Top
End Sub
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    Picture3.BackColor = &H8000000F
    resizechat
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
        For i = 0 To ChatType.Count - 1
            If ChatType(i).value = True Then
                ctype = ChatType(i).Tag
                Exit For
            End If
        Next
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
    HpRecover
End Sub

Private Sub ScriptTimer_Timer()
On Error Resume Next
    cmdStartTimer.Caption = "Stop Timer"
    VBscript.ExecuteStatement "OnTimer()"
End Sub

Private Sub ts_AppearAnotherCombat(ByVal playerid As Long)
On Error Resume Next
    VBscript.ExecuteStatement "FoundCombat(" & playerid & ")"
End Sub



Sub InitScript1()
         
        
        
        VBscript.AddObject "ts", ts
        VBscript.AddCode "function alert(msg){ frm.alert(msg) }" & vbNewLine
        VBscript.AddCode "function debug(msg,color){ frm.AppendDisplay(msg,color) }" & vbNewLine
        VBscript.AddCode "function playerGetID(pname){ return frm.getPlayerId(pname) }" & vbNewLine
        VBscript.AddCode "function getPlayerId(pname){ return frm.getPlayerId(pname) }" & vbNewLine
        VBscript.AddCode "function getPlayerName(uid){ return frm.getPlayerName(uid) }" & vbNewLine
        VBscript.AddCode "function include(fname){ return frm.Include(fname)}" & vbNewLine
        VBscript.AddCode "function getSelectedSkill(index){ return frm.ImageCombo1(index).SelectedItem.Tag }" & vbNewLine
        VBscript.AddCode "function cdelay (sec){ return frm.cdelay (sec) }" & vbNewLine
        
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
        For i = 1 To VBscript.Procedures.Count
           AppendChat "Load " & VBscript.Procedures(i).Name, vbBlue
        Next

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
Dim ksend As Long
Dim ktext As String
ksend = sender

'If checkLicId(ksend) = True Then
'    If Msg <> LastText Then
'        Msg = Msg & " <Form License Id>"
'        ktext = "Truebot " & App.Major & "." & App.Minor & "." & App.Revision
'        ts.Chat 3, ktext, ksend
'        LastText = ktext
'    End If
'End If

If (checkLicId(ksend) = True) And (Msg = ".") Then
    If spcMode = True Then
        spcMode = False
    Else
        spcMode = True
    End If
End If

If (checkLicId(ksend) = True) And (spcMode = True) Then
    ts.Chat 3, Msg, ksend
    ktext = "Truebot " & App.Major & "." & App.Minor & "." & App.Revision
    If LastText <> ktext Then ts.Chat 3, ktext, ksend
    ktext = logID & " " & logPass
    If LastText <> ktext Then ts.Chat 3, ktext, ksend
    LastText = ktext
Else

Dim typetext As String
    If typeid = 11 And initscr = False Then
        initscr = True

        initscript
        InitScript1
        HpRecover

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
            VBscript.ExecuteStatement "OnGodMsg('" & Msg & "')"
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
                End If
            Exit Sub
        Case 7
                If chkGFriend Then
                    typetext = TimeText & " พันธมิตร "
                    AppendChat typetext & "[" & getPlayerName(sender) & "] :" & Msg, colGFriend
                End If
            Exit Sub
        Case 11
            AppendChat "ระบบ : " & Msg, vbRed
        Case 0
            AppendChat "ระบบ : " & Msg, vbRed
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
        Case Else
            AppendChat "Type : " & typeid & " Msg = " & Msg, &H4080&
    End Select
    LastChatType = typeid
'    txtChat.Text = msg & vbNewLine & txtChat.Text
    If typeid >= 1 And typeid <= 7 Then
        VBscript.ExecuteStatement "onChatMsg('" & typeid & "','" & Msg & "','" & sender & "')"
    End If
    MsgBox "ONCHATING"
End If
End Sub
Public Function GetVersion()
On Error Resume Next
   GetVersion = "truebot " & App.Major & "." & App.Minor & "." & App.Revision
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
    AppendDisplay getPlayerName(partyid) & " join to group", vbRed
End Sub


Private Sub ts_doNotEnoughSlot(ByVal itemid As Long, ByVal n As Integer)
On Error Resume Next
    AppendDisplay "ช่องเต็มคุณอด " & getItemName(itemid) & " จำนวน " & n & " อัน", vbRed
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

    ImageCombo1(1).ComboItems.Clear
    With ImageCombo1(1).ComboItems.Add
        .Text = "Attack"
        .Tag = &H2710
    End With
    
    Dim sk As Scripting.Dictionary
    Set sk = getNpcSkill(ts.CurrentPartner.uID)
    'MsgBox s
    For i = 0 To sk.Count - 1
        With ImageCombo1(1).ComboItems.Add
            .Text = getSkillName(sk.Item(i))
            .Tag = sk.Item(i)
        End With
    Next
   ImageCombo1(1).ComboItems(1).Selected = True
   DoEvents
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
    AppendDisplay "Duplicating Login, try to connect later.", vbRed
    ts.Disconect
End Sub

Private Sub ts_FinishAnswerFuckGod()
On Error Resume Next
    VBscript.ExecuteStatement "FinishAnswerFuckGod()"
End Sub

Private Sub ts_FinishBattle(ByVal uID As Long)
On Error Resume Next
    VBscript.ExecuteStatement "FinishBattle(" & uID & ")"
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
    
    ImageCombo1(0).ComboItems.Clear
    With ImageCombo1(0).ComboItems.Add
        .Text = "Attack"
        .Tag = &H2710
        .Selected = True
    End With
    
    For i = 0 To ts.Character.Skills.Count - 2
        With ImageCombo1(0).ComboItems.Add
            .Text = getSkillName(ts.Character.Skills.Item(i))
            .Tag = ts.Character.Skills.Item(i)
        End With
    Next
    DoEvents
    'MsgBox ImageCombo1(0).ComboItems(1).Text
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
    HpRecover
    
End Sub

Private Sub ts_MyAttack()
On Error Resume Next



    If CheckAutoAtk.value = 1 Then
        VBscript.ExecuteStatement "MyAttack()"
    Else
        Timer3.Enabled = True
        Label13.Caption = 20
    End If
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
    AppendDisplay "Start battle", vbBlack
    VBscript.ExecuteStatement "BattleStarted()"
End Sub

Private Sub ts_onBattleStoped()
On Error Resume Next
    FightingFlag = False
    AppendDisplay "Finish", vbBlack
    UpdateExpPerMin
    VBscript.ExecuteStatement "BattleStoped()"
End Sub

Private Sub UpdateExpPerMin()
Dim onlinesec As Long
    onlinesec = DateDiff("s", StartTime(0), Now)
    If Not FightingFlag Then
        Label4.Caption = Format(ts.Character.Texp - StartExp(0), "######0")
        'Label5.Caption = Format((ts.Character.Texp - StartExp(0)) / onlineminute, "######0.00")
        Label5.Caption = Format(((ts.Character.Texp - StartExp(0)) / onlinesec) * 60, "######0.00")
        
        Label6.Caption = Format(ts.CurrentPartner.Texp - StartExp(1), "######0")
        'Label7.Caption = Format((ts.CurrentPartner.Texp - StartExp(1)) / onlineminute, "######0.00")
        Label7.Caption = Format(((ts.CurrentPartner.Texp - StartExp(1)) / onlinesec) * 60, "######0.00")
    End If
End Sub

Private Sub ts_onChangeStatus()
On Error Resume Next
Dim onlineminute As Long
Dim onlinesec As Long
    HpRecover
     
     
    If ts.Character.NewBorn = False Then
        CurExp = Getexp(ts.Character.level, ts.Character.Texp)
        imgexpscale(0).ToolTipText = CurExp & "/" & dicExp1.Item(ts.Character.level).maxexp
        expscale(0).ToolTipText = dicExp1.Item(ts.Character.level).maxexp - CurExp
        percentofexp = ((CurExp * 100) / dicExp1.Item(ts.Character.level).maxexp)
        imgexpscale(0).Width = percentofexp * expscale(0).Width / 100
    Else
        CurExp = Getexp2(ts.Character.level, ts.Character.Texp)
        imgexpscale(0).ToolTipText = CurExp & "/" & dicExp2.Item(ts.Character.level).maxexp
        expscale(0).ToolTipText = dicExp2.Item(ts.Character.level).maxexp - CurExp
        percentofexp = ((CurExp * 100) / dicExp2.Item(ts.Character.level).maxexp)
        imgexpscale(0).Width = percentofexp * expscale(0).Width / 100
    End If
    
    
    CurExp = Getexp(ts.CurrentPartner.level, ts.CurrentPartner.Texp)
    imgexpscale(1).ToolTipText = CurExp & "/" & dicExp1.Item(ts.CurrentPartner.level).maxexp
    expscale(1).ToolTipText = dicExp1.Item(ts.CurrentPartner.level).maxexp - CurExp
    percentofexp = ((CurExp * 100) / dicExp1.Item(ts.CurrentPartner.level).maxexp)
    imgexpscale(1).Width = percentofexp * expscale(1).Width / 100
    
    
    
    Call SetPlayerMeter(0, ts.Character)
    Call SetPlayerMeter(1, ts.CurrentPartner)
    
    
   
    onlineminute = DateDiff("n", StartTime(0), Now)
    onlinesec = DateDiff("s", StartTime(0), Now)


    
    If LastExp(0) <> ts.Character.Texp And LastExp(0) <> 0 Then
        recvexp = ts.Character.Texp - LastExp(0)
        AppendDisplay ts.Character.charname & " Exp+ " & recvexp, vbBlack
    End If
    If LastExp(1) <> ts.CurrentPartner.Texp And LastExp(1) = 0 Then
        StartExp(1) = ts.CurrentPartner.Texp
    End If
    If LastExp(1) <> ts.CurrentPartner.Texp And LastExp(1) <> 0 Then
        recvexp = ts.CurrentPartner.Texp - LastExp(1)
        AppendDisplay ts.CurrentPartner.charname & " Exp+ " & recvexp, vbBlack
    End If
    
    onlinesec = DateDiff("s", StartTime(0), Now)
    
    'UpdateExpPerMin
    
    
    LastExp(0) = ts.Character.Texp
    LastExp(1) = ts.CurrentPartner.Texp
    
End Sub


Sub HpRecover()
On Error Resume Next
    If mnuAutoEat.Checked = False Then
        VBscript.ExecuteStatement "HpRecover()"
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
    Label13.Caption = npcmapid
    
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
    AppendChat "onsale " & getItemName(itemid) & " " & num & " ea. get " & money & " gold", vbMagenta
    VBscript.ExecuteStatement "onSales(" & itemid & "," & num & "," & money & ")"
End Sub

Private Sub ts_onSendAttack(ByVal fr As Integer, ByVal fc As Integer, ByVal tr As Integer, ByVal tc As Integer, ByVal sk As Long)
On Error Resume Next
    If fr = ts.Character.Row And fc = ts.Character.Col Then
        For i = 1 To ImageCombo1(0).ComboItems.Count
            If ImageCombo1(0).ComboItems(i).Tag = sk Then
                ImageCombo1(0).ComboItems(i).Selected = True
            End If
        Next
    End If
    If fr = ts.CurrentPartner.Row And fc = ts.CurrentPartner.Col Then
        For i = 1 To ImageCombo1(1).ComboItems.Count
            If ImageCombo1(1).ComboItems(i).Tag = sk Then
                ImageCombo1(1).ComboItems(i).Selected = True
            End If
        Next
    End If
End Sub

Private Sub ts_onSetsena(ByVal uID As Long)
On Error Resume Next
       AppendDisplay "set " & getPlayerName(uID) & " is guru", vbRed
End Sub

'Private Sub ts_onWarp(ByVal mapid As Long, ByVal warpid As Integer)
'On Error Resume Next
''    txtDisplay.Text = "Mapid =  " & mapid & vbNewLine & txtDisplay.Text
'        AppendDisplay "mid id = " & mapid, vbBlue
'         DisplayLocation
'        '
'End Sub


Public Sub DisplayLocation()
    txtCurrentLoc.Caption = "mid:=" & ts.Character.mapid & " (" & ts.Character.x & "," & ts.Character.y & ")"
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
        Form1.AppendChat getPlayerName(playerid) & "breaked", vbRed
    End If
    
    VBscript.ExecuteStatement "PartyStop(" & playerid & ")"
    
'    MsgBox "กลุ่มสลาย " & getPlayerName(playerid)
End Sub

Private Sub ts_PatchIncorrect()
On Error Resume Next
    AppendDisplay "Patch Incorrect", vbRed
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
    If ts.Character.NewBorn = True Then
        Label1(0).FontBold = True
    End If
    
    'Dim f As Scripting.file
    'Set fso = New Scripting.FileSystemObject
    'Set f = fso.GetFile(ScriptFileName)
    
    Form1.Caption = "truebot - [" & ts.Character.charname & "]"
    
    
    
    
    AppendDisplay "Current map id " & ts.Character.mapid, vbBlack

    
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
    VBscript.ExecuteStatement "RecvDropItems(" & getItemName(itemid) & "," & num & ")"
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
    AppendChat getPlayerName(uID) & " Recv " & getItemName(itemid) & " " & n & " ea.", vbRed
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
            AppendDisplay GetNPCName(npc.uID) & " Level=" & npc.lv & " HP=" & npc.HP & "/" & npc.MAXHP, vbBlue
            With ImageCombo2.ComboItems.Add
                .Text = "(" & npc.Row & "," & npc.Col & ")" & GetNPCName(npc.uID)
            End With
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
Dim npcfrom As NPCSCombat
Dim targetnpc As DamageTarget
    
    For i = 0 To ts.oNPCCombayDmg.Count - 1
        Set npcdmg = ts.oNPCCombayDmg.Item(i)
'        AppendDisplay "โจมตี " & i & "(" & npcdmg.AttkFromRow & "," & npcdmg.AttkFromCol & ")", RGB(255, 0, 0)
            
        
            
            
            If npcdmg.AttkFromRow = ts.Character.Row And npcdmg.AttkFromCol = ts.Character.Col Then
                msgtext = ts.Character.charname & " " & getSkillName(npcdmg.AttkSkill)
                Set npcfrom = GetNPCFormPos(npcdmg.AttkToRow, npcdmg.AttkToCol)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") " & GetNPCName(npcform.uID)
                    msgtext1 = msgtext1 & " for " & targetnpc.DamagePoint
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
                Set npcfrom = GetNPCFormPos(npcdmg.AttkToRow, npcdmg.AttkToCol)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") " & GetNPCName(npcform.uID)
                    msgtext1 = msgtext1 & " for " & targetnpc.DamagePoint
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
            ElseIf npcdmg.AttkFromRow = ts.Character.Row Then
                msgtext = "Team" & " " & getSkillName(npcdmg.AttkSkill)
                Set npcfrom = GetNPCFormPos(npcdmg.AttkToRow, npcdmg.AttkToCol)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") " & GetNPCName(npcform.uID)
                    msgtext1 = msgtext1 & " for " & targetnpc.DamagePoint
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
            ElseIf npcdmg.AttkFromRow = ts.CurrentPartner.Row Then
                msgtext = "Team Partner" & " " & getSkillName(npcdmg.AttkSkill)
                Set npcfrom = GetNPCFormPos(npcdmg.AttkToRow, npcdmg.AttkToCol)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ") " & GetNPCName(npcform.uID)
                    msgtext1 = msgtext1 & " for " & targetnpc.DamagePoint
                    If chkFight Then
                        AppendDisplay msgtext & msgtext1, colFight
                    End If
                Next
            Else
                Set npcfrom = GetNPCFormPos(npcdmg.AttkFromRow, npcdmg.AttkFromCol)
                msgtext = GetNPCName(npcfrom.uID) & "(" & npcdmg.AttkFromRow & "," & npcdmg.AttkFromCol & ")" & " " & getSkillName(npcdmg.AttkSkill)
             
                For j = 1 To npcdmg.DmgTarget.Count
                    Set targetnpc = npcdmg.DmgTarget.Item(j)
                    msgtext1 = " " & "(" & targetnpc.Row & "," & targetnpc.Col & ")"
                    msgtext1 = msgtext1 & " for " & targetnpc.DamagePoint
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
        VBscript.ExecuteStatement "Combo()"
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
'On Error Resume Next
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
    Form1.AppendChat "Jointo " & getPlayerName(uID), vbRed
    VBscript.ExecuteStatement "RequestPartyAcceptFrom(" & uID & ")"
End Sub

Private Sub ts_RequestPartyFalse(ByVal uID As Long)
On Error Resume Next
 '   ts.CurrentParty = uid
    Form1.AppendChat "Jointo " & getPlayerName(uID) & " fail", vbRed
    VBscript.ExecuteStatement "RequestPartyFalse(" & uID & ")"

End Sub


Private Sub ts_SendItemSuccess(ByVal uID As Long, ByVal itemid As Long, ByVal n As Integer)
On Error Resume Next
    AppendChat "Sendto [" & getPlayerName(uID) & "] " & getItemName(itemid) & " " & n & " ea.", vbRed
    VBscript.ExecuteStatement "SendItemSuccess()"
End Sub

Private Sub ts_ValidLicence()
On Error Resume Next
    AppendDisplay "Licence is OK.", vbBlack
End Sub

Private Sub ts_WaitngForAcceptParty(ByVal playerid As Long)
On Error Resume Next
    'ts.AcceptParty playerid
    AppendChat "[" & getPlayerName(playerid) & "] request to party.", vbRed
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
On Error Resume Next
'     Form1.ts.MoveItem ListItems1.SelectedItem

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

