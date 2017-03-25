VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChatSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option Setting"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmCharSetting.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox dataLine 
      Height          =   285
      Left            =   2820
      TabIndex        =   31
      Text            =   "1000"
      Top             =   420
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   8
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   28
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   7
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   27
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Value           =   1  'Checked
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cdlgColor 
      Left            =   3480
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox chatLine 
      Height          =   285
      Left            =   2820
      TabIndex        =   20
      Text            =   "1000"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   6
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   5
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   4
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   3
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   2
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   1
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   0
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "DataBox ลบทุก "
      Height          =   195
      Left            =   2820
      TabIndex        =   33
      Top             =   180
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "บรรทัด"
      Height          =   195
      Left            =   3540
      TabIndex        =   32
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "ข้อมูลการต่อสู้"
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   30
      Top             =   3000
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Item Drop"
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   29
      Top             =   2640
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "เปิด/ปิด"
      Height          =   195
      Left            =   60
      TabIndex        =   23
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "บรรทัด"
      Height          =   195
      Left            =   3540
      TabIndex        =   22
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ChatBox ลบทุก "
      Height          =   195
      Left            =   2820
      TabIndex        =   21
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "พันธมิตร"
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "กองทัพ"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   11
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "กลุ่ม"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "ส่วนตัว"
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "กระซิบ"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "ทั้งหมด"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "พื้นหลัง"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmChatSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    Form1.colBack = Picture1(0).BackColor
    Form1.colAll = Picture1(1).BackColor
    Form1.colPublic = Picture1(2).BackColor
    Form1.colWhisper = Picture1(3).BackColor
    Form1.colTeam = Picture1(4).BackColor
    Form1.colGuild = Picture1(5).BackColor
    Form1.colGFriend = Picture1(6).BackColor
    Form1.colDrop = Picture1(7).BackColor
    Form1.colFight = Picture1(8).BackColor
    
    Form1.ChatLength = chatLine.Text
    Form1.DataLength = dataLine.Text
    
    Form1.chkAll = Check1(0).value
    Form1.chkPublic = Check1(1).value
    Form1.chkWhisper = Check1(2).value
    Form1.chkTeam = Check1(3).value
    Form1.chkGuild = Check1(4).value
    Form1.chkGFriend = Check1(5).value
    Form1.chkDrop = Check1(6).value
    Form1.chkFight = Check1(7).value
    
    
    Call Form1.SaveConfig
    MsgBox "จัดเก็บเรียบร้อย", vbOKOnly, "Save"
    Me.Hide
End Sub

Private Sub Form_Load()
    
    chatLine.Text = Form1.ChatLength
    dataLine.Text = Form1.DataLength

    Picture1(0).BackColor = Form1.colBack
    Picture1(1).BackColor = Form1.colAll
    Picture1(2).BackColor = Form1.colPublic
    Picture1(3).BackColor = Form1.colWhisper
    Picture1(4).BackColor = Form1.colTeam
    Picture1(5).BackColor = Form1.colGuild
    Picture1(6).BackColor = Form1.colGFriend
    Picture1(7).BackColor = Form1.colDrop
    Picture1(8).BackColor = Form1.colFight

    Check1(0).value = IIf(Form1.chkAll, 1, 0)
    Check1(1).value = IIf(Form1.chkPublic, 1, 0)
    Check1(2).value = IIf(Form1.chkWhisper, 1, 0)
    Check1(3).value = IIf(Form1.chkTeam, 1, 0)
    Check1(4).value = IIf(Form1.chkGuild, 1, 0)
    Check1(5).value = IIf(Form1.chkGFriend, 1, 0)
    Check1(6).value = IIf(Form1.chkDrop, 1, 0)
    Check1(7).value = IIf(Form1.chkFight, 1, 0)
    
End Sub

Private Sub Picture1_Click(Index As Integer)
    cdlgColor.CancelError = True
    On Error Resume Next
    cdlgColor.ShowColor

    If Err.number = cdlCancel Then
        ' The user canceled. Do nothing.
    ElseIf Err.number <> 0 Then
        ' Unexpected error.
        MsgBox "Error " & Format$(Err.number) & _
            " selecting color." & vbCrLf & _
            Err.Description, _
            vbExclamation Or vbOKOnly, "Error"
    Else
        ' Set the new color.
        Picture1(Index).BackColor = cdlgColor.Color
    End If

End Sub
