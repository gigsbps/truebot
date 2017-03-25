VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form dlgPlayerOnline 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Player online"
   ClientHeight    =   6960
   ClientLeft      =   12645
   ClientTop       =   4095
   ClientWidth     =   4695
   Icon            =   "dlgPlayerOnline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11880
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อ"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "level"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "dlgPlayerOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
   ' CreatePlayerList
End Sub

Private Sub Form_Load()
    CreatePlayerList
    
    
End Sub
Public Sub CreatePlayerListx()
On Error Resume Next
Dim oTs As Dictionary
Dim uID
Dim itm As ListSubItem
    
    Set oTs = Form1.ts.ol
 
    ListView1.ListItems.Clear
    For Each uID In oTs.Keys
        With ListView1.ListItems.Add
            .Tag = uID
            .Text = oTs.Item(uID).charname
            Set itm = .ListSubItems.Add(, , oTs.Item(uID).level)
                
            If oTs.Item(uID).NewBorn = True Then
                
                .Bold = True
            Else
                
            End If
        End With
    Next
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
Dim i As Integer
    If columheader = 2 Then
        For i = 1 To ListView1.ListItems.Count
            While Len(ListView1.ListItems(i).ListSubItems.Item(1).Text) < 3
                ListView1.ListItems(i).ListSubItems.Item(1).Text = 0 & ListView1.ListItems(i).ListSubItems.Item(1).Text
            Wend
        Next i
    End If
    
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    ListView1.Sorted = True
    ListView1.SortKey = ColumnHeader.Index - 1
    
    If columheader = 2 Then
        For i = 1 To ListView1.ListItems.Count
            While Left(ListView1.ListItems(i).ListSubItems.Item(1).Text, 1) = "0"
                ListView1.ListItems(i).ListSubItems.Item(1).Text = Mid(ListView1.ListItems(i).ListSubItems.Item(1).Text, 2, Len(ListView1.ListItems(i).ListSubItems.Item(1).Text) - 1)
            Wend
        Next i
    End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Form1.Text3.Text = Item.Text
    Form1.Text3.Tag = Item.Tag
End Sub
