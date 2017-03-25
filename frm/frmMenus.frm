VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "Form3"
   ClientHeight    =   1875
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7365
   LinkTopic       =   "Form3"
   ScaleHeight     =   1875
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuItem 
      Caption         =   "ItemMenu"
      Begin VB.Menu mnuUseItem 
         Caption         =   "ใช้ Item"
         Begin VB.Menu mnuUseItemPlayer 
            Caption         =   "Player"
         End
         Begin VB.Menu mnuUseItemPartner 
            Caption         =   "Partner"
         End
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
