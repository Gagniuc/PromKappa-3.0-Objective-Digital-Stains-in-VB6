VERSION 5.00
Begin VB.Form Save_data 
   Caption         =   "Save"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox list_gene 
      Height          =   5535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   14895
   End
   Begin VB.CommandButton ok 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton save 
      Caption         =   "Save text"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   5760
      Width           =   1815
   End
End
Attribute VB_Name = "Save_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ok_Click()
    Unload Me
End Sub

