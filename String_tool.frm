VERSION 5.00
Begin VB.Form String_tool 
   Caption         =   "Tools"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton len_original_text 
      Caption         =   "Len text"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cut 
      Caption         =   "Cut"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   5640
      Width           =   2055
   End
   Begin VB.TextBox len_cut 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "599"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox from 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox process_text 
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3240
      Width           =   9135
   End
   Begin VB.TextBox original_text 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label len_cut_txt 
      Alignment       =   1  'Right Justify
      Caption         =   "Lent text: 0"
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Len text:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "String_tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cut_Click()


original_text.Text = Replace(original_text.Text, Chr(10), "")
original_text.Text = Replace(original_text.Text, Chr(13), "")

process_text.Text = Mid(original_text.Text, Val(from.Text), Val(len_cut.Text) + 1)

len_cut_txt.Caption = "Lent text: " & Len(process_text.Text)

End Sub

Private Sub len_original_text_Click()
original_text.Text = Replace(original_text.Text, Chr(10), "")
original_text.Text = Replace(original_text.Text, Chr(13), "")

MsgBox Len(original_text.Text)
End Sub
