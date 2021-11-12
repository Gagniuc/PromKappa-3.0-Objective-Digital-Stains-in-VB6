VERSION 5.00
Begin VB.Form Pattern_Numbers 
   Caption         =   "Pattern Data"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox graf_m 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Mask Pen
      Height          =   4695
      Left            =   480
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   30
      Top             =   360
      Width           =   5055
      Begin VB.Line ZeroLine 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   3
         X1              =   8
         X2              =   320
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Shape Focus_Shape 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Height          =   135
         Left            =   5040
         Top             =   0
         Width           =   135
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   336
         X2              =   0
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   312
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "C+G%"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4440
         TabIndex        =   32
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "KIC"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Width           =   735
      End
      Begin VB.Line Line1_x_min 
         BorderColor     =   &H000080FF&
         BorderStyle     =   2  'Dash
         X1              =   56
         X2              =   56
         Y1              =   0
         Y2              =   304
      End
      Begin VB.Line Line2_x_max 
         BorderColor     =   &H000080FF&
         BorderStyle     =   2  'Dash
         X1              =   280
         X2              =   280
         Y1              =   0
         Y2              =   312
      End
      Begin VB.Line Line2_y_max 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   2  'Dash
         X1              =   0
         X2              =   336
         Y1              =   48
         Y2              =   48
      End
      Begin VB.Line Line1_y_min 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   2  'Dash
         X1              =   0
         X2              =   336
         Y1              =   256
         Y2              =   256
      End
      Begin VB.Line best_fit 
         BorderColor     =   &H008080FF&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   3
         X1              =   8
         X2              =   320
         Y1              =   216
         Y2              =   216
      End
   End
   Begin VB.CommandButton Add_data_of_the_pattern_to_mem 
      Caption         =   "Add Data"
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Save_Data_HTML 
      Caption         =   "Save"
      Height          =   495
      Left            =   5760
      TabIndex        =   24
      Top             =   5520
      Width           =   3375
   End
   Begin VB.CommandButton L_R 
      Caption         =   "Linear regression on pattern"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   5520
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pattern Data"
      Height          =   4695
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   5175
      Begin VB.CheckBox R_line 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox E_line 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   34
         Top             =   1320
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox PS_lines 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox Angle_0 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Text            =   "0"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox pattern_surface 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Angle_360 
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   3960
         TabIndex        =   19
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Angle_180 
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Y_diff 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Text            =   "0"
         ToolTipText     =   "Height of the pattern (Y max - Y min)"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox X_diff 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "Width of the pattern (X max - X min)"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Y_limax 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Y_limin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Text            =   "0"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox X_limax 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   3240
         TabIndex        =   7
         Text            =   "0"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox X_limin 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "0"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox CG_mean_txt_m 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "0"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox KIC_mean_txt_m 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Angle 90  0  -90:"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label PattS 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pattern surface (%):"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   2775
         Left            =   600
         Shape           =   2  'Oval
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label A360 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Angle 360°:"
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label A180 
         BackStyle       =   0  'Transparent
         Caption         =   "Angle 180°:"
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label YMaxL 
         BackStyle       =   0  'Transparent
         Caption         =   "Y max limit:"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label YMinL 
         BackStyle       =   0  'Transparent
         Caption         =   "Y min limit:"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label XMaxL 
         BackStyle       =   0  'Transparent
         Caption         =   "X max limit:"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label XMinL 
         BackStyle       =   0  'Transparent
         Caption         =   "X min limit:"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label ACGV 
         BackStyle       =   0  'Transparent
         Caption         =   "Average C+G values:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label AKIC 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Average KIC values:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton ok 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   44
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4240
      TabIndex        =   43
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2920
      TabIndex        =   42
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1720
      TabIndex        =   41
      Top             =   5170
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   3865
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   1465
      Width           =   255
   End
   Begin VB.Line Line12 
      X1              =   360
      X2              =   480
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line11 
      X1              =   360
      X2              =   480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line10 
      X1              =   1800
      X2              =   1800
      Y1              =   5160
      Y2              =   5040
   End
   Begin VB.Line Line9 
      X1              =   4320
      X2              =   4320
      Y1              =   5160
      Y2              =   5040
   End
   Begin VB.Line Line8 
      X1              =   5520
      X2              =   5520
      Y1              =   5160
      Y2              =   5040
   End
   Begin VB.Line Line7 
      X1              =   3000
      X2              =   3000
      Y1              =   5160
      Y2              =   5040
   End
   Begin VB.Line Line6 
      X1              =   480
      X2              =   360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label6 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2665
      Width           =   255
   End
   Begin VB.Line Line4 
      X1              =   480
      X2              =   480
      Y1              =   5160
      Y2              =   5040
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   480
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   5160
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   265
      Width           =   255
   End
   Begin VB.Label GNX 
      Caption         =   "none"
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label GName 
      Caption         =   "Gene name:"
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.Label angle_regression 
      Caption         =   "Angle:"
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   5040
      Width           =   5055
   End
End
Attribute VB_Name = "Pattern_Numbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /            PromKappa           \________________________/       v3.00        |
' |                                                                               |
' |            Name:  PromKappa V3.0                                              |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |  ____________________________________________________________________________ |
' |                                                                               |
' |    Date Created:  September 2013                                              |
' |       Tested On:  WinXP, WinVista, Win7, Win8                                 |
' |             Use:  Analysis of gene promoters                                  |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Option Explicit

Dim Sx As Double
Dim Sxy As Double
Dim Sy As Double
Dim Sx2 As Double
Dim n As Long
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
Private Const XMAX = 400
Private Const YMAX = 400
Private Const XMIN = -400
Private Const YMIN = -400
Dim patt_gene_name As String

Private Const PI As Double = 3.14159265358979

Private Sub Form_Load()
On Error Resume Next
patt_gene_name = Split(PromKappa.seq_name_from_file.Caption, "(+)")(1)
patt_gene_name = Split(patt_gene_name, "_")(0)
GNX.Caption = patt_gene_name

L_R_Click
End Sub

Public Sub L_R_Click()

Call border_pattern

graf_m.Cls
Sxy = 0
Sx2 = 0
Sx = 0
Sy = 0
n = 0

Dim m As Double
Dim b As Double

Dim ax() As String
Dim ay() As String
Dim i As Variant
Dim X As Variant
Dim Y As Variant

Dim X1X As Long
Dim Y1Y As Long
Dim X2X As Long
Dim Y2Y As Long

Dim general_angle As Variant
Dim general_PI As Variant

ax() = Split(Sir_CG_regression, ",")
ay() = Split(Sir_IC_regression, ",")

    
    For i = 1 To UBound(ay()) - 1

    
        X = (graf_m.ScaleWidth / 100) * Val(ax(i))
        Y = (graf_m.ScaleHeight / 100) * Val(100 - ay(i))


        'graf_m.PSet (x, y)          'replotare sablon
        n = n + 1                    'n is the number of points
        Sx = Sx + X                  'Sx is the SUM of all X coordinates
        Sy = Sy + Y                  'Sy is the SUM of all Y coordinates
        Sxy = Sxy + X * Y            'Sxy is the SUM of all x*y
        Sx2 = Sx2 + X * X            'Sx2 is the SUM of all x*x
        If n > 1 Then
            If n > 2 Then graf_m.Line (X1, Y1)-(X2, Y2) 'erase old line
                'calculate new line using least squares
                'm and b define the line using y=m*x+b
                m = (n * Sxy - Sx * Sy) / (n * Sx2 - Sx * Sx)
                b = (Sy - m * Sx) / n
                Me.Caption = "Y = " & Round(m, 2) & " * X + " & Round(b, 2)
                CalcLine m, b, X1, Y1, X2, Y2
                graf_m.Line (X1, Y1)-(X2, Y2)
        End If


    Next i


best_fit.X1 = X1
best_fit.Y1 = Y1
best_fit.X2 = X2
best_fit.Y2 = Y2


X1X = X1
Y1Y = Y1
X2X = X2
Y2Y = Y2

general_angle = Round(Rad2Deg(AngleBetween(X1X, Y1Y, X2X, Y2Y)), 1)

If Val(general_angle) > 180 Then general_PI = Val(general_angle) - 180 Else general_PI = Val(general_angle)


angle_regression.Caption = "Angle 2PI: " & general_angle & "° | Angle PI: " & general_PI & "°"


Angle_360.Text = general_angle
Angle_180.Text = general_PI

Dim zero As Variant
If general_PI < 90 Then zero = general_PI
If general_PI > 90 Then zero = general_PI - 180
Angle_0.Text = Round(zero, 2)


'#####################################

Angle_RL (-general_angle + 270)

End Sub

Public Sub Angle_RL(AngleZero As Long)

 Dim Diameter As Integer
 Dim Radius As Integer
 Dim CCenterX As Integer
 Dim CCenterY As Integer
 Dim Radian As Variant
 Dim X As Variant
 Dim Y As Variant
 
 Diameter = 200 'Diameter in pixels
 Radius = 100 'Radius in pixels

CCenterX = (graf_m.ScaleWidth / 100) * Val(CG_mean_txt_m.Text)
CCenterY = graf_m.ScaleHeight - ((graf_m.ScaleHeight / 100) * Val(KIC_mean_txt_m.Text))

Radian = (PI / 180) * AngleZero ' Convert degrees to radians
X = Cos(Radian) * Radius + CCenterX 'Sets the x end point of the line
Y = Sin(Radian) * Radius + CCenterY 'Sets the y end point of the line

ZeroLine.X1 = CCenterX
ZeroLine.Y1 = CCenterY
ZeroLine.X2 = X
ZeroLine.Y2 = Y
End Sub



Function border_pattern()

Dim latura_a As Variant
Dim latura_b As Variant
Dim Surface_graph As Variant
Dim pattern_surface_in_pixels As Variant

graf_m.Picture = PromKappa.graf_point.Image

KIC_mean_txt_m.Text = Round(PromKappa.KIC_mean_txt.Text, 2)
CG_mean_txt_m.Text = Round(PromKappa.CG_mean_txt.Text, 2)

X_limin.Text = Round((100 / graf_m.ScaleWidth) * Last_x_lim_min, 2)
X_limax.Text = Round((100 / graf_m.ScaleWidth) * Last_x_lim_max, 2)

Y_limax.Text = Round(100 - ((100 / graf_m.ScaleHeight) * Last_y_lim_min), 2)
Y_limin.Text = Round(100 - ((100 / graf_m.ScaleHeight) * Last_y_lim_max), 2)

X_diff.Text = Round(Val(X_limax.Text) - Val(X_limin.Text), 2)
Y_diff.Text = Round(Val(Y_limax.Text) - Val(Y_limin.Text), 2)

'show pattern boundaries on x and y
Line1_x_min.X1 = Last_x_lim_min
Line1_x_min.X2 = Last_x_lim_min

Line2_x_max.X1 = Last_x_lim_max
Line2_x_max.X2 = Last_x_lim_max


Line1_y_min.Y1 = Last_y_lim_min
Line1_y_min.Y2 = Last_y_lim_min

Line2_y_max.Y1 = Last_y_lim_max
Line2_y_max.Y2 = Last_y_lim_max


'calculate total graph area in pixels
Surface_graph = Val(graf_m.ScaleHeight) * Val(graf_m.ScaleWidth)

'calculate pattern area in pixels
latura_a = Last_y_lim_max - Last_y_lim_min
latura_b = Last_x_lim_max - Last_x_lim_min
pattern_surface_in_pixels = (latura_a * latura_b)

'show pattern area in percentage
pattern_surface.Text = Round((100 / Surface_graph) * pattern_surface_in_pixels, 2)

End Function


Private Sub ok_Click()
Unload Me
End Sub


Private Sub CalcLine(ByVal m#, ByVal b#, X1#, Y1#, X2#, Y2#)
'given a line defined by m and b, calculates two points on that line
'that intersect the bounding box of (100,100)-(-100,-100)
Dim X#, Y#

Y = m * XMIN + b
If (Y >= YMIN) And (Y <= YMAX) Then
    X1 = XMIN
    Y1 = Y
Else
    If m > 0 Then
        Y1 = YMIN
    Else
        Y1 = YMAX
    End If
    X1 = (Y1 - b) / m
End If
Y = m * XMAX + b
If (Y >= YMIN) And (Y <= YMAX) Then
    X2 = XMAX
    Y2 = Y
Else
    If m > 0 Then
        Y2 = YMAX
    Else
        Y2 = YMIN
    End If
    X2 = (Y2 - b) / m
End If
End Sub


Public Sub Add_data_of_the_pattern_to_mem_Click()
patt_gene_name = Split(PromKappa.seq_name_from_file.Caption, "(+)")(1)
patt_gene_name = Split(patt_gene_name, "_")(0)

html = html & "<tr>" & vbCrLf
html = html & "<td>" & patt_gene_name & "</td>" & vbCrLf
html = html & "<td>" & KIC_mean_txt_m.Text & "</td>" & vbCrLf
html = html & "<td>" & CG_mean_txt_m.Text & "</td>" & vbCrLf
html = html & "<td>" & pattern_surface.Text & "</td>" & vbCrLf
html = html & "<td>" & Angle_0.Text & "</td>" & vbCrLf
html = html & "<td>" & Angle_360.Text & "</td>" & vbCrLf
html = html & "<td>" & Angle_180.Text & "</td>" & vbCrLf
html = html & "<td>" & Y_limax.Text & "</td>" & vbCrLf
html = html & "<td>" & Y_limin.Text & "</td>" & vbCrLf
html = html & "<td>" & X_limin.Text & "</td>" & vbCrLf
html = html & "<td>" & X_limax.Text & "</td>" & vbCrLf
html = html & "<td>" & Y_diff.Text & "</td>" & vbCrLf
html = html & "<td>" & X_diff.Text & "</td>" & vbCrLf
html = html & "</tr>" & vbCrLf

'MsgBox "Data for (" & patt_gene_name & ") has been added into memory"
End Sub

Public Sub Save_Data_HTML_Click()
Kill App.Path & "\results.htm"

    Dim intFileHandle As Integer
    intFileHandle = FreeFile
    
    Open App.Path & "\results.htm" For Append As #intFileHandle
    Print #intFileHandle, html & "</table></body></HTML>" & vbCrLf
    Close #intFileHandle

End Sub


Private Sub R_line_Click()
If best_fit.Visible = True Then
best_fit.Visible = False
Else
best_fit.Visible = True
End If
End Sub


Private Sub PS_lines_Click()
If Line2_y_max.Visible = True Then
Line2_y_max.Visible = False
Line2_x_max.Visible = False
Line1_y_min.Visible = False
Line1_x_min.Visible = False
Else
Line2_y_max.Visible = True
Line2_x_max.Visible = True
Line1_y_min.Visible = True
Line1_x_min.Visible = True
End If
End Sub

Private Sub E_line_Click()
If ZeroLine.Visible = True Then
ZeroLine.Visible = False
Else
ZeroLine.Visible = True
End If
End Sub
