Attribute VB_Name = "CurveFunctions"
Option Explicit

'Some essential constants
Private Const PI As Double = 3.14159265358979
Private Const PIM2 As Double = 2 * PI
Private Const PID2 As Double = PI / 2
Private Const PI32 As Double = PI * (3 / 2)
Private Const PI180 As Double = PI / 180

Public html As String

'Conversion from radians to degress
Public Function Rad2Deg(dblRad As Double) As Single
    Rad2Deg = dblRad / PI180
End Function

'This function finds angle between to points (returned value is in radians)
Public Function AngleBetween(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Double
    Dim X As Single, Y As Single
    X = X2 - X1
    Y = Y2 - Y1
    
    If Y = 0 Then
        If X1 < X2 Then
            AngleBetween = 0
        Else
            AngleBetween = PI
        End If
    Else
        If Y < 0 Then
            AngleBetween = Atn(X / Y) + PID2
        Else
            AngleBetween = Atn(X / Y) + PI32
        End If
        
    End If

End Function



