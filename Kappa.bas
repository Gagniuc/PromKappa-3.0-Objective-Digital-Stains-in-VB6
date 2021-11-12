Attribute VB_Name = "Kappa"
Public total_val_x As Variant
Public total_val_y As Variant

Public Last_x_lim_min As Variant
Public Last_x_lim_max As Variant

Public Last_y_lim_min As Variant
Public Last_y_lim_max As Variant

Public Sir_IC_regression As String
Public Sir_CG_regression As String

Function IC(ByVal sequence As String) As Variant
    Dim count, i, u, total As Long
    Dim S1, s2 As String
    Dim max As Integer
    S1 = sequence
    max = Len(S1) - 1
    For u = 1 To max
        s2 = Mid(S1, u + 1)
        For i = 1 To Len(s2)
            If Mid(S1, i, 1) = Mid(s2, i, 1) Then
                count = count + 1
            End If
        Next i
        total = total + (count / Len(s2) * 100)
        count = 0
    Next u
    IC = Round((total / max), 2)
End Function




Function strand2(ByVal strand1 As String) As String

    For j = 1 To Len(strand1)
    
        nucleotida = LCase(Mid(strand1, j, 1))
        
        If nucleotida = "a" Then
            nucleotida = "t"
            GoTo 1
        End If
        
        
        If nucleotida = "t" Then
            nucleotida = "a"
            GoTo 1
        End If
        
        If nucleotida = "c" Then
            nucleotida = "g"
            GoTo 1
        End If
        
        If nucleotida = "g" Then
            nucleotida = "c"
            GoTo 1
        End If
        
1:
        fereastra_continut = fereastra_continut & nucleotida
    
    Next j
    
    strand2 = UCase(fereastra_continut)

End Function
