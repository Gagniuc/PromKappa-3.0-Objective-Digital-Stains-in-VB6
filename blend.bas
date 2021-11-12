Attribute VB_Name = "Blend"
Sub LongToRGB(RGBcol As Long, ByRef R As Single, ByRef G As Single, ByRef B As Single)

    R = RGBcol And &HFF                     ' set red
    G = (RGBcol And &H100FF00) / &H100      ' set green
    B = (RGBcol And &HFF0000) / &H10000     ' set blue

    If G > 255 Then Stop

End Sub


Function BlendColors(C1 As Long, C2 As Long, P As Single) As Long
    Dim R1             As Single
    Dim G1             As Single
    Dim B1             As Single

    Dim R2             As Single
    Dim G2             As Single
    Dim B2             As Single

    Dim P2             As Single
    P2 = 1 - P

    If C1 < 0 Then C1 = 0
    If C2 < 0 Then C2 = 0


    LongToRGB C1, R1, G1, B1
    LongToRGB C2, R2, G2, B2

    R1 = R1 * P + R2 * P2
    G1 = G1 * P + G2 * P2
    B1 = B1 * P + B2 * P2

    BlendColors = RGB(R1, G1, B1)

End Function

Function BlendColorsRGB(R1, G1, B1, R2, G2, B2, P As Single) As Long
    Dim R              As Single
    Dim G              As Single
    Dim B              As Single

    Dim P2             As Single
    P2 = 1 - P


    R = R1 * P + R2 * P2
    G = G1 * P + G2 * P2
    B = B1 * P + B2 * P2

    BlendColorsRGB = RGB(R, G, B)

End Function
