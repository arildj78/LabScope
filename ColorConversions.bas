Attribute VB_Name = "ColorConversions"
Option Explicit




'Formulas below is from http://www.easyrgb.com/en/math.php

'-------------------------------------
'|    Conversion LAB <--> XYZ        |
'-------------------------------------

Sub LAB2XYZ(ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double, ByRef x As Double, ByRef y As Double, ByRef z As Double)
Dim var_X As Double
Dim var_Y As Double
Dim var_Z As Double
    
    var_Y = (LabL + 16) / 116
    var_X = LabA / 500 + var_Y
    var_Z = var_Y - LabB / 200
    
    If (var_Y ^ 3 > 0.008856) Then var_Y = var_Y ^ 3 _
                              Else var_Y = (var_Y - 16 / 116) / 7.787
    If (var_X ^ 3 > 0.008856) Then var_X = var_X ^ 3 _
                              Else var_X = (var_X - 16 / 116) / 7.787
    If (var_Z ^ 3 > 0.008856) Then var_Z = var_Z ^ 3 _
                              Else var_Z = (var_Z - 16 / 116) / 7.787

    
    'Ref D50
    x = var_X * 96.422 ' Ref_X D50
    y = var_Y * 100#   ' Ref_Y D50
    z = var_Z * 82.521 ' Ref_Z D50
End Sub

Sub XYZ2LAB(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double)
    'Reference-X, Y and Z refer to specific illuminants and observers.
    'Common reference values are available below in this same page.
    
Dim var_X As Double
Dim var_Y As Double
Dim var_Z As Double
    
    'Ref D50
    var_X = x / 96.422
    var_Y = y / 100#
    var_Z = z / 82.521
    
    If (var_X > 0.008856) Then var_X = var_X ^ (1 / 3) _
                          Else var_X = (7.787 * var_X) + (16 / 116)
    If (var_Y > 0.008856) Then var_Y = var_Y ^ (1 / 3) _
                          Else var_Y = (7.787 * var_Y) + (16 / 116)
    If (var_Z > 0.008856) Then var_Z = var_Z ^ (1 / 3) _
                          Else var_Z = (7.787 * var_Z) + (16 / 116)
    
    LabL = (116 * var_Y) - 16
    LabA = 500 * (var_X - var_Y)
    LabB = 200 * (var_Y - var_Z)
End Sub





'-------------------------------------
'|    Conversion XYZ <--> sRGB       |
'-------------------------------------

Sub XYZ2sRGB(ByRef x As Double, ByRef y As Double, ByRef z As Double, ByRef sR As Double, ByRef sG As Double, ByRef sB As Double, Optional ByRef InGamut As Boolean)
    'X, Y and Z input refer to a D65/2° standard illuminant.
    'sR, sG and sB (standard RGB) output range = 0 ÷ 255
Dim var_X As Double
Dim var_Y As Double
Dim var_Z As Double
Dim var_R As Double
Dim var_G As Double
Dim var_B As Double
    
    var_X = x / 100
    var_Y = y / 100
    var_Z = z / 100
    
    
    'Bradford XYZ D50 to sRGB D65
    var_R = var_X * 3.1338561 + var_Y * (-1.6168667) + var_Z * (-0.4906146)
    var_G = var_X * (-0.9787684) + var_Y * 1.9161415 + var_Z * 0.033454
    var_B = var_X * 0.0719453 + var_Y * (-0.2289914) + var_Z * 1.4052427
    
    'Companding
    InGamut = Not ((var_R < 0#) Or (var_G < 0#) Or (var_B <= 0#)) ' No RGB values fall below 0

    If (var_R < 0) Then var_R = 0 Else If (var_R > 0.0031308) Then var_R = 1.055 * (var_R ^ (1 / 2.4)) - 0.055 _
                                       Else var_R = 12.92 * var_R
    If (var_G < 0) Then var_G = 0 Else If (var_G > 0.0031308) Then var_G = 1.055 * (var_G ^ (1 / 2.4)) - 0.055 _
                                       Else var_G = 12.92 * var_G
    If (var_B < 0) Then var_B = 0 Else If (var_B > 0.0031308) Then var_B = 1.055 * (var_B ^ (1 / 2.4)) - 0.055 _
                                       Else var_B = 12.92 * var_B

    var_R = var_R * 255
    var_G = var_G * 255
    var_B = var_B * 255
   
    InGamut = InGamut And (Not ((var_R > 255#) Or (var_G > 255#) Or (var_B > 255#))) 'No RGB values extend above 255
    
    sR = Min(var_R, 255)
    sG = Min(var_G, 255)
    sB = Min(var_B, 255)
End Sub

Sub sRGB2XYZ(ByVal sR As Double, ByVal sG As Double, ByVal sB As Double, ByRef x As Double, ByRef y As Double, ByRef z As Double)
    'sR, sG and sB (Standard RGB) input range = 0 ÷ 255
    'X, Y and Z output refer to a D50/2° standard illuminant.
Dim var_X As Double
Dim var_Y As Double
Dim var_Z As Double
Dim var_R As Double
Dim var_G As Double
Dim var_B As Double
    
    var_R = (sR / 255#)
    var_G = (sG / 255#)
    var_B = (sB / 255#)
    
    If (var_R > 0.04045) Then var_R = ((var_R + 0.055) / 1.055) ^ 2.4 _
                         Else var_R = var_R / 12.92
    If (var_G > 0.04045) Then var_G = ((var_G + 0.055) / 1.055) ^ 2.4 _
                         Else var_G = var_G / 12.92
    If (var_B > 0.04045) Then var_B = ((var_B + 0.055) / 1.055) ^ 2.4 _
                         Else var_B = var_B / 12.92
    
    
    Dim m(1 To 3, 1 To 3) As Double
    
    'Bradford sRGB D65 to XYZ D50
    m(1, 1) = 0.4360747
    m(1, 2) = 0.3850649
    m(1, 3) = 0.1430804
    
    m(2, 1) = 0.2225045
    m(2, 2) = 0.7168786
    m(2, 3) = 0.0606169
    
    m(3, 1) = 0.0139322
    m(3, 2) = 0.0971045
    m(3, 3) = 0.7141733
    
    
    var_X = var_R * m(1, 1) + var_G * m(1, 2) + var_B * m(1, 3)
    var_Y = var_R * m(2, 1) + var_G * m(2, 2) + var_B * m(2, 3)
    var_Z = var_R * m(3, 1) + var_G * m(3, 2) + var_B * m(3, 3)
    
    
    'XYZ in [0, 100]
    x = var_X * 100#
    y = var_Y * 100#
    z = var_Z * 100#
End Sub






'-------------------------------------
'|    Conversion XYZ <--> aRGB       |
'-------------------------------------



Sub XYZ2aRGB(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef aR As Double, ByRef aG As Double, ByRef ab As Double)
    'X, Y and Z input refer to a D50/2° standard illuminant.
    'aR, aG and aB (RGB Adobe 1998) output range = 0 ÷ 255
Dim var_X As Double
Dim var_Y As Double
Dim var_Z As Double
Dim var_R As Double
Dim var_G As Double
Dim var_B As Double
    var_X = x / 100#
    var_Y = y / 100#
    var_Z = z / 100#
    
    Dim m(1 To 3, 1 To 3) As Double
    
    'Bradford XYZ D50 to aRGB D65
    m(1, 1) = 1.9624274
    m(1, 2) = -0.6105343
    m(1, 3) = -0.3413404
    
    m(2, 1) = -0.9787684
    m(2, 2) = 1.9161415
    m(2, 3) = 0.033454
    
    m(3, 1) = 0.0286869
    m(3, 2) = -0.1406752
    m(3, 3) = 1.3487655
    
    var_R = var_X * m(1, 1) + var_Y * m(1, 2) + var_Z * m(1, 3)
    var_G = var_X * m(2, 1) + var_Y * m(2, 2) + var_Z * m(2, 3)
    var_B = var_X * m(3, 1) + var_Y * m(3, 2) + var_Z * m(3, 3)

    If var_R < 0 Then var_R = 0 Else var_R = var_R ^ (1 / 2.19921875)
    If var_G < 0 Then var_G = 0 Else var_G = var_G ^ (1 / 2.19921875)
    If var_B < 0 Then var_B = 0 Else var_B = var_B ^ (1 / 2.19921875)
    
    aR = var_R * 255#
    aG = var_G * 255#
    ab = var_B * 255#
End Sub

Sub aRGB2XYZ(ByVal aR As Double, ByVal aG As Double, ByVal ab As Double, ByRef x As Double, ByRef y As Double, ByRef z As Double)
    'aR, aG and aB (RGB Adobe 1998) input range = 0 ÷ 255
    'X, Y and Z output refer to a D50/2° standard illuminant.
Dim var_R As Double
Dim var_G As Double
Dim var_B As Double
    
    var_R = (aR / 255#)
    var_G = (aG / 255#)
    var_B = (ab / 255#)
    
    var_R = var_R ^ 2.19921875
    var_G = var_G ^ 2.19921875
    var_B = var_B ^ 2.19921875
    
    var_R = var_R * 100#
    var_G = var_G * 100#
    var_B = var_B * 100#
    
    Dim m(1 To 3, 1 To 3) As Double
    'Bradford XYZ D50 to aRGB D65
    m(1, 1) = 0.6097559
    m(1, 2) = 0.2052401
    m(1, 3) = 0.149224
    
    m(2, 1) = 0.3111242
    m(2, 2) = 0.625656
    m(2, 3) = 0.0632197
    
    m(3, 1) = 0.0194811
    m(3, 2) = 0.0608902
    m(3, 3) = 0.7448387
    
    
    x = var_R * m(1, 1) + var_G * m(1, 2) + var_B * m(1, 3)
    y = var_R * m(2, 1) + var_G * m(2, 2) + var_B * m(2, 3)
    z = var_R * m(3, 1) + var_G * m(3, 2) + var_B * m(3, 3)
End Sub









'Wrapper functions

'***************
'    sRGB
'***************
Public Sub LAB2sRGB(ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double, ByRef r As Double, ByRef g As Double, ByRef b As Double, Optional InGamut As Boolean)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    LAB2XYZ LabL, LabA, LabB, x, y, z
    XYZ2sRGB x, y, z, r, g, b, InGamut

End Sub
Public Sub sRGB2LAB(ByVal r As Double, ByVal g As Double, ByVal b As Double, ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    sRGB2XYZ r, g, b, x, y, z
    XYZ2LAB x, y, z, LabL, LabA, LabB

End Sub
Public Function InGamut_sRGB(LabL As Double, LabA As Double, LabB As Double) As Boolean
Dim var_X As Double
Dim var_Y As Double
Dim var_Z As Double
Dim var_R As Double
Dim var_G As Double
Dim var_B As Double
Dim lo As Boolean
Dim hi As Boolean

    var_Y = (LabL + 16) / 116
    var_X = LabA / 500 + var_Y
    var_Z = var_Y - LabB / 200
    

    If (var_Y ^ 3 > 0.008856) Then var_Y = var_Y ^ 3 _
                              Else var_Y = (var_Y - 16 / 116) / 7.787
    If (var_X ^ 3 > 0.008856) Then var_X = var_X ^ 3 _
                              Else var_X = (var_X - 16 / 116) / 7.787
    If (var_Z ^ 3 > 0.008856) Then var_Z = var_Z ^ 3 _
                              Else var_Z = (var_Z - 16 / 116) / 7.787

    
    'Ref D50
    var_X = var_X * 0.96422 ' Ref_X D50
   'var_Y = var_Y * 1#      ' Ref_Y D50
    var_Z = var_Z * 0.82521 ' Ref_Z D50
    
    
    'Bradford XYZ D50 to sRGB D65
    var_R = var_X * 3.1338561 + var_Y * (-1.6168667) + var_Z * (-0.4906146)
    var_G = var_X * (-0.9787684) + var_Y * 1.9161415 + var_Z * 0.033454
    var_B = var_X * 0.0719453 + var_Y * (-0.2289914) + var_Z * 1.4052427
    
    lo = (var_R < 0#) Or (var_G < 0#) Or (var_B <= 0#)    ' Out of gamut on low values
    
    'Companding
    If (var_R < 0) Then var_R = 0 Else If (var_R > 0.0031308) Then var_R = 1.055 * (var_R ^ (1 / 2.4)) - 0.055 _
                                       Else var_R = 12.92 * var_R
    If (var_G < 0) Then var_G = 0 Else If (var_G > 0.0031308) Then var_G = 1.055 * (var_G ^ (1 / 2.4)) - 0.055 _
                                       Else var_G = 12.92 * var_G
    If (var_B < 0) Then var_B = 0 Else If (var_B > 0.0031308) Then var_B = 1.055 * (var_B ^ (1 / 2.4)) - 0.055 _
                                       Else var_B = 12.92 * var_B

        
    hi = (var_R > 1#) Or (var_G > 1#) Or (var_B > 1#) 'Out of gamut on high end
    InGamut_sRGB = Not (hi Or lo)
End Function


'***************
'    aRGB
'***************
Public Sub LAB2aRGB(ByVal LabL As Double, ByVal LabA As Double, ByVal LabB As Double, ByRef r As Double, ByRef g As Double, ByRef b As Double)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    LAB2XYZ LabL, LabA, LabB, x, y, z
    XYZ2aRGB x, y, z, r, g, b

End Sub
Public Sub aRGB2LAB(ByVal r As Double, ByVal g As Double, ByVal b As Double, ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    aRGB2XYZ r, g, b, x, y, z
    XYZ2LAB x, y, z, LabL, LabA, LabB
End Sub
Public Function InGamut_aRGB(LabL As Double, LabA As Double, LabB As Double) As Boolean
Dim var_X As Double
Dim var_Y As Double
Dim var_Z As Double
Dim var_R As Double
Dim var_G As Double
Dim var_B As Double
Dim lo As Boolean
Dim hi As Boolean
    
    var_Y = (LabL + 16) / 116
    var_X = LabA / 500 + var_Y
    var_Z = var_Y - LabB / 200
    
    If (var_Y ^ 3 > 0.008856) Then var_Y = var_Y ^ 3 _
                              Else var_Y = (var_Y - 16 / 116) / 7.787
    If (var_X ^ 3 > 0.008856) Then var_X = var_X ^ 3 _
                              Else var_X = (var_X - 16 / 116) / 7.787
    If (var_Z ^ 3 > 0.008856) Then var_Z = var_Z ^ 3 _
                              Else var_Z = (var_Z - 16 / 116) / 7.787

    
    'Ref D50
    var_X = var_X * 0.96422 ' Ref_X D50
   'var_Y = var_Y * 1#      ' Ref_Y D50
    var_Z = var_Z * 0.82521 ' Ref_Z D50
    
    
    'Bradford XYZ D50 to aRGB D65
    var_R = var_X * 1.9624274 + var_Y * (-0.6105343) + var_Z * (-0.3413404)
    var_G = var_X * (-0.9787684) + var_Y * 1.9161415 + var_Z * 0.033454
    var_B = var_X * 0.0286869 + var_Y * (-0.1406752) + var_Z * 1.3487655

    lo = (var_R < 0#) Or (var_G < 0#) Or (var_B <= 0#)    ' Out of gamut on low values
    
    If var_R < 0 Then var_R = 0 Else var_R = var_R ^ (1 / 2.19921875)
    If var_G < 0 Then var_G = 0 Else var_G = var_G ^ (1 / 2.19921875)
    If var_B < 0 Then var_B = 0 Else var_B = var_B ^ (1 / 2.19921875)
    
    hi = (var_R > 1#) Or (var_G > 1#) Or (var_B > 1#) 'Out of gamut on high end
    InGamut_aRGB = Not (hi Or lo)
End Function





Sub CalcLAB()
Attribute CalcLAB.VB_ProcData.VB_Invoke_Func = "a\n14"
Dim r As Double
Dim g As Double
Dim b As Double
Dim LabL As Double
Dim LabA As Double
Dim LabB As Double

    r = Selection.Cells(1, 1).value
    g = Selection.Cells(1, 2).value
    b = Selection.Cells(1, 3).value
    
    sRGB2LAB r, g, b, LabL, LabA, LabB
    Selection.Cells(1, 3).offset(0, 1).Cells.value = LabL
    Selection.Cells(1, 3).offset(0, 2).Cells.value = LabA
    Selection.Cells(1, 3).offset(0, 3).Cells.value = LabB
    
End Sub

Sub CalcRGB()
Attribute CalcRGB.VB_ProcData.VB_Invoke_Func = " \n14"
Dim r As Double
Dim g As Double
Dim b As Double
Dim LabL As Double
Dim LabA As Double
Dim LabB As Double

    LabL = Selection.Cells(1, 1).value
    LabA = Selection.Cells(1, 2).value
    LabB = Selection.Cells(1, 3).value
    
    LAB2aRGB LabL, LabA, LabB, r, g, b
    Selection.Cells(1, 3).offset(0, 1).Cells.value = r
    Selection.Cells(1, 3).offset(0, 2).Cells.value = g
    Selection.Cells(1, 3).offset(0, 3).Cells.value = b
    
End Sub

Private Function Min(ParamArray values() As Variant) As Variant
   Dim minValue, value As Variant
   minValue = values(0)
   For Each value In values
       If value < minValue Then minValue = value
   Next
   Min = minValue
End Function

Private Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue, value As Variant
   maxValue = values(0)
   For Each value In values
       If value > maxValue Then maxValue = value
   Next
   Max = maxValue
End Function






Function DE00(L_1 As Double, a_1 As Double, b_1 As Double, L_2 As Double, a_2 As Double, b_2 As Double) As Double
Dim Brk1 As Double
Dim Brk2 As Double
Dim Brk3 As Double
Dim Brk4 As Double
Dim Brk5 As Double
    
    Dim Lmerk As Double
    Lmerk = (L_1 + L_2) / 2
    
    Dim C_1 As Double
    Dim C_2 As Double
    Dim Csnitt As Double
    C_1 = Sqr(a_1 ^ 2 + b_1 ^ 2)
    C_2 = Sqr(a_2 ^ 2 + b_2 ^ 2)
    Csnitt = (C_1 + C_2) / 2
    
    Dim g As Double
    g = 0.5 * (1 - Sqr(Csnitt ^ 7 / (Csnitt ^ 7 + 25 ^ 7)))
    
    Dim a_1merk As Double
    Dim a_2merk As Double
    a_1merk = a_1 * (1 + g)
    a_2merk = a_2 * (1 + g)
    
    Dim C_1merk As Double
    Dim C_2merk As Double
    C_1merk = Sqr(a_1merk ^ 2 + b_1 ^ 2)
    C_2merk = Sqr(a_2merk ^ 2 + b_2 ^ 2)
    
    Dim Csnittmerk As Double
    Csnittmerk = (C_1merk + C_2merk) / 2
    
    Dim h_1merk As Double
    Dim h_2merk As Double
    
    If a_1merk <> 0 And b_1 <> 0 Then h_1merk = r2d(WorksheetFunction.Atan2(a_1merk, b_1)) _
                                 Else h_1merk = 0
    If h_1merk < 0 Then h_1merk = h_1merk + 360
    If a_2merk <> 0 And b_2 <> 0 Then h_2merk = r2d(WorksheetFunction.Atan2(a_2merk, b_2)) _
                                 Else h_2merk = 0
    If h_2merk < 0 Then h_2merk = h_2merk + 360
    
    Dim Hsnittmerk As Double
    If Abs(h_1merk - h_2merk) > 180 _
        Then Hsnittmerk = (h_1merk + h_2merk + 360) / 2 _
        Else Hsnittmerk = (h_1merk + h_2merk) / 2
        
    Dim T As Double
    T = 1 - 0.17 * Cos(d2r(Hsnittmerk - 30)) + 0.24 * Cos(d2r(2 * Hsnittmerk)) + 0.32 * Cos(d2r(3 * Hsnittmerk + 6)) - 0.2 * Cos(d2r(4 * Hsnittmerk - 63))
    
    Dim dhmerk As Double
    If Abs(h_2merk - h_1merk) <= 180 Then
        dhmerk = h_2merk - h_1merk
    ElseIf (Abs(h_2merk - h_1merk) > 180) And (h_2merk <= h_1merk) Then
        dhmerk = h_2merk - h_1merk + 360
    Else
        dhmerk = h_2merk - h_1merk - 360
    End If
    
    Dim dLmerk As Double
    dLmerk = L_2 - L_1
    
    Dim dCmerk As Double
    dCmerk = C_2merk - C_1merk
    
    Dim dUcHmerk As Double
    dUcHmerk = 2 * Sqr(C_1merk * C_2merk) * Sin(d2r(dhmerk / 2))
    
    Dim S_L As Double
    S_L = 1 + (0.015 * (Lmerk - 50) ^ 2) / Sqr(20 + (Lmerk - 50) ^ 2)
    
    Dim S_C As Double
    S_C = 1 + 0.045 * Csnittmerk
    
    Dim S_H As Double
    S_H = 1 + 0.015 * Csnittmerk * T
    
    
    Dim dRho As Double
    dRho = 30 * Exp(-((Hsnittmerk - 275) / 25) ^ 2)
    
    Dim R_C As Double
    R_C = 2 * Sqr(Csnittmerk ^ 7 / (Csnittmerk ^ 7 + 25 ^ 7))
        
    Dim R_T As Double
    R_T = -R_C * Sin(d2r(2 * dRho))
    
    Dim K_L As Double
    K_L = 1
    
    Dim K_C As Double
    K_C = 1
    
    Dim K_H As Double
    K_H = 1
    
    
    Brk1 = dLmerk / (K_L * S_L)
    Brk2 = dCmerk / (K_C * S_C)
    Brk3 = dUcHmerk / (K_H * S_H)
    Brk4 = dCmerk / (K_C * S_C)
    Brk5 = dUcHmerk / (K_H * S_H)
    DE00 = Sqr(Brk1 ^ 2 + Brk2 ^ 2 + Brk3 ^ 2 + R_T * Brk4 * Brk5)
End Function

Function d2r(d As Double) As Double
    'Degrees to radians
    d2r = 2 * WorksheetFunction.Pi * d / 360
End Function
Function r2d(r As Double) As Double
    'Radians to degrees
    r2d = 360 * r / (2 * WorksheetFunction.Pi)
End Function




'ICC
Sub asdf()
    Dim ICCprofile As String
    Dim filenumber As Integer
    Dim bytTemp As Byte
    Dim wrdTemp As Integer
    Dim arrTemp(1 To 25) As Byte
        ICCprofile = "C:\Windows\System32\spool\drivers\color\HP DesignJet Z6200ps 42in Photo_PolyPropylene Banner 200 - HP Matte Polypropylene.icc"
        
        filenumber = FreeFile
        Open ICCprofile For Binary Access Read As filenumber
        
        GetBigE filenumber, 2, wrdTemp
        GetBigE filenumber, 3, wrdTemp
        GetBigE filenumber, 1, wrdTemp
        GetBigE filenumber, 1, wrdTemp
        GetBigE filenumber, 1, wrdTemp
        GetBigE filenumber, 1, wrdTemp
        
        Close filenumber
        

End Sub


Function GetBigE(ByRef filenumber As Integer, ByVal recnumber As Long, ByRef varname As Integer) As Byte
Dim LittleEByte() As Byte
Dim BigEByte() As Byte
Dim n As Integer
Dim length As Integer
    length = Len(varname)
    ReDim LittleEByte(1 To length) As Byte
    ReDim BigEByte(1 To length) As Byte
    Get filenumber, recnumber, LittleEByte
    
    For n = 1 To length
        BigEByte(length - n + 1) = LittleEByte(n)
    Next n

    Mem_Copy varname, BigEByte(1), length
End Function




