Attribute VB_Name = "ColorManagement"
Option Explicit

'**********************************************'
'                                              '
'     Read different datatypes from buffer     '
'                                              '
'**********************************************'
'-------------------------------------
'            BASIC TYPES             '
'-------------------------------------
Function GetBig_uInt32Number(Buffer() As Byte, offset As Long) As Long
    'Returns the unsigned 32 bit integer stored at offset in 0 based buffer.
    'Return value is limited to is [0, 2^31-1] due to Long being a signed
    '32 bit data type. The data is stored in Big Endian.
    
    If Buffer(offset) >= &H80 Then Err.Raise 13, "GetBig_uInt32Number", "GetBig_uInt32Number: Input value to large for Long."
        
    GetBig_uInt32Number = Buffer(offset) * &H1000000 Or _
                          Buffer(offset + 1) * &H10000 Or _
                          Buffer(offset + 2) * &H100& Or _
                          Buffer(offset + 3)
End Function
Function GetBig_uInt16Number(Buffer() As Byte, offset As Long) As Long
    GetBig_uInt16Number = Buffer(offset) * &H100& Or _
                          Buffer(offset + 1)
End Function
Function GetBig_uInt8Number(Buffer() As Byte, offset As Long) As Integer
    GetBig_uInt8Number = Buffer(offset)
End Function


'-------------------------------------
'            ARRAYS                  '
'-------------------------------------
Sub Get_Struct(Buffer() As Byte, offset As Long, output() As Byte)
Dim n As Long
Dim lo As Long
    lo = LBound(output)
    
    For n = lo To UBound(output)
        output(n) = Buffer(offset + n - lo)
    Next n
End Sub



'-------------------------------------
'            STRUCTURES              '
'-------------------------------------
Function ReadType_curveType(tagData() As Byte, offset As Long) As CurveType
'   Structure is defined in
'   ICC v4.3 specification
'   section 10.5
Dim n As Long

    With ReadType_curveType
        .signature = GetBig_uInt32Number(tagData, offset + 0)
        
        If .signature <> &H63757276 Then Err.Raise 11, "ReadType_curveType", "curveType signature not found."
        
        .n = GetBig_uInt32Number(tagData, offset + 8)
        
        ReDim .curve(0 To .n - 1) As Long
        
        For n = 0 To .n - 1
            .curve(n) = GetBig_uInt16Number(tagData, offset + 12 + 2 * n)
        Next n
        
        .FunctionType = FunctionType.curve    'Implementation specific, not part of ICC standard.
        .size = 12 + 2 * .n
        .size = 4 * (Int(.size / 4) - (.size / 4 - Int(.size / 4) > 0)) 'Round up to 4 byte boundary
    End With
End Function


Function ReadType_parametricCurveType(tagData() As Byte, offset As Long) As CurveType
'   Structure is defined in
'   ICC v4.3 specification
'   section 10.16
Dim n As Long

    With ReadType_parametricCurveType
        .signature = GetBig_uInt32Number(tagData, offset + 0)
        
        If .signature <> &H70617261 Then Err.Raise 11, "ReadType_parametricCurveType", "parametricCurveType signature not found."
        
        .FunctionType = GetBig_uInt32Number(tagData, offset + 8)
        
        .g = GetBig_uInt32Number(tagData, offset + 12)
        .size = 16
        
        If .FunctionType >= FunctionType.PARAMgab Then
            .a = GetBig_uInt32Number(tagData, offset + 16)
            .b = GetBig_uInt32Number(tagData, offset + 20)
            .size = 24
        End If
        If .FunctionType >= FunctionType.PARAMgabc Then
            .c = GetBig_uInt32Number(tagData, offset + 24)
            .size = 28
        End If
        If .FunctionType >= FunctionType.PARAMgabcd Then
            .d = GetBig_uInt32Number(tagData, offset + 28)
            .size = 32
        End If
        If .FunctionType >= FunctionType.PARAMgabcdef Then
            .e = GetBig_uInt32Number(tagData, offset + 32)
            .f = GetBig_uInt32Number(tagData, offset + 36)
            .size = 40
        End If
        .size = 4 * (Int(.size / 4) - (.size / 4 - Int(.size / 4) > 0)) 'Round up to 4 byte boundary
    End With
End Function



Function ReadType_lut8Type(tagData() As Byte, offset As Long) As lutBToAType
'   Structure is defined in
'   ICC v4.3 specification
'   section 10.9
Dim n As Long
Dim i As Long
'Dim w As WORD
Dim inLo As Long
Dim inHi As Long
Dim CLUTLo As Long
Dim CLUTHi As Long
Dim outLo As Long
Dim outHi As Long

    'lut16type and lut8type data structures is equal in this implementation. lut8 does not use the .m and .n parameter
    
    With ReadType_lut8Type
        .signature = GetBig_uInt32Number(tagData, offset + 0)
        
        If .signature <> &H6D667431 Then Err.Raise 11, "ReadType_lut8Type", "lut8Type signature not found."
        
        .i = GetBig_uInt8Number(tagData, offset + 8)
        .o = GetBig_uInt8Number(tagData, offset + 9)
        .g = GetBig_uInt8Number(tagData, offset + 10)
        
        
        '*****************************
        '    Process Matrix
        '*****************************
        .Matrix(1) = GetBig_uInt32Number(tagData, offset + 12)
        .Matrix(2) = GetBig_uInt32Number(tagData, offset + 16)
        .Matrix(3) = GetBig_uInt32Number(tagData, offset + 20)
        .Matrix(4) = GetBig_uInt32Number(tagData, offset + 24)
        .Matrix(5) = GetBig_uInt32Number(tagData, offset + 28)
        .Matrix(6) = GetBig_uInt32Number(tagData, offset + 32)
        .Matrix(7) = GetBig_uInt32Number(tagData, offset + 36)
        .Matrix(8) = GetBig_uInt32Number(tagData, offset + 40)
        .Matrix(9) = GetBig_uInt32Number(tagData, offset + 44)
        
        
        'Prepare for the curves
        inLo = 48
        inHi = 47 + 256 * .i
        CLUTLo = inHi + 1
        CLUTHi = inHi + .o * .g ^ .i
        outLo = CLUTHi + 1
        outHi = CLUTHi + 256 * .o

        
        
        '*****************************
        '    Process Mcurve
        '*****************************
        'lut8Type does not have Mcurve
        
        
        
        '*****************************
        '    Process Bcurve
        '*****************************
        ReDim .Bcurve(0 To .i - 1) As CurveType 'InTables
        .n = 256                                'Entries in InTables
        For n = 0 To (.i - 1) 'Cycle through the input curves
            ReDim .Bcurve(n).curve(0 To .n - 1) As Long 'Redim array to hold each curve. Source data is byte, stored as long
            .Bcurve(n).n = .n
            .Bcurve(n).FunctionType = FunctionType.lut16Table
            For i = 0 To .n - 1
                .Bcurve(n).curve(i) = 257& * GetBig_uInt8Number(tagData, offset + inLo + .n * n + i)
            Next i
        Next n
        
        
        '*****************************
        '    Process CLUT
        '*****************************
        ReDim .CLUT(0 To .o * .g ^ .i - 1) As Long
        .CLUTchannels = .i
        For n = 0 To .i - 1
            .CLUTgridPoints(n) = .g
        Next n
        For n = LBound(.CLUT) To UBound(.CLUT)
            .CLUT(n) = 257& * GetBig_uInt8Number(tagData, offset + CLUTLo + n)
        Next n
        

        '*****************************
        '    Process Acurve
        '*****************************
        ReDim .Acurve(0 To .o - 1) As CurveType 'OutTables
        .m = 256                                'Entries in outTables
        For n = 0 To (.o - 1) 'Cycle through the output curves
            ReDim .Acurve(n).curve(0 To .m - 1) As Long 'Redim array to hold each curve. Source data is byte, stored as long
            .Acurve(n).n = .n
            .Acurve(n).FunctionType = FunctionType.lut16Table
            For i = 0 To .m - 1
                .Acurve(n).curve(i) = 257& * GetBig_uInt8Number(tagData, offset + outLo + .m * n + i)
            Next i
        Next n
        
        .legacyPCS = False
    
    End With
End Function

Function ReadType_lut16Type(tagData() As Byte, offset As Long) As lutBToAType
'   Structure is defined in
'   ICC v4.3 specification
'   section 10.8
Dim n As Long
Dim i As Long
Dim inLo As Long
Dim inHi As Long
Dim CLUTLo As Long
Dim CLUTHi As Long
Dim outLo As Long
Dim outHi As Long

    With ReadType_lut16Type
        .signature = GetBig_uInt32Number(tagData, offset + 0)
        
        If .signature <> &H6D667432 Then Err.Raise 11, "ReadType_lut16Type", "lut16Type signature not found."
        
        .i = GetBig_uInt8Number(tagData, offset + 8)
        .o = GetBig_uInt8Number(tagData, offset + 9)
        .g = GetBig_uInt8Number(tagData, offset + 10)
        .Matrix(1) = GetBig_uInt32Number(tagData, offset + 12)
        .Matrix(2) = GetBig_uInt32Number(tagData, offset + 16)
        .Matrix(3) = GetBig_uInt32Number(tagData, offset + 20)
        .Matrix(4) = GetBig_uInt32Number(tagData, offset + 24)
        .Matrix(5) = GetBig_uInt32Number(tagData, offset + 28)
        .Matrix(6) = GetBig_uInt32Number(tagData, offset + 32)
        .Matrix(7) = GetBig_uInt32Number(tagData, offset + 36)
        .Matrix(8) = GetBig_uInt32Number(tagData, offset + 40)
        .Matrix(9) = GetBig_uInt32Number(tagData, offset + 44)
        
        .n = GetBig_uInt16Number(tagData, offset + 48)
        .m = GetBig_uInt16Number(tagData, offset + 50)
        
        inLo = 52
        inHi = 51 + 2 * .n * .i
        CLUTLo = inHi + 1
        CLUTHi = inHi + 2 * .o * .g ^ .i
        outLo = CLUTHi + 1
        outHi = CLUTHi + 2 * .m * .o
        
        
        
        '*****************************
        '    Process Bcurve
        '*****************************
        ReDim .Bcurve(0 To .i - 1) As CurveType 'InTables
        For n = 0 To (.i - 1)   'Cycle through the input curves
            ReDim .Bcurve(n).curve(0 To .n - 1) As Long 'Redim array to hold each curve. Source data is byte, stored as long
            .Bcurve(n).n = .n
            .Bcurve(n).FunctionType = FunctionType.lut16Table
            For i = 0 To .n - 1
                .Bcurve(n).curve(i) = GetBig_uInt16Number(tagData, offset + inLo + 2 * (.n * n + i))
            Next i
        Next n
        
        
        '*****************************
        '    Process CLUT
        '*****************************
        ReDim .CLUT(0 To .o * .g ^ .i - 1) As Long
        .CLUTchannels = .i
        For n = 0 To .i - 1
            .CLUTgridPoints(n) = .g
        Next n
        For n = LBound(.CLUT) To UBound(.CLUT)
            .CLUT(n) = GetBig_uInt16Number(tagData, offset + CLUTLo + 2 * n)
        Next n
        
        

        '*****************************
        '    Process Acurve
        '*****************************
        ReDim .Acurve(0 To .o - 1) As CurveType 'OutTables
        For n = 0 To (.o - 1) 'Cycle through the output curves
            ReDim .Acurve(n).curve(0 To .m - 1) As Long 'Redim array to hold each curve. Source data is byte, stored as long
            .Acurve(n).n = .m
            .Acurve(n).FunctionType = FunctionType.lut16Table
            For i = 0 To .m - 1
                .Acurve(n).curve(i) = GetBig_uInt16Number(tagData, offset + outLo + 2 * (.m * n + i))
            Next i
        Next n
        
        .legacyPCS = True 'Tables 39 and 40 in the standard
    
    End With
End Function

Function ReadType_lutBToAType(tagData() As Byte, offset As Long) As lutBToAType
'   Structure is defined in
'   ICC v4.3 specification
'   section 10.10
Dim readBytes As Long
Dim n As Long
    With ReadType_lutBToAType
        
        .signature = GetBig_uInt32Number(tagData, offset + 0)
        
        If .signature <> &H6D424120 Then Err.Raise 11, "ReadType_lutBToAType", "lutBToAType signature not found."
        
        .i = GetBig_uInt8Number(tagData, offset + 8)
        .o = GetBig_uInt8Number(tagData, offset + 9)
        .offsetBcurve = GetBig_uInt32Number(tagData, offset + 12)
        .offsetMatrix = GetBig_uInt32Number(tagData, offset + 16)
        .offsetMcurve = GetBig_uInt32Number(tagData, offset + 20)
        .offsetCLUT = GetBig_uInt32Number(tagData, offset + 24)
        .offsetAcurve = GetBig_uInt32Number(tagData, offset + 28)
            
        '*****************************
        '    Process Bcurve
        '*****************************
        ReDim .Bcurve(0 To .i - 1) As CurveType
        readBytes = 0
        
        If .offsetBcurve <> 0 Then
            For n = LBound(.Bcurve) To UBound(.Bcurve)
                Select Case GetBig_uInt32Number(tagData, offset + .offsetBcurve)
                    Case Is = &H63757276: .Bcurve(n) = ReadType_curveType(tagData, offset + .offsetBcurve + readBytes)
                    Case Is = &H70617261: .Bcurve(n) = ReadType_parametricCurveType(tagData, offset + .offsetBcurve + readBytes)
                    Case Else: Err.Raise 11, "ReadType_lutBToAType", "Bcurve tag not recognized."
                End Select
                readBytes = readBytes + .Bcurve(n).size
            Next n
        End If
    
        
        '*****************************
        '    Process Mcurve
        '*****************************
        ReDim .Mcurve(0 To .i - 1) As CurveType
        readBytes = 0
        
        If .offsetMcurve <> 0 Then
            For n = LBound(.Mcurve) To UBound(.Mcurve)
                Select Case GetBig_uInt32Number(tagData, offset + .offsetMcurve)
                    Case Is = &H63757276: .Mcurve(n) = ReadType_curveType(tagData, offset + .offsetMcurve + readBytes)
                    Case Is = &H70617261: .Mcurve(n) = ReadType_parametricCurveType(tagData, offset + .offsetMcurve + readBytes)
                    Case Else: Err.Raise 11, "ReadType_lutBToAType", "Mcurve tag not recognized."
                End Select
                readBytes = readBytes + .Mcurve(n).size
            Next n
        End If
        
        '*****************************
        '    Process Acurve
        '*****************************
        ReDim .Acurve(0 To .o - 1) As CurveType
        readBytes = 0
        
        If .offsetAcurve <> 0 Then
            For n = LBound(.Acurve) To UBound(.Acurve)
                Select Case GetBig_uInt32Number(tagData, offset + .offsetAcurve)
                    Case Is = &H63757276: .Acurve(n) = ReadType_curveType(tagData, offset + .offsetAcurve + readBytes)
                    Case Is = &H70617261: .Acurve(n) = ReadType_parametricCurveType(tagData, offset + .offsetAcurve + readBytes)
                    Case Else: Err.Raise 11, "ReadType_lutBToAType", "Acurve tag not recognized."
                End Select
                readBytes = readBytes + .Acurve(n).size
            Next n
        End If
        
        '*****************************
        '    Process Matrix
        '*****************************
        If .offsetMatrix <> 0 Then
            For n = 1 To 12
                .Matrix(n) = GetBig_uInt32Number(tagData, offset + .offsetMatrix + (n - 1) * 4)
            Next n
        End If
        
        
        '*****************************
        '    Process CLUT
        '*****************************
        If .offsetCLUT <> 0 Then
            'Find the number of values in each of the 16 possible dimensions
            For n = LBound(.CLUTgridPoints) To UBound(.CLUTgridPoints)
                .CLUTgridPoints(n) = GetBig_uInt8Number(tagData, offset + .offsetCLUT + n)
            Next n
            
            'Look up table is either lut8Type or lut16Type
            .CLUTbitCount = GetBig_uInt8Number(tagData, offset + .offsetCLUT + 16) * 8
            
            'The number of entries in CLUT (nGridPoints) are equal to all the
            'dimensions multiplied with eachother (excluding the 'zero' entries)
            Dim nGridPoints As Long
            nGridPoints = 1
            While .CLUTgridPoints(.CLUTchannels) > 0
                nGridPoints = nGridPoints * .CLUTgridPoints(.CLUTchannels)
                .CLUTchannels = .CLUTchannels + 1
            Wend
            
            ReDim .CLUT(0 To nGridPoints * .o - 1) As Long
            
            If .CLUTbitCount = 8 Then
                '8 bit CLUT
                For n = 0 To nGridPoints * .o - 1
                    .CLUT(n) = 257& * GetBig_uInt8Number(tagData, offset + .offsetCLUT + 20 + 2)
                Next n
            Else
                '16 bit CLUT
                For n = 0 To nGridPoints * .o - 1
                    .CLUT(n) = GetBig_uInt16Number(tagData, offset + .offsetCLUT + 20 + 2 * n)
                Next n
            End If
        End If
        
        
        
        .legacyPCS = False

    End With
End Function



'************************************'
'                                    '
'       Transform                    '
'                                    '
'************************************'
Public Function TransformLab2Tag(transform As lutBToAType, LabL As Double, LabA As Double, LabB As Double) As RGBCOLOR
Dim i As Long
Dim k As Long
Dim j As Long
Dim pt(0 To 1, 0 To 1, 0 To 1) As LabCOLOR
Dim dLookUpL As Double
Dim lLookUpL() As Long 'L to be looked up. Two values - one below and one above
Dim dLookUpA As Double
Dim lLookUpA() As Long 'A to be looked up. Two values - one below and one above
Dim dLookUpB As Double
Dim lLookUpB() As Long 'B to be looked up. Two values - one below and one above

Dim varL As Double
Dim varA As Double
Dim varB As Double
Dim delta(0 To 2) As Long
Dim target(0 To 2) As Double

    '*****************************
    ' Input conversion
    '*****************************
    If transform.legacyPCS Then
'    If False Then
        'Use legacy PCSLAB encoding
        dLookUpL = LabL / (100# + 25500 / 65280)
        dLookUpA = (LabA + 128#) / (255# + 255 / 256)
        dLookUpB = (LabB + 128#) / (255# + 255 / 256)
        'dLookUpL = LabL / (100# + 25500 / 65280)
        'dLookUpA = (LabA + 128#) / 256#
        'dLookUpB = (LabB + 128#) / 256#
    Else
        'Use ICC v4.3 PCSLAB encoding
        dLookUpL = LabL / (100#)
        dLookUpA = (LabA + 128) / 255
        dLookUpB = (LabB + 128) / 255
    End If
    
    
    
    '*****************************
    ' Bcurve
    '*****************************
    OverUnderReal dLookUpL, transform.Bcurve(0).n, lLookUpL, target(0)
    OverUnderReal dLookUpA, transform.Bcurve(1).n, lLookUpA, target(1)
    OverUnderReal dLookUpB, transform.Bcurve(2).n, lLookUpB, target(2)
    
    delta(0) = transform.Bcurve(0).curve(lLookUpL(1)) - transform.Bcurve(0).curve(lLookUpL(0))
    delta(1) = transform.Bcurve(1).curve(lLookUpA(1)) - transform.Bcurve(1).curve(lLookUpA(0))
    delta(2) = transform.Bcurve(2).curve(lLookUpB(1)) - transform.Bcurve(2).curve(lLookUpB(0))
    
    varL = transform.Bcurve(0).curve(lLookUpL(0)) + target(0) * delta(0)
    varA = transform.Bcurve(1).curve(lLookUpA(0)) + target(1) * delta(1)
    varB = transform.Bcurve(2).curve(lLookUpB(0)) + target(2) * delta(2)
    
    varL = varL / 65535
    varA = varA / 65535
    varB = varB / 65535
    
    
    
    '*****************************
    ' CLUT
    '*****************************
    OverUnderReal varL, transform.CLUTgridPoints(0), lLookUpL, target(0)
    OverUnderReal varA, transform.CLUTgridPoints(1), lLookUpA, target(1)
    OverUnderReal varB, transform.CLUTgridPoints(2), lLookUpB, target(2)
    
    For i = 0 To 1
        For k = 0 To 1
            For j = 0 To 1
                pt(i, k, j).L = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, lLookUpL(i), lLookUpA(k), lLookUpB(j)) + 0)
                pt(i, k, j).a = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, lLookUpL(i), lLookUpA(k), lLookUpB(j)) + 1)
                pt(i, k, j).b = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, lLookUpL(i), lLookUpA(k), lLookUpB(j)) + 2)
                'Debug.Print i; j; k; pt(i, k, j).L
                Debug.Print pt(i, k, j).L; pt(i, k, j).a; pt(i, k, j).b
            Next j
        Next k
    Next i
    
    Dim CLUTresult As LabCOLOR
    CLUTresult = (TriLinInterpolation(pt, target))
    
    varL = CLUTresult.L / 65535
    varA = CLUTresult.a / 65535
    varB = CLUTresult.b / 65535
    
    
    '*****************************
    ' Acurve
    '*****************************
    OverUnderReal varL, transform.Acurve(0).n, lLookUpL, target(0)
    OverUnderReal varA, transform.Acurve(1).n, lLookUpA, target(1)
    OverUnderReal varB, transform.Acurve(2).n, lLookUpB, target(2)
    
    delta(0) = transform.Acurve(0).curve(lLookUpL(1)) - transform.Acurve(0).curve(lLookUpL(0))
    delta(1) = transform.Acurve(1).curve(lLookUpA(1)) - transform.Acurve(1).curve(lLookUpA(0))
    delta(2) = transform.Acurve(2).curve(lLookUpB(1)) - transform.Acurve(2).curve(lLookUpB(0))
    
    varL = transform.Acurve(0).curve(lLookUpL(0)) + target(0) * delta(0)
    varA = transform.Acurve(1).curve(lLookUpA(0)) + target(1) * delta(1)
    varB = transform.Acurve(2).curve(lLookUpB(0)) + target(2) * delta(2)
    
    varL = varL / 65535
    varA = varA / 65535
    varB = varB / 65535
    
    'TransformLab2Tag.red = Int(varL * 255)
    'TransformLab2Tag.green = Int(varA * 255)      JIPPIIII                             JIPPIIII
    'TransformLab2Tag.blue = Int(varB * 255)       'LAB relativ       LAB abs          'iPro2 LAB rel     'iPro2 LAB abs
    TransformLab2Tag.red = Round(varL * 255, 0)    '.5489             .5813            .5772               .5755
    TransformLab2Tag.green = Round(varA * 255, 0)  '.6510             .6715            .6418               .6610
    TransformLab2Tag.blue = Round(varB * 255, 0)   '.1480             .1660            .1586               .1334
End Function

Sub testtesttest()
Dim aaaaa As Variant
Dim RGB As RGBCOLOR
Dim b2a1 As lutBToAType
Dim gamt As lutBToAType
    Get_B2A1_gamt "C:\Windows\System32\spool\drivers\color\HP DesignJet Z6200ps 42in Photo_PolyPropylene Banner 200 - HP Matte Polypropylene.icc", b2a1, gamt
    'Get_B2A1_gamt "C:\Windows\System32\spool\drivers\color\z6200 PolyPropylene 200 M1 for D50 i1Pro2.icc", b2a1, gamt
    
    RGB = TransformLab2Tag(b2a1, 60, -10, 30)
    
    Debug.Print RGB.red; RGB.green; RGB.blue
End Sub

'************************************'
'                                    '
'       Conversion                   '
'                                    '
'************************************'
Function LongToChars(FourBytes As Long) As String
Dim b(0 To 3) As Byte
    
    b(3) = (FourBytes And &HFF000000) / &H1000000
    b(2) = (FourBytes And &HFF0000) / &H10000
    b(1) = (FourBytes And &HFF00&) / &H100
    b(0) = (FourBytes And &HFF&)
    
    LongToChars = Chr(b(3)) & Chr(b(2)) & Chr(b(1)) & Chr(b(0))
End Function

Function GetArrayAddress(GridPoints() As Byte, ParamArray input0to15() As Variant) As Long
Dim n As Integer
Dim dimensions As Long
    dimensions = UBound(input0to15)
    For n = 0 To dimensions
        GetArrayAddress = GetArrayAddress + input0to15(dimensions - n) * GridPoints(dimensions - n) ^ n
    Next n
    GetArrayAddress = GetArrayAddress * (dimensions + 1)
End Function

Function TriLinInterpolation(pt() As LabCOLOR, delta() As Double) As LabCOLOR
Dim c000 As Long
Dim c001 As Long
Dim c010 As Long
Dim c011 As Long
Dim c100 As Long
Dim c101 As Long
Dim c110 As Long
Dim c111 As Long

Dim c00 As Double
Dim c01 As Double
Dim c10 As Double
Dim c11 As Double

Dim c0 As Double
Dim c1 As Double

Dim c As Double
'**********************************
'   L-channel
'**********************************
    c000 = pt(0, 0, 0).L
    c001 = pt(0, 0, 1).L
    c010 = pt(0, 1, 0).L
    c011 = pt(0, 1, 1).L
    c100 = pt(1, 0, 0).L
    c101 = pt(1, 0, 1).L
    c110 = pt(1, 1, 0).L
    c111 = pt(1, 1, 1).L
    
    c00 = delta(0) * (c100 - c000) + c000
    c01 = delta(0) * (c101 - c001) + c001
    c10 = delta(0) * (c110 - c010) + c010
    c11 = delta(0) * (c111 - c011) + c011
    
    c0 = delta(1) * (c10 - c00) + c00
    c1 = delta(1) * (c11 - c01) + c01
    
    c = delta(2) * (c1 - c0) + c0
    
    TriLinInterpolation.L = c
    
'**********************************
'   a-channel
'**********************************
    c000 = pt(0, 0, 0).a
    c001 = pt(0, 0, 1).a
    c010 = pt(0, 1, 0).a
    c011 = pt(0, 1, 1).a
    c100 = pt(1, 0, 0).a
    c101 = pt(1, 0, 1).a
    c110 = pt(1, 1, 0).a
    c111 = pt(1, 1, 1).a
    
    c00 = delta(0) * (c100 - c000) + c000
    c01 = delta(0) * (c101 - c001) + c001
    c10 = delta(0) * (c110 - c010) + c010
    c11 = delta(0) * (c111 - c011) + c011
    
    c0 = delta(1) * (c10 - c00) + c00
    c1 = delta(1) * (c11 - c01) + c01
    
    c = delta(2) * (c1 - c0) + c0
    
    TriLinInterpolation.a = c

'**********************************
'   b-channel
'**********************************
    c000 = pt(0, 0, 0).b
    c001 = pt(0, 0, 1).b
    c010 = pt(0, 1, 0).b
    c011 = pt(0, 1, 1).b
    c100 = pt(1, 0, 0).b
    c101 = pt(1, 0, 1).b
    c110 = pt(1, 1, 0).b
    c111 = pt(1, 1, 1).b
    
    c00 = delta(0) * (c100 - c000) + c000
    c01 = delta(0) * (c101 - c001) + c001
    c10 = delta(0) * (c110 - c010) + c010
    c11 = delta(0) * (c111 - c011) + c011
    
    c0 = delta(1) * (c10 - c00) + c00
    c1 = delta(1) * (c11 - c01) + c01
    
    c = delta(2) * (c1 - c0) + c0
    
    TriLinInterpolation.b = c
End Function


Sub OverUnderReal(value As Double, ByVal n As Long, output() As Long, target As Double)
'Return is array with integer above and below the exact value.
'If input represents integer then both returned values is equal
ReDim output(0 To 1)
    output(0) = Int(value * (n - 1))
    output(1) = Int(value * (n - 1) + 0.9999999999999)
    
    target = value * (n - 1) - output(0)
End Sub



'************************************'
'                                    '
'       LookUp from structure        '
'                                    '
'************************************'
Function GetIccHeader(ICCprofile() As Byte) As tIccHeader
    GetIccHeader.ProfileSize = GetBig_uInt32Number(ICCprofile, 0)
    GetIccHeader.PreferredCMM = GetBig_uInt32Number(ICCprofile, 4)
    GetIccHeader.ProfileVersion = GetBig_uInt32Number(ICCprofile, 8)
    GetIccHeader.DeviceClass = GetBig_uInt32Number(ICCprofile, 12)
    GetIccHeader.ColorSpace = GetBig_uInt32Number(ICCprofile, 16)
    GetIccHeader.PCS = GetBig_uInt32Number(ICCprofile, 20)
    
    Get_Struct ICCprofile, 24, GetIccHeader.CreatedDateTime
    
    GetIccHeader.signature = GetBig_uInt32Number(ICCprofile, 36)
    GetIccHeader.PrimaryPlatformSignature = GetBig_uInt32Number(ICCprofile, 40)
    GetIccHeader.Flags = GetBig_uInt32Number(ICCprofile, 44)
    GetIccHeader.DeviceManufacturer = GetBig_uInt32Number(ICCprofile, 48)
    GetIccHeader.DeviceModel = GetBig_uInt32Number(ICCprofile, 52)
    
    Get_Struct ICCprofile, 56, GetIccHeader.DeviceAttributes
    GetIccHeader.RenderingIntent = GetBig_uInt32Number(ICCprofile, 64)
    Get_Struct ICCprofile, 68, GetIccHeader.IlluminantCIEXYZ
    
    GetIccHeader.ProfileCreator = GetBig_uInt32Number(ICCprofile, 80)
    Get_Struct ICCprofile, 84, GetIccHeader.ProfileID
    Get_Struct ICCprofile, 100, GetIccHeader.reserved
End Function

Function GetTag(tag As String, ICCtagTable As tIccTagTable) As tIccTagEntry
Dim n As Long
    For n = LBound(ICCtagTable.tagEntries) To UBound(ICCtagTable.tagEntries)
        If ICCtagTable.tagEntries(n).StringSig = tag Then
            GetTag = ICCtagTable.tagEntries(n)
            Exit Function
        End If
    Next n
    Err.Raise 5000, "GetTag", "Tag was not found"
End Function

Function GetTagTable(ICCprofile() As Byte) As tIccTagTable
Dim n As Long
Const TagTableOffset = 128
    
    GetTagTable.count = GetBig_uInt32Number(ICCprofile, 0 + TagTableOffset)
    ReDim GetTagTable.tagEntries(1 To GetTagTable.count) As tIccTagEntry
    For n = 1 To GetTagTable.count
        GetTagTable.tagEntries(n).signature = GetBig_uInt32Number(ICCprofile, TagTableOffset + 4 + 12 * (n - 1))
        GetTagTable.tagEntries(n).StringSig = LongToChars(GetTagTable.tagEntries(n).signature)
        GetTagTable.tagEntries(n).offset = GetBig_uInt32Number(ICCprofile, TagTableOffset + 8 + 12 * (n - 1))
        GetTagTable.tagEntries(n).size = GetBig_uInt32Number(ICCprofile, TagTableOffset + 12 + 12 * (n - 1))
    Next n
End Function



'************************************'
'                                    '
'       Read IccProfile from file    '
'                                    '
'************************************'
Function GetIccProfile(filename As String) As Byte()
Dim filenumber As Integer
Dim ICCprofileLength As Long
    filenumber = FreeFile
    Open filename For Binary Access Read As filenumber
    
    ICCprofileLength = LOF(filenumber)
    ReDim GetIccProfile(0 To ICCprofileLength - 1) As Byte
    Get filenumber, , GetIccProfile
    Close filenumber
End Function



Sub Get_B2A1_gamt(ICCpath As String, b2a1 As lutBToAType, gamt As lutBToAType)
Dim ICCprofile() As Byte
Dim ICCheader As tIccHeader
Dim ICCtagTable As tIccTagTable
Dim myICCtagEntry As tIccTagEntry
    
    
    ICCprofile = GetIccProfile(ICCpath)
    ICCheader = GetIccHeader(ICCprofile)
    ICCtagTable = GetTagTable(ICCprofile)
    
    
    'Get Gamut tables
    myICCtagEntry = GetTag("gamt", ICCtagTable)
    Select Case LongToChars(GetBig_uInt32Number(ICCprofile, myICCtagEntry.offset))
        Case "mft1": gamt = ReadType_lut8Type(ICCprofile, myICCtagEntry.offset)     'mft2' lut16Type   encoding - page 64
        Case "mft2": gamt = ReadType_lut16Type(ICCprofile, myICCtagEntry.offset)     'mft2' lut16Type   encoding - page 64
        Case "mBA ": gamt = ReadType_lutBToAType(ICCprofile, myICCtagEntry.offset)  'mBA ' lutBToAType encoding - page 73
    End Select
    
    
    'Get B2A1 tables
    myICCtagEntry = GetTag("B2A1", ICCtagTable)
    Select Case LongToChars(GetBig_uInt32Number(ICCprofile, myICCtagEntry.offset))
        Case "mft1": b2a1 = ReadType_lut8Type(ICCprofile, myICCtagEntry.offset)     'mft2' lut16Type   encoding - page 64
        Case "mft2": b2a1 = ReadType_lut16Type(ICCprofile, myICCtagEntry.offset)     'mft2' lut16Type   encoding - page 64
        Case "mBA ": b2a1 = ReadType_lutBToAType(ICCprofile, myICCtagEntry.offset)  'mBA ' lutBToAType encoding - page 73
    End Select
    
End Sub

