Attribute VB_Name = "ColorManagement"
Option Explicit

'Graphic methods
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, ByRef lpnTabStopPositions As Any, ByVal nTabOrigin As Long) As Long


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



Function ReadType_lut8Type(tagData() As Byte, offset As Long) As lut16Type
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
        .e(1) = GetBig_uInt32Number(tagData, offset + 12)
        .e(2) = GetBig_uInt32Number(tagData, offset + 16)
        .e(3) = GetBig_uInt32Number(tagData, offset + 20)
        .e(4) = GetBig_uInt32Number(tagData, offset + 24)
        .e(5) = GetBig_uInt32Number(tagData, offset + 28)
        .e(6) = GetBig_uInt32Number(tagData, offset + 32)
        .e(7) = GetBig_uInt32Number(tagData, offset + 36)
        .e(8) = GetBig_uInt32Number(tagData, offset + 40)
        .e(9) = GetBig_uInt32Number(tagData, offset + 44)
        
        
        'Prepare for the curves
        inLo = 48
        inHi = 47 + 256 * .i
        CLUTLo = inHi + 1
        CLUTHi = inHi + .o * .g ^ .i
        outLo = CLUTHi + 1
        outHi = CLUTHi + 256 * .o

        
        
        
        '*****************************
        '    Process inTables
        '*****************************
        .n = 256                                'Entries in InTables
        ReDim .inTables(0 To .i * .n - 1) 'InTables
        For n = 0 To (.i - 1) 'Cycle through the input curves
            For i = 0 To .n - 1
                .inTables(.n * n + i) = &H100& * GetBig_uInt8Number(tagData, offset + inLo + .n * n + i)
            Next i
        Next n
        
        
        '*****************************
        '    Process CLUT
        '*****************************
        ReDim .CLUT(0 To .o * .g ^ .i - 1) As Long
        For n = LBound(.CLUT) To UBound(.CLUT)
            .CLUT(n) = &H100& * GetBig_uInt8Number(tagData, offset + CLUTLo + n)
        Next n
        

        '*****************************
        '    Process outTables
        '*****************************
        .m = 256                                'Entries in outTables
        ReDim .outTables(0 To .o * .m - 1)            'OutTables
        For n = 0 To (.o - 1) 'Cycle through the output curves
            For i = 0 To .m - 1
                .outTables(.m * n + i) = &H100& * GetBig_uInt8Number(tagData, offset + outLo + .m * n + i)
            Next i
        Next n
        
        .legacyPCS = True  'Tables 39 and 40 in the standard
    
    End With
End Function

Function ReadType_lut16Type(tagData() As Byte, offset As Long) As lut16Type
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
        .e(1) = GetBig_uInt32Number(tagData, offset + 12)
        .e(2) = GetBig_uInt32Number(tagData, offset + 16)
        .e(3) = GetBig_uInt32Number(tagData, offset + 20)
        .e(4) = GetBig_uInt32Number(tagData, offset + 24)
        .e(5) = GetBig_uInt32Number(tagData, offset + 28)
        .e(6) = GetBig_uInt32Number(tagData, offset + 32)
        .e(7) = GetBig_uInt32Number(tagData, offset + 36)
        .e(8) = GetBig_uInt32Number(tagData, offset + 40)
        .e(9) = GetBig_uInt32Number(tagData, offset + 44)
        
        .n = GetBig_uInt16Number(tagData, offset + 48)
        .m = GetBig_uInt16Number(tagData, offset + 50)
        
        inLo = 52
        inHi = 51 + 2 * .n * .i
        CLUTLo = inHi + 1
        CLUTHi = inHi + 2 * .o * .g ^ .i
        outLo = CLUTHi + 1
        outHi = CLUTHi + 2 * .m * .o
        
        
        
        
        
        '*****************************
        '    Process inTables
        '*****************************
        ReDim .inTables(0 To .i * .n - 1) 'InTables
        For n = 0 To (.i - 1)   'Cycle through the input curves
            For i = 0 To .n - 1
                .inTables(.n * n + i) = GetBig_uInt16Number(tagData, offset + inLo + 2 * (.n * n + i))
            Next i
        Next n
        
        
        
        
        '*****************************
        '    Process CLUT
        '*****************************
        ReDim .CLUT(0 To .o * .g ^ .i - 1) As Long
        For n = LBound(.CLUT) To UBound(.CLUT)
            .CLUT(n) = GetBig_uInt16Number(tagData, offset + CLUTLo + 2 * n)
        Next n
        
        

        '*****************************
        '    Process outTables
        '*****************************
        ReDim .outTables(0 To .o * .m - 1) 'OutTables
        For n = 0 To (.o - 1) 'Cycle through the output curves
            For i = 0 To .m - 1
                .outTables(.m * n + i) = GetBig_uInt16Number(tagData, offset + outLo + 2 * (.m * n + i))
            Next i
        Next n
        
        .legacyPCS = True 'Tables 39 and 40 in the standard
    
    
    
    End With
End Function

Function ReadType_lutAToBType(tagData() As Byte, offset As Long) As lutAToBType
'Wrapper
Dim a2b As lutAToBType
Dim b2a As lutBToAType
Dim i As Integer
    
Dim test(0 To 199) As Byte

    b2a = ReadType_lutBToAType(tagData, offset)
    
    'Mem_Copy a2b, b2a, Len(b2a)
    'a2b.m = -1 'initialize a2b to avoid application error when function ends
    
    
    a2b.Acurve = b2a.Acurve
    a2b.Bcurve = b2a.Bcurve
    a2b.Mcurve = b2a.Mcurve
    a2b.CLUT = b2a.CLUT
    a2b.signature = b2a.signature
    a2b.offsetBcurve = b2a.offsetBcurve
    a2b.offsetMatrix = b2a.offsetMatrix
    a2b.offsetMcurve = b2a.offsetMcurve
    a2b.offsetCLUT = b2a.offsetCLUT
    a2b.offsetAcurve = b2a.offsetAcurve
    a2b.n = b2a.n
    a2b.m = b2a.m
    a2b.i = b2a.i
    a2b.o = b2a.o
    a2b.g = b2a.g
    a2b.CLUTchannels = b2a.CLUTchannels
    a2b.CLUTbitCount = b2a.CLUTbitCount
    
    For i = 1 To 12
        a2b.Matrix(i) = b2a.Matrix(i)
    Next i
    For i = 0 To 15
        a2b.CLUTgridPoints(i) = b2a.CLUTgridPoints(i)
    Next i

    ReadType_lutAToBType = a2b
End Function

Function ReadType_lutBToAType(tagData() As Byte, offset As Long) As lutBToAType
'   Structure is defined in
'   ICC v4.3 specification
'   section 10.10
Dim readBytes As Long
Dim n As Long
    With ReadType_lutBToAType
        
        .signature = GetBig_uInt32Number(tagData, offset + 0)
        
        If .signature <> &H6D414220 And _
           .signature <> &H6D424120 Then Err.Raise 11, "ReadType_lutABA", "lutAToBType or lutBToAType signature not found."
        
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
                    .CLUT(n) = &H101& * GetBig_uInt8Number(tagData, offset + .offsetCLUT + 20 + 2)
                Next n
            Else
                '16 bit CLUT
                For n = 0 To nGridPoints * .o - 1
                    .CLUT(n) = GetBig_uInt16Number(tagData, offset + .offsetCLUT + 20 + 2 * n)
                Next n
            End If
        End If
        
        
        
    End With
End Function



'************************************'
'                                    '
'       Transform                    '
'                                    '
'************************************'
Public Function TransformLut16(transform As lut16Type, inColors As dTriChannel) As dTriChannel
Dim i As Long
Dim k As Long
Dim j As Long
Dim pt(0 To 1, 0 To 1, 0 To 1) As dTriChannel

'Dim dLookUp0 As Double
Dim lLookUp0() As Long 'Channel 0 to be looked up. Two values - one below and one above
'Dim dLookUp1 As Double
Dim lLookUp1() As Long 'Channel 1 to be looked up. Two values - one below and one above
'Dim dLookUp2 As Double
Dim lLookUp2() As Long 'Channel 2 to be looked up. Two values - one below and one above

Dim var(0 To 2) As Double
Dim delta(0 To 2) As Long
Dim target(0 To 2) As Double

    With transform
        '*****************************
        ' InCurve
        '*****************************
        OverUnderReal inColors.ch0, .n, lLookUp0, target(0)
        OverUnderReal inColors.ch1, .n, lLookUp1, target(1)
        OverUnderReal inColors.ch2, .n, lLookUp2, target(2)
        
        delta(0) = .inTables(0 * .n + lLookUp0(1)) - .inTables(0 * .n + lLookUp0(0))
        delta(1) = .inTables(1 * .n + lLookUp1(1)) - .inTables(1 * .n + lLookUp1(0))
        delta(2) = .inTables(2 * .n + lLookUp2(1)) - .inTables(2 * .n + lLookUp2(0))
        
        var(0) = .inTables(0 * .n + lLookUp0(0)) + target(0) * delta(0)
        var(1) = .inTables(1 * .n + lLookUp1(0)) + target(1) * delta(1)
        var(2) = .inTables(2 * .n + lLookUp2(0)) + target(2) * delta(2)
        
        var(0) = var(0) / 65535
        var(1) = var(1) / 65535
        var(2) = var(2) / 65535
        
        
        
        '*****************************
        ' CLUT
        '*****************************
        OverUnderReal var(0), .g, lLookUp0, target(0)
        OverUnderReal var(1), .g, lLookUp1, target(1)
        OverUnderReal var(2), .g, lLookUp2, target(2)
        
        For i = 0 To 1
            For k = 0 To 1
                For j = 0 To 1
                    pt(i, k, j).ch0 = .CLUT((lLookUp0(i) * .g ^ 2 + lLookUp1(k) * .g + lLookUp2(j)) * .o + 0)
                    pt(i, k, j).ch1 = .CLUT((lLookUp0(i) * .g ^ 2 + lLookUp1(k) * .g + lLookUp2(j)) * .o + 1)
                    pt(i, k, j).ch2 = .CLUT((lLookUp0(i) * .g ^ 2 + lLookUp1(k) * .g + lLookUp2(j)) * .o + 2)
                Next j
            Next k
        Next i
        
        Dim CLUTresult As dTriChannel
        CLUTresult = (TriLinInterpolation(pt, target))
        
        var(0) = CLUTresult.ch0 / 65535
        var(1) = CLUTresult.ch1 / 65535
        var(2) = CLUTresult.ch2 / 65535
        
        
        
        '*****************************
        ' OutCurve
        '*****************************
        OverUnderReal var(0), .m, lLookUp0, target(0)
        OverUnderReal var(1), .m, lLookUp1, target(1)
        OverUnderReal var(2), .m, lLookUp2, target(2)
        
        delta(0) = .outTables(0 * .m + lLookUp0(1)) - .outTables(0 * .m + lLookUp0(0))
        delta(1) = .outTables(1 * .m + lLookUp1(1)) - .outTables(1 * .m + lLookUp1(0))
        delta(2) = .outTables(2 * .m + lLookUp2(1)) - .outTables(2 * .m + lLookUp2(0))
        
        var(0) = .outTables(0 * .m + lLookUp0(0)) + target(0) * delta(0)
        var(1) = .outTables(1 * .m + lLookUp1(0)) + target(1) * delta(1)
        var(2) = .outTables(2 * .m + lLookUp2(0)) + target(2) * delta(2)
    End With
    
    TransformLut16.ch0 = var(0) / 65535
    TransformLut16.ch1 = var(1) / 65535
    TransformLut16.ch2 = var(2) / 65535
End Function

Public Function TransformLut16_b2a(transform As lut16Type, inColors As dLabCOLOR) As dRGBCOLOR
Dim inClr As dTriChannel
Dim outClr As dTriChannel

    '*****************************
    ' Input conversion
    '*****************************
    
    'Use legacy PCSLAB encoding
    inClr.ch0 = inColors.L / (100# + 25500 / 65280)
    inClr.ch1 = (inColors.a + 128#) / (255# + 255 / 256)
    inClr.ch2 = (inColors.b + 128#) / (255# + 255 / 256)
    
    outClr = TransformLut16(transform, inClr)

    TransformLut16_b2a.red = outClr.ch0 * 255#
    TransformLut16_b2a.green = outClr.ch1 * 255#
    TransformLut16_b2a.blue = outClr.ch2 * 255#

    
    'TransformLab2Tag.red = Round(varL * 255, 0)    '.5489             .5813            .5772               .5755
    'TransformLab2Tag.green = Round(varA * 255, 0)  '.6510             .6715            .6418               .6610
    'TransformLab2Tag.blue = Round(varB * 255, 0)   '.1480             .1660            .1586               .1334
End Function

Public Function TransformLut16_a2b(transform As lut16Type, inColors As dRGBCOLOR) As dLabCOLOR
Dim inClr As dTriChannel
Dim outClr As dTriChannel

    '*****************************
    ' Input conversion
    '*****************************
    inClr.ch0 = inColors.red / 255#
    inClr.ch1 = inColors.green / 255#
    inClr.ch2 = inColors.blue / 255#
    
    outClr = TransformLut16(transform, inClr)

    'Use legacy PCSLAB encoding
    TransformLut16_a2b.L = outClr.ch0 * (100# + 25500 / 65280)
    TransformLut16_a2b.a = outClr.ch1 * (255# + 255 / 256) - 128#
    TransformLut16_a2b.b = outClr.ch2 * (255# + 255 / 256) - 128#

    'TransformLab2Tag.red = Round(varL * 255, 0)    '.5489             .5813            .5772               .5755
    'TransformLab2Tag.green = Round(varA * 255, 0)  '.6510             .6715            .6418               .6610
    'TransformLab2Tag.blue = Round(varB * 255, 0)   '.1480             .1660            .1586               .1334
End Function

Public Function TransformB2A(transform As lutBToAType, LAB As dLabCOLOR) As dRGBCOLOR
Dim i As Long
Dim k As Long
Dim j As Long
Dim pt(0 To 1, 0 To 1, 0 To 1) As dTriChannel
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
    'Use ICC v4.3 PCSLAB encoding
    dLookUpL = LAB.L / (100#)
    dLookUpA = (LAB.a + 128#) / 255#
    dLookUpB = (LAB.b + 128#) / 255#
    
    
    
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
                pt(i, k, j).ch0 = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, transform.o, lLookUpL(i), lLookUpA(k), lLookUpB(j)) + 0)
                pt(i, k, j).ch1 = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, transform.o, lLookUpL(i), lLookUpA(k), lLookUpB(j)) + 1)
                pt(i, k, j).ch2 = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, transform.o, lLookUpL(i), lLookUpA(k), lLookUpB(j)) + 2)
            Next j
        Next k
    Next i
    
    Dim CLUTresult As dTriChannel
    CLUTresult = (TriLinInterpolation(pt, target))
    
    varL = CLUTresult.ch0 / 65535
    varA = CLUTresult.ch1 / 65535
    varB = CLUTresult.ch2 / 65535
    
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
    
    TransformB2A.red = varL * 255#    '.5489             .5813            .5772               .5755
    TransformB2A.green = varA * 255#  '.6510             .6715            .6418               .6610
    TransformB2A.blue = varB * 255#   '.1480             .1660            .1586               .1334
End Function
Public Function TransformA2B(transform As lutAToBType, RGB As dRGBCOLOR) As dLabCOLOR
Dim i As Long
Dim k As Long
Dim j As Long
Dim pt(0 To 1, 0 To 1, 0 To 1) As dTriChannel
Dim dLookUpR As Double
Dim lLookUpR() As Long 'L to be looked up. Two values - one below and one above
Dim dLookUpG As Double
Dim lLookUpG() As Long 'A to be looked up. Two values - one below and one above
Dim dLookUpB As Double
Dim lLookUpB() As Long 'B to be looked up. Two values - one below and one above

Dim varR As Double
Dim varG As Double
Dim varB As Double
Dim delta(0 To 2) As Long
Dim target(0 To 2) As Double
    
    
    '*****************************
    ' Input conversion
    '*****************************
    dLookUpR = RGB.red / 255#
    dLookUpG = RGB.green / 255#
    dLookUpB = RGB.blue / 255#
    
    
    '*****************************
    ' Acurve
    '*****************************
    OverUnderReal dLookUpR, transform.Acurve(0).n, lLookUpR, target(0)
    OverUnderReal dLookUpG, transform.Acurve(1).n, lLookUpG, target(1)
    OverUnderReal dLookUpB, transform.Acurve(2).n, lLookUpB, target(2)
    
    delta(0) = transform.Acurve(0).curve(lLookUpR(1)) - transform.Acurve(0).curve(lLookUpR(0))
    delta(1) = transform.Acurve(1).curve(lLookUpG(1)) - transform.Acurve(1).curve(lLookUpG(0))
    delta(2) = transform.Acurve(2).curve(lLookUpB(1)) - transform.Acurve(2).curve(lLookUpB(0))
    
    varR = transform.Acurve(0).curve(lLookUpR(0)) + target(0) * delta(0)
    varG = transform.Acurve(1).curve(lLookUpG(0)) + target(1) * delta(1)
    varB = transform.Acurve(2).curve(lLookUpB(0)) + target(2) * delta(2)
    
    varR = varR / 65535
    varG = varG / 65535
    varB = varB / 65535
    
    
    '*****************************
    ' CLUT
    '*****************************
    OverUnderReal varR, transform.CLUTgridPoints(0), lLookUpR, target(0)
    OverUnderReal varG, transform.CLUTgridPoints(1), lLookUpG, target(1)
    OverUnderReal varB, transform.CLUTgridPoints(2), lLookUpB, target(2)
    
    For i = 0 To 1
        For k = 0 To 1
            For j = 0 To 1
                pt(i, k, j).ch0 = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, transform.o, lLookUpR(i), lLookUpG(k), lLookUpB(j)) + 0)
                pt(i, k, j).ch1 = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, transform.o, lLookUpR(i), lLookUpG(k), lLookUpB(j)) + 1)
                pt(i, k, j).ch2 = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, transform.o, lLookUpR(i), lLookUpG(k), lLookUpB(j)) + 2)
            Next j
        Next k
    Next i
    
    Dim CLUTresult As dTriChannel
    CLUTresult = (TriLinInterpolation(pt, target))
    
    varR = CLUTresult.ch0 / 65535
    varG = CLUTresult.ch1 / 65535
    varB = CLUTresult.ch2 / 65535
    
    
    
    
    '*****************************
    ' Bcurve
    '*****************************
    OverUnderReal varR, transform.Bcurve(0).n, lLookUpR, target(0)
    OverUnderReal varG, transform.Bcurve(1).n, lLookUpG, target(1)
    OverUnderReal varB, transform.Bcurve(2).n, lLookUpB, target(2)
    
    delta(0) = transform.Bcurve(0).curve(lLookUpR(1)) - transform.Bcurve(0).curve(lLookUpR(0))
    delta(1) = transform.Bcurve(1).curve(lLookUpG(1)) - transform.Bcurve(1).curve(lLookUpG(0))
    delta(2) = transform.Bcurve(2).curve(lLookUpB(1)) - transform.Bcurve(2).curve(lLookUpB(0))
    
    varR = transform.Bcurve(0).curve(lLookUpR(0)) + target(0) * delta(0)
    varG = transform.Bcurve(1).curve(lLookUpG(0)) + target(1) * delta(1)
    varB = transform.Bcurve(2).curve(lLookUpB(0)) + target(2) * delta(2)
    
    varR = varR / 65535
    varG = varG / 65535
    varB = varB / 65535
    
    
    
    
    '*****************************
    ' Output conversion
    '*****************************
    'Use ICC v4.3 PCSLAB encoding
    TransformA2B.L = varR * 100#
    TransformA2B.a = varG * 255# - 128#
    TransformA2B.b = varB * 255# - 128#
    
End Function

Public Function InGamutLab2Tag(transform As lutBToAType, LAB As dLabCOLOR) As dTriChannel
Dim i As Long
Dim k As Long
Dim j As Long
Dim pt(0 To 1, 0 To 1, 0 To 1) As dTriChannel
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
    'Use ICC v4.3 PCSLAB encoding
    dLookUpL = LAB.L / (100#)
    dLookUpA = (LAB.a + 128#) / 255#
    dLookUpB = (LAB.b + 128#) / 255#
    
    
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
                pt(i, k, j).ch0 = transform.CLUT(GetArrayAddress(transform.CLUTgridPoints, transform.o, lLookUpL(i), lLookUpA(k), lLookUpB(j)))
            Next j
        Next k
    Next i
    
    Dim CLUTresult As dTriChannel
    CLUTresult = (TriLinInterpolation(pt, target))
    
    varL = CLUTresult.ch0 / 65535
    
    
    '*****************************
    ' Acurve
    '*****************************
    OverUnderReal varL, transform.Acurve(0).n, lLookUpL, target(0)
    
    delta(0) = transform.Acurve(0).curve(lLookUpL(1)) - transform.Acurve(0).curve(lLookUpL(0))
    
    varL = transform.Acurve(0).curve(lLookUpL(0)) + target(0) * delta(0)
    
    InGamutLab2Tag.ch0 = 255# * (varL / 65535)
    'If varL = 65535 Then InGamutLab2Tag = True
    '                Else: InGamutLab2Tag = False
End Function


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

Function GetArrayAddress(GridPoints() As Byte, nOutChannels As Byte, ParamArray input0to15() As Variant) As Long
Dim n As Integer
Dim dimensions As Long
    dimensions = UBound(input0to15)
    For n = 0 To dimensions
        GetArrayAddress = GetArrayAddress + input0to15(dimensions - n) * GridPoints(dimensions - n) ^ n
    Next n
    GetArrayAddress = GetArrayAddress * nOutChannels
End Function

Function TriLinInterpolation(pt() As dTriChannel, delta() As Double) As dTriChannel
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
'   Channel 0
'**********************************
    c000 = pt(0, 0, 0).ch0
    c001 = pt(0, 0, 1).ch0
    c010 = pt(0, 1, 0).ch0
    c011 = pt(0, 1, 1).ch0
    c100 = pt(1, 0, 0).ch0
    c101 = pt(1, 0, 1).ch0
    c110 = pt(1, 1, 0).ch0
    c111 = pt(1, 1, 1).ch0
    
    c00 = delta(0) * (c100 - c000) + c000
    c01 = delta(0) * (c101 - c001) + c001
    c10 = delta(0) * (c110 - c010) + c010
    c11 = delta(0) * (c111 - c011) + c011
    
    c0 = delta(1) * (c10 - c00) + c00
    c1 = delta(1) * (c11 - c01) + c01
    
    c = delta(2) * (c1 - c0) + c0
    
    TriLinInterpolation.ch0 = c
    
'**********************************
'   Channel 1
'**********************************
    c000 = pt(0, 0, 0).ch1
    c001 = pt(0, 0, 1).ch1
    c010 = pt(0, 1, 0).ch1
    c011 = pt(0, 1, 1).ch1
    c100 = pt(1, 0, 0).ch1
    c101 = pt(1, 0, 1).ch1
    c110 = pt(1, 1, 0).ch1
    c111 = pt(1, 1, 1).ch1
    
    c00 = delta(0) * (c100 - c000) + c000
    c01 = delta(0) * (c101 - c001) + c001
    c10 = delta(0) * (c110 - c010) + c010
    c11 = delta(0) * (c111 - c011) + c011
    
    c0 = delta(1) * (c10 - c00) + c00
    c1 = delta(1) * (c11 - c01) + c01
    
    c = delta(2) * (c1 - c0) + c0
    
    TriLinInterpolation.ch1 = c

'**********************************
'   Channel 2
'**********************************
    c000 = pt(0, 0, 0).ch2
    c001 = pt(0, 0, 1).ch2
    c010 = pt(0, 1, 0).ch2
    c011 = pt(0, 1, 1).ch2
    c100 = pt(1, 0, 0).ch2
    c101 = pt(1, 0, 1).ch2
    c110 = pt(1, 1, 0).ch2
    c111 = pt(1, 1, 1).ch2
    
    c00 = delta(0) * (c100 - c000) + c000
    c01 = delta(0) * (c101 - c001) + c001
    c10 = delta(0) * (c110 - c010) + c010
    c11 = delta(0) * (c111 - c011) + c011
    
    c0 = delta(1) * (c10 - c00) + c00
    c1 = delta(1) * (c11 - c01) + c01
    
    c = delta(2) * (c1 - c0) + c0
    
    TriLinInterpolation.ch2 = c
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
        GetTagTable.tagEntries(n).datatype = LongToChars(GetBig_uInt32Number(ICCprofile, GetTagTable.tagEntries(n).offset))
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




'************************************'
'                                    '
'       Translate colors             '
'                                    '
'************************************'

Function Translate(ICCprofile() As Byte, tag As tIccTagEntry, RGB As dRGBCOLOR, LAB As dLabCOLOR)
Dim lut16 As lut16Type
Dim lutA2B As lutAToBType
Dim lutB2A As lutBToAType

    'Select the correct direction for the translation
    Select Case tag.StringSig
        
        
        'Tag is conversion from RGB to PCS with relative colorimetric rendering intent
        Case "A2B1":
            Select Case LongToChars(GetBig_uInt32Number(ICCprofile, tag.offset))
                Case "mft1":
                    lut16 = ReadType_lut8Type(ICCprofile, tag.offset)     'mft1' lut8Type   encoding - page 52
                    LAB = TransformLut16_a2b(lut16, RGB)
                Case "mft2":
                    lut16 = ReadType_lut16Type(ICCprofile, tag.offset)     'mft2' lut16Type   encoding - page 64
                    LAB = TransformLut16_a2b(lut16, RGB)
                Case "mAB ":
                    lutA2B = ReadType_lutAToBType(ICCprofile, tag.offset)  'mAB ' lutAToBType encoding - page 54
                    LAB = TransformA2B(lutA2B, RGB)
            End Select
            
        
        'Tag is conversion from PCS to RGB with relative colorimetric rendering intent
        Case "B2A1":
            Select Case LongToChars(GetBig_uInt32Number(ICCprofile, tag.offset))
                Case "mft1":
                    lut16 = ReadType_lut8Type(ICCprofile, tag.offset)     'mft1' lut8Type   encoding - page 52
                    RGB = TransformLut16_b2a(lut16, LAB)
                Case "mft2":
                    lut16 = ReadType_lut16Type(ICCprofile, tag.offset)     'mft2' lut16Type   encoding - page 64
                    RGB = TransformLut16_b2a(lut16, LAB)
                Case "mBA ":
                    lutB2A = ReadType_lutBToAType(ICCprofile, tag.offset)  'mBA ' lutAToBType encoding - page 57
                    RGB = TransformB2A(lutB2A, LAB)
            End Select
        
        
        Case "gamt":
            'Tag is conversion from PCS to gamut
             Select Case LongToChars(GetBig_uInt32Number(ICCprofile, tag.offset))
                Case "mft1":
                    lut16 = ReadType_lut8Type(ICCprofile, tag.offset)     'mft1' lut8Type   encoding - page 52
                    RGB = TransformLut16_b2a(lut16, LAB)
                Case "mft2":
                    lut16 = ReadType_lut16Type(ICCprofile, tag.offset)     'mft2' lut16Type   encoding - page 64
                    RGB = TransformLut16_b2a(lut16, LAB)
                Case "mBA ":
                    lutB2A = ReadType_lutBToAType(ICCprofile, tag.offset)  'mBA ' lutAToBType encoding - page 57
                    Dim gmt As Double
                    gmt = InGamutLab2Tag(lutB2A, LAB).ch0
                    RGB.red = gmt
                    RGB.green = gmt
                    RGB.blue = gmt
            End Select
       
        
    Case Else:
        Err.Raise 5000, "Translate()", "Tag was not found"
    End Select
End Function





'*************************************************************'
'                                                             '
'       Fill DeltaE00(L, a, b) and save to disk  from file    '
'                                                             '
'       Contains DeltaE00 values for conversion from LAB      '
'       via device and back to LAB                            '
'                                                             '
'*************************************************************'


Sub CreateGamutFile(ColorProfile As String)
Const ColorFolder = "C:\Windows\System32\spool\drivers\color\"

Dim ICCprofile() As Byte
Dim ICCtagTable As tIccTagTable
Dim ICCtag_A2B As tIccTagEntry
Dim ICCtag_B2A As tIccTagEntry

Dim L As Integer
Dim a As Integer
Dim b As Integer

Dim clr1 As dLabCOLOR
Dim clr2 As dLabCOLOR
Dim RGB As dRGBCOLOR

Dim gamutfile As Long
Dim DeltaE00(0 To 100, -128 To 127, -128 To 127) As Single
    
Dim sStatus As String
Dim hScreenDC As LongPtr

Dim lut16_a2b As lut16Type
Dim lut16_b2a As lut16Type
Dim lutA2B As lutAToBType
Dim lutB2A As lutBToAType

    hScreenDC = GetDC(0)
    
    ICCprofile = GetIccProfile(ColorFolder & ColorProfile & ".icc")
    ICCtagTable = GetTagTable(ICCprofile)
    
    ICCtag_B2A = GetTag("B2A1", ICCtagTable)
    ICCtag_A2B = GetTag("A2B1", ICCtagTable)

    Select Case ICCtag_A2B.datatype
        Case "mft1":
            lut16_a2b = ReadType_lut8Type(ICCprofile, ICCtag_A2B.offset)     'mft1' lut8Type   encoding - page 52
            lut16_b2a = ReadType_lut8Type(ICCprofile, ICCtag_B2A.offset)     'mft1' lut8Type   encoding - page 52
        Case "mft2":
            lut16_a2b = ReadType_lut16Type(ICCprofile, ICCtag_A2B.offset)     'mft1' lut8Type   encoding - page 52
            lut16_b2a = ReadType_lut16Type(ICCprofile, ICCtag_B2A.offset)     'mft1' lut8Type   encoding - page 52
        Case "mAB ":
            lutA2B = ReadType_lutAToBType(ICCprofile, ICCtag_A2B.offset)  'mAB ' lutAToBType encoding - page 54
            lutB2A = ReadType_lutBToAType(ICCprofile, ICCtag_B2A.offset)  'mBA ' lutAToBType encoding - page
        Case Else
            Err.Raise 5000, "CreateGamutFile", "Unknown datatype"
    End Select
    
    If ICCtag_A2B.datatype = "mAB " Then
        'Transform for lutAtoBType and lutBtoAType
        For L = 0 To 100
            For a = -128 To 127
                For b = -128 To 127
                    clr1.L = L
                    clr1.a = a
                    clr1.b = b
                    
                    RGB = TransformB2A(lutB2A, clr1)
                    clr2 = TransformA2B(lutA2B, RGB)
                    DeltaE00(L, a, b) = DE00(clr1.L, clr1.a, clr1.b, clr2.L, clr2.a, clr2.b)
    
                    sStatus = "lutBtoAType (" & Format(L, "000") & ", " & Format(a, "000") & ")"
                    TabbedTextOut hScreenDC, 800, 0, sStatus, Len(sStatus), 0, 0, 0
                Next b
            Next a
        Next L
    Else
        'Transform for lut8Type  and lut16Type
        For L = 0 To 100
            For a = -128 To 127
                For b = -128 To 127
                    clr1.L = L
                    clr1.a = a
                    clr1.b = b
                    
                    RGB = TransformLut16_b2a(lut16_b2a, clr1)
                    clr2 = TransformLut16_a2b(lut16_a2b, RGB)
                    DeltaE00(L, a, b) = DE00(clr1.L, clr1.a, clr1.b, clr2.L, clr2.a, clr2.b)
    
                    sStatus = "lut16Type (" & Format(L, "000") & ", " & Format(a, "000") & ")"
                    TabbedTextOut hScreenDC, 800, 0, sStatus, Len(sStatus), 0, 0, 0
                Next b
            Next a
        Next L
    End If
    
    gamutfile = FreeFile
    Open ColorFolder & ColorProfile & ".gmt" For Binary Access Write As gamutfile
    
    Put gamutfile, , DeltaE00
    'Get gamutfile, , DeltaE00
    
    Close gamutfile

    ReleaseDC 0, hScreenDC
End Sub


Sub CreateGamutFiles()
    CreateGamutFile "HP DesignJet Z6200ps 42in Photo_PolyPropylene Banner 200 - HP Matte Polypropylene"
    CreateGamutFile "z6200 PolyPropylene 200 M1 for D50 i1Pro2"
End Sub

