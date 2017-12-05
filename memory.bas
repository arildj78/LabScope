Attribute VB_Name = "memory"
#If Win64 Then
    Public Const PTR_LENGTH As Long = 8
#Else
    Public Const PTR_LENGTH As Long = 4
#End If


Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" ( _
ByRef Destination As Any, _
ByRef Source As Any, _
ByVal length As Long)
Public Declare PtrSafe Sub Mem_CopyPtr Lib "kernel32" Alias "RtlMoveMemory" ( _
ByVal Destination As LongPtr, _
ByVal Source As LongPtr, _
ByVal length As Long)

'Platform-independent method to return the full zero-padded
'hexadecimal representation of a pointer value
Function HexPtr(ByVal Ptr As LongPtr) As String
    HexPtr = Hex$(Ptr)
    HexPtr = String$((PTR_LENGTH * 2) - Len(HexPtr), "0") & HexPtr
End Function

Public Function Mem_ReadHex(ByVal Ptr As LongPtr, ByVal length As Long) As String
Dim bBuffer() As Byte, strBytes() As String, i As Long, ub As Long, b As Byte
    
    ub = length - 1
    ReDim bBuffer(ub)
    ReDim strBytes(ub)
    
    Mem_Copy bBuffer(0), ByVal Ptr, length
    For i = 0 To ub
        b = bBuffer(i)
        strBytes(i) = IIf(b < 16, "0", "") & Hex$(b)
    Next
    Mem_ReadHex = Join(strBytes, " ")
End Function

Public Function Mem_ReadString(ByVal Ptr As LongPtr, ByVal length As Long) As String
Dim bBuffer() As Byte, strBytes() As String, i As Long, ub As Long, b As Byte
    
    ub = length - 1
    ReDim bBuffer(ub)
    ReDim strBytes(ub)
    
    Mem_Copy bBuffer(0), ByVal Ptr, length
    For i = 0 To ub
        b = bBuffer(i)
        strBytes(i) = Chr(b)
    Next
    Mem_ReadString = Join(strBytes, "")
End Function



Sub Performance()
Dim t1 As Long
Dim t2 As Long
Dim n As Long
Dim i As Long


'***************************************************************
'  Declarations and inititalisation for code under test
'***************************************************************
    Dim filenumber As Integer
    Dim arrTemp() As Byte
    Dim arrTemp2() As Long
    Dim a As Long
        ICCprofile = "C:\Windows\System32\spool\drivers\color\HP DesignJet Z6200ps 42in Photo_PolyPropylene Banner 200 - HP Matte Polypropylene.icc"
        filenumber = FreeFile
        Open ICCprofile For Binary Access Read As filenumber
        ReDim arrTemp(0 To LOF(filenumber) - 1) As Byte
        ReDim arrTemp2(0 To LOF(filenumber) / 4 - 1) As Long
        Get filenumber, , arrTemp


'***************************************************************
    
   
    
    t1 = GetTickCount
    For n = 1 To 10 ^ 0
        
        
        
        '**************************************
        '    Code to be tested
        '**************************************
       
        For i = 0 To LOF(filenumber) / 4 - 1
        
            arrTemp2(i) = GetBig_uInt32Number(arrTemp, i * 2)
        
        Next i

        '**************************************
        
        
        
        
    Next n
    
    
    
    
    '**************************************
    '    Cleanup after test
    '**************************************
    Close filenumber
    
    '**************************************
 
 
 
    t2 = GetTickCount
    MsgBox "Operation took " & t2 - t1 & " ticks"
End Sub
