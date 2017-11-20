Attribute VB_Name = "memory"
#If Win64 Then
    Public Const PTR_LENGTH As Long = 8
#Else
    Public Const PTR_LENGTH As Long = 4
#End If

Public Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" ( _
ByRef Destination As Any, _
ByRef Source As Any, _
ByVal Length As Long)

'Platform-independent method to return the full zero-padded
'hexadecimal representation of a pointer value
Function HexPtr(ByVal Ptr As LongPtr) As String
    HexPtr = Hex$(Ptr)
    HexPtr = String$((PTR_LENGTH * 2) - Len(HexPtr), "0") & HexPtr
End Function

Public Function Mem_ReadHex(ByVal Ptr As LongPtr, ByVal Length As Long) As String
Dim bBuffer() As Byte, strBytes() As String, i As Long, ub As Long, b As Byte
    
    ub = Length - 1
    ReDim bBuffer(ub)
    ReDim strBytes(ub)
    
    Mem_Copy bBuffer(0), ByVal Ptr, Length
    For i = 0 To ub
        b = bBuffer(i)
        strBytes(i) = IIf(b < 16, "0", "") & Hex$(b)
    Next
    Mem_ReadHex = Join(strBytes, " ")
End Function

Public Function Mem_ReadString(ByVal Ptr As LongPtr, ByVal Length As Long) As String
Dim bBuffer() As Byte, strBytes() As String, i As Long, ub As Long, b As Byte
    
    ub = Length - 1
    ReDim bBuffer(ub)
    ReDim strBytes(ub)
    
    Mem_Copy bBuffer(0), ByVal Ptr, Length
    For i = 0 To ub
        b = bBuffer(i)
        strBytes(i) = Chr(b)
    Next
    Mem_ReadString = Join(strBytes, "")
End Function

