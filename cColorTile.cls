VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cColorTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hdcDest As LongPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Boolean
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Boolean
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As LongPtr
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal fnObject As Long) As LongPtr
Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal iBkMode As Long) As Long
Private Declare PtrSafe Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As LongPtr, ByVal nXStart As Long, ByVal nYStart As Long, ByVal lpString As String, ByVal cchString As Long) As Boolean

'Drawing methods
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, ByRef lpRect As Any, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function Polygon Lib "gdi32" (ByVal hDC As LongPtr, ByVal lpPoints As LongPtr, ByVal nCount As Long) As Boolean
Private Declare PtrSafe Function Rectangle Lib "gdi32" (ByVal hDC As LongPtr, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Boolean
Private Declare PtrSafe Function MoveToEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpPoint As LongPtr) As Boolean
Private Declare PtrSafe Function LineTo Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long) As Boolean


'Memory management methods
Private Declare PtrSafe Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As LongPtr
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Boolean
Private Declare PtrSafe Function HeapDestroy Lib "kernel32" (ByVal hHeap As LongPtr) As Boolean






Private pParent As Object
Private WithEvents pDisplaySurface As MSForms.Image
Attribute pDisplaySurface.VB_VarHelpID = -1
Private hDcDrawing As LongPtr
Private bmpDrawingOrig As LongPtr
Private bmpDrawing As LongPtr
Private BrushDrawingTgt As LongPtr
Private BrushDrawingHit As LongPtr
Private BrushDrawingOrig As LongPtr
Private BrushTransparent As LongPtr
Private PenDrawingOrig As LongPtr
Private PenFrame As LongPtr
Private PenXout As LongPtr
Private PenSelectedFrame As LongPtr
Private PenNull As LongPtr

Private TgtRGB As Long
Private HitRGB As Long

Private pxWidth As Long
Private pxHeight As Long
Private GlobalRect As RECT
Private LocalRect As RECT

Private hHeap As LongPtr

Private hWndForm As LongPtr
Private hDcDisplaySurface As LongPtr

Private Triangle(0 To 2) As POINT


Private result As Long 'result from various operations used for debugging

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Properties
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private pLabScope As cLabScope
Private pLScope As cLScope
Private pHitLabL As Double
Private pHitLabA As Double
Private pHitLabB As Double
Private pTgtLabL As Double
Private pTgtLabA As Double
Private pTgtLabB As Double
Private pTileGroup As Collection
Private pFrame As Boolean
Private pFrameThickness As Long
Private pXout As Boolean


''''''''''''''''''''''
' LabScope property
''''''''''''''''''''''
Public Property Get LabScope() As cLabScope
    Set LabScope = pLabScope
End Property
Public Property Let LabScope(value As cLabScope)
    Set pLabScope = value
End Property
''''''''''''''''''''''
' LScope property
''''''''''''''''''''''
Public Property Get LScope() As cLScope
    Set LScope = pLScope
End Property
Public Property Let LScope(value As cLScope)
    Set pLScope = value
End Property
''''''''''''''''''''''
' TgtLabL property
''''''''''''''''''''''
Public Property Get TgtLabL() As Double
    TgtLabL = pTgtLabL
End Property
Public Property Let TgtLabL(value As Double)
    pTgtLabL = value
    UpdateTgtBrush
    RedrawControl
End Property
''''''''''''''''''''''
' TgtLabA property
''''''''''''''''''''''
Public Property Get TgtLabA() As Double
    TgtLabA = pTgtLabA
End Property
Public Property Let TgtLabA(value As Double)
    pTgtLabA = value
    UpdateTgtBrush
    RedrawControl
End Property
''''''''''''''''''''''
' TgtLabB property
''''''''''''''''''''''
Public Property Get TgtLabB() As Double
    TgtLabB = pTgtLabB
End Property
Public Property Let TgtLabB(value As Double)
    pTgtLabB = value
    UpdateTgtBrush
    RedrawControl
End Property

''''''''''''''''''''''
' HitLabL property
''''''''''''''''''''''
Public Property Get HitLabL() As Double
    HitLabL = pHitLabL
End Property
Public Property Let HitLabL(value As Double)
    pHitLabL = value
    UpdateHitBrush
    RedrawControl
End Property
''''''''''''''''''''''
' HitLabA property
''''''''''''''''''''''
Public Property Get HitLabA() As Double
    HitLabA = pHitLabA
End Property
Public Property Let HitLabA(value As Double)
    pHitLabA = value
    UpdateHitBrush
    RedrawControl
End Property
''''''''''''''''''''''
' HitLabB property
''''''''''''''''''''''
Public Property Get HitLabB() As Double
    HitLabB = pHitLabB
End Property
Public Property Let HitLabB(value As Double)
    pHitLabB = value
    UpdateHitBrush
    RedrawControl
End Property
''''''''''''''''''''''
' Frame property
''''''''''''''''''''''
Public Property Get Frame() As Boolean
    Frame = pFrame
End Property
Public Property Let Frame(value As Boolean)
    If pFrame <> value Then
        pFrame = value
        RedrawControl
    End If
End Property
''''''''''''''''''''''
' FrameThickness property
''''''''''''''''''''''
Public Property Get FrameThickness() As Long
    FrameThickness = pFrameThickness
End Property
Public Property Let FrameThickness(value As Long)
    pFrameThickness = value
    RedrawControl
End Property
''''''''''''''''''''''
' TileGroup property
''''''''''''''''''''''
Public Property Get TileGroup() As Collection
    Set TileGroup = pTileGroup
End Property
Public Property Let TileGroup(value As Collection)
    Set pTileGroup = value
End Property
''''''''''''''''''''''
' xOut property
''''''''''''''''''''''
Public Property Get xOut() As Boolean
    xOut = pXout
End Property
Public Property Let xOut(value As Boolean)
    pXout = value
    RedrawControl
End Property
''''''''''''''''''''''
' AverageDeltaE00 property
''''''''''''''''''''''
Public Property Get AverageDeltaE00() As Double
    If pTileGroup Is Nothing Then
        'Tile is not assigned to any group.
        AverageDeltaE00 = 0
    Else
        'Tile is assigned to a group. Calculate average of group.
        AverageDeltaE00 = 0
    End If
End Property



Public Sub InitiateProperties(ByRef Parent As Object, _
                              ByRef DisplaySurface As MSForms.Image, _
                              Optional LabScope As cLabScope = Nothing, _
                              Optional LScope As cLScope = Nothing, _
                              Optional TgtLabL As Double = 60#, _
                              Optional TgtLabA As Double = 0#, _
                              Optional TgtLabB As Double = 0#, _
                              Optional HitLabL As Double = 50#, _
                              Optional HitLabA As Double = 0#, _
                              Optional HitLabB As Double = 0#, _
                              Optional Frame As Boolean = False, _
                              Optional FrameThickness As Long = 1, _
                              Optional xOut As Boolean)
    
    
    hHeap = HeapCreate(0, 0, 0)  ' Create a heap memory object to be used for allocating memory to usercontrol images
    
    
    'Set connection between me and my display surface
    Set pParent = Parent
    Set pDisplaySurface = DisplaySurface

    Set pLabScope = LabScope
    Set pLScope = LScope
        
    pTgtLabL = TgtLabL
    pTgtLabA = TgtLabA
    pTgtLabB = TgtLabB
    pHitLabL = HitLabL
    pHitLabA = HitLabA
    pHitLabB = HitLabB
    pFrame = Frame
    pFrameThickness = FrameThickness
    pXout = xOut
    
    Dim sR As Double
    Dim sG As Double
    Dim sB As Double
    LAB2sRGB pTgtLabL, pTgtLabA, pTgtLabB, sR, sG, sB
    TgtRGB = RGB(sR, sG, sB)
    
    LAB2sRGB pHitLabL, pHitLabA, pHitLabB, sR, sG, sB
    HitRGB = RGB(sR, sG, sB)
    
    pFrame = Frame
    pFrameThickness = FrameThickness
    
    
    Find_Window_handle
    hDcDisplaySurface = GetDC(hWndForm)
    hDcDrawing = CreateCompatibleDC(0)            ' Create DeviceContext that matches the screen that image1 is shown on
    
    BrushTransparent = GetStockObject(NULL_BRUSH And &HFFFFFFF)
    PenFrame = CreatePen(PenStyle.PS_Solid, 1, 0)
    PenXout = CreatePen(PenStyle.PS_Solid, 3, &HFF&)
    PenSelectedFrame = CreatePen(PenStyle.PS_Solid, pFrameThickness, 0)
    PenNull = GetStockObject(StockObject.NULL_PEN)
    PenDrawingOrig = SelectObject(hDcDrawing, PenNull)      'Set Pen to NULL_PEN
    DefineDimensions
    
    bmpDrawing = CreateCompatibleBitmap(hDcDisplaySurface, pxWidth, pxHeight)  'This line creates a blank bitmap for doublebuffered drawing compatible with pDisplaySurface
    Set pDisplaySurface.Picture = HandleToPicture(bmpDrawing, PICTYPE_BITMAP)

    BrushDrawingTgt = CreateSolidBrush(TgtRGB)
    BrushDrawingHit = CreateSolidBrush(HitRGB)
    

    RedrawControl
End Sub


Private Sub Class_Terminate()
    'Clean up
    HeapDestroy hHeap                  ' Heap memory object
    
    SelectObject hDcDrawing, PenDrawingOrig     'Set original Pen before destroying

    DeleteObject BrushDrawingTgt
    DeleteObject BrushDrawingHit
    DeleteObject BrushTransparent
    
    DeleteObject bmpDrawing
    DeleteObject PenFrame
    DeleteObject PenSelectedFrame
    DeleteObject PenNull
    
    ReleaseDC hWndForm, hDcDisplaySurface
    DeleteDC hDcDrawing
    
End Sub

Private Sub pDisplaySurface_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button <> 1 Then
        'Right button clicked, x Out tile
        Me.xOut = Not Me.xOut
        Exit Sub
    End If
    
    If pTileGroup Is Nothing Then Exit Sub
    
    Dim NewState As Boolean
    Dim tile As cColorTile
    Dim i As Integer
    
    NewState = Not pFrame
    
    If NewState Then
        'This tile is beeing selected. Deselect all others
        For i = 1 To pTileGroup.count
            Set tile = pTileGroup(i)
            If tile Is Me Then
                LabScope.SetColors pTgtLabL, pTgtLabA, pTgtLabB, pHitLabA, pHitLabB
                LScope.SetColors pTgtLabL, pTgtLabA, pTgtLabB, pHitLabL
            Else
                tile.Frame = False
            End If
        Next i
        Frame = True
    End If
    
    
End Sub



Public Sub RedrawControl()
    'Select resources
    bmpDrawingOrig = SelectObject(hDcDrawing, bmpDrawing)
    
    
    
    'Redraw the control
   
    BrushDrawingOrig = SelectObject(hDcDrawing, BrushDrawingHit)
    

    result = FillRect(hDcDrawing, LocalRect, BrushDrawingTgt)
    
    result = Polygon(hDcDrawing, VarPtr(Triangle(0)), 3)
    
    SelectObject hDcDrawing, BrushTransparent

    'Draw xOut
    If pXout Then
        SelectObject hDcDrawing, PenXout
        MoveToEx hDcDrawing, 0, 0, 0
        LineTo hDcDrawing, pxWidth, pxHeight
    End If

    If pFrame Then
        'Draw selection border
        SelectObject hDcDrawing, PenSelectedFrame
        Rectangle hDcDrawing, 0, 0, pxWidth, pxHeight
    ElseIf pDisplaySurface.BorderStyle = fmBorderStyleSingle Then
        'Draw border
        SelectObject hDcDrawing, PenFrame
        Rectangle hDcDrawing, 0, 0, pxWidth, pxHeight
    End If
    
    SelectObject hDcDrawing, PenNull
    SelectObject hDcDrawing, BrushDrawingOrig
    
    
    'Draw text
    SetBkMode hDcDrawing, BkMode.TRANSPARENT
    result = TextOut(hDcDrawing, 0, 0, "Tgt", 3)
    result = TextOut(hDcDrawing, pxWidth - 25, pxHeight - 20, "Hit", 3)
    
    
    
    result = BitBlt(hDcDisplaySurface, GlobalRect.topleft.x, GlobalRect.topleft.y, pxWidth, pxHeight, hDcDrawing, 0, 0, SRCCOPY)
    
    'Release resources
    SelectObject hDcDrawing, bmpDrawingOrig
End Sub

Private Sub DefineDimensions()
Dim xPixPrPt As Double
Dim yPixPrPt As Double

    
    'Set Scope rectangle for use when invalidating paint area
    xPixPrPt = GetDeviceCaps(hDcDrawing, LOGPIXELSX) / 72
    yPixPrPt = GetDeviceCaps(hDcDrawing, LOGPIXELSY) / 72
    GlobalRect.topleft.x = pDisplaySurface.Left * xPixPrPt
    GlobalRect.topleft.y = pDisplaySurface.Top * yPixPrPt
    GlobalRect.btmRight.x = (pDisplaySurface.Left + pDisplaySurface.Width) * xPixPrPt
    GlobalRect.btmRight.y = (pDisplaySurface.Top + pDisplaySurface.Height) * yPixPrPt
    
    pxWidth = pDisplaySurface.Width * xPixPrPt
    pxHeight = pDisplaySurface.Height * yPixPrPt
  
    LocalRect.topleft.x = 0
    LocalRect.topleft.y = 0
    LocalRect.btmRight.x = pxWidth
    LocalRect.btmRight.y = pxHeight
    
    Triangle(0).x = 0
    Triangle(0).y = pxHeight
    Triangle(1).x = pxWidth
    Triangle(1).y = 0
    Triangle(2).x = pxWidth
    Triangle(2).y = pxHeight
    
End Sub

Sub UpdateTgtBrush()
    'After a change of target color, convert the color from LAB to RGB, delete the old brush and create a new one
    Dim sR As Double
    Dim sG As Double
    Dim sB As Double
    LAB2sRGB pTgtLabL, pTgtLabA, pTgtLabB, sR, sG, sB
    TgtRGB = RGB(sR, sG, sB)
    
    DeleteObject BrushDrawingTgt
    BrushDrawingTgt = CreateSolidBrush(TgtRGB)
End Sub
Sub UpdateHitBrush()
    'After a change of hit color, convert the color from LAB to RGB, delete the old brush and create a new one
    Dim sR As Double
    Dim sG As Double
    Dim sB As Double
    LAB2sRGB pHitLabL, pHitLabA, pHitLabB, sR, sG, sB
    HitRGB = RGB(sR, sG, sB)
    
    DeleteObject BrushDrawingHit
    BrushDrawingHit = CreateSolidBrush(HitRGB)
End Sub


Private Sub Find_Window_handle()
Dim s As String
Dim pS As LongPtr
Dim strLength As Long
    pS = HeapAlloc(hHeap, 0, 255)

    hWndForm = FindWindow("ThunderDFrame", pParent.Caption)
    strLength = GetWindowText(hWndForm, pS, 255)
    
    s = Mem_ReadString(pS, strLength)
    
    HeapFree hHeap, 0, pS
End Sub


''------------------------------------------------------
' Procedure : HandleToPicture
' Purpose   : Creates a StdPicture object to wrap a GDI
'             image handle
' Credit    : Eduardo A. Morcillo @ https://www.mvps.org/emorcillo/en/code/vb6/loadimagegdip.shtml
'------------------------------------------------------
'
Private Function HandleToPicture( _
   ByVal hGDIHandle As LongPtr, _
   ByVal ObjectType As PictureTypeConstants, _
   Optional ByVal hPal As LongPtr = 0) As StdPicture
Dim tPictDesc As PICTDESC
Dim IID_IPicture As GUID
Dim oPicture As IPicture
    
   ' Initialize the PICTDESC structure
   With tPictDesc
      .cbSizeOfStruct = Len(tPictDesc)
      .picType = ObjectType
      .hgdiObj = hGDIHandle
      .hPalOrXYExt = hPal
   End With
    
   ' Initialize the IPicture interface ID
   With IID_IPicture
      .Data1 = &H7BF80981
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(3) = &HAA
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
   End With
    
   ' Create the object
   OleCreatePictureIndirect tPictDesc, IID_IPicture, _
                            True, oPicture
    
   ' Return the picture object
   Set HandleToPicture = oPicture
        
End Function

Public Function AddToGroup(ColorTile As cColorTile)
    If ColorTile Is Nothing Then Err.Raise vbObjectError + errInvalidPropertyValue, "cColorTile.AddToGroup", "No object supplied"
    If Not (ColorTile.TileGroup Is Nothing) Then Err.Raise vbObjectError + errAlreadyMemberOfGroup, "cColorTile.AddToGroup", "ColorTile already member of group"
    
    If pTileGroup Is Nothing Then
        Set pTileGroup = New Collection
        pTileGroup.Add Me
    End If
    pTileGroup.Add ColorTile
    
    ColorTile.TileGroup = pTileGroup
End Function
