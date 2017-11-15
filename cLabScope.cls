VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLabScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Graphic methods
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function MoveToEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpPoint As LongPtr) As Boolean
Private Declare PtrSafe Function LineTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Boolean
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal hbmp As LongPtr, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As LongPtr, ByRef lpbi As Any, ByVal uUsage As Long) As Long
Private Declare PtrSafe Function SetDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal hbmp As LongPtr, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As LongPtr, ByRef lpbmi As Any, ByVal fuColorUse As Long) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As Any, bErase As Boolean) As Boolean
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hdcDest As LongPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Boolean
Private Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hdcDest As LongPtr, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As LongPtr, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal dwRop As Long) As Boolean
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, ByRef lpRect As Any, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Boolean
Private Declare PtrSafe Function CopyImage Lib "gdi32" (ByVal hImage As LongPtr, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuFlags As Long) As LongPtr
Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function Ellipse Lib "gdi32" (ByVal hdc As LongPtr, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Boolean
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal fnObject As Long) As LongPtr
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As LongPtr
Private Declare PtrSafe Function Rectangle Lib "gdi32" (ByVal hdc As LongPtr, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Boolean
Private Declare PtrSafe Function TabbedTextOut Lib "user32" Alias "TabbedTextOutW" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, ByRef lpnTabStopPositions As Any, ByVal nTabOrigin As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal iBkMode As Long) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long

'Memory management methods
Private Declare PtrSafe Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As LongPtr
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Boolean
Private Declare PtrSafe Function HeapDestroy Lib "kernel32" (ByVal hHeap As LongPtr) As Boolean

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr


Private Type PICTDESC     'For use when creating OLE pictureobject
   cbSizeOfStruct As Long
   picType As Long
   hgdiObj As LongPtr
   hPalOrXYExt As LongPtr
End Type

Private Type GUID          'For use when creating OLE pictureobject
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7)  As Byte
End Type

    
Private Type BITMAPINFOHEADER ' 40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD           '4 bytes
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO       ' 44 bytes
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type POINT
    x As Long
    y As Long
End Type
Private Type RECT
    topleft As POINT
    btmRight As POINT
End Type
    
    
Enum DIBColors
   DIB_RGB_COLORS = &H0
End Enum

Enum DeviceCap
    LOGPIXELSX = 88 ' Logical pixels inch in X
    LOGPIXELSY = 90 ' Logical pixels inch in Y
End Enum

Enum TernaryRasterOperations
    SRCCOPY = &HCC0020
    SRCPAINT = &HEE0086
    SRCAND = &H8800C6
    SRCINVERT = &H660046
    SRCERASE = &H440328
    NOTSRCCOPY = &H330008
    NOTSRCERASE = &H1100A6
    MERGECOPY = &HC000CA
    MERGEPAINT = &HBB0226
    PATCOPY = &HF00021
    PATPAINT = &HFB0A09
    PATINVERT = &H5A0049
    DSTINVERT = &H550009
    BLACKNESS = &H42
    WHITENESS = &HFF0062
End Enum
Enum PenStyle
    PS_COSMETIC = &H0
    PS_ENDCAP_ROUND = &H0
    PS_JOIN_ROUND = &H0
    PS_Solid = &H0
    PS_DASH = &H1
    PS_DOT = &H2
    PS_DASHDOT = &H3
    PS_DASHDOTDOT = &H4
    PS_NULL = &H5
    PS_INSIDEFRAME = &H6
    PS_USERSTYLE = &H7
    PS_ALTERNATE = &H8
    PS_ENDCAP_SQUARE = &H100
    PS_ENDCAP_FLAT = &H200
    PS_JOIN_BEVEL = &H1000
    PS_JOIN_MITER = &H2000
    PS_GEOMETRIC = &H10000
End Enum

Enum StockObject
    WHITE_BRUSH = &H0
    LTGRAY_BRUSH = &H1
    GRAY_BRUSH = &H2
    DKGRAY_BRUSH = &H3
    BLACK_BRUSH = &H4
    NULL_BRUSH = &H5
    HOLLOW_BRUSH = &H5
    WHITE_PEN = &H6
    BLACK_PEN = &H7
    NULL_PEN = &H8
    OEM_FIXED_FONT = &HA
    ANSI_FIXED_FONT = &HB
    ANSI_VAR_FONT = &HC
    SYSTEM_FONT = &HD
    DEVICE_DEFAULT_FONT = &HE
    DEFAULT_PALETTE = &HF
    SYSTEM_FIXED_FONT = &H10
    DEFAULT_GUI_FONT = &H11
    DC_BRUSH = &H12
    DC_PEN = &H13
End Enum

Enum PictureTypeConstants
    PICTYPE_UNINITIALIZED = -1
    PICTYPE_NONE = 0
    PICTYPE_BITMAP = 1
    PICTYPE_METAFILE = 2
    PICTYPE_ICON = 3
    PICTYPE_ENHMETAFILE = 4
End Enum
 
Enum BkMode
    OPAQUE = 0
    TRANSPARENT = 1
End Enum

Enum LocalErrors
    errInvalidPropertyValue = 1
End Enum

Private pParent As Object
Private pDisplaySurface As MSForms.Image
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Property variables
Private pBackColor As Long
Private pForeColor As Long
Private pSpokes As Long
Private pRings As Long
Private pPadding As Long
Private pLineWidth As Long
Private pReticleSize As Double
Private pReticleColor As Long
Private pUnitsPrDivision As Double
Private pHitLabA As Double
Private pHitLabB As Double
Private pTgtLabL As Double
Private pTgtLabA As Double
Private pTgtLabB As Double
Private pLabExaggeration As Double
Private pScrollBeyondLimit As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private hDcDisplaySurface As LongPtr
Private hDcDrawing As LongPtr
Private hDcBackground As LongPtr
Private hDcLabSurface As LongPtr

Private bmpLabSurfaceOrig As LongPtr
Private bmpBackgroundOrig As LongPtr


Private BackgroundBrush As LongPtr
Private ReticleBrush As LongPtr
Private ForegroundPen As LongPtr
Private LightPen As LongPtr



Private bmpDrawingOrig As LongPtr
Private hHeap As LongPtr
Private bmpLabSurface As LongPtr
Private bmpBackground As LongPtr
Private bmpDrawing As LongPtr
Private hWndForm As LongPtr
Private GlobalRect As RECT
Private LocalRect As RECT
Private CenterPixelX As Long
Private CenterPixelY As Long
Private Radius As Double
Private pxWidth As Long
Private pxHeight As Long
Private pxSize As Long
Private ReticleSizeHalf As Double
Private ringDist As Double
Private TgtAmin As Double
Private TgtAmax As Double
Private TgtBmin As Double
Private TgtBmax As Double


Private bi As BITMAPINFO




''''''''''''''''''''''
' HitLabA property
''''''''''''''''''''''
Public Property Get HitLabA() As Double
    HitLabA = pHitLabA
End Property
Public Property Let HitLabA(Value As Double)
    pHitLabA = Value
    RedrawControl
End Property


''''''''''''''''''''''
' HitLabB property
''''''''''''''''''''''
Public Property Get HitLabB() As Double
    HitLabB = pHitLabB
End Property
Public Property Let HitLabB(Value As Double)
    pHitLabB = Value
    RedrawControl
End Property

''''''''''''''''''''''
' Background color property
''''''''''''''''''''''
Public Property Get BackColor() As Double
    BackColor = pBackColor
End Property
Public Property Let BackColor(Value As Double)
    pBackColor = Value
    RedrawControl
End Property

''''''''''''''''''''''
' Foreground color property
''''''''''''''''''''''
Public Property Get ForeColor() As Double
    ForeColor = pForeColor
End Property
Public Property Let ForeColor(Value As Double)
    pForeColor = Value
    DeleteObject (ForegroundPen)
    ForegroundPen = CreatePen(PS_Solid, pLineWidth, pForeColor)
    Draw_Background
    RedrawControl
End Property


''''''''''''''''''''''
' Spokes property
''''''''''''''''''''''
Public Property Get Spokes() As Double
    Spokes = pSpokes
End Property
Public Property Let Spokes(Value As Double)
    pSpokes = Value
    Draw_Background
    RedrawControl
End Property


''''''''''''''''''''''
' Rings property
''''''''''''''''''''''
Public Property Get Rings() As Double
    Rings = pRings
End Property
Public Property Let Rings(Value As Double)
    pRings = Value
    DefineDimensions
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property


''''''''''''''''''''''
' Padding property
''''''''''''''''''''''
Public Property Get Padding() As Double
    Padding = pPadding
End Property
Public Property Let Padding(Value As Double)
    If Value < Radius Then pPadding = Value _
                      Else pPadding = Radius
    DefineDimensions
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property



''''''''''''''''''''''
' LineWidth property
''''''''''''''''''''''
Public Property Get LineWidth() As Double
    LineWidth = pLineWidth
End Property
Public Property Let LineWidth(Value As Double)
    pLineWidth = Value
    DeleteObject (ForegroundPen)
    ForegroundPen = CreatePen(PS_Solid, pLineWidth, pForeColor)
    Draw_Background
    RedrawControl
End Property

''''''''''''''''''''''
' Reticlesize property
''''''''''''''''''''''
Public Property Get ReticleSize() As Double
    ReticleSize = pReticleSize
End Property
Public Property Let ReticleSize(Value As Double)
    pReticleSize = Value
    ReticleSizeHalf = pReticleSize / 2
    RedrawControl
End Property

''''''''''''''''''''''
' TgtLabL property
''''''''''''''''''''''
Public Property Get TgtLabL() As Double
    TgtLabL = pTgtLabL
End Property
Public Property Let TgtLabL(Value As Double)
    pTgtLabL = Value
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property
''''''''''''''''''''''
' TgtLabA property
''''''''''''''''''''''
Public Property Get TgtLabA() As Double
    TgtLabA = pTgtLabA
End Property
Public Property Let TgtLabA(Value As Double)
    pTgtLabA = Value
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property
''''''''''''''''''''''
' TgtLabB property
''''''''''''''''''''''
Public Property Get TgtLabB() As Double
    TgtLabB = pTgtLabB
End Property
Public Property Let TgtLabB(Value As Double)
    pTgtLabB = Value
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property


''''''''''''''''''''''
' LabExaggeration property
''''''''''''''''''''''
Public Property Get LabExaggeration() As Double
    LabExaggeration = pLabExaggeration
End Property
Public Property Let LabExaggeration(Value As Double)
    pLabExaggeration = Value
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property


''''''''''''''''''''''
' UnitsPrDivision property
''''''''''''''''''''''
Public Property Get UnitsPrDivision() As Double
    UnitsPrDivision = pUnitsPrDivision
End Property
Public Property Let UnitsPrDivision(Value As Double)
    If Value > 0 Then pUnitsPrDivision = Value _
                 Else Err.Raise vbObjectError + errInvalidPropertyValue, "cScope.UnitsPrDivision", "UnitsPrDivision should be a double greater than 0"
    
    DefineDimensions
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property

''''''''''''''''''''''
' ScrollBeyondLimit property
''''''''''''''''''''''
Public Property Get ScrollBeyondLimit() As Double
    ScrollBeyondLimit = pScrollBeyondLimit
End Property
Public Property Let ScrollBeyondLimit(Value As Double)
    pScrollBeyondLimit = Value
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Property



Public Sub InitiateProperties(ByRef Parent As Object, _
                              ByRef DisplaySurface As MSForms.Image, _
                              Optional HitLabA = 0#, _
                              Optional HitLabB = 0#, _
                              Optional BackColor = &H8000000F, _
                              Optional ForeColor = &H80000012, _
                              Optional ReticleColor = &HFF7F00, _
                              Optional Spokes = 4, _
                              Optional Rings = 3, _
                              Optional Padding = 0, _
                              Optional UnitsPrDivision = 5#, _
                              Optional TgtLabL = 60#, _
                              Optional TgtLabA = 0#, _
                              Optional TgtLabB = 0#, _
                              Optional LabExaggeration = 5#, _
                              Optional ScrollBeyondLimit = True)
Dim bufferSize As Long
Dim result As Long
    hHeap = HeapCreate(0, 0, 0)  ' Create a heap memory object to be used for allocating memory to usercontrol images (Max size 2^23 = 8Mb)
    
    'Set connection between me and my display surface
    Set pParent = Parent
    Set pDisplaySurface = DisplaySurface
    
    'Set properties
    pHitLabA = HitLabA
    pHitLabB = HitLabB
    pBackColor = BackColor
    pForeColor = ForeColor
    pReticleColor = ReticleColor
    pSpokes = Spokes
    pRings = Rings
    pPadding = Padding
    pUnitsPrDivision = UnitsPrDivision
    pTgtLabL = TgtLabL
    pTgtLabA = TgtLabA
    pTgtLabB = TgtLabB
    pLabExaggeration = LabExaggeration
    pScrollBeyondLimit = ScrollBeyondLimit

    'Convert system colors to RGB values
    If (pBackColor And &H80000000) = &H80000000 Then pBackColor = GetSysColor(pBackColor And &HFFFFFFF)      'Convert system color to RGB
    If (pForeColor And &H80000000) = &H80000000 Then pForeColor = GetSysColor(pForeColor And &HFFFFFFF)      'Convert system color to RGB
    If (pReticleColor And &H80000000) = &H80000000 Then pReticleColor = GetSysColor(pReticleColor And &HFFFFFFF)      'Convert system color to RGB
    

    
    'Create a Device Context to draw on and allocate memory for background image
    Find_Window_handle
    hDcDisplaySurface = GetDC(hWndForm)
    hDcDrawing = CreateCompatibleDC(0)            ' Create DeviceContext that matches the screen that image1 is shown on
    hDcBackground = CreateCompatibleDC(0)         ' Create DeviceContext that matches the screen that image1 is shown on
    hDcLabSurface = CreateCompatibleDC(0)
    
    SetBkMode hDcBackground, BkMode.TRANSPARENT
    
    DefineDimensions 'Set the controls dimensions. Requires hDcDrawing to be initialised
    
    bmpDrawing = CreateCompatibleBitmap(hDcDisplaySurface, pxWidth, pxHeight)  'This line creates a blank bitmap for doublebuffered drawing compatible with pDisplaySurface
    bmpBackground = CreateCompatibleBitmap(hDcDisplaySurface, pxWidth, pxHeight) 'This line creates a blank bitmap for background compatible with pDisplaySurface
    bmpLabSurface = CreateCompatibleBitmap(hDcDisplaySurface, 100, 100) 'This line creates a blank bitmap for background compatible with pDisplaySurface
    Set pDisplaySurface.Picture = HandleToPicture(bmpDrawing, PICTYPE_BITMAP)
    
    
    BackgroundBrush = CreateSolidBrush(pBackColor)
    ReticleBrush = CreateSolidBrush(pReticleColor)
    ForegroundPen = CreatePen(PS_Solid, 1, pForeColor)
    LightPen = CreatePen(PS_Solid, pLineWidth, RGB(255, 255, 255))

    
    'Create the background for the control
    CalculateLabSpace
    Draw_Background
    RedrawControl
End Sub



Private Sub Class_Terminate()
    'Clean up
    HeapDestroy hHeap                  ' Heap memory object
    
    DeleteObject BackgroundBrush
    DeleteObject ReticleBrush
    DeleteObject ForegroundPen
    DeleteObject LightPen
    
    DeleteObject bmpDrawing
    DeleteObject bmpBackground
    DeleteObject bmpLabSurface

    ReleaseDC hWndForm, hDcDisplaySurface
    DeleteDC hDcDrawing
    DeleteDC hDcBackground                       ' Gdi32 Device Context created earlier
    DeleteDC hDcLabSurface                       ' Gdi32 Device Context created earlier
End Sub



Public Sub RedrawControl()
    Dim result As Long
    Dim i As Integer
    'Point the Device Context at a copy of the original background picture
    
    bmpDrawingOrig = SelectObject(hDcDrawing, bmpDrawing)
    bmpBackgroundOrig = SelectObject(hDcBackground, bmpBackground)
    
    result = BitBlt(hDcDrawing, 0, 0, pxWidth, pxHeight, hDcBackground, 0, 0, SRCCOPY)

    
    'Redraw the control
    Dim BrushDrawingOrig As LongPtr
    Dim hitPixelX As Long
    Dim hitPixelY As Long
    hitPixelX = r2r(TgtAmin, TgtAmax, pHitLabA, pPadding, pxWidth - pPadding)
    hitPixelY = r2r(TgtBmin, TgtBmax, pHitLabB, pxHeight - pPadding, pPadding)
    
    BrushDrawingOrig = SelectObject(hDcDrawing, ReticleBrush)
    result = Ellipse(hDcDrawing, hitPixelX - ReticleSizeHalf, _
                                 hitPixelY - ReticleSizeHalf, _
                                 hitPixelX + ReticleSizeHalf, _
                                 hitPixelY + ReticleSizeHalf)
    SelectObject hDcDrawing, BrushDrawingOrig
    
    
    
    'Draw text
    Dim TabStops(1 To 2) As Long
    TabStops(1) = 160
    TabStops(2) = pxSize / 2
    Dim s As String
    Dim sW As String
    
    SetBkMode hDcDrawing, BkMode.TRANSPARENT
    s = "Hit a*" & vbTab & Format(pHitLabA, "0.00")
    sW = StrConv(s, vbUnicode)
    result = TabbedTextOut(hDcDrawing, 0, pxSize - 40, sW, Len(s), 2, TabStops(1), 0)
    
    s = "Hit b*" & vbTab & Format(pHitLabB, "0.00")
    sW = StrConv(s, vbUnicode)
    result = TabbedTextOut(hDcDrawing, 0, pxSize - 20, sW, Len(s), 2, TabStops(1), 0)
    
    
    result = BitBlt(hDcDisplaySurface, GlobalRect.topleft.x, GlobalRect.topleft.y, pxWidth, pxHeight, hDcDrawing, 0, 0, SRCCOPY)
    
    SelectObject hDcDrawing, bmpDrawingOrig
    SelectObject hDcBackground, bmpBackgroundOrig
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
    pxSize = pxWidth
    If pxHeight < pxSize Then pxSize = pxHeight
    CenterPixelX = pxSize / 2
    CenterPixelY = pxSize / 2
    
    LocalRect.topleft.x = 0
    LocalRect.topleft.x = 0
    LocalRect.btmRight.x = pxWidth
    LocalRect.btmRight.y = pxHeight
    
    
    'Set radius equal to min(CenterPixelX,CenterPixelY)
    Radius = pxSize / 2
    pReticleSize = 20
    ReticleSizeHalf = 10
    
    If pRings > 0 Then ringDist = (Radius - pPadding) / pRings _
                  Else ringDist = Radius - pPadding
                    
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


Private Sub Draw_Background()
Dim result As Long
Dim i As Integer
    Dim BrushBackgroundOrig As LongPtr
    Dim PenBackgroundOrig As LongPtr
    Dim backgroundGradient As LongPtr
    Dim r As Long
    Dim c As Long
   
   
    
    
    'Draw the gradient background
    Dim BytesPrScanline As Long
    Dim RGBtuple(0 To 3) As Byte
    bi.bmiHeader.biSize = Len(bi.bmiHeader)

    
    result = GetDIBits(hDcLabSurface, bmpLabSurface, 0, 0, 0, bi, DIB_RGB_COLORS)
    BytesPrScanline = Ceiling(bi.bmiHeader.biWidth * bi.bmiHeader.biBitCount / 8, 8) 'Number of Double Words to each line.

    backgroundGradient = HeapAlloc(hHeap, 0, BytesPrScanline * bi.bmiHeader.biHeight)
    
    result = GetDIBits(hDcLabSurface, bmpLabSurface, 0, bi.bmiHeader.biHeight, backgroundGradient, bi, DIB_RGB_COLORS)
    
    Dim xFract As Double
    Dim yFract As Double
    Dim tgtLabARange As Double
    Dim tgtLabBRange As Double
    Dim LabA As Double
    Dim LabB As Double
    Dim sR As Double
    Dim sG As Double
    Dim sB As Double
    
    tgtLabARange = TgtAmax - TgtAmin
    tgtLabBRange = TgtBmax - TgtBmin
    
    For r = 0 To 99
       yFract = r / 99
       LabB = pTgtLabB + pLabExaggeration * (TgtBmin + yFract * tgtLabBRange - pTgtLabB)
       For c = 0 To 99
            xFract = c / 99
            LabA = pTgtLabA + pLabExaggeration * (TgtAmin + xFract * tgtLabARange - pTgtLabA)
           
            LAB2sRGB pTgtLabL, LabA, LabB, sR, sG, sB
            RGBtuple(0) = sB 'Blue
            RGBtuple(1) = sG 'Green
            RGBtuple(2) = sR 'Red
            RGBtuple(3) = &H0  '***Reserved***
            Mem_Copy ByVal backgroundGradient + BytesPrScanline * r + c * 4, RGBtuple(0), 4
        Next c
    Next r
    
        


   
    result = SetDIBits(hDcLabSurface, bmpLabSurface, 0, bi.bmiHeader.biHeight, backgroundGradient, bi, DIB_RGB_COLORS)
    
    
    bmpBackgroundOrig = SelectObject(hDcBackground, bmpBackground)
    bmpLabSurfaceOrig = SelectObject(hDcLabSurface, bmpLabSurface)
    
    result = FillRect(hDcBackground, LocalRect, BackgroundBrush)
    result = StretchBlt(hDcBackground, 0, 0, pxSize, pxSize, hDcLabSurface, 0, 0, 100, 100, SRCCOPY)

    HeapFree hHeap, 0, backgroundGradient
    


    If pTgtLabL >= 30 Then
        PenBackgroundOrig = SelectObject(hDcBackground, ForegroundPen)
        result = SetTextColor(hDcBackground, pForeColor)
    Else
        PenBackgroundOrig = SelectObject(hDcBackground, LightPen)
        result = SetTextColor(hDcBackground, &HFFFFFF)
    End If
    
    If pTgtLabL >= 45 Then
        result = SetTextColor(hDcDrawing, pForeColor)
    Else
        result = SetTextColor(hDcDrawing, &HFFFFFF)
    End If
    
    
    'Draw the Rings
    Dim TransparentBrush As LongPtr
    TransparentBrush = GetStockObject(NULL_BRUSH And &HFFFFFFF)
    BrushBackgroundOrig = SelectObject(hDcBackground, TransparentBrush)
    
    Dim tgtPixelX As Long
    Dim tgtPixelY As Long
    tgtPixelX = r2r(TgtAmin, TgtAmax, pTgtLabA, Padding, pxSize - Padding)
    tgtPixelY = r2r(TgtBmin, TgtBmax, pTgtLabB, pxSize - Padding, Padding)

    For i = 1 To pRings
        result = Ellipse(hDcBackground, tgtPixelX - i * ringDist, _
                                        tgtPixelY - i * ringDist, _
                                        tgtPixelX + i * ringDist, _
                                        tgtPixelY + i * ringDist)
    Next i
    
    
    
    
    'Draw spokes
    Dim SpokeAngleRad As Double
    
    For i = 0 To pSpokes - 1
        SpokeAngleRad = i * (6.28318530718 / pSpokes)
        MoveToEx hDcBackground, tgtPixelX, tgtPixelY, 0
        LineTo hDcBackground, tgtPixelX + 2 * Radius * (Math.Sin(SpokeAngleRad)), tgtPixelY - 2 * Radius * (Math.Cos(SpokeAngleRad))
    Next i
    
    
    
    'Draw text
    Dim TabStops(1 To 2) As Long
    TabStops(1) = 160
    TabStops(2) = pxSize / 2
    Dim s As String
    Dim sW As String
    
    s = ChrW(&H394) & "E* pr division" & vbTab & Format(UnitsPrDivision, "0.00")
    sW = StrConv(s, vbUnicode)
    result = TabbedTextOut(hDcBackground, 0, 0, sW, Len(s), 2, TabStops(1), 0)
    
    s = "Color Exaggeration" & vbTab & Format(pLabExaggeration, "0.00")
    sW = StrConv(s, vbUnicode)
    result = TabbedTextOut(hDcBackground, 0, 20, sW, Len(s), 2, TabStops(1), 0)
    
    s = "Target a*" & vbTab & Format(pTgtLabA, "0.00")
    sW = StrConv(s, vbUnicode)
    result = TabbedTextOut(hDcBackground, 0, 40, sW, Len(s), 2, TabStops(1), 0)
    
    s = "Target b*" & vbTab & Format(pTgtLabB, "0.00")
    sW = StrConv(s, vbUnicode)
    result = TabbedTextOut(hDcBackground, 0, 60, sW, Len(s), 2, TabStops(1), 0)
    
    
    
    
    'Draw border
    If pDisplaySurface.BorderStyle = fmBorderStyleSingle Then
        Rectangle hDcBackground, 0, 0, pxWidth, pxHeight
    End If
    
    SelectObject hDcBackground, bmpBackgroundOrig
    SelectObject hDcLabSurface, bmpLabSurfaceOrig
    
    SelectObject hDcBackground, BrushBackgroundOrig
    SelectObject hDcBackground, PenBackgroundOrig

End Sub

Sub CalculateLabSpace()
    Dim tgtArange As Double
    Dim tgtBrange As Double
  
    tgtArange = pRings * pUnitsPrDivision * 2
    tgtBrange = pRings * pUnitsPrDivision * 2
    
    'A Component
    If (Not pScrollBeyondLimit) And ((pTgtLabA - tgtArange / 2) < -128) Then
        'Offset due to target beeing close to lower end
        TgtAmin = -128
    
    ElseIf (Not pScrollBeyondLimit) And ((pTgtLabA + tgtArange / 2) > 127) Then
        'Offset due to target beeing close to upper end
        TgtAmin = 127 - tgtArange

    Else
        'No offset
        TgtAmin = pTgtLabA - tgtArange / 2
    End If
    

    'B Component
    If (Not pScrollBeyondLimit) And ((pTgtLabB - tgtBrange / 2) < -128) Then
        'Offset due to target beeing close to lower end
        TgtBmin = -128
    
    ElseIf (Not pScrollBeyondLimit) And ((pTgtLabB + tgtBrange / 2) > 127) Then
        'Offset due to target beeing close to upper end
        TgtBmin = 127 - tgtBrange

    Else
        'No offset
        TgtBmin = pTgtLabB - tgtBrange / 2
    End If


    TgtAmax = TgtAmin + tgtArange
    TgtBmax = TgtBmin + tgtBrange
End Sub

Function r2r(sourceMin As Double, sourceMax As Double, sourceValue As Double, ByVal targetMin As Double, ByVal targetMax As Double) As Double
    'Range to Range - Returns a value a value from a source range to a target range
    r2r = (targetMax - targetMin) * (sourceValue - sourceMin) / (sourceMax - sourceMin) + targetMin
End Function





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

Private Function Ceiling(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(x / Factor) - (x / Factor - Int(x / Factor) > 0)) * Factor
End Function

Private Function Min(ParamArray values() As Variant) As Variant
   Dim minValue, Value As Variant
   minValue = values(0)
   For Each Value In values
       If Value < minValue Then minValue = Value
   Next
   Min = minValue
End Function

Private Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue, Value As Variant
   maxValue = values(0)
   For Each Value In values
       If Value > maxValue Then maxValue = Value
   Next
   Max = maxValue
End Function
