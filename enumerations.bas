Attribute VB_Name = "Enumerations"
Public Enum DeviceCap
    LOGPIXELSX = 88 ' Logical pixels inch in X
    LOGPIXELSY = 90 ' Logical pixels inch in Y
End Enum

Public Enum TernaryRasterOperations
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

Public Type PICTDESC     'For use when creating OLE pictureobject
   cbSizeOfStruct As Long
   picType As Long
   hgdiObj As LongPtr
   hPalOrXYExt As LongPtr
End Type

Public Type GUID          'For use when creating OLE pictureobject
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7)  As Byte
End Type

Public Type POINT
    X As Long
    Y As Long
End Type

Public Type RECT
    topleft As POINT
    btmRight As POINT
End Type

Public Enum StockObject
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

Public Enum PictureTypeConstants
    PICTYPE_UNINITIALIZED = -1
    PICTYPE_NONE = 0
    PICTYPE_BITMAP = 1
    PICTYPE_METAFILE = 2
    PICTYPE_ICON = 3
    PICTYPE_ENHMETAFILE = 4
End Enum

Public Enum PenStyle
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

Public Enum BkMode
    OPAQUE = 0
    TRANSPARENT = 1
End Enum

Public Type BITMAPINFOHEADER ' 40 bytes
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

Public Type RGBQUAD           '4 bytes
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO       ' 44 bytes
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Enum DIBColors
   DIB_RGB_COLORS = &H0
End Enum

Public Enum LocalErrors
    errInvalidPropertyValue = 1
End Enum

