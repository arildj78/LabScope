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
    x As Long
    y As Long
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
    errObjectMissing = 2
    errAlreadyMemberOfGroup = 3
End Enum



'ColorManagement

Public Type LabCOLOR
  L As Long
  a As Long
  b As Long
End Type
Public Type dLabCOLOR
  L As Double
  a As Double
  b As Double
End Type

Public Enum tColorSpace
    LAB = 1
    ProPhoto = 2
    AdobeRGB = 3
    sRGB = 4
End Enum



Public Enum GamutMapping
    LCS_GM_BUSINESS = 1&
    LCS_GM_GRAPHICS = 2&
    LCS_GM_IMAGES = 4&
    LCS_GM_ABS_COLORIMETRIC = 8&
    LCS_GM_Saturation = 1&
    LCS_GM_Rel_Colorimetric = 2&
    LCS_GM_Perceptual = 4&
End Enum

Public Type CIEXYZ
    ciexyzX As Long  'FXPT2DOT30
    ciexyzY As Long  'FXPT2DOT30
    ciexyzZ As Long  'FXPT2DOT30
End Type

Public Type CIEXYZTRIPLE
    ciexyzRed   As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyzBlue  As CIEXYZ
End Type



Public Type ColorProfile
    dwType As Long
    pProfileData As String
    cbDataSize As Long
End Type



'*************************************************
'      From minwindef.h
'*************************************************

Public Const MAX_PATH = 260

'*************************************************
'      From wingdi.h
'*************************************************
'/* Logcolorspace lcsType values */
Public Enum tagLOGCOLORSPACE
    LCS_CALIBRATED_RGB = 0&
    LCS_sRGB = &H73524742                 'ASCII  'sRGB'
    LCS_WINDOWS_COLOR_SPACE = &H57696E20  'ASCII  'Win '    '// Windows default color space
End Enum

Public Type tLOGCOLORSPACE
    lcsSignature As Long
    lcsVersion As Long
    lcsSize As Long
    lcsCSType As Long
    lcsIntent As GamutMapping
    lcsEndpoints As CIEXYZTRIPLE
    lcsGammaRed As Long      'DWORD
    lcsGammaGreen As Long    'DWORD
    lcsGammaBlue As Long     'DWORD
    lcsFilename(1 To MAX_PATH) As Byte    'TCHAR
End Type

'/* Logcolorspace signature */
Public Const LCS_SIGNATURE = &H50534F43    'PSOC'




'*************************************************
'      From Icm.h
'*************************************************
'//
'// Profile types to be used in the PROFILE structure
'//
Public Enum tagPROFILE_TYPE
    PROFILE_FILENAME = 1      '// profile data is NULL terminated filename
    PROFILE_MEMBUFFER = 2     '// profile data is a buffer containing the profile
End Enum

'//
'// Desired access mode for opening profiles
'//
Public Enum tagDesiredAccess
    PROFILE_READ = 1&        '// opened for read access
    PROFILE_READWRITE = 2&   '// opened for read and write access
End Enum

Public Const DONT_USE_EMBEDDED_WCS_PROFILES = 1&

Public Type RGBCOLOR
    red As Integer
    green As Integer
    blue As Integer
    inGamut As Boolean
End Type
Public Type dRGBCOLOR
    red As Double
    green As Double
    blue As Double
End Type
Public Type dTriChannel
    ch0 As Double
    ch1 As Double
    ch2 As Double
End Type

Public Enum COLORTYPE
  COLOR_GRAY = 1
  COLOR_RGB = 2
  COLOR_XYZ = 3
  COLOR_Yxy = 4
  COLOR_Lab = 5
  COLOR_3_CHANNEL = 6
  COLOR_CMYK = 7
  COLOR_5_CHANNEL = 8
  COLOR_6_CHANNEL = 9
  COLOR_7_CHANNEL = 10
  COLOR_8_CHANNEL = 11
  COLOR_NAMED = 12
End Enum

'//
'// Device color data type
'//
Public Enum COLORDATATYPE
    COLOR_BYTE = 1                 '// BYTE per channel. data range [0, 255]
    COLOR_WORD = 2                 '// WORD per channel. data range [0, 65535]
    COLOR_FLOAT = 3                '// FLOAT per channel. IEEE 32-bit floating point
    COLOR_S2DOT13FIXED = 4         '// WORD per channel. data range [-4, +4] using s2.13
    COLOR_10b_R10G10B10A2 = 5      '// Packed WORD per channel.  data range [0, 1]
    COLOR_10b_R10G10B10A2_XR = 6   '// Packed extended range WORD per channel.  data range [-1, 3]
                                   '// using 4.0 scale and -1.0 bias.
    COLOR_FLOAT16 = 7              '// FLOAT16 per channel.
End Enum


'*************************************************
'      From winnt.h
'*************************************************
Public Enum tagShareMode
    FILE_SHARE_READ = 1&        '// opened for read access
    FILE_SHARE_WRITE = 2&   '// opened for read and write access
End Enum




'*************************************************
'      From fileapi.h
'*************************************************
Public Enum tagCreationMode
    CREATE_NEW = 1&           '// opened for read access
    CREATE_ALWAYS = 2&        '// opened for read and write access
    OPEN_EXISTING = 3&        '// opened for read access
    OPEN_ALWAYS = 4&          '// opened for read and write access
    TRUNCATE_EXISTING = 5&    '// opened for read and write access
End Enum





'*************************************************
'      From ICC v4 standard
'*************************************************
Public Type tIccHeader
    ProfileSize As Long
    PreferredCMM As Long
    ProfileVersion As Long
    DeviceClass As Long
    ColorSpace As Long
    PCS As Long
    
    CreatedDateTime(1 To 12) As Byte
    
    signature As Long
    PrimaryPlatformSignature As Long
    Flags As Long
    DeviceManufacturer As Long
    DeviceModel As Long
    
    DeviceAttributes(1 To 8) As Byte
    
    RenderingIntent As Long
    
    IlluminantCIEXYZ(1 To 12) As Byte

    ProfileCreator As Long
    
    ProfileID(1 To 16) As Byte
    reserved(1 To 28) As Byte
End Type


Public Type tIccTagEntry
    'tag As String
    signature As Long
    StringSig As String
    offset As Long
    size As Long
    datatype As String
End Type
Public Type tIccTagTable
    count As Long
    tagEntries() As tIccTagEntry
End Type

Public Type CurveType
    signature As Long
    
    'parametricCurveType
    FunctionType As Integer
    g As Long
    a As Long
    b As Long
    c As Long
    d As Long
    e As Long
    f As Long
    
    'CurveType
    n As Long
    curve() As Long
    
    'implementation specicfic
    size As Long
End Type

Public Type lut16Type
    signature As Long
    i As Byte
    o As Byte
    g As Byte
    e(1 To 9) As Long
    n As Long
    m As Long
    inTables() As Long
    CLUT() As Long
    outTables() As Long
    
    'implementation specicfic
    legacyPCS As Boolean
End Type

Public Type lutBToAType '133 byte
    Acurve() As CurveType             ' 8 byte
    Bcurve() As CurveType             ' 8 byte
    Mcurve() As CurveType             ' 8 byte
    CLUT() As Long                    ' 8 byte
    Matrix(1 To 12) As Long           '48 byte
    signature As Long                 ' 4 byte
    offsetBcurve As Long              ' 4 byte
    offsetMatrix As Long              ' 4 byte
    offsetMcurve As Long              ' 4 byte
    offsetCLUT As Long                ' 4 byte
    offsetAcurve As Long              ' 4 byte
    n As Long                         ' 4 byte
    m As Long                         ' 4 byte
    CLUTgridPoints(0 To 15) As Byte   '16 byte
    i As Byte                         ' 1 byte
    o As Byte                         ' 1 byte
    g As Byte                         ' 1 byte
    CLUTchannels As Byte              ' 1 byte
    CLUTbitCount As Byte              ' 1 byte
End Type
Public Type lutAToBType '133 byte
    Acurve() As CurveType             ' 8 byte
    Bcurve() As CurveType             ' 8 byte
    Mcurve() As CurveType             ' 8 byte
    CLUT() As Long                    ' 8 byte
    Matrix(1 To 12) As Long           '48 byte
    signature As Long                 ' 4 byte
    offsetBcurve As Long              ' 4 byte
    offsetMatrix As Long              ' 4 byte
    offsetMcurve As Long              ' 4 byte
    offsetCLUT As Long                ' 4 byte
    offsetAcurve As Long              ' 4 byte
    n As Long                         ' 4 byte
    m As Long                         ' 4 byte
    CLUTgridPoints(0 To 15) As Byte   '16 byte
    i As Byte                         ' 1 byte
    o As Byte                         ' 1 byte
    g As Byte                         ' 1 byte
    CLUTchannels As Byte              ' 1 byte
    CLUTbitCount As Byte              ' 1 byte
End Type


'Implementation specific
Public Enum FunctionType
    lut16Table = -2
    curve = -1
    PARAMg = 0
    PARAMgab = 1
    PARAMgabc = 2
    PARAMgabcd = 3
    PARAMgabcdef = 4
End Enum

Public Enum RenderingIntent
    perceptual = 0
    relative = 1
    saturation = 2
    absolute = 3
End Enum
    
Public Enum TranslateDirection
    a2b = 0
    b2a = 1
End Enum

