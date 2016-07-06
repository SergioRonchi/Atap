Attribute VB_Name = "mdlLocalInfo"
Option Explicit
Public Declare Function GetLocaleInfo Lib "kernel32" _
    Alias "GetLocaleInfoA" (ByVal Locale As Long, _
    ByVal LCType As Long, ByVal lpLCData As String, _
    ByVal cchData As Long) As Long

Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_IDATE = &H21 ' short date format ordering
Public Const LOCALE_SLANGUAGE = &H2 ' localized name of language
Public Const LOCALE_SCOUNTRY = &H6 ' localized name of country
Public Const LOCALE_SCURRENCY = &H14 ' local monetary symbol
Public Const LOCALE_ILDATE = &H22 ' Long date format ordering
Public Const LOCALE_SDECIMAL              As Long = &HE     'decimal separator
Public Const LOCALE_STHOUSAND             As Long = &HF     'thousand separator

Public Const LOCALE_SDATE As Long = &H1D    'date separator
Public Const LOCALE_SSHORTDATE            As Long = &H1F    'short date format string
Public Const LOCALE_SLONGDATE             As Long = &H20    'long date format string

'In a module
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&
Private Const TMPF_FIXED_PITCH = &H1
Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4
Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type
Type NEWTEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
   ntmFlags As Long
   ntmSizeEM As Long
   ntmCellHeight As Long
   ntmAveWidth As Long
End Type
Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, LParam As Any) As Long
Private m_Fonts As Collection
Private m_countFont As Long

Public Function GetInstalledFontFamilies(hdc As Long)
  If m_Fonts Is Nothing Then
    Set m_Fonts = New Collection
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, ByVal 0&
    pSort
End If
   Set GetInstalledFontFamilies = m_Fonts
End Function
Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, LParam As Long) As Long
   Dim FaceName As String
   
   
  'convert the returned string to Unicode
   FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
   
   m_Fonts.Add Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
   m_countFont = m_countFont + 1
  'continue enumeration
   EnumFontFamProc = 1
End Function


Sub GetTheLocaleInfo(ByRef LongDate As String, ByRef ShortDate As String, _
                     ByRef Lang As String, ByRef Country As String, _
                     ByRef Money As String, ByRef DecSep As String, ByRef MiglSep As String)
    Dim strBuffer As String * 100
    Dim lngReturn As Long
    Dim strResult As String
    Dim msg As String
    
    lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, _
        strBuffer, 99)
    ShortDate = LCase(LPSTRToVBString(strBuffer))
    
    
    
    lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE, _
        strBuffer, 99)
    LongDate = LPSTRToVBString(strBuffer)
    
    
    lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLANGUAGE, _
        strBuffer, 99)
    strResult = LPSTRToVBString(strBuffer)
    Lang = strResult
    
    lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCOUNTRY, _
        strBuffer, 99)
    strResult = LPSTRToVBString(strBuffer)
    Country = strResult
    
    lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, _
        strBuffer, 99)
    strResult = LPSTRToVBString(strBuffer)
    Money = strResult
    
    
    lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, _
        strBuffer, 99)
    strResult = LPSTRToVBString(strBuffer)
    DecSep = strResult
    
  
    
    lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, _
        strBuffer, 99)
    strResult = LPSTRToVBString(strBuffer)
    MiglSep = strResult
End Sub

Public Function LPSTRToVBString(ByVal s As String) As String
    Dim nullpos As Integer
    nullpos = InStr(s, Chr(0))
    If nullpos > 0 Then
        LPSTRToVBString = Left(s, nullpos - 1)
    Else
        LPSTRToVBString = ""
    End If
End Function



Private Sub pSort()

Dim bMoved As Boolean
Dim nCtr As Integer
Dim strTemp As String

Do
   bMoved = False
   For nCtr = 1 To m_countFont - 1
       If m_Fonts(nCtr) > m_Fonts(nCtr + 1) Then
           bMoved = True
           strTemp = m_Fonts(nCtr)
           m_Fonts.Remove nCtr
           'm_Fonts.Add m_Fonts(nCtr + 1), , nCtr
           m_Fonts.Add strTemp, , , nCtr
       End If
   Next
Loop Until bMoved = False
   
End Sub
