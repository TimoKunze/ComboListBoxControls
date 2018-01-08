Attribute VB_Name = "basMain"
Option Explicit

  Type LOGFONT1
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
'    lfItalic As Byte
'    lfUnderline As Byte
'    lfStrikeOut As Byte
'    lfCharSet As Byte
    lf1 As Long
'    lfOutPrecision As Byte
'    lfClipPrecision As Byte
'    lfQuality As Byte
'    lfPitchAndFamily As Byte
    lf2 As Long
    lfFaceName(1 To 64) As Byte
  End Type

  Type LOGFONT2
'    lfHeight As Long
    lfHeight1 As Byte
    lfHeight2 As Byte
    lfHeight3 As Byte
    lfHeight4 As Byte
'    lfWidth As Long
    lfWidth1 As Byte
    lfWidth2 As Byte
    lfWidth3 As Byte
    lfWidth4 As Byte
'    lfEscapement As Long
    lfEscapement1 As Byte
    lfEscapement2 As Byte
    lfEscapement3 As Byte
    lfEscapement4 As Byte
'    lfOrientation As Long
    lfOrientation1 As Byte
    lfOrientation2 As Byte
    lfOrientation3 As Byte
    lfOrientation4 As Byte
'    lfWeight As Long
    lfWeight1 As Byte
    lfWeight2 As Byte
    lfWeight3 As Byte
    lfWeight4 As Byte
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 64) As Byte
  End Type

  Private Type ENUMLOGFONTEX
    elfLogFont As LOGFONT1
    elfFullName(1 To 128) As Byte
    elfStyle(1 To 64) As Byte
    elfScript(1 To 64) As Byte
  End Type


  Private lastFont As String
  Private targetCtl As Object


  Global lfDefault As LOGFONT1


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT1) As Long
  Private Declare Function EnumFontFamiliesEx Lib "gdi32.dll" Alias "EnumFontFamiliesExW" (ByVal hDC As Long, ByRef lpLogFont As LOGFONT2, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
  Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
  Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long


Sub ByteArray2String(ByteArray() As Byte, OutputString As String)
  OutputString = String$(250, 0)
  lstrcpy StrPtr(OutputString), VarPtr(ByteArray(LBound(ByteArray)))
  OutputString = Left$(OutputString, lstrlen(StrPtr(OutputString)))
End Sub

Function EnumFontFamExProc(ByRef lpElfe As ENUMLOGFONTEX, ByVal lpntme As Long, ByVal FontType As Long, ByVal lParam As Long) As Long
  Dim arrBytes(1 To 348) As Byte
  Dim elfFullName(1 To 128) As Byte
  Dim FaceName As String
  Dim hFont As Long

  CopyMemory VarPtr(arrBytes(1)), VarPtr(lpElfe), 348
  CopyMemory VarPtr(elfFullName(1)), VarPtr(arrBytes(93)), 128

  ByteArray2String elfFullName, FaceName
  If lastFont <> FaceName Then
    lpElfe.elfLogFont.lfWidth = lfDefault.lfWidth
    lpElfe.elfLogFont.lfHeight = lfDefault.lfHeight

    hFont = CreateFontIndirect(lpElfe.elfLogFont)
    If TypeOf targetCtl Is CBLCtlsLibUCtl.ListBox Then
      targetCtl.ListItems.Add FaceName, , hFont
    ElseIf TypeOf targetCtl Is CBLCtlsLibUCtl.ComboBox Then
      targetCtl.ComboItems.Add FaceName, , hFont
    End If
  End If
  lastFont = FaceName

  EnumFontFamExProc = 1
End Function

Sub EnumFonts(ByVal hDC As Long, targetControl As Object, ByVal CharSet As Long)
  Dim lf As LOGFONT2

  lastFont = ""
  Set targetCtl = targetControl
  lf.lfCharSet = CharSet
  EnumFontFamiliesEx hDC, lf, AddressOf EnumFontFamExProc, 0&, 0&
End Sub

' merges <Lo> and <Hi> to a 16 bit number
Function MakeWord(ByVal Lo As Byte, ByVal Hi As Byte) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(Hi), LenB(Hi)

  MakeWord = ret
End Function
