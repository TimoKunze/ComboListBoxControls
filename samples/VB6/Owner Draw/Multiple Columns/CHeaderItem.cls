VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHeaderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Public Caption As String
  Public ItemID As Long


  Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal nXEnd As Long, ByVal nYEnd As Long) As Long
  Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long


Public Sub OwnerDrawItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Const COLOR_BTNFACE = 15
  Const COLOR_BTNHIGHLIGHT = 20
  Const COLOR_BTNSHADOW = 16
  Const COLOR_3DFACE = COLOR_BTNFACE
  Const COLOR_3DSHADOW = COLOR_BTNSHADOW
  Const COLOR_BTNTEXT = 18
  Const DT_CENTER = &H1
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const PS_SOLID = 0
  Const TRANSPARENT = 1
  Dim flags As Long
  Dim hPen As Long
  Dim hPreviousPen As Long
  Dim oldBkMode As Long
  Dim oldTextColor As Long
  Dim rcItem As RECT

  LSet rcItem = drawingRectangle
  If Not (comboItem Is Nothing) Then
    ' draw item background
    Call FillRect(hDC, rcItem, GetSysColorBrush(COLOR_3DFACE))

    ' now draw the text
    oldTextColor = SetTextColor(hDC, GetSysColor(COLOR_BTNTEXT))
    oldBkMode = SetBkMode(hDC, TRANSPARENT)
    flags = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Call DrawText(hDC, StrPtr(Caption), Len(Caption), rcItem, flags)
    ' draw 3D border
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT))
    If hPen Then
      hPreviousPen = SelectObject(hDC, hPen)
      Call MoveToEx(hDC, rcItem.Left, rcItem.Bottom, 0)
      Call LineTo(hDC, rcItem.Left, rcItem.Top)
      Call LineTo(hDC, rcItem.Right, rcItem.Top)
      Call SelectObject(hDC, hPreviousPen)
      Call DeleteObject(hPen)
    End If
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DSHADOW))
    If hPen Then
      hPreviousPen = SelectObject(hDC, hPen)
      Call MoveToEx(hDC, rcItem.Right - 1, rcItem.Top, 0)
      Call LineTo(hDC, rcItem.Right - 1, rcItem.Bottom - 1)
      Call LineTo(hDC, rcItem.Left - 1, rcItem.Bottom - 1)
      Call SelectObject(hDC, hPreviousPen)
      Call DeleteObject(hPen)
    End If

    Call SetTextColor(hDC, oldTextColor)
    Call SetBkMode(hDC, oldBkMode)
  End If
End Sub
