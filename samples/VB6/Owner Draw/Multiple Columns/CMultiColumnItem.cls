VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultiColumnItem"
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


  Private m_ColumnCount As Long
  Private m_ColumnWidths() As Long
  Private m_Texts() As String

  Public ItemData As Long
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


Private Sub Class_Initialize()
  m_ColumnCount = 1
  ReDim Preserve m_ColumnWidths(1 To m_ColumnCount) As Long
  ReDim Preserve m_Texts(1 To m_ColumnCount) As String
End Sub

Private Sub Class_Terminate()
  Erase m_ColumnWidths
  Erase m_Texts
End Sub


Public Property Get ColumnCount() As Long
  ColumnCount = m_ColumnCount
End Property

Public Property Let ColumnCount(ByVal newValue As Long)
  m_ColumnCount = newValue
  ReDim Preserve m_ColumnWidths(1 To m_ColumnCount) As Long
  ReDim Preserve m_Texts(1 To m_ColumnCount) As String
End Property

Public Property Get ColumnWidth(ByVal Index As Long) As Long
  ColumnWidth = m_ColumnWidths(Index)
End Property

Public Property Let ColumnWidth(ByVal Index As Long, ByVal newValue As Long)
  m_ColumnWidths(Index) = newValue
End Property

Public Property Get Text(ByVal Index As Long) As String
  Text = m_Texts(Index)
End Property

Public Property Let Text(ByVal Index As Long, newValue As String)
  m_Texts(Index) = newValue
End Property

Public Function CalculateColumnWidth(ByVal hDC As Long, ByVal Index As Long) As Long
  Const DT_CALCRECT = &H400
  Const DT_LEFT = &H0
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Dim flags As Long
  Dim rc As RECT

  flags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_CALCRECT
  rc.Right = 10000
  rc.Bottom = 10000
  Call DrawText(hDC, StrPtr(m_Texts(Index)), Len(m_Texts(Index)), rc, flags)
  CalculateColumnWidth = rc.Right - rc.Left + 7
End Function

Public Sub SetColumnWidths(ParamArray widths() As Variant)
  Dim i As Long
  Dim j As Long

  j = 1
  For i = LBound(widths) To UBound(widths)
    m_ColumnWidths(j) = CLng(widths(i))
    j = j + 1
  Next i
End Sub

Public Sub SetTexts(ParamArray texts() As Variant)
  Dim i As Long
  Dim j As Long

  j = 1
  For i = LBound(texts) To UBound(texts)
    m_Texts(j) = CStr(texts(i))
    j = j + 1
  Next i
End Sub

Public Sub OwnerDrawItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Const COLOR_BTNFACE = 15
  Const COLOR_BTNSHADOW = 16
  Const COLOR_3DFACE = COLOR_BTNFACE
  Const COLOR_3DSHADOW = COLOR_BTNSHADOW
  Const COLOR_BTNTEXT = 18
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14
  Const COLOR_WINDOW = 5
  Const COLOR_WINDOWTEXT = 8
  Const DT_LEFT = &H0
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const PS_SOLID = 0
  Const TRANSPARENT = 1
  Dim flags As Long
  Dim hPen As Long
  Dim hPreviousPen As Long
  Dim i As Long
  Dim newBkColor As Long
  Dim newLineColor As Long
  Dim newTextColor As Long
  Dim oldBkMode As Long
  Dim oldTextColor As Long
  Dim rcCol As RECT
  Dim rcItem As RECT
  Dim rcText As RECT

  LSet rcItem = drawingRectangle
  If Not (comboItem Is Nothing) Then
    If itemState And OwnerDrawItemStateConstants.odisSelected Then
      If itemState And OwnerDrawItemStateConstants.odisFocus Then
        newBkColor = COLOR_HIGHLIGHT
        newLineColor = COLOR_3DSHADOW
        newTextColor = COLOR_HIGHLIGHTTEXT
      Else
        newBkColor = COLOR_3DFACE
        newLineColor = COLOR_3DSHADOW
        newTextColor = COLOR_BTNTEXT
      End If
    Else
      newBkColor = COLOR_WINDOW
      newLineColor = COLOR_3DSHADOW
      newTextColor = COLOR_WINDOWTEXT
    End If

    ' draw item background
    Call FillRect(hDC, rcItem, GetSysColorBrush(newBkColor))

    ' draw the columns
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(newLineColor))
    If hPen Then
      hPreviousPen = SelectObject(hDC, hPen)
      oldTextColor = SetTextColor(hDC, GetSysColor(newTextColor))
      oldBkMode = SetBkMode(hDC, TRANSPARENT)
      rcCol = rcItem
      If itemState And odisDrawingIntoSelectionField Then
        ' make the grid lines in the selection field match the grid lines in the drop-down list
        rcCol.Left = rcCol.Left - 2
      End If
      flags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
      For i = 1 To m_ColumnCount
        rcCol.Right = rcCol.Left + m_ColumnWidths(i)
        If i > 1 Then
          Call MoveToEx(hDC, rcCol.Left, rcCol.Top, 0)
          Call LineTo(hDC, rcCol.Left, rcCol.Bottom)
        End If
        rcText = rcCol
        rcText.Left = rcText.Left + 4
        rcText.Right = rcText.Right - 3
        Call DrawText(hDC, StrPtr(m_Texts(i)), Len(m_Texts(i)), rcText, flags)

        rcCol.Left = rcCol.Right
      Next i
      Call SelectObject(hDC, hPreviousPen)
      Call SetTextColor(hDC, oldTextColor)
      Call SetBkMode(hDC, oldBkMode)
      Call DeleteObject(hPen)
    End If
  End If
End Sub
