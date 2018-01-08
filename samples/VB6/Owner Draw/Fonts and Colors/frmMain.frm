VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ComboListBoxControls 1.5 - OwnerDraw sample"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   2  'Bildschirmmitte
   Begin CBLCtlsLibUCtl.ListBox lstColor 
      Height          =   3615
      Left            =   3728
      TabIndex        =   3
      Top             =   600
      Width           =   3375
      _cx             =   5953
      _cy             =   6376
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048811
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertMarkStyle =   1
      IntegralHeight  =   0   'False
      ItemHeight      =   -1
      Locale          =   1024
      MousePointer    =   0
      MultiColumn     =   0   'False
      MultiSelect     =   0
      OLEDragImageStyle=   0
      OwnerDrawItems  =   1
      ProcessContextMenuKeys=   -1  'True
      ProcessTabs     =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   0
      UseSystemFont   =   -1  'True
      VirtualMode     =   0   'False
   End
   Begin CBLCtlsLibUCtl.ComboBox cmbColor 
      Height          =   360
      Left            =   3735
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      _cx             =   5953
      _cy             =   635
      AcceptNumbersOnly=   0   'False
      Appearance      =   3
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      BorderStyle     =   0
      CharacterConversion=   0
      DisabledEvents  =   267503
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragDropDownTime=   -1
      DropDownKey     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      IntegralHeight  =   -1  'True
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListBackColor   =   -2147483643
      ListDragScrollTimeBase=   -1
      ListForeColor   =   -2147483640
      ListHeight      =   -1
      ListInsertMarkColor=   0
      ListScrollableWidth=   0
      ListWidth       =   0
      Locale          =   1024
      MaxTextLength   =   -1
      MinVisibleItems =   30
      MousePointer    =   0
      OwnerDrawItems  =   1
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   0   'False
      Style           =   1
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0000
      Text            =   "frmMain.frx":0020
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2948
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin CBLCtlsLibUCtl.ListBox lstFont 
      Height          =   3615
      Left            =   128
      TabIndex        =   1
      Top             =   600
      Width           =   3375
      _cx             =   5953
      _cy             =   6376
      AllowDragDrop   =   0   'False
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   1048683
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertMarkStyle =   1
      IntegralHeight  =   0   'False
      ItemHeight      =   -1
      Locale          =   1024
      MousePointer    =   0
      MultiColumn     =   0   'False
      MultiSelect     =   0
      OLEDragImageStyle=   0
      OwnerDrawItems  =   2
      ProcessContextMenuKeys=   -1  'True
      ProcessTabs     =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   -1  'True
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   2
      UseSystemFont   =   -1  'True
      VirtualMode     =   0   'False
   End
   Begin CBLCtlsLibUCtl.ComboBox cmbFont 
      Height          =   360
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _cx             =   5953
      _cy             =   635
      AcceptNumbersOnly=   0   'False
      Appearance      =   3
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      BorderStyle     =   0
      CharacterConversion=   0
      DisabledEvents  =   267375
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragDropDownTime=   -1
      DropDownKey     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HasStrings      =   -1  'True
      HoverTime       =   -1
      IMEMode         =   -1
      IntegralHeight  =   -1  'True
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListBackColor   =   -2147483643
      ListDragScrollTimeBase=   -1
      ListForeColor   =   -2147483640
      ListHeight      =   -1
      ListInsertMarkColor=   0
      ListScrollableWidth=   0
      ListWidth       =   0
      Locale          =   1024
      MaxTextLength   =   -1
      MinVisibleItems =   30
      MousePointer    =   0
      OwnerDrawItems  =   2
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   -1  'True
      Style           =   1
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0040
      Text            =   "frmMain.frx":0060
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Const ANSI_CHARSET = 0
  Private Const BALTIC_CHARSET = 186
  Private Const CHINESEBIG5_CHARSET = 136
  Private Const DEFAULT_CHARSET = 1
  Private Const EASTEUROPE_CHARSET = 238
  Private Const GB2312_CHARSET = 134
  Private Const GREEK_CHARSET = 161
  Private Const HANGUL_CHARSET = 129
  Private Const MAC_CHARSET = 77
  Private Const OEM_CHARSET = 255
  Private Const RUSSIAN_CHARSET = 204
  Private Const SHIFTJIS_CHARSET = 128
  Private Const SYMBOL_CHARSET = 2
  Private Const TURKISH_CHARSET = 162


  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private themeableOS As Boolean
  Private ToolTip As clsToolTip


  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
  Private Declare Function DrawRect Lib "gdi32.dll" Alias "Rectangle" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectW" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, lpvObject As Any) As Long
  Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long


Private Sub cmbColor_OwnerDrawItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Const COLOR_BTNFACE = 15
  Const COLOR_3DFACE = COLOR_BTNFACE
  Const COLOR_BTNTEXT = 18
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const PS_SOLID = 0
  Dim flags As Long
  Dim hBorderPen As Long
  Dim hBrush As Long
  Dim hPreviousBrush As Long
  Dim hPreviousPen As Long
  Dim newBkColor As Long
  Dim newTextColor As Long
  Dim oldBkColor As Long
  Dim oldTextColor As Long
  Dim rcColor As RECT
  Dim rcItem As RECT
  Dim rcText As RECT
  Dim str As String

  LSet rcItem = drawingRectangle
  rcText = rcItem
  rcText.Left = rcText.Left + 30
  If Not (comboItem Is Nothing) Then
    ' set all needed colors
    If itemState And OwnerDrawItemStateConstants.odisSelected Then
      If itemState And OwnerDrawItemStateConstants.odisFocus Then
        newBkColor = GetSysColor(COLOR_HIGHLIGHT)
        newTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
      Else
        newBkColor = GetSysColor(COLOR_3DFACE)
        newTextColor = GetSysColor(COLOR_BTNTEXT)
      End If
    Else
      newBkColor = TranslateColor(cmbColor.BackColor)
      newTextColor = TranslateColor(cmbColor.ForeColor)
    End If
    If newTextColor = newBkColor Then
      newTextColor = RGB(&HC0, &HC0, &HC0)
    End If

    ' draw item background
    hBrush = CreateSolidBrush(newBkColor)
    If hBrush Then
      FillRect hDC, rcItem, hBrush
      DeleteObject hBrush
    End If

    ' now calculate the dimensions of the color rectangle
    rcColor.Left = rcItem.Left + 3
    rcColor.Top = rcItem.Top + 2
    rcColor.Bottom = rcItem.Bottom - 2
    rcColor.Right = rcColor.Left + 24

    ' draw the color rectangle
    hPreviousBrush = SelectObject(hDC, GetSysColorBrush(comboItem.ItemData And &HFF))
    hBorderPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0))
    hPreviousPen = SelectObject(hDC, hBorderPen)
    DrawRect hDC, rcColor.Left, rcColor.Top, rcColor.Right, rcColor.Bottom
    SelectObject hDC, hPreviousPen
    SelectObject hDC, hPreviousBrush

    ' now draw the text
    oldTextColor = SetTextColor(hDC, newTextColor)
    oldBkColor = SetBkColor(hDC, newBkColor)
    flags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    If cmbColor.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    str = comboItem.Text
    DrawText hDC, StrPtr(str), Len(str), rcText, flags

    SetTextColor hDC, oldTextColor
    SetBkColor hDC, oldBkColor

    ' draw the focus rectangle
    If (itemState And (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus Or OwnerDrawItemStateConstants.odisNoFocusRectangle)) = (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus) Then
      DrawFocusRect hDC, rcItem
    End If
  End If
End Sub

Private Sub cmbColor_RecreatedControlWindow(ByVal hWnd As Long)
  InsertComboItems_Color
End Sub

Private Sub cmbFont_FreeItemData(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem)
  Const OBJ_FONT = 6
  Dim h As Long

  h = comboItem.ItemData
  If GetObjectType(h) = OBJ_FONT Then DeleteObject h
End Sub

Private Sub cmbFont_MeasureItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ItemHeight As stdole.OLE_YSIZE_PIXELS)
  Const DT_CALCRECT = &H400
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const OBJ_FONT = 6
  Dim flags As Long
  Dim hDC As Long
  Dim hDCCompatible As Long
  Dim hFont As Long
  Dim hOldFont As Long
  Dim rc As RECT

  If Not (comboItem Is Nothing) Then
    hDCCompatible = GetDC(0)
    If hDCCompatible Then
      hDC = CreateCompatibleDC(hDCCompatible)
      If hDC Then
        ' measure item text
        hFont = comboItem.ItemData
        If GetObjectType(hFont) = OBJ_FONT Then
          hOldFont = SelectObject(hDC, hFont)
        End If
        flags = DT_LEFT Or DT_SINGLELINE
        If cmbFont.RightToLeft And RightToLeftConstants.rtlText Then
          flags = flags Or DT_RTLREADING
        End If
        DrawText hDC, StrPtr(comboItem.Text), -1, rc, flags Or DT_CALCRECT
        ItemHeight = rc.Bottom - rc.Top

        If hOldFont Then
          SelectObject hDC, hOldFont
        End If
        DeleteDC hDC
      End If
      ReleaseDC 0, hDCCompatible
    End If
  End If
End Sub

Private Sub cmbFont_OwnerDrawItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Const COLOR_BTNFACE = 15
  Const COLOR_3DFACE = COLOR_BTNFACE
  Const COLOR_BTNTEXT = 18
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const OBJ_FONT = 6
  Dim backClr As Long
  Dim flags As Long
  Dim foreClr As Long
  Dim hBrush As Long
  Dim hFont As Long
  Dim hOldFont As Long
  Dim oldBkColor As Long
  Dim oldTextColor As Long
  Dim rc As RECT

  If Not (comboItem Is Nothing) Then
    If itemState And OwnerDrawItemStateConstants.odisSelected Then
      If itemState And OwnerDrawItemStateConstants.odisFocus Then
        backClr = GetSysColor(COLOR_HIGHLIGHT)
        foreClr = GetSysColor(COLOR_HIGHLIGHTTEXT)
      Else
        backClr = GetSysColor(COLOR_3DFACE)
        foreClr = GetSysColor(COLOR_BTNTEXT)
      End If
    Else
      backClr = TranslateColor(cmbFont.BackColor)
      foreClr = TranslateColor(cmbFont.ForeColor)
    End If

    ' draw item background
    LSet rc = drawingRectangle
    hBrush = CreateSolidBrush(backClr)
    If hBrush Then
      FillRect hDC, rc, hBrush
      DeleteObject hBrush
    End If

    ' draw item text
    hFont = comboItem.ItemData
    If GetObjectType(hFont) = OBJ_FONT Then
      hOldFont = SelectObject(hDC, hFont)
    End If
    oldBkColor = SetBkColor(hDC, backClr)
    oldTextColor = SetTextColor(hDC, foreClr)
    flags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    If cmbFont.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    DrawText hDC, StrPtr(comboItem.Text), -1, rc, flags

    SetTextColor hDC, oldTextColor
    SetBkColor hDC, oldBkColor
    If hOldFont Then
      SelectObject hDC, hOldFont
    End If

    ' draw the focus rectangle
    If (itemState And (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus Or OwnerDrawItemStateConstants.odisNoFocusRectangle)) = (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus) Then
      DrawFocusRect hDC, rc
    End If
  End If
End Sub

Private Sub cmbFont_RecreatedControlWindow(ByVal hWnd As Long)
  InsertComboItems_Font
End Sub

Private Sub cmdAbout_Click()
  cmbFont.About
End Sub

Private Sub Form_Initialize()
  Dim hMod As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  Set ToolTip = New clsToolTip
End Sub

Private Sub Form_Load()
  Const WM_GETFONT = &H31
  Dim hDefaultFont As Long

  hDefaultFont = SendMessage(cmbFont.hWnd, WM_GETFONT, 0, 0)
  GetObjectAPI hDefaultFont, LenB(lfDefault), lfDefault

  InsertComboItems_Font
  InsertComboItems_Color
  InsertListItems_Font
  InsertListItems_Color
End Sub

Private Sub Form_Unload(cancel As Integer)
  Const OBJ_FONT = 6
  Dim comboItem As ComboBoxItem
  Dim h As Long
  Dim listItem As ListBoxItem

  ' The FreeItemData won't get fired on program termination (actually it gets fired, but we won't
  ' receive it anymore). So ensure the fonts get freed.
  For Each comboItem In cmbFont.ComboItems
    h = comboItem.ItemData
    If GetObjectType(h) = OBJ_FONT Then
      DeleteObject h
    End If
  Next comboItem
  For Each listItem In lstFont.ListItems
    h = listItem.ItemData
    If GetObjectType(h) = OBJ_FONT Then
      DeleteObject h
    End If
  Next listItem
  ToolTip.Detach
  Set ToolTip = Nothing
End Sub

Private Sub lstColor_OwnerDrawItem(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Const COLOR_BTNFACE = 15
  Const COLOR_3DFACE = COLOR_BTNFACE
  Const COLOR_BTNTEXT = 18
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const PS_SOLID = 0
  Dim flags As Long
  Dim hBorderPen As Long
  Dim hBrush As Long
  Dim hPreviousBrush As Long
  Dim hPreviousPen As Long
  Dim newBkColor As Long
  Dim newTextColor As Long
  Dim oldBkColor As Long
  Dim oldTextColor As Long
  Dim rcColor As RECT
  Dim rcItem As RECT
  Dim rcText As RECT
  Dim str As String

  LSet rcItem = drawingRectangle
  rcText = rcItem
  rcText.Left = rcText.Left + 30
  If Not (listItem Is Nothing) Then
    ' set all needed colors
    If itemState And OwnerDrawItemStateConstants.odisSelected Then
      If itemState And OwnerDrawItemStateConstants.odisFocus Then
        newBkColor = GetSysColor(COLOR_HIGHLIGHT)
        newTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
      Else
        newBkColor = GetSysColor(COLOR_3DFACE)
        newTextColor = GetSysColor(COLOR_BTNTEXT)
      End If
    Else
      newBkColor = TranslateColor(lstColor.BackColor)
      newTextColor = TranslateColor(lstColor.ForeColor)
    End If
    If newTextColor = newBkColor Then
      newTextColor = RGB(&HC0, &HC0, &HC0)
    End If

    ' draw item background
    hBrush = CreateSolidBrush(newBkColor)
    If hBrush Then
      FillRect hDC, rcItem, hBrush
      DeleteObject hBrush
    End If

    ' now calculate the dimensions of the color rectangle
    rcColor.Left = rcItem.Left + 3
    rcColor.Top = rcItem.Top + 2
    rcColor.Bottom = rcItem.Bottom - 2
    rcColor.Right = rcColor.Left + 24

    ' draw the color rectangle
    hPreviousBrush = SelectObject(hDC, GetSysColorBrush(listItem.ItemData And &HFF))
    hBorderPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0))
    hPreviousPen = SelectObject(hDC, hBorderPen)
    DrawRect hDC, rcColor.Left, rcColor.Top, rcColor.Right, rcColor.Bottom
    SelectObject hDC, hPreviousPen
    SelectObject hDC, hPreviousBrush

    ' now draw the text
    oldTextColor = SetTextColor(hDC, newTextColor)
    oldBkColor = SetBkColor(hDC, newBkColor)
    flags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    If lstColor.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    str = listItem.Text
    DrawText hDC, StrPtr(str), Len(str), rcText, flags

    SetTextColor hDC, oldTextColor
    SetBkColor hDC, oldBkColor

    ' draw the focus rectangle
    If (itemState And (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus Or OwnerDrawItemStateConstants.odisNoFocusRectangle)) = (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus) Then
      DrawFocusRect hDC, rcItem
    End If
  End If
End Sub

Private Sub lstColor_RecreatedControlWindow(ByVal hWnd As Long)
  InsertListItems_Color
End Sub

Private Sub lstFont_DestroyedControlWindow(ByVal hWnd As Long)
  ToolTip.Detach
End Sub

Private Sub lstFont_FreeItemData(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem)
  Const OBJ_FONT = 6
  Dim h As Long

  h = listItem.ItemData
  If GetObjectType(h) = OBJ_FONT Then DeleteObject h
End Sub

Private Sub lstFont_ItemGetInfoTipText(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Const OBJ_FONT = 6

  If GetObjectType(listItem.ItemData) = OBJ_FONT Then
    infoTipText = listItem.Text
  End If
End Sub

Private Sub lstFont_MeasureItem(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ItemHeight As stdole.OLE_YSIZE_PIXELS)
  Const DT_CALCRECT = &H400
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const OBJ_FONT = 6
  Dim flags As Long
  Dim hDC As Long
  Dim hDCCompatible As Long
  Dim hFont As Long
  Dim hOldFont As Long
  Dim rc As RECT

  If Not (listItem Is Nothing) Then
    hDCCompatible = GetDC(0)
    If hDCCompatible Then
      hDC = CreateCompatibleDC(hDCCompatible)
      If hDC Then
        ' measure item text
        hFont = listItem.ItemData
        If GetObjectType(hFont) = OBJ_FONT Then
          hOldFont = SelectObject(hDC, hFont)
        End If
        flags = DT_LEFT Or DT_SINGLELINE
        If lstFont.RightToLeft And RightToLeftConstants.rtlText Then
          flags = flags Or DT_RTLREADING
        End If
        DrawText hDC, StrPtr(listItem.Text), -1, rc, flags Or DT_CALCRECT
        ItemHeight = rc.Bottom - rc.Top

        If hOldFont Then
          SelectObject hDC, hOldFont
        End If
        DeleteDC hDC
      End If
      ReleaseDC 0, hDCCompatible
    End If
  End If
End Sub

Private Sub lstFont_OwnerDrawItem(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Const COLOR_BTNFACE = 15
  Const COLOR_3DFACE = COLOR_BTNFACE
  Const COLOR_BTNTEXT = 18
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_HIGHLIGHTTEXT = 14
  Const DT_LEFT = &H0
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const OBJ_FONT = 6
  Dim backClr As Long
  Dim flags As Long
  Dim foreClr As Long
  Dim hBrush As Long
  Dim hFont As Long
  Dim hOldFont As Long
  Dim oldBkColor As Long
  Dim oldTextColor As Long
  Dim rc As RECT

  If Not (listItem Is Nothing) Then
    If itemState And OwnerDrawItemStateConstants.odisSelected Then
      If itemState And OwnerDrawItemStateConstants.odisFocus Then
        backClr = GetSysColor(COLOR_HIGHLIGHT)
        foreClr = GetSysColor(COLOR_HIGHLIGHTTEXT)
      Else
        backClr = GetSysColor(COLOR_3DFACE)
        foreClr = GetSysColor(COLOR_BTNTEXT)
      End If
    Else
      backClr = TranslateColor(lstFont.BackColor)
      foreClr = TranslateColor(lstFont.ForeColor)
    End If

    ' draw item background
    LSet rc = drawingRectangle
    hBrush = CreateSolidBrush(backClr)
    If hBrush Then
      FillRect hDC, rc, hBrush
      DeleteObject hBrush
    End If

    ' draw item text
    hFont = listItem.ItemData
    If GetObjectType(hFont) = OBJ_FONT Then
      hOldFont = SelectObject(hDC, hFont)
    End If
    oldBkColor = SetBkColor(hDC, backClr)
    oldTextColor = SetTextColor(hDC, foreClr)
    flags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    If lstFont.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    DrawText hDC, StrPtr(listItem.Text), -1, rc, flags

    SetTextColor hDC, oldTextColor
    SetBkColor hDC, oldBkColor
    If hOldFont Then
      SelectObject hDC, hOldFont
    End If

    ' draw the focus rectangle
    If (itemState And (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus Or OwnerDrawItemStateConstants.odisNoFocusRectangle)) = (OwnerDrawItemStateConstants.odisSelected Or OwnerDrawItemStateConstants.odisFocus) Then
      DrawFocusRect hDC, rc
    End If
  End If
End Sub

Private Sub lstFont_RecreatedControlWindow(ByVal hWnd As Long)
  InsertListItems_Font
End Sub


Private Sub InsertComboItems_Color()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme cmbColor.hWnd, StrPtr("explorer"), 0
  End If

  With cmbColor.ComboItems
    .Add "ActiveBorder", , &H8000000A
    .Add "ActiveCaption", , &H80000002
    .Add "ActiveCaptionText", , &H80000009
    .Add "ActiveCaptionGradient", , &H8000001B
    .Add "AppWorkSpace", , &H8000000C
    .Add "Control", , &H8000000F
    .Add "ControlDark", , &H80000010
    .Add "ControlDarkDark", , &H80000015
    .Add "ControlLight", , &H80000016
    .Add "ControlLightLight", , &H80000014
    .Add "ControlText", , &H80000012
    .Add "Desktop", , &H80000001
    .Add "GrayText", , &H80000011
    .Add "Highlight", , &H8000000D
    .Add "HighlightText", , &H8000000E
    .Add "Hot", , &H8000001A
    .Add "InactiveBorder", , &H8000000B
    .Add "InactiveCaption", , &H80000003
    .Add "InactiveCaptionGradient", , &H8000001C
    .Add "InactiveCaptionText", , &H80000013
    .Add "Info", , &H80000018
    .Add "InfoText", , &H80000017
    .Add "Menu", , &H80000004
    .Add "MenuBar", , &H8000001E
    .Add "MenuHighlight", , &H8000001D
    .Add "MenuText", , &H80000007
    .Add "ScrollBar", , &H80000000
    .Add "Window", , &H80000005
    .Add "WindowFrame", , &H80000006
    .Add "WindowText", , &H80000008
  End With
End Sub

Private Sub InsertComboItems_Font()
  Dim hDC As Long

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme cmbFont.hWnd, StrPtr("explorer"), 0
  End If

  With cmbFont.ComboItems
    hDC = GetDC(0)

'    EnumFonts hDC, cmbFont, ANSI_CHARSET
'    EnumFonts hDC, cmbFont, BALTIC_CHARSET
'    EnumFonts hDC, cmbFont, CHINESEBIG5_CHARSET
    EnumFonts hDC, cmbFont, DEFAULT_CHARSET
'    EnumFonts hDC, cmbFont, EASTEUROPE_CHARSET
'    EnumFonts hDC, cmbFont, GB2312_CHARSET
'    EnumFonts hDC, cmbFont, GREEK_CHARSET
'    EnumFonts hDC, cmbFont, HANGUL_CHARSET
'    EnumFonts hDC, cmbFont, MAC_CHARSET
'    EnumFonts hDC, cmbFont, OEM_CHARSET
'    EnumFonts hDC, cmbFont, RUSSIAN_CHARSET
'    EnumFonts hDC, cmbFont, SHIFTJIS_CHARSET
'    EnumFonts hDC, cmbFont, SYMBOL_CHARSET
'    EnumFonts hDC, cmbFont, TURKISH_CHARSET

    ReleaseDC 0, hDC
  End With
End Sub

Private Sub InsertListItems_Color()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme lstColor.hWnd, StrPtr("explorer"), 0
  End If

  With lstColor.ListItems
    .Add "ActiveBorder", , &H8000000A
    .Add "ActiveCaption", , &H80000002
    .Add "ActiveCaptionText", , &H80000009
    .Add "ActiveCaptionGradient", , &H8000001B
    .Add "AppWorkSpace", , &H8000000C
    .Add "Control", , &H8000000F
    .Add "ControlDark", , &H80000010
    .Add "ControlDarkDark", , &H80000015
    .Add "ControlLight", , &H80000016
    .Add "ControlLightLight", , &H80000014
    .Add "ControlText", , &H80000012
    .Add "Desktop", , &H80000001
    .Add "GrayText", , &H80000011
    .Add "Highlight", , &H8000000D
    .Add "HighlightText", , &H8000000E
    .Add "Hot", , &H8000001A
    .Add "InactiveBorder", , &H8000000B
    .Add "InactiveCaption", , &H80000003
    .Add "InactiveCaptionGradient", , &H8000001C
    .Add "InactiveCaptionText", , &H80000013
    .Add "Info", , &H80000018
    .Add "InfoText", , &H80000017
    .Add "Menu", , &H80000004
    .Add "MenuBar", , &H8000001E
    .Add "MenuHighlight", , &H8000001D
    .Add "MenuText", , &H80000007
    .Add "ScrollBar", , &H80000000
    .Add "Window", , &H80000005
    .Add "WindowFrame", , &H80000006
    .Add "WindowText", , &H80000008
  End With
End Sub

Private Sub InsertListItems_Font()
  Dim hDC As Long

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme lstFont.hWnd, StrPtr("explorer"), 0
  End If

  ToolTip.Attach lstFont.hWndToolTip
  ToolTip.BalloonStyle = True
  ToolTip.TitleText = "Font name:"
  ToolTip.TitleIcon = ToolTipTitleIconConstants.tttiInfo

  With lstFont.ListItems
    hDC = GetDC(0)

'    EnumFonts hDC, lstFont, ANSI_CHARSET
'    EnumFonts hDC, lstFont, BALTIC_CHARSET
'    EnumFonts hDC, lstFont, CHINESEBIG5_CHARSET
    EnumFonts hDC, lstFont, DEFAULT_CHARSET
'    EnumFonts hDC, lstFont, EASTEUROPE_CHARSET
'    EnumFonts hDC, lstFont, GB2312_CHARSET
'    EnumFonts hDC, lstFont, GREEK_CHARSET
'    EnumFonts hDC, lstFont, HANGUL_CHARSET
'    EnumFonts hDC, lstFont, MAC_CHARSET
'    EnumFonts hDC, lstFont, OEM_CHARSET
'    EnumFonts hDC, lstFont, RUSSIAN_CHARSET
'    EnumFonts hDC, lstFont, SHIFTJIS_CHARSET
'    EnumFonts hDC, lstFont, SYMBOL_CHARSET
'    EnumFonts hDC, lstFont, TURKISH_CHARSET

    ReleaseDC 0, hDC
  End With
End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional ByVal hPal As Long = 0) As Long
  Const CLR_INVALID = &HFFFF&

  If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function
