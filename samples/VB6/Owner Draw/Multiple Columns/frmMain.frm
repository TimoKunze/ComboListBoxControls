VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ComboListBoxControls 1.5 - Multiple Columns sample"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
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
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   StartUpPosition =   2  'Bildschirmmitte
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
      Left            =   1965
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin CBLCtlsLibUCtl.ComboBox cmbColumns 
      Height          =   360
      Left            =   945
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
      DisabledEvents  =   267367
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Const ColumnCount = 3


  Private themeableOS As Boolean


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal Size As Long)
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal Length As Long)


Private Sub cmbColumns_FreeItemData(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem)
  Call FreeComboTag(comboItem)
End Sub

Private Sub cmbColumns_KeyDown(keyCode As Integer, ByVal shift As Integer)
  Dim searchUp As Boolean
  Dim nextItem As ComboBoxItem
  Dim selItem As ComboBoxItem
  Dim tagObject As Object

  If (keyCode = vbKeyDown) Or (keyCode = vbKeyRight) Then
    ' jump to next non-header item
    searchUp = False
    Set selItem = cmbColumns.SelectedItem
    If selItem Is Nothing Then
      Set nextItem = cmbColumns.ComboItems(0)
    ElseIf selItem.Index < cmbColumns.ComboItems.Count - 1 Then
      Set nextItem = cmbColumns.ComboItems(selItem.Index + 1)
    Else
      Exit Sub
    End If
  ElseIf (keyCode = vbKeyUp) Or (keyCode = vbKeyLeft) Then
    ' jump to previous non-header item
    searchUp = True
    Set selItem = cmbColumns.SelectedItem
    If selItem Is Nothing Then
      Set nextItem = cmbColumns.ComboItems(0)
      searchUp = False
    ElseIf selItem.Index > 0 Then
      Set nextItem = cmbColumns.ComboItems(selItem.Index - 1)
    Else
      Exit Sub
    End If
  ElseIf keyCode = vbKeyPageDown Then
    ' jump to next non-header item, skipping at least 6 items
    searchUp = False
    Set selItem = cmbColumns.SelectedItem
    If selItem Is Nothing Then
      Set nextItem = cmbColumns.ComboItems(6)
    ElseIf selItem.Index < cmbColumns.ComboItems.Count - 7 Then
      Set nextItem = cmbColumns.ComboItems(selItem.Index + 7)
    ElseIf selItem.Index < cmbColumns.ComboItems.Count - 1 Then
      Set nextItem = cmbColumns.ComboItems(cmbColumns.ComboItems.Count - 1)
      searchUp = True
    Else
      Exit Sub
    End If
  ElseIf keyCode = vbKeyPageUp Then
    ' jump to previous non-header item, skipping at least 6 items
    searchUp = True
    Set selItem = cmbColumns.SelectedItem
    If selItem Is Nothing Then
      Set nextItem = cmbColumns.ComboItems(0)
      searchUp = False
    ElseIf selItem.Index > 6 Then
      Set nextItem = cmbColumns.ComboItems(selItem.Index - 7)
    ElseIf selItem.Index > 0 Then
      Set nextItem = cmbColumns.ComboItems(0)
      searchUp = False
    Else
      Exit Sub
    End If
  ElseIf keyCode = vbKeyEnd Then
    ' jump to last non-header item
    Set nextItem = cmbColumns.ComboItems(cmbColumns.ComboItems.Count - 1)
    searchUp = True
  ElseIf keyCode = vbKeyHome Then
    ' jump to first non-header item
    Set nextItem = cmbColumns.ComboItems(0)
    searchUp = False
  End If
  If Not (nextItem Is Nothing) Then
    Set tagObject = GetComboTag(nextItem)
  End If

  If Not (tagObject Is Nothing) Then
    Do
      If TypeOf tagObject Is CHeaderItem Then
        If searchUp And (nextItem.Index > 0) Then
          Set nextItem = cmbColumns.ComboItems(nextItem.Index - 1)
        ElseIf (Not searchUp) And (nextItem.Index < cmbColumns.ComboItems.Count - 1) Then
          Set nextItem = cmbColumns.ComboItems(nextItem.Index + 1)
        Else
          Exit Do
        End If
      End If
      Set tagObject = GetComboTag(nextItem)
      If tagObject Is Nothing Then
        Exit Do
      ElseIf Not (TypeOf tagObject Is CHeaderItem) Then
        Set cmbColumns.SelectedItem = nextItem
        Exit Do
      End If
    Loop
    keyCode = 0
  End If
End Sub

Private Sub cmbColumns_OwnerDrawItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Dim tagObject As Object

  Set tagObject = GetComboTag(comboItem)
  If Not (tagObject Is Nothing) Then
    Call tagObject.OwnerDrawItem(comboItem, requiredAction, itemState, hDC, drawingRectangle)
  End If
End Sub

Private Sub cmbColumns_RecreatedControlWindow(ByVal hWnd As Long)
  Call InsertComboItems
End Sub

Private Sub cmbColumns_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
  Dim tagObject As Object

  If Not (newSelectedItem Is Nothing) Then
    Set tagObject = GetComboTag(newSelectedItem)
    If Not (tagObject Is Nothing) Then
      If TypeOf tagObject Is CHeaderItem Then
        If newSelectedItem.Index < cmbColumns.ComboItems.Count - 1 Then
          Set cmbColumns.SelectedItem = cmbColumns.ComboItems(newSelectedItem.Index + 1)
        Else
          Set cmbColumns.SelectedItem = previousSelectedItem
        End If
      End If
    End If
  End If
End Sub

Private Sub cmdAbout_Click()
  Call cmbColumns.About
End Sub

Private Sub Form_Initialize()
  Dim hMod As Long

  Call InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    Call FreeLibrary(hMod)
  End If
End Sub

Private Sub Form_Load()
  Call InsertComboItems
End Sub

Private Sub Form_Unload(cancel As Integer)
  ' The FreeItemData won't get fired on program termination (actually it gets fired, but we won't
  ' receive it anymore). So ensure everything get freed.
  Call cmbColumns.ComboItems.RemoveAll
End Sub


Private Sub FreeComboTag(item As CBLCtlsLibUCtl.IComboBoxItem)
  Dim tagObject As Object

  If Not (item Is Nothing) Then
    If item.ItemData <> 0 Then
      Call CopyMemory(VarPtr(tagObject), VarPtr(item.ItemData), 4)
      item.ItemData = 0
      Set tagObject = Nothing
    End If
  End If
End Sub

Private Function GetComboTag(item As CBLCtlsLibUCtl.IComboBoxItem) As Object
  Dim tagObject As Object

  If Not (item Is Nothing) Then
    If item.ItemData <> 0 Then
      Call CopyMemory(VarPtr(tagObject), VarPtr(item.ItemData), 4)
      Set GetComboTag = tagObject
      Call ZeroMemory(VarPtr(tagObject), 4)
    End If
  End If
End Function

Private Sub InsertComboItems()
  Dim headerTag As CHeaderItem
  Dim i As Long
  Dim multiColumnTag As CMultiColumnItem

  If themeableOS Then
    ' for Windows Vista
    Call SetWindowTheme(cmbColumns.hWnd, StrPtr("explorer"), 0)
  End If

  With cmbColumns.ComboItems
    For i = 1 To 5
      Set headerTag = New CHeaderItem
      headerTag.Caption = "Group " & CStr(i)
      headerTag.ItemID = .Add(headerTag.Caption, ItemData:=MakeComboTag(headerTag)).ID
      Call ZeroMemory(VarPtr(headerTag), 4)

      Set multiColumnTag = New CMultiColumnItem
      multiColumnTag.ColumnCount = ColumnCount
      Call multiColumnTag.SetTexts("Group " & CStr(i) & ", Item 1", "Col 2", "Col 3")
      multiColumnTag.ItemID = .Add(multiColumnTag.Text(1), ItemData:=MakeComboTag(multiColumnTag)).ID
      Call ZeroMemory(VarPtr(multiColumnTag), 4)

      Set multiColumnTag = New CMultiColumnItem
      multiColumnTag.ColumnCount = ColumnCount
      Call multiColumnTag.SetTexts("Group " & CStr(i) & ", Item 2", "Col 2", "Col 3")
      multiColumnTag.ItemID = .Add(multiColumnTag.Text(1), ItemData:=MakeComboTag(multiColumnTag)).ID
      Call ZeroMemory(VarPtr(multiColumnTag), 4)

      Set multiColumnTag = New CMultiColumnItem
      multiColumnTag.ColumnCount = ColumnCount
      Call multiColumnTag.SetTexts("Group " & CStr(i) & ", Item 3", "Col 2", "Col 3")
      multiColumnTag.ItemID = .Add(multiColumnTag.Text(1), ItemData:=MakeComboTag(multiColumnTag)).ID
      Call ZeroMemory(VarPtr(multiColumnTag), 4)
    Next i
  End With

  Call SetColumnWidths
End Sub

Private Function MakeComboTag(Tag As Object) As Long
  If Not (Tag Is Nothing) Then
    MakeComboTag = ObjPtr(Tag)
    ' will be done in InsertItems
    'Call ZeroMemory(VarPtr(Tag), 4)
  End If
End Function

Private Sub SetColumnWidths()
  Const WM_GETFONT = &H31
  Dim columnWidths() As Long
  Dim comboItem As ComboBoxItem
  Dim cx As Long
  Dim hDC As Long
  Dim hDCMem As Long
  Dim hPreviousFont As Long
  Dim i As Long
  Dim tagObject As Object
  Dim totalWidth As Long

  ReDim columnWidths(1 To ColumnCount) As Long
  ' calculate the column widths
  hDC = GetDC(cmbColumns.hWnd)
  If hDC Then
    hDCMem = CreateCompatibleDC(hDC)
    If hDCMem Then
      hPreviousFont = SelectObject(hDCMem, SendMessageAsLong(cmbColumns.hWnd, WM_GETFONT, 0, 0))

      For Each comboItem In cmbColumns.ComboItems
        Set tagObject = GetComboTag(comboItem)
        If Not (tagObject Is Nothing) Then
          If TypeOf tagObject Is CMultiColumnItem Then
            For i = 1 To ColumnCount
              cx = tagObject.CalculateColumnWidth(hDCMem, i)
              If cx > columnWidths(i) Then columnWidths(i) = cx
            Next i
          End If
        End If
      Next comboItem
      totalWidth = 3
      If cmbColumns.ComboItems.Count * cmbColumns.ItemHeight > cmbColumns.ListHeight Then
        totalWidth = totalWidth + 19
      ElseIf cmbColumns.ListAlwaysShowVerticalScrollBar Then
        totalWidth = totalWidth + 19
      End If
      For i = 1 To ColumnCount
        totalWidth = totalWidth + columnWidths(i)
      Next i
      If totalWidth > cmbColumns.Width Then
        cmbColumns.ListWidth = totalWidth
      End If

      Call SelectObject(hDCMem, hPreviousFont)
      Call DeleteDC(hDCMem)
    End If
    Call ReleaseDC(cmbColumns.hWnd, hDC)
  End If

  ' update each item
  For Each comboItem In cmbColumns.ComboItems
    Set tagObject = GetComboTag(comboItem)
    If Not (tagObject Is Nothing) Then
      If TypeOf tagObject Is CMultiColumnItem Then
        For i = 1 To ColumnCount
          tagObject.ColumnWidth(i) = columnWidths(i)
        Next i
      End If
    End If
  Next comboItem
End Sub
