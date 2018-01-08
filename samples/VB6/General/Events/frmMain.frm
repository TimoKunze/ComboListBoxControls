VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Object = "{EE7B09EE-4DEB-47AA-8B0F-FA832AF08A0F}#1.5#0"; "CBLCtlsA.ocx"
Begin VB.Form frmMain 
   Caption         =   "ComboListBoxControls 1.5 - Events Sample"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
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
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'Bildschirmmitte
   Begin CBLCtlsLibACtl.ImageComboBox ImgCmbA 
      Height          =   330
      Left            =   3000
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
      _cx             =   3201
      _cy             =   582
      AcceptNumbersOnly=   0   'False
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      CaseSensitiveItemSearching=   0   'False
      DisabledEvents  =   0
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
      HoverTime       =   -1
      IconVisibility  =   0
      IMEMode         =   -1
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListDragScrollTimeBase=   -1
      ListHeight      =   -1
      ListInsertMarkColor=   0
      ListWidth       =   0
      MaxTextLength   =   -1
      MousePointer    =   0
      OLEDragImageStyle=   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   0   'False
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextEndEllipsis =   -1  'True
      UseShellWordBreakFunction=   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0000
      Text            =   "frmMain.frx":0020
   End
   Begin CBLCtlsLibACtl.DriveComboBox DrvCmbA 
      Height          =   330
      Left            =   3000
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
      _cx             =   3201
      _cy             =   582
      BackColor       =   -2147483643
      CaseSensitiveItemSearching=   0   'False
      DisabledEvents  =   0
      DisplayedDriveTypes=   127
      DisplayNameStyle=   4
      DisplayOverlayImages=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragDropDownTime=   -1
      DriveTypesWithVolumeName=   68
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
      HandleOLEDragDrop=   3
      HoverTime       =   -1
      IconVisibility  =   0
      IMEMode         =   -1
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListDragScrollTimeBase=   -1
      ListHeight      =   -1
      ListWidth       =   0
      MousePointer    =   0
      OLEDragImageStyle=   1
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      UseSystemImageList=   1
      CueBanner       =   "frmMain.frx":0040
   End
   Begin CBLCtlsLibACtl.ComboBox CmbA 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
      _cx             =   3201
      _cy             =   556
      AcceptNumbersOnly=   0   'False
      Appearance      =   3
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      BorderStyle     =   0
      CharacterConversion=   0
      DisabledEvents  =   0
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
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   0   'False
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0060
      Text            =   "frmMain.frx":0080
   End
   Begin CBLCtlsLibACtl.ListBox LstA 
      Height          =   2790
      Left            =   0
      TabIndex        =   4
      Top             =   2895
      Width           =   2895
      _cx             =   5106
      _cy             =   4921
      AllowDragDrop   =   -1  'True
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   0
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
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      ProcessTabs     =   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   3
      UseSystemFont   =   -1  'True
      VirtualMode     =   0   'False
   End
   Begin CBLCtlsLibUCtl.ImageComboBox ImgCmbU 
      Height          =   330
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
      _cx             =   3201
      _cy             =   582
      AcceptNumbersOnly=   0   'False
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      CaseSensitiveItemSearching=   0   'False
      DisabledEvents  =   0
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
      HoverTime       =   -1
      IconVisibility  =   0
      IMEMode         =   -1
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListDragScrollTimeBase=   -1
      ListHeight      =   -1
      ListInsertMarkColor=   0
      ListWidth       =   0
      MaxTextLength   =   -1
      MousePointer    =   0
      OLEDragImageStyle=   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   0   'False
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextEndEllipsis =   -1  'True
      UseShellWordBreakFunction=   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":00A0
      Text            =   "frmMain.frx":00C0
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   5400
      Width           =   975
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
      Left            =   6150
      TabIndex        =   8
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   9
      Top             =   120
      Width           =   4455
   End
   Begin CBLCtlsLibUCtl.DriveComboBox DrvCmbU 
      Height          =   330
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _cx             =   3201
      _cy             =   582
      BackColor       =   -2147483643
      CaseSensitiveItemSearching=   0   'False
      DisabledEvents  =   0
      DisplayedDriveTypes=   127
      DisplayNameStyle=   4
      DisplayOverlayImages=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragDropDownTime=   -1
      DriveTypesWithVolumeName=   68
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
      HandleOLEDragDrop=   3
      HoverTime       =   -1
      IconVisibility  =   0
      IMEMode         =   -1
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListDragScrollTimeBase=   -1
      ListHeight      =   -1
      ListWidth       =   0
      MousePointer    =   0
      OLEDragImageStyle=   1
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      UseSystemImageList=   1
      CueBanner       =   "frmMain.frx":00E0
   End
   Begin CBLCtlsLibUCtl.ComboBox CmbU 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      _cx             =   3201
      _cy             =   556
      AcceptNumbersOnly=   0   'False
      Appearance      =   3
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      BorderStyle     =   0
      CharacterConversion=   0
      DisabledEvents  =   0
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
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   0   'False
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CueBanner       =   "frmMain.frx":0100
      Text            =   "frmMain.frx":0120
   End
   Begin CBLCtlsLibUCtl.ListBox LstU 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _cx             =   5106
      _cy             =   4921
      AllowDragDrop   =   -1  'True
      AllowItemSelection=   -1  'True
      AlwaysShowVerticalScrollBar=   0   'False
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   0
      ColumnWidth     =   -1
      DisabledEvents  =   0
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
      OwnerDrawItems  =   0
      ProcessContextMenuKeys=   -1  'True
      ProcessTabs     =   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ScrollableWidth =   0
      Sorted          =   0   'False
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      ToolTips        =   3
      UseSystemFont   =   -1  'True
      VirtualMode     =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type POINTAPI
    x As Long
    y As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private bLog As Boolean
  Private hImgLst As Long
  Private objActiveCtl As Object
  Private ptStartDrag As POINTAPI


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim iconsDir As String
  Dim iconPath As String

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconsDir = iconsDir & "16x16" & IIf(bComctl32Version600OrNewer, "x32bpp\", "x8bpp\")
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
    iconPath = Dir$
  Wend
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked

  ' this is required to make the control work as expected
  Subclass

  InsertComboItemsA
  InsertComboItemsU
  DrvCmbA.hWndShellUIParentWindow = Me.hWnd
  DrvCmbU.hWndShellUIParentWindow = Me.hWnd
  InsertImageComboItemsA
  InsertImageComboItemsU
  InsertListItemsA
  InsertListItemsU
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    LstU.Move 0, 5, LstU.Width, (Me.ScaleHeight - 11) / 2
    CmbU.Move LstU.Left + LstU.Width + 5, LstU.Top
    DrvCmbU.Move CmbU.Left, CmbU.Top + CmbU.Height + 5
    ImgCmbU.Move CmbU.Left, DrvCmbU.Top + DrvCmbU.Height + 5
    LstA.Move 0, LstU.Top + LstU.Height + 5, LstU.Width, LstU.Height
    CmbA.Move LstU.Left + LstU.Width + 5, LstA.Top
    DrvCmbA.Move CmbA.Left, CmbA.Top + CmbA.Height + 5
    ImgCmbA.Move CmbA.Left, DrvCmbA.Top + DrvCmbA.Height + 5
    txtLog.Move CmbU.Left + CmbU.Width + 5, 0, Me.ScaleWidth - CmbU.Width - CmbU.Left - 5, Me.ScaleHeight - cmdAbout.Height - 10
    chkLog.Move txtLog.Left, txtLog.Top + txtLog.Height + (Me.ScaleHeight - (txtLog.Top + txtLog.Height) - chkLog.Height) / 2
    cmdAbout.Move chkLog.Left + chkLog.Width + ((txtLog.Width - chkLog.Width) - cmdAbout.Width) / 2, txtLog.Top + txtLog.Height + (Me.ScaleHeight - (txtLog.Top + txtLog.Height) - cmdAbout.Height) / 2
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub CmbA_BeforeDrawText()
  AddLogEntry "CmbA_BeforeDrawText"
End Sub

Private Sub CmbA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_CompareItems(ByVal firstItem As CBLCtlsLibACtl.IComboBoxItem, ByVal secondItem As CBLCtlsLibACtl.IComboBoxItem, ByVal Locale As Long, result As CBLCtlsLibACtl.CompareResultConstants)
  Dim str As String

  str = "CmbA_CompareItems: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & firstItem.Index & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Index
  End If
  str = str & ", locale=" & Locale & ", result=" & result

  AddLogEntry str
End Sub

Private Sub CmbA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "CmbA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub CmbA_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "CmbA_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub CmbA_CreatedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "CmbA_CreatedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub CmbA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CmbA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub CmbA_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "CmbA_DestroyedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub CmbA_DestroyedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "CmbA_DestroyedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub CmbA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "CmbA_DragDrop"
End Sub

Private Sub CmbA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "CmbA_DragOver"
End Sub

Private Sub CmbA_FreeItemData(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "CmbA_FreeItemData: comboItem=Nothing"
  Else
    AddLogEntry "CmbA_FreeItemData: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub CmbA_GotFocus()
  Set objActiveCtl = CmbA
  AddLogEntry "CmbA_GotFocus"
End Sub

Private Sub CmbA_InsertedItem(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "CmbA_InsertedItem: comboItem=Nothing"
  Else
    AddLogEntry "CmbA_InsertedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub CmbA_InsertingItem(ByVal comboItem As CBLCtlsLibACtl.IVirtualComboBoxItem, cancelInsertion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "CmbA_InsertingItem: comboItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "CmbA_InsertingItem: comboItem=" & comboItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub CmbA_ItemMouseEnter(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "CmbA_ItemMouseEnter: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbA_ItemMouseLeave(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "CmbA_ItemMouseLeave: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CmbA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CmbA_KeyPress(keyAscii As Integer)
  AddLogEntry "CmbA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub CmbA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CmbA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CmbA_ListCloseUp()
  AddLogEntry "CmbA_ListCloseUp"
End Sub

Private Sub CmbA_ListDropDown()
  AddLogEntry "CmbA_ListDropDown"
End Sub

Private Sub CmbA_ListMouseDown(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "CmbA_ListMouseDown: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbA_ListMouseMove(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "CmbA_ListMouseMove: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub CmbA_ListMouseUp(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "CmbA_ListMouseUp: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbA_ListMouseWheel(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "CmbA_ListMouseWheel: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbA_ListOLEDragDrop(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CmbA_ListOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    CmbA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub CmbA_ListOLEDragEnter(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CmbA_ListOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub CmbA_ListOLEDragLeave(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "CmbA_ListOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub CmbA_ListOLEDragMouseMove(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CmbA_ListOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub CmbA_LostFocus()
  AddLogEntry "CmbA_LostFocus"
End Sub

Private Sub CmbA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MeasureItem(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ItemHeight As stdole.OLE_YSIZE_PIXELS)
  Dim str As String

  str = "CmbA_MeasureItem: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", ItemHeight=" & ItemHeight

  AddLogEntry str
  ItemHeight = 17
End Sub

Private Sub CmbA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  'AddLogEntry "CmbA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_OLEDragDrop(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CmbA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    CmbA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub CmbA_OLEDragEnter(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "CmbA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub CmbA_OLEDragLeave(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CmbA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbA_OLEDragMouseMove(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "CmbA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub CmbA_OutOfMemory()
  AddLogEntry "CmbA_OutOfMemory"
End Sub

Private Sub CmbA_OwnerDrawItem(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibACtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibACtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibACtl.RECTANGLE)
  Dim str As String

  str = "CmbA_OwnerDrawItem: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", requiredAction=0x" & Hex(requiredAction) & ", itemState=0x" & Hex(itemState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & ", " & drawingRectangle.Top & ")-(" & drawingRectangle.Right & ", " & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub CmbA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CmbA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertComboItemsA
End Sub

Private Sub CmbA_RemovedItem(ByVal comboItem As CBLCtlsLibACtl.IVirtualComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "CmbA_RemovedItem: comboItem=Nothing"
  Else
    AddLogEntry "CmbA_RemovedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub CmbA_RemovingItem(ByVal comboItem As CBLCtlsLibACtl.IComboBoxItem, cancelDeletion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "CmbA_RemovingItem: comboItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "CmbA_RemovingItem: comboItem=" & comboItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub CmbA_ResizedControlWindow()
  AddLogEntry "CmbA_ResizedControlWindow"
End Sub

Private Sub CmbA_SelectionCanceled()
  AddLogEntry "CmbA_SelectionCanceled"
End Sub

Private Sub CmbA_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibACtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibACtl.IComboBoxItem)
  Dim str As String

  str = "CmbA_SelectionChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Text & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub CmbA_SelectionChanging()
  AddLogEntry "CmbA_SelectionChanging"
End Sub

Private Sub CmbA_TextChanged()
  AddLogEntry "CmbA_TextChanged"
End Sub

Private Sub CmbA_TruncatedText()
  AddLogEntry "CmbA_TruncatedText"
End Sub

Private Sub CmbA_Validate(Cancel As Boolean)
  AddLogEntry "CmbA_Validate"
End Sub

Private Sub CmbA_WritingDirectionChanged(ByVal newWritingDirection As CBLCtlsLibACtl.WritingDirectionConstants)
  AddLogEntry "CmbA_WritingDirectionChanged: newWritingDirection=" & newWritingDirection
End Sub

Private Sub CmbA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "CmbA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_BeforeDrawText()
  AddLogEntry "CmbU_BeforeDrawText"
End Sub

Private Sub CmbU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_CompareItems(ByVal firstItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal secondItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal Locale As Long, result As CBLCtlsLibUCtl.CompareResultConstants)
  Dim str As String

  str = "CmbU_CompareItems: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & firstItem.Index & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Index
  End If
  str = str & ", locale=" & Locale & ", result=" & result

  AddLogEntry str
End Sub

Private Sub CmbU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "CmbU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub CmbU_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "CmbU_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub CmbU_CreatedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "CmbU_CreatedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub CmbU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CmbU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub CmbU_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "CmbU_DestroyedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub CmbU_DestroyedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "CmbU_DestroyedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub CmbU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "CmbU_DragDrop"
End Sub

Private Sub CmbU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "CmbU_DragOver"
End Sub

Private Sub CmbU_FreeItemData(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "CmbU_FreeItemData: comboItem=Nothing"
  Else
    AddLogEntry "CmbU_FreeItemData: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub CmbU_GotFocus()
  Set objActiveCtl = CmbU
  AddLogEntry "CmbU_GotFocus"
End Sub

Private Sub CmbU_InsertedItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "CmbU_InsertedItem: comboItem=Nothing"
  Else
    AddLogEntry "CmbU_InsertedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub CmbU_InsertingItem(ByVal comboItem As CBLCtlsLibUCtl.IVirtualComboBoxItem, cancelInsertion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "CmbU_InsertingItem: comboItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "CmbU_InsertingItem: comboItem=" & comboItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub CmbU_ItemMouseEnter(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "CmbU_ItemMouseEnter: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbU_ItemMouseLeave(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "CmbU_ItemMouseLeave: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CmbU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CmbU_KeyPress(keyAscii As Integer)
  AddLogEntry "CmbU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub CmbU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CmbU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CmbU_ListCloseUp()
  AddLogEntry "CmbU_ListCloseUp"
End Sub

Private Sub CmbU_ListDropDown()
  AddLogEntry "CmbU_ListDropDown"
End Sub

Private Sub CmbU_ListMouseDown(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "CmbU_ListMouseDown: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbU_ListMouseMove(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "CmbU_ListMouseMove: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub CmbU_ListMouseUp(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "CmbU_ListMouseUp: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbU_ListMouseWheel(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "CmbU_ListMouseWheel: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbU_ListOLEDragDrop(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CmbU_ListOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    CmbU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub CmbU_ListOLEDragEnter(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CmbU_ListOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub CmbU_ListOLEDragLeave(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "CmbU_ListOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub CmbU_ListOLEDragMouseMove(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CmbU_ListOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub CmbU_LostFocus()
  AddLogEntry "CmbU_LostFocus"
End Sub

Private Sub CmbU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MeasureItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ItemHeight As stdole.OLE_YSIZE_PIXELS)
  Dim str As String

  str = "CmbU_MeasureItem: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", ItemHeight=" & ItemHeight

  AddLogEntry str
  ItemHeight = 17
End Sub

Private Sub CmbU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  'AddLogEntry "CmbU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_OLEDragDrop(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CmbU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    CmbU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub CmbU_OLEDragEnter(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "CmbU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub CmbU_OLEDragLeave(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CmbU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub CmbU_OLEDragMouseMove(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "CmbU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub CmbU_OutOfMemory()
  AddLogEntry "CmbU_OutOfMemory"
End Sub

Private Sub CmbU_OwnerDrawItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Dim str As String

  str = "CmbU_OwnerDrawItem: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", requiredAction=0x" & Hex(requiredAction) & ", itemState=0x" & Hex(itemState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & ", " & drawingRectangle.Top & ")-(" & drawingRectangle.Right & ", " & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub CmbU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CmbU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertComboItemsU
End Sub

Private Sub CmbU_RemovedItem(ByVal comboItem As CBLCtlsLibUCtl.IVirtualComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "CmbU_RemovedItem: comboItem=Nothing"
  Else
    AddLogEntry "CmbU_RemovedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub CmbU_RemovingItem(ByVal comboItem As CBLCtlsLibUCtl.IComboBoxItem, cancelDeletion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "CmbU_RemovingItem: comboItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "CmbU_RemovingItem: comboItem=" & comboItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub CmbU_ResizedControlWindow()
  AddLogEntry "CmbU_ResizedControlWindow"
End Sub

Private Sub CmbU_SelectionCanceled()
  AddLogEntry "CmbU_SelectionCanceled"
End Sub

Private Sub CmbU_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IComboBoxItem)
  Dim str As String

  str = "CmbU_SelectionChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Text & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub CmbU_SelectionChanging()
  AddLogEntry "CmbU_SelectionChanging"
End Sub

Private Sub CmbU_TextChanged()
  AddLogEntry "CmbU_TextChanged"
End Sub

Private Sub CmbU_TruncatedText()
  AddLogEntry "CmbU_TruncatedText"
End Sub

Private Sub CmbU_Validate(Cancel As Boolean)
  AddLogEntry "CmbU_Validate"
End Sub

Private Sub CmbU_WritingDirectionChanged(ByVal newWritingDirection As CBLCtlsLibUCtl.WritingDirectionConstants)
  AddLogEntry "CmbU_WritingDirectionChanged: newWritingDirection=" & newWritingDirection
End Sub

Private Sub CmbU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub CmbU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CmbU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_BeginSelectionChange()
  AddLogEntry "DrvCmbA_BeginSelectionChange"
End Sub

Private Sub DrvCmbA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "DrvCmbA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub DrvCmbA_CreatedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "DrvCmbA_CreatedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub DrvCmbA_CreatedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "DrvCmbA_CreatedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub DrvCmbA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_DestroyedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "DrvCmbA_DestroyedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub DrvCmbA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "DrvCmbA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub DrvCmbA_DestroyedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "DrvCmbA_DestroyedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub DrvCmbA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "DrvCmbA_DragDrop"
End Sub

Private Sub DrvCmbA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "DrvCmbA_DragOver"
End Sub

Private Sub DrvCmbA_FreeItemData(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbA_FreeItemData: comboItem=Nothing"
  Else
    AddLogEntry "DrvCmbA_FreeItemData: comboItem=" & comboItem.Path
  End If
End Sub

Private Sub DrvCmbA_GotFocus()
  Set objActiveCtl = DrvCmbA
  AddLogEntry "DrvCmbA_GotFocus"
End Sub

Private Sub DrvCmbA_InsertedItem(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbA_InsertedItem: comboItem=Nothing"
  Else
    AddLogEntry "DrvCmbA_InsertedItem: comboItem=" & comboItem.Path
  End If
End Sub

Private Sub DrvCmbA_InsertingItem(ByVal comboItem As CBLCtlsLibACtl.IVirtualDriveComboBoxItem, cancelInsertion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbA_InsertingItem: comboItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "DrvCmbA_InsertingItem: comboItem=" & comboItem.Path & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub DrvCmbA_ItemBeginDrag(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ItemBeginDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ItemBeginRDrag(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ItemBeginRDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ItemGetDisplayInfo(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal requestedInfo As CBLCtlsLibACtl.RequestedInfoConstants, IconIndex As Long, SelectedIconIndex As Long, OverlayIndex As Long, Indent As Long, ByVal maxItemTextLength As Long, itemText As String, dontAskAgain As Boolean)
  Dim str As String

  str = "DrvCmbA_ItemGetDisplayInfo: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Index
  End If
  str = str & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", SelectedIconIndex=" & SelectedIconIndex & ", OverlayIndex=" & OverlayIndex & ", Indent=" & Indent & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  AddLogEntry str
End Sub

Private Sub DrvCmbA_ItemMouseEnter(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ItemMouseEnter: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ItemMouseLeave(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ItemMouseLeave: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "DrvCmbA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub DrvCmbA_KeyPress(keyAscii As Integer)
  AddLogEntry "DrvCmbA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub DrvCmbA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "DrvCmbA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub DrvCmbA_ListClick(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ListClick: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ListCloseUp()
  AddLogEntry "DrvCmbA_ListCloseUp"
End Sub

Private Sub DrvCmbA_ListDropDown()
  AddLogEntry "DrvCmbA_ListDropDown"
End Sub

Private Sub DrvCmbA_ListMouseDown(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ListMouseDown: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ListMouseMove(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ListMouseMove: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub DrvCmbA_ListMouseUp(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ListMouseUp: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ListMouseWheel(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbA_ListMouseWheel: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ListOLEDragDrop(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_ListOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    DrvCmbA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub DrvCmbA_ListOLEDragEnter(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_ListOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ListOLEDragLeave(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_ListOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub DrvCmbA_ListOLEDragMouseMove(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_ListOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub DrvCmbA_LostFocus()
  AddLogEntry "DrvCmbA_LostFocus"
End Sub

Private Sub DrvCmbA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  'AddLogEntry "DrvCmbA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_OLECompleteDrag(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal performedEffect As CBLCtlsLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub DrvCmbA_OLEDragDrop(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    DrvCmbA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub DrvCmbA_OLEDragEnter(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub DrvCmbA_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "DrvCmbA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub DrvCmbA_OLEDragLeave(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbA_OLEDragLeavePotentialTarget()
  AddLogEntry "DrvCmbA_OLEDragLeavePotentialTarget"
End Sub

Private Sub DrvCmbA_OLEDragMouseMove(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub DrvCmbA_OLEGiveFeedback(ByVal effect As CBLCtlsLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "DrvCmbA_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub DrvCmbA_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As CBLCtlsLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "DrvCmbA_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub DrvCmbA_OLEReceivedNewData(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub DrvCmbA_OLESetData(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub DrvCmbA_OLEStartDrag(ByVal data As CBLCtlsLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "DrvCmbA_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub DrvCmbA_OutOfMemory()
  AddLogEntry "DrvCmbA_OutOfMemory"
End Sub

Private Sub DrvCmbA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "DrvCmbA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  DrvCmbA.hWndShellUIParentWindow = Me.hWnd
End Sub

Private Sub DrvCmbA_RemovedItem(ByVal comboItem As CBLCtlsLibACtl.IVirtualDriveComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbA_RemovedItem: comboItem=Nothing"
  Else
    AddLogEntry "DrvCmbA_RemovedItem: comboItem=" & comboItem.Path
  End If
End Sub

Private Sub DrvCmbA_RemovingItem(ByVal comboItem As CBLCtlsLibACtl.IDriveComboBoxItem, cancelDeletion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbA_RemovingItem: comboItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "DrvCmbA_RemovingItem: comboItem=" & comboItem.Path & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub DrvCmbA_ResizedControlWindow()
  AddLogEntry "DrvCmbA_ResizedControlWindow"
End Sub

Private Sub DrvCmbA_SelectedDriveChanged(ByVal previousSelectedItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal newSelectedItem As CBLCtlsLibACtl.IDriveComboBoxItem)
  Dim str As String

  str = "DrvCmbA_SelectedDriveChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Path & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Path
  End If

  AddLogEntry str
End Sub

Private Sub DrvCmbA_SelectionCanceled()
  AddLogEntry "DrvCmbA_SelectionCanceled"
End Sub

Private Sub DrvCmbA_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal newSelectedItem As CBLCtlsLibACtl.IDriveComboBoxItem)
  Dim str As String

  str = "DrvCmbA_SelectionChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Path & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Path
  End If

  AddLogEntry str
End Sub

Private Sub DrvCmbA_SelectionChanging(ByVal newSelectedItem As CBLCtlsLibACtl.IDriveComboBoxItem, ByVal selectionFieldText As String, ByVal selectionFieldHasBeenEdited As Boolean, ByVal selectionChangeReason As CBLCtlsLibACtl.SelectionChangeReasonConstants, cancelChange As Boolean)
  Dim str As String

  str = "DrvCmbA_SelectionChanging: newSelectedItem="
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Path
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", selectionFieldHasBeenEdited=" & selectionFieldHasBeenEdited & ", selectionChangeReason=" & selectionChangeReason & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub DrvCmbA_Validate(Cancel As Boolean)
  AddLogEntry "DrvCmbA_Validate"
End Sub

Private Sub DrvCmbA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "DrvCmbA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_BeginSelectionChange()
  AddLogEntry "DrvCmbU_BeginSelectionChange"
End Sub

Private Sub DrvCmbU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "DrvCmbU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub DrvCmbU_CreatedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "DrvCmbU_CreatedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub DrvCmbU_CreatedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "DrvCmbU_CreatedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub DrvCmbU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_DestroyedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "DrvCmbU_DestroyedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub DrvCmbU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "DrvCmbU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub DrvCmbU_DestroyedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "DrvCmbU_DestroyedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub DrvCmbU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "DrvCmbU_DragDrop"
End Sub

Private Sub DrvCmbU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "DrvCmbU_DragOver"
End Sub

Private Sub DrvCmbU_FreeItemData(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbU_FreeItemData: comboItem=Nothing"
  Else
    AddLogEntry "DrvCmbU_FreeItemData: comboItem=" & comboItem.Path
  End If
End Sub

Private Sub DrvCmbU_GotFocus()
  Set objActiveCtl = DrvCmbU
  AddLogEntry "DrvCmbU_GotFocus"
End Sub

Private Sub DrvCmbU_InsertedItem(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbU_InsertedItem: comboItem=Nothing"
  Else
    AddLogEntry "DrvCmbU_InsertedItem: comboItem=" & comboItem.Path
  End If
End Sub

Private Sub DrvCmbU_InsertingItem(ByVal comboItem As CBLCtlsLibUCtl.IVirtualDriveComboBoxItem, cancelInsertion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbU_InsertingItem: comboItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "DrvCmbU_InsertingItem: comboItem=" & comboItem.Path & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub DrvCmbU_ItemBeginDrag(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ItemBeginDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ItemBeginRDrag(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ItemBeginRDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ItemGetDisplayInfo(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal requestedInfo As CBLCtlsLibUCtl.RequestedInfoConstants, IconIndex As Long, SelectedIconIndex As Long, OverlayIndex As Long, Indent As Long, ByVal maxItemTextLength As Long, itemText As String, dontAskAgain As Boolean)
  Dim str As String

  str = "DrvCmbU_ItemGetDisplayInfo: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Index
  End If
  str = str & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", SelectedIconIndex=" & SelectedIconIndex & ", OverlayIndex=" & OverlayIndex & ", Indent=" & Indent & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  AddLogEntry str
End Sub

Private Sub DrvCmbU_ItemMouseEnter(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ItemMouseEnter: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ItemMouseLeave(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ItemMouseLeave: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "DrvCmbU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub DrvCmbU_KeyPress(keyAscii As Integer)
  AddLogEntry "DrvCmbU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub DrvCmbU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "DrvCmbU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub DrvCmbU_ListClick(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ListClick: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ListCloseUp()
  AddLogEntry "DrvCmbU_ListCloseUp"
End Sub

Private Sub DrvCmbU_ListDropDown()
  AddLogEntry "DrvCmbU_ListDropDown"
End Sub

Private Sub DrvCmbU_ListMouseDown(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ListMouseDown: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ListMouseMove(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ListMouseMove: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub DrvCmbU_ListMouseUp(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ListMouseUp: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ListMouseWheel(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "DrvCmbU_ListMouseWheel: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ListOLEDragDrop(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_ListOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    DrvCmbU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub DrvCmbU_ListOLEDragEnter(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_ListOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ListOLEDragLeave(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_ListOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub DrvCmbU_ListOLEDragMouseMove(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_ListOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub DrvCmbU_LostFocus()
  AddLogEntry "DrvCmbU_LostFocus"
End Sub

Private Sub DrvCmbU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  'AddLogEntry "DrvCmbU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_OLECompleteDrag(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal performedEffect As CBLCtlsLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub DrvCmbU_OLEDragDrop(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    DrvCmbU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub DrvCmbU_OLEDragEnter(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub DrvCmbU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "DrvCmbU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub DrvCmbU_OLEDragLeave(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub DrvCmbU_OLEDragLeavePotentialTarget()
  AddLogEntry "DrvCmbU_OLEDragLeavePotentialTarget"
End Sub

Private Sub DrvCmbU_OLEDragMouseMove(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Path
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub DrvCmbU_OLEGiveFeedback(ByVal effect As CBLCtlsLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "DrvCmbU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub DrvCmbU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As CBLCtlsLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "DrvCmbU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub DrvCmbU_OLEReceivedNewData(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub DrvCmbU_OLESetData(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub DrvCmbU_OLEStartDrag(ByVal data As CBLCtlsLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "DrvCmbU_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub DrvCmbU_OutOfMemory()
  AddLogEntry "DrvCmbU_OutOfMemory"
End Sub

Private Sub DrvCmbU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "DrvCmbU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  DrvCmbU.hWndShellUIParentWindow = Me.hWnd
End Sub

Private Sub DrvCmbU_RemovedItem(ByVal comboItem As CBLCtlsLibUCtl.IVirtualDriveComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbU_RemovedItem: comboItem=Nothing"
  Else
    AddLogEntry "DrvCmbU_RemovedItem: comboItem=" & comboItem.Path
  End If
End Sub

Private Sub DrvCmbU_RemovingItem(ByVal comboItem As CBLCtlsLibUCtl.IDriveComboBoxItem, cancelDeletion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "DrvCmbU_RemovingItem: comboItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "DrvCmbU_RemovingItem: comboItem=" & comboItem.Path & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub DrvCmbU_ResizedControlWindow()
  AddLogEntry "DrvCmbU_ResizedControlWindow"
End Sub

Private Sub DrvCmbU_SelectedDriveChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IDriveComboBoxItem)
  Dim str As String

  str = "DrvCmbU_SelectedDriveChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Path & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Path
  End If

  AddLogEntry str
End Sub

Private Sub DrvCmbU_SelectionCanceled()
  AddLogEntry "DrvCmbU_SelectionCanceled"
End Sub

Private Sub DrvCmbU_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IDriveComboBoxItem)
  Dim str As String

  str = "DrvCmbU_SelectionChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Path & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Path
  End If

  AddLogEntry str
End Sub

Private Sub DrvCmbU_SelectionChanging(ByVal newSelectedItem As CBLCtlsLibUCtl.IDriveComboBoxItem, ByVal selectionFieldText As String, ByVal selectionFieldHasBeenEdited As Boolean, ByVal selectionChangeReason As CBLCtlsLibUCtl.SelectionChangeReasonConstants, cancelChange As Boolean)
  Dim str As String

  str = "DrvCmbU_SelectionChanging: newSelectedItem="
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Path
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", selectionFieldHasBeenEdited=" & selectionFieldHasBeenEdited & ", selectionChangeReason=" & selectionChangeReason & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub DrvCmbU_Validate(Cancel As Boolean)
  AddLogEntry "DrvCmbU_Validate"
End Sub

Private Sub DrvCmbU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub DrvCmbU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "DrvCmbU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_BeforeDrawText()
  AddLogEntry "ImgCmbA_BeforeDrawText"
End Sub

Private Sub ImgCmbA_BeginSelectionChange()
  AddLogEntry "ImgCmbA_BeginSelectionChange"
End Sub

Private Sub ImgCmbA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "ImgCmbA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub ImgCmbA_CreatedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "ImgCmbA_CreatedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub ImgCmbA_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ImgCmbA_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ImgCmbA_CreatedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "ImgCmbA_CreatedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub ImgCmbA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_DestroyedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "ImgCmbA_DestroyedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub ImgCmbA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ImgCmbA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ImgCmbA_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ImgCmbA_DestroyedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ImgCmbA_DestroyedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "ImgCmbA_DestroyedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub ImgCmbA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ImgCmbA_DragDrop"
End Sub

Private Sub ImgCmbA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "ImgCmbA_DragOver"
End Sub

Private Sub ImgCmbA_FreeItemData(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbA_FreeItemData: comboItem=Nothing"
  Else
    AddLogEntry "ImgCmbA_FreeItemData: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub ImgCmbA_GotFocus()
  Set objActiveCtl = ImgCmbA
  AddLogEntry "ImgCmbA_GotFocus"
End Sub

Private Sub ImgCmbA_InsertedItem(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbA_InsertedItem: comboItem=Nothing"
  Else
    AddLogEntry "ImgCmbA_InsertedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub ImgCmbA_InsertingItem(ByVal comboItem As CBLCtlsLibACtl.IVirtualImageComboBoxItem, cancelInsertion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbA_InsertingItem: comboItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ImgCmbA_InsertingItem: comboItem=" & comboItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ImgCmbA_ItemBeginDrag(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ItemBeginDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  'ImgCmbA.BeginDrag ImgCmbA.CreateItemContainer(comboItem), -1
  ImgCmbA.OLEDrag , , , ImgCmbA.CreateItemContainer(comboItem)
End Sub

Private Sub ImgCmbA_ItemBeginRDrag(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ItemBeginRDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ItemGetDisplayInfo(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal requestedInfo As CBLCtlsLibACtl.RequestedInfoConstants, IconIndex As Long, SelectedIconIndex As Long, OverlayIndex As Long, Indent As Long, ByVal maxItemTextLength As Long, itemText As String, dontAskAgain As Boolean)
  Dim str As String

  str = "ImgCmbA_ItemGetDisplayInfo: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Index
  End If
  str = str & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", SelectedIconIndex=" & SelectedIconIndex & ", OverlayIndex=" & OverlayIndex & ", Indent=" & Indent & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  AddLogEntry str

  If requestedInfo And RequestedInfoConstants.riItemText Then
    itemText = "Item " & (comboItem.Index + 1)
  End If
  If requestedInfo And RequestedInfoConstants.riIconIndex Then
    IconIndex = 0
  End If
End Sub

Private Sub ImgCmbA_ItemMouseEnter(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ItemMouseEnter: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ItemMouseLeave(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ItemMouseLeave: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ImgCmbA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ImgCmbA_KeyPress(keyAscii As Integer)
  AddLogEntry "ImgCmbA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ImgCmbA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ImgCmbA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ImgCmbA_ListClick(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ListClick: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ListCloseUp()
  AddLogEntry "ImgCmbA_ListCloseUp"
End Sub

Private Sub ImgCmbA_ListDropDown()
  AddLogEntry "ImgCmbA_ListDropDown"
End Sub

Private Sub ImgCmbA_ListMouseDown(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ListMouseDown: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ListMouseMove(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ListMouseMove: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub ImgCmbA_ListMouseUp(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ListMouseUp: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ListMouseWheel(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbA_ListMouseWheel: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ListOLEDragDrop(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_ListOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ImgCmbA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ImgCmbA_ListOLEDragEnter(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_ListOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ListOLEDragLeave(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_ListOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub ImgCmbA_ListOLEDragMouseMove(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_ListOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ImgCmbA_LostFocus()
  AddLogEntry "ImgCmbA_LostFocus"
End Sub

Private Sub ImgCmbA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  'AddLogEntry "ImgCmbA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_OLECompleteDrag(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal performedEffect As CBLCtlsLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ImgCmbA_OLEDragDrop(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ImgCmbA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ImgCmbA_OLEDragEnter(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub ImgCmbA_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ImgCmbA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ImgCmbA_OLEDragLeave(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbA_OLEDragLeavePotentialTarget()
  AddLogEntry "ImgCmbA_OLEDragLeavePotentialTarget"
End Sub

Private Sub ImgCmbA_OLEDragMouseMove(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub ImgCmbA_OLEGiveFeedback(ByVal effect As CBLCtlsLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ImgCmbA_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ImgCmbA_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As CBLCtlsLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "ImgCmbA_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ImgCmbA_OLEReceivedNewData(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ImgCmbA_OLESetData(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ImgCmbA_OLEStartDrag(ByVal data As CBLCtlsLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ImgCmbA_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ImgCmbA_OutOfMemory()
  AddLogEntry "ImgCmbA_OutOfMemory"
End Sub

Private Sub ImgCmbA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ImgCmbA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertImageComboItemsA
End Sub

Private Sub ImgCmbA_RemovedItem(ByVal comboItem As CBLCtlsLibACtl.IVirtualImageComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbA_RemovedItem: comboItem=Nothing"
  Else
    AddLogEntry "ImgCmbA_RemovedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub ImgCmbA_RemovingItem(ByVal comboItem As CBLCtlsLibACtl.IImageComboBoxItem, cancelDeletion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbA_RemovingItem: comboItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ImgCmbA_RemovingItem: comboItem=" & comboItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ImgCmbA_ResizedControlWindow()
  AddLogEntry "ImgCmbA_ResizedControlWindow"
End Sub

Private Sub ImgCmbA_SelectionCanceled()
  AddLogEntry "ImgCmbA_SelectionCanceled"
End Sub

Private Sub ImgCmbA_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal newSelectedItem As CBLCtlsLibACtl.IImageComboBoxItem)
  Dim str As String

  str = "ImgCmbA_SelectionChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Text & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ImgCmbA_SelectionChanging(ByVal newSelectedItem As CBLCtlsLibACtl.IImageComboBoxItem, ByVal selectionFieldText As String, ByVal selectionFieldHasBeenEdited As Boolean, ByVal selectionChangeReason As CBLCtlsLibACtl.SelectionChangeReasonConstants, cancelChange As Boolean)
  Dim str As String

  str = "ImgCmbA_SelectionChanging: newSelectedItem="
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Text
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", selectionFieldHasBeenEdited=" & selectionFieldHasBeenEdited & ", selectionChangeReason=" & selectionChangeReason & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ImgCmbA_TextChanged()
  AddLogEntry "ImgCmbA_TextChanged"
End Sub

Private Sub ImgCmbA_TruncatedText()
  AddLogEntry "ImgCmbA_TruncatedText"
End Sub

Private Sub ImgCmbA_Validate(Cancel As Boolean)
  AddLogEntry "ImgCmbA_Validate"
End Sub

Private Sub ImgCmbA_WritingDirectionChanged(ByVal newWritingDirection As CBLCtlsLibACtl.WritingDirectionConstants)
  AddLogEntry "ImgCmbA_WritingDirectionChanged: newWritingDirection=" & newWritingDirection
End Sub

Private Sub ImgCmbA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  AddLogEntry "ImgCmbA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_BeforeDrawText()
  AddLogEntry "ImgCmbU_BeforeDrawText"
End Sub

Private Sub ImgCmbU_BeginSelectionChange()
  AddLogEntry "ImgCmbU_BeginSelectionChange"
End Sub

Private Sub ImgCmbU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "ImgCmbU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub ImgCmbU_CreatedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "ImgCmbU_CreatedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub ImgCmbU_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ImgCmbU_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ImgCmbU_CreatedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "ImgCmbU_CreatedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub ImgCmbU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_DestroyedComboBoxControlWindow(ByVal hWndComboBox As Long)
  AddLogEntry "ImgCmbU_DestroyedComboBoxControlWindow: hWndComboBox=0x" & Hex(hWndComboBox)
End Sub

Private Sub ImgCmbU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ImgCmbU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ImgCmbU_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ImgCmbU_DestroyedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ImgCmbU_DestroyedListBoxControlWindow(ByVal hWndListBox As Long)
  AddLogEntry "ImgCmbU_DestroyedListBoxControlWindow: hWndListBox=0x" & Hex(hWndListBox)
End Sub

Private Sub ImgCmbU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ImgCmbU_DragDrop"
End Sub

Private Sub ImgCmbU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "ImgCmbU_DragOver"
End Sub

Private Sub ImgCmbU_FreeItemData(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbU_FreeItemData: comboItem=Nothing"
  Else
    AddLogEntry "ImgCmbU_FreeItemData: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub ImgCmbU_GotFocus()
  Set objActiveCtl = ImgCmbU
  AddLogEntry "ImgCmbU_GotFocus"
End Sub

Private Sub ImgCmbU_InsertedItem(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbU_InsertedItem: comboItem=Nothing"
  Else
    AddLogEntry "ImgCmbU_InsertedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub ImgCmbU_InsertingItem(ByVal comboItem As CBLCtlsLibUCtl.IVirtualImageComboBoxItem, cancelInsertion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbU_InsertingItem: comboItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ImgCmbU_InsertingItem: comboItem=" & comboItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ImgCmbU_ItemBeginDrag(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ItemBeginDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  'ImgCmbU.BeginDrag ImgCmbU.CreateItemContainer(comboItem), -1
  ImgCmbU.OLEDrag , , , ImgCmbU.CreateItemContainer(comboItem)
End Sub

Private Sub ImgCmbU_ItemBeginRDrag(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal selectionFieldText As String, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ItemBeginRDrag: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ItemGetDisplayInfo(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal requestedInfo As CBLCtlsLibUCtl.RequestedInfoConstants, IconIndex As Long, SelectedIconIndex As Long, OverlayIndex As Long, Indent As Long, ByVal maxItemTextLength As Long, itemText As String, dontAskAgain As Boolean)
  Dim str As String

  str = "ImgCmbU_ItemGetDisplayInfo: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Index
  End If
  str = str & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", SelectedIconIndex=" & SelectedIconIndex & ", OverlayIndex=" & OverlayIndex & ", Indent=" & Indent & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  AddLogEntry str

  If requestedInfo And RequestedInfoConstants.riItemText Then
    itemText = "Item " & (comboItem.Index + 1)
  End If
  If requestedInfo And RequestedInfoConstants.riIconIndex Then
    IconIndex = 0
  End If
End Sub

Private Sub ImgCmbU_ItemMouseEnter(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ItemMouseEnter: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ItemMouseLeave(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ItemMouseLeave: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ImgCmbU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ImgCmbU_KeyPress(keyAscii As Integer)
  AddLogEntry "ImgCmbU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ImgCmbU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ImgCmbU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ImgCmbU_ListClick(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ListClick: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ListCloseUp()
  AddLogEntry "ImgCmbU_ListCloseUp"
End Sub

Private Sub ImgCmbU_ListDropDown()
  AddLogEntry "ImgCmbU_ListDropDown"
End Sub

Private Sub ImgCmbU_ListMouseDown(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ListMouseDown: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ListMouseMove(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ListMouseMove: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub ImgCmbU_ListMouseUp(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ListMouseUp: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ListMouseWheel(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "ImgCmbU_ListMouseWheel: comboItem="
  If comboItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & comboItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ListOLEDragDrop(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_ListOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ImgCmbU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ImgCmbU_ListOLEDragEnter(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_ListOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ListOLEDragLeave(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_ListOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub ImgCmbU_ListOLEDragMouseMove(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_ListOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ImgCmbU_LostFocus()
  AddLogEntry "ImgCmbU_LostFocus"
End Sub

Private Sub ImgCmbU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  'AddLogEntry "ImgCmbU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_OLECompleteDrag(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal performedEffect As CBLCtlsLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ImgCmbU_OLEDragDrop(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ImgCmbU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ImgCmbU_OLEDragEnter(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub ImgCmbU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ImgCmbU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ImgCmbU_OLEDragLeave(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ImgCmbU_OLEDragLeavePotentialTarget()
  AddLogEntry "ImgCmbU_OLEDragLeavePotentialTarget"
End Sub

Private Sub ImgCmbU_OLEDragMouseMove(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub ImgCmbU_OLEGiveFeedback(ByVal effect As CBLCtlsLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ImgCmbU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ImgCmbU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As CBLCtlsLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "ImgCmbU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ImgCmbU_OLEReceivedNewData(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ImgCmbU_OLESetData(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ImgCmbU_OLEStartDrag(ByVal data As CBLCtlsLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ImgCmbU_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ImgCmbU_OutOfMemory()
  AddLogEntry "ImgCmbU_OutOfMemory"
End Sub

Private Sub ImgCmbU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ImgCmbU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertImageComboItemsU
End Sub

Private Sub ImgCmbU_RemovedItem(ByVal comboItem As CBLCtlsLibUCtl.IVirtualImageComboBoxItem)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbU_RemovedItem: comboItem=Nothing"
  Else
    AddLogEntry "ImgCmbU_RemovedItem: comboItem=" & comboItem.Text
  End If
End Sub

Private Sub ImgCmbU_RemovingItem(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem, cancelDeletion As Boolean)
  If comboItem Is Nothing Then
    AddLogEntry "ImgCmbU_RemovingItem: comboItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ImgCmbU_RemovingItem: comboItem=" & comboItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ImgCmbU_ResizedControlWindow()
  AddLogEntry "ImgCmbU_ResizedControlWindow"
End Sub

Private Sub ImgCmbU_SelectionCanceled()
  AddLogEntry "ImgCmbU_SelectionCanceled"
End Sub

Private Sub ImgCmbU_SelectionChanged(ByVal previousSelectedItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal newSelectedItem As CBLCtlsLibUCtl.IImageComboBoxItem)
  Dim str As String

  str = "ImgCmbU_SelectionChanged: previousSelectedItem="
  If previousSelectedItem Is Nothing Then
    str = str & "Nothing, newSelectedItem="
  Else
    str = str & previousSelectedItem.Text & ", newSelectedItem="
  End If
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ImgCmbU_SelectionChanging(ByVal newSelectedItem As CBLCtlsLibUCtl.IImageComboBoxItem, ByVal selectionFieldText As String, ByVal selectionFieldHasBeenEdited As Boolean, ByVal selectionChangeReason As CBLCtlsLibUCtl.SelectionChangeReasonConstants, cancelChange As Boolean)
  Dim str As String

  str = "ImgCmbU_SelectionChanging: newSelectedItem="
  If newSelectedItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newSelectedItem.Text
  End If
  str = str & ", selectionFieldText=" & selectionFieldText & ", selectionFieldHasBeenEdited=" & selectionFieldHasBeenEdited & ", selectionChangeReason=" & selectionChangeReason & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ImgCmbU_TextChanged()
  AddLogEntry "ImgCmbU_TextChanged"
End Sub

Private Sub ImgCmbU_TruncatedText()
  AddLogEntry "ImgCmbU_TruncatedText"
End Sub

Private Sub ImgCmbU_Validate(Cancel As Boolean)
  AddLogEntry "ImgCmbU_Validate"
End Sub

Private Sub ImgCmbU_WritingDirectionChanged(ByVal newWritingDirection As CBLCtlsLibUCtl.WritingDirectionConstants)
  AddLogEntry "ImgCmbU_WritingDirectionChanged: newWritingDirection=" & newWritingDirection
End Sub

Private Sub ImgCmbU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ImgCmbU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  AddLogEntry "ImgCmbU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub LstA_AbortedDrag()
  AddLogEntry "LstA_AbortedDrag"
End Sub

Private Sub LstA_CaretChanged(ByVal previousCaretItem As CBLCtlsLibACtl.IListBoxItem, ByVal newCaretItem As CBLCtlsLibACtl.IListBoxItem)
  Dim str As String

  str = "LstA_CaretChanged: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub LstA_Click(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_Click: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_CompareItems(ByVal firstItem As CBLCtlsLibACtl.IListBoxItem, ByVal secondItem As CBLCtlsLibACtl.IListBoxItem, ByVal Locale As Long, result As CBLCtlsLibACtl.CompareResultConstants)
  Dim str As String

  str = "LstA_CompareItems: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & firstItem.Index & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Index
  End If
  str = str & ", locale=" & Locale & ", result=" & result

  AddLogEntry str
End Sub

Private Sub LstA_ContextMenu(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_ContextMenu: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_DblClick(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_DblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "LstA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub LstA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "LstA_DragDrop"
End Sub

Private Sub LstA_DragMouseMove(dropTarget As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim str As String

  str = "LstA_DragMouseMove: dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub LstA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "LstA_DragOver"
End Sub

Private Sub LstA_Drop(ByVal dropTarget As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_Drop: dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_FreeItemData(ByVal listItem As CBLCtlsLibACtl.IListBoxItem)
  If listItem Is Nothing Then
    AddLogEntry "LstA_FreeItemData: listItem=Nothing"
  Else
    AddLogEntry "LstA_FreeItemData: listItem=" & listItem.Text
  End If
End Sub

Private Sub LstA_GotFocus()
  Set objActiveCtl = LstA
  AddLogEntry "LstA_GotFocus"
End Sub

Private Sub LstA_InsertedItem(ByVal listItem As CBLCtlsLibACtl.IListBoxItem)
  If listItem Is Nothing Then
    AddLogEntry "LstA_InsertedItem: listItem=Nothing"
  Else
    AddLogEntry "LstA_InsertedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub LstA_InsertingItem(ByVal listItem As CBLCtlsLibACtl.IVirtualListBoxItem, cancelInsertion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "LstA_InsertingItem: listItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "LstA_InsertingItem: listItem=" & listItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub LstA_ItemBeginDrag(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim col As CBLCtlsLibACtl.ListBoxItems

  If listItem Is Nothing Then
    AddLogEntry "LstA_ItemBeginDrag: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "LstA_ItemBeginDrag: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If

  On Error Resume Next
  Set col = LstA.ListItems
  col.FilterType(CBLCtlsLibACtl.FilteredPropertyConstants.fpSelected) = CBLCtlsLibACtl.FilterTypeConstants.ftIncluding
  col.Filter(CBLCtlsLibACtl.FilteredPropertyConstants.fpSelected) = Array(True)

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  LstA.BeginDrag LstA.CreateItemContainer(col), -1
End Sub

Private Sub LstA_ItemBeginRDrag(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  If listItem Is Nothing Then
    AddLogEntry "LstA_ItemBeginRDrag: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "LstA_ItemBeginRDrag: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub LstA_ItemGetDisplayInfo(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal requestedInfo As CBLCtlsLibACtl.RequestedInfoConstants, IconIndex As Long, OverlayIndex As Long)
  If listItem Is Nothing Then
    AddLogEntry "LstA_ItemGetDisplayInfo: listItem=Nothing, requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", OverlayIndex=" & OverlayIndex
  Else
    AddLogEntry "LstA_ItemGetDisplayInfo: listItem=" & listItem.Index & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", OverlayIndex=" & OverlayIndex
  End If

  If requestedInfo And RequestedInfoConstants.riIconIndex Then
    IconIndex = 0
  End If
End Sub

Private Sub LstA_ItemGetInfoTipText(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "LstA_ItemGetInfoTipText: listItem=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "LstA_ItemGetInfoTipText: listItem=" & listItem.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If infoTipText <> "" Then
      infoTipText = infoTipText & vbNewLine & "Index: " & listItem.Index & vbNewLine & "ItemData: " & listItem.ItemData
    Else
      infoTipText = "Index: " & listItem.Index & vbNewLine & "ItemData: " & listItem.ItemData
    End If
  End If
End Sub

Private Sub LstA_ItemMouseEnter(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_ItemMouseEnter: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_ItemMouseLeave(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_ItemMouseLeave: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "LstA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (LstA.DraggedItems Is Nothing) Then LstA.EndDrag True
  End If
End Sub

Private Sub LstA_KeyPress(keyAscii As Integer)
  AddLogEntry "LstA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub LstA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "LstA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub LstA_LostFocus()
  AddLogEntry "LstA_LostFocus"
End Sub

Private Sub LstA_MClick(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_MDblClick(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_MeasureItem(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ItemHeight As stdole.OLE_YSIZE_PIXELS)
  Dim str As String

  str = "LstA_MeasureItem: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", ItemHeight=" & ItemHeight

  AddLogEntry str
  ItemHeight = 17
End Sub

Private Sub LstA_MouseDown(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MouseDown: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_MouseEnter(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MouseEnter: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_MouseHover(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MouseHover: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_MouseLeave(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MouseLeave: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_MouseMove(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MouseMove: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub LstA_MouseUp(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MouseUp: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If button = MouseButtonConstants.vbLeftButton Then
    If Not (LstA.DraggedItems Is Nothing) Then
      ' are we within the client area?
      If ((HitTestConstants.htAbove Or HitTestConstants.htBelow Or HitTestConstants.htToLeft Or HitTestConstants.htToRight) And hitTestDetails) = 0 Then
        ' yes
        LstA.EndDrag False
      Else
        ' no
        LstA.EndDrag True
      End If
    End If
  End If
End Sub

Private Sub LstA_MouseWheel(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_MouseWheel: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_OLECompleteDrag(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal performedEffect As CBLCtlsLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "LstA_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub LstA_OLEDragDrop(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "LstA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    LstA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub LstA_OLEDragEnter(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "LstA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub LstA_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "LstA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub LstA_OLEDragLeave(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "LstA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_OLEDragLeavePotentialTarget()
  AddLogEntry "LstA_OLEDragLeavePotentialTarget"
End Sub

Private Sub LstA_OLEDragMouseMove(ByVal data As CBLCtlsLibACtl.IOLEDataObject, effect As CBLCtlsLibACtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "LstA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub LstA_OLEGiveFeedback(ByVal effect As CBLCtlsLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "LstA_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub LstA_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As CBLCtlsLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "LstA_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub LstA_OLEReceivedNewData(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "LstA_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub LstA_OLESetData(ByVal data As CBLCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "LstA_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub LstA_OLEStartDrag(ByVal data As CBLCtlsLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "LstA_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub LstA_OutOfMemory()
  AddLogEntry "LstA_OutOfMemory"
End Sub

Private Sub LstA_OwnerDrawItem(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal requiredAction As CBLCtlsLibACtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibACtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibACtl.RECTANGLE)
  Dim str As String

  str = "LstA_OwnerDrawItem: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", requiredAction=0x" & Hex(requiredAction) & ", itemState=0x" & Hex(itemState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & ", " & drawingRectangle.Top & ")-(" & drawingRectangle.Right & ", " & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub LstA_ProcessCharacterInput(ByVal keyAscii As Integer, ByVal shift As Integer, itemToSelect As Long)
  AddLogEntry "LstA_ProcessCharacterInput: keyAscii=" & keyAscii & ", shift=" & shift & ", itemToSelect=" & itemToSelect
End Sub

Private Sub LstA_ProcessKeyStroke(ByVal keyCode As Integer, ByVal shift As Integer, itemToSelect As Long)
  AddLogEntry "LstA_ProcessKeyStroke: keyCode=" & keyCode & ", shift=" & shift & ", itemToSelect=" & itemToSelect
End Sub

Private Sub LstA_RClick(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_RClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_RDblClick(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_RDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "LstA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertListItemsA
End Sub

Private Sub LstA_RemovedItem(ByVal listItem As CBLCtlsLibACtl.IVirtualListBoxItem)
  If listItem Is Nothing Then
    AddLogEntry "LstA_RemovedItem: listItem=Nothing"
  Else
    AddLogEntry "LstA_RemovedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub LstA_RemovingItem(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, cancelDeletion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "LstA_RemovingItem: listItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "LstA_RemovingItem: listItem=" & listItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub LstA_ResizedControlWindow()
  AddLogEntry "LstA_ResizedControlWindow"
End Sub

Private Sub LstA_SelectionChanged()
  AddLogEntry "LstA_SelectionChanged"
End Sub

Private Sub LstA_Validate(Cancel As Boolean)
  AddLogEntry "LstA_Validate"
End Sub

Private Sub LstA_XClick(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_XClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstA_XDblClick(ByVal listItem As CBLCtlsLibACtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibACtl.HitTestConstants)
  Dim str As String

  str = "LstA_XDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_AbortedDrag()
  AddLogEntry "LstU_AbortedDrag"
End Sub

Private Sub LstU_CaretChanged(ByVal previousCaretItem As CBLCtlsLibUCtl.IListBoxItem, ByVal newCaretItem As CBLCtlsLibUCtl.IListBoxItem)
  Dim str As String

  str = "LstU_CaretChanged: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub LstU_Click(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_Click: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_CompareItems(ByVal firstItem As CBLCtlsLibUCtl.IListBoxItem, ByVal secondItem As CBLCtlsLibUCtl.IListBoxItem, ByVal Locale As Long, result As CBLCtlsLibUCtl.CompareResultConstants)
  Dim str As String

  str = "LstU_CompareItems: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & firstItem.Index & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Index
  End If
  str = str & ", locale=" & Locale & ", result=" & result

  AddLogEntry str
End Sub

Private Sub LstU_ContextMenu(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_ContextMenu: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_DblClick(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_DblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "LstU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub LstU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "LstU_DragDrop"
End Sub

Private Sub LstU_DragMouseMove(dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim str As String

  str = "LstU_DragMouseMove: dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub LstU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "LstU_DragOver"
End Sub

Private Sub LstU_Drop(ByVal dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_Drop: dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_FreeItemData(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem)
  If listItem Is Nothing Then
    AddLogEntry "LstU_FreeItemData: listItem=Nothing"
  Else
    AddLogEntry "LstU_FreeItemData: listItem=" & listItem.Text
  End If
End Sub

Private Sub LstU_GotFocus()
  Set objActiveCtl = LstU
  AddLogEntry "LstU_GotFocus"
End Sub

Private Sub LstU_InsertedItem(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem)
  If listItem Is Nothing Then
    AddLogEntry "LstU_InsertedItem: listItem=Nothing"
  Else
    AddLogEntry "LstU_InsertedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub LstU_InsertingItem(ByVal listItem As CBLCtlsLibUCtl.IVirtualListBoxItem, cancelInsertion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "LstU_InsertingItem: listItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "LstU_InsertingItem: listItem=" & listItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub LstU_ItemBeginDrag(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim col As CBLCtlsLibUCtl.ListBoxItems

  If listItem Is Nothing Then
    AddLogEntry "LstU_ItemBeginDrag: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "LstU_ItemBeginDrag: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If

  On Error Resume Next
  Set col = LstU.ListItems
  col.FilterType(CBLCtlsLibUCtl.FilteredPropertyConstants.fpSelected) = CBLCtlsLibUCtl.FilterTypeConstants.ftIncluding
  col.Filter(CBLCtlsLibUCtl.FilteredPropertyConstants.fpSelected) = Array(True)

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  LstU.BeginDrag LstU.CreateItemContainer(col), -1
End Sub

Private Sub LstU_ItemBeginRDrag(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  If listItem Is Nothing Then
    AddLogEntry "LstU_ItemBeginRDrag: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "LstU_ItemBeginRDrag: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub LstU_ItemGetDisplayInfo(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal requestedInfo As CBLCtlsLibUCtl.RequestedInfoConstants, IconIndex As Long, OverlayIndex As Long)
  If listItem Is Nothing Then
    AddLogEntry "LstU_ItemGetDisplayInfo: listItem=Nothing, requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", OverlayIndex=" & OverlayIndex
  Else
    AddLogEntry "LstU_ItemGetDisplayInfo: listItem=" & listItem.Index & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", OverlayIndex=" & OverlayIndex
  End If

  If requestedInfo And RequestedInfoConstants.riIconIndex Then
    IconIndex = 0
  End If
End Sub

Private Sub LstU_ItemGetInfoTipText(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "LstU_ItemGetInfoTipText: listItem=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "LstU_ItemGetInfoTipText: listItem=" & listItem.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If infoTipText <> "" Then
      infoTipText = infoTipText & vbNewLine & "Index: " & listItem.Index & vbNewLine & "ItemData: " & listItem.ItemData
    Else
      infoTipText = "Index: " & listItem.Index & vbNewLine & "ItemData: " & listItem.ItemData
    End If
  End If
End Sub

Private Sub LstU_ItemMouseEnter(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_ItemMouseEnter: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_ItemMouseLeave(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_ItemMouseLeave: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "LstU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (LstU.DraggedItems Is Nothing) Then LstU.EndDrag True
  End If
End Sub

Private Sub LstU_KeyPress(keyAscii As Integer)
  AddLogEntry "LstU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub LstU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "LstU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub LstU_LostFocus()
  AddLogEntry "LstU_LostFocus"
End Sub

Private Sub LstU_MClick(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_MDblClick(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_MeasureItem(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ItemHeight As stdole.OLE_YSIZE_PIXELS)
  Dim str As String

  str = "LstU_MeasureItem: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", ItemHeight=" & ItemHeight

  AddLogEntry str
  ItemHeight = 17
End Sub

Private Sub LstU_MouseDown(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MouseDown: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_MouseEnter(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MouseEnter: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_MouseHover(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MouseHover: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_MouseLeave(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MouseLeave: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_MouseMove(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MouseMove: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub LstU_MouseUp(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MouseUp: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If button = MouseButtonConstants.vbLeftButton Then
    If Not (LstU.DraggedItems Is Nothing) Then
      ' are we within the client area?
      If ((HitTestConstants.htAbove Or HitTestConstants.htBelow Or HitTestConstants.htToLeft Or HitTestConstants.htToRight) And hitTestDetails) = 0 Then
        ' yes
        LstU.EndDrag False
      Else
        ' no
        LstU.EndDrag True
      End If
    End If
  End If
End Sub

Private Sub LstU_MouseWheel(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As CBLCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_MouseWheel: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_OLECompleteDrag(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal performedEffect As CBLCtlsLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "LstU_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub LstU_OLEDragDrop(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "LstU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    LstU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub LstU_OLEDragEnter(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "LstU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub LstU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "LstU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub LstU_OLEDragLeave(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "LstU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_OLEDragLeavePotentialTarget()
  AddLogEntry "LstU_OLEDragLeavePotentialTarget"
End Sub

Private Sub LstU_OLEDragMouseMove(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, effect As CBLCtlsLibUCtl.OLEDropEffectConstants, dropTarget As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "LstU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub LstU_OLEGiveFeedback(ByVal effect As CBLCtlsLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "LstU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub LstU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As CBLCtlsLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "LstU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub LstU_OLEReceivedNewData(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "LstU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub LstU_OLESetData(ByVal data As CBLCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "LstU_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub LstU_OLEStartDrag(ByVal data As CBLCtlsLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "LstU_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub LstU_OutOfMemory()
  AddLogEntry "LstU_OutOfMemory"
End Sub

Private Sub LstU_OwnerDrawItem(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal requiredAction As CBLCtlsLibUCtl.OwnerDrawActionConstants, ByVal itemState As CBLCtlsLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As CBLCtlsLibUCtl.RECTANGLE)
  Dim str As String

  str = "LstU_OwnerDrawItem: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", requiredAction=0x" & Hex(requiredAction) & ", itemState=0x" & Hex(itemState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & ", " & drawingRectangle.Top & ")-(" & drawingRectangle.Right & ", " & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub LstU_ProcessCharacterInput(ByVal keyAscii As Integer, ByVal shift As Integer, itemToSelect As Long)
  AddLogEntry "LstU_ProcessCharacterInput: keyAscii=" & keyAscii & ", shift=" & shift & ", itemToSelect=" & itemToSelect
End Sub

Private Sub LstU_ProcessKeyStroke(ByVal keyCode As Integer, ByVal shift As Integer, itemToSelect As Long)
  AddLogEntry "LstU_ProcessKeyStroke: keyCode=" & keyCode & ", shift=" & shift & ", itemToSelect=" & itemToSelect
End Sub

Private Sub LstU_RClick(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_RClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_RDblClick(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_RDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "LstU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertListItemsU
End Sub

Private Sub LstU_RemovedItem(ByVal listItem As CBLCtlsLibUCtl.IVirtualListBoxItem)
  If listItem Is Nothing Then
    AddLogEntry "LstU_RemovedItem: listItem=Nothing"
  Else
    AddLogEntry "LstU_RemovedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub LstU_RemovingItem(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, cancelDeletion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "LstU_RemovingItem: listItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "LstU_RemovingItem: listItem=" & listItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub LstU_ResizedControlWindow()
  AddLogEntry "LstU_ResizedControlWindow"
End Sub

Private Sub LstU_SelectionChanged()
  AddLogEntry "LstU_SelectionChanged"
End Sub

Private Sub LstU_Validate(Cancel As Boolean)
  AddLogEntry "LstU_Validate"
End Sub

Private Sub LstU_XClick(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_XClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub LstU_XDblClick(ByVal listItem As CBLCtlsLibUCtl.IListBoxItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As CBLCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "LstU_XDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

Private Sub InsertComboItemsA()
  Dim i As Long

  With CmbA.ComboItems
    For i = 1 To 20
      .Add "Item " & CStr(i), , i
    Next i
  End With
End Sub

Private Sub InsertComboItemsU()
  Dim i As Long

  With CmbU.ComboItems
    For i = 1 To 20
      .Add "Item " & CStr(i), , i
    Next i
  End With
End Sub

Private Sub InsertImageComboItemsA()
  Dim cImages As Long
  Dim i As Long

  ImgCmbA.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ImgCmbA.ComboItems
    For i = 1 To 20
      .Add "Item " & CStr(i), , (i - 1) Mod cImages, i Mod cImages, , , i
    Next i
  End With
End Sub

Private Sub InsertImageComboItemsU()
  Dim cImages As Long
  Dim i As Long

  ImgCmbU.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ImgCmbU.ComboItems
    For i = 1 To 20
      .Add "Item " & CStr(i), , (i - 1) Mod cImages, i Mod cImages, , , i
    Next i
  End With
End Sub

Private Sub InsertListItemsA()
  Dim i As Long

  With LstA.ListItems
    For i = 1 To 20
      .Add "Item " & CStr(i), , i
    Next i
  End With
End Sub

Private Sub InsertListItemsU()
  Dim i As Long

  With LstU.ListItems
    For i = 1 To 20
      .Add "Item " & CStr(i), , i
    Next i
  End With
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong CmbU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong DrvCmbU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ImgCmbU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong LstU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong CmbA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong DrvCmbA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ImgCmbA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong LstA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
