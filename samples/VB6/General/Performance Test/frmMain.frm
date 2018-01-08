VERSION 5.00
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.5#0"; "CBLCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ComboListBoxControls 1.5 - Performance Test Sample"
   ClientHeight    =   5040
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5580
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin CBLCtlsLibUCtl.ListBox LstU 
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      _cx             =   6588
      _cy             =   8493
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
      OwnerDrawItems  =   0
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
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill ListBox"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtItems 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Text            =   "50000"
      Top             =   150
      Width           =   735
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
      Left            =   3960
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Fill time: 0 ms"
      Height          =   195
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ite&ms:"
      Height          =   195
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private cItems As Long


  Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
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


Private Sub cmdAbout_Click()
  LstU.About
End Sub

Private Sub cmdFill_Click()
  Dim endTime As Long
  Dim i As Long
  Dim startTime As Long

  cmdFill.Enabled = False
  DoEvents
  DoEvents

  startTime = GetTickCount
  LstU.DontRedraw = True
  LstU.PrepareForItemInsertions cItems, Len("Item " & CStr(cItems))
  With LstU.ListItems
    .RemoveAll
    For i = 1 To cItems
      Call .Add("Item " & CStr(i))
    Next i
  End With
  LstU.DontRedraw = False
  endTime = GetTickCount

  lblTime.Caption = "Fill time: " & CStr(endTime - startTime) & " ms"
  cmdFill.Enabled = True
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  cItems = CLng(txtItems.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub txtItems_Change()
  On Error Resume Next
  cItems = CLng(txtItems.Text)
End Sub

Private Sub txtItems_LostFocus()
  txtItems.Text = cItems
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
    SendMessageAsLong LstU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
