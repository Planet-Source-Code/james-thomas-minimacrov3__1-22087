VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1125
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   2280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider sldMaxLoop 
      Height          =   255
      Left            =   600
      TabIndex        =   3
      ToolTipText     =   "Controls the Number of times a Macro Repeats"
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.TextBox txtLoop 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "1"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox lblData 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2265
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1518
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":196A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":1F04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblButtons 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Play Macro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Macro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Close"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Configure MiniMacro"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Edit HotKeys"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Properties"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "main"
      Visible         =   0   'False
      Begin VB.Menu mnuRecord 
         Caption         =   "&Record"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowEdit 
         Caption         =   "Show Editor"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Show Properties"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHotKeys 
         Caption         =   "Show HotKeys"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_SETTEXT = &HC
Private Const EM_SETREADONLY = &HCF

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetTheFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private CloseWhenDone As Boolean 'When true the program ends after completing the macro
Public Sub SelectOption(iButton As Integer)
Dim lLoop As Long
Select Case iButton
    Case 1 'Record
        bKeyStop = False
        StartRecord
        Call WatchForCancel
    Case 2 'Play
        bKeyStop = False
        For lLoop = 0 To CLng(txtLoop.Text) - 1
            UnLoadHotKeys
            OpenMacro (True)
            StartPlay
            Call WatchForCancel
            DoEvents
            If bKeyStop = True Then
                Exit For
            End If
        Next lLoop
    Case 3 ' Stop
        StopMacro
    Case 4 'Open
        OpenMacro
    Case 5
        Call ShellOn(False)
        Unload Me
    Case 6
        frmProp.Show
    Case 7 'Show Editor from the Systray
        frmEditor.Visible = True
    Case 8 'Show the Hotkey editor
        frmHotKeys.Show
End Select
End Sub
Private Sub Form_Activate()
Dim lngResult As Long
Call frmPosition(Me, 4)
Call OnTop(Me)
sldMaxLoop.Max = lMaxLoop
If bStartUp = True Then
    Me.Visible = False
    frmEditor.mnuShowEdit.Checked = False
Else
    frmEditor.mnuShowEdit.Checked = True
End If
End Sub



Private Sub Form_Load()
If App.PrevInstance Then
    End
End If
strOSVER = GetOSVer()
Me.Caption = App.Title
sStopKey = GetSetting(App.ProductName, "Key", "Key", "S")
sEXT = GetSetting(App.ProductName, "Extentions", 0, ".mac")
bStartUp = GetSetting(App.ProductName, "Settings", "SysTray", False)
lMaxLoop = GetSetting(App.ProductName, "Settings", "MaxLoop", 30)
PC_SPEED = GetSetting(App.ProductName, "Settings", "PCSpeed", 60)
tblButtons.Buttons(2).Enabled = False
frmEditor.mnuPlay.Enabled = False
tblButtons.Buttons(3).Enabled = False
frmEditor.mnuStop.Enabled = False
frmEditor.sldMaxLoop.Max = lMaxLoop
frmEditor.txtLoop = 1
frmEditor.mnuProp.Checked = False
frmEditor.mnuHotKeys.Checked = False

'Check the Command Line for a pass macro name
If Command <> "" Then
    CloseWhenDone = True
End If
If CloseWhenDone = False Then
    tblButtons.Buttons(2).Enabled = False
    ShellOn (True)
    Show
    If bStartUp = True Then
        Me.Visible = False
        frmEditor.mnuShowEdit.Checked = True
    Else
        frmEditor.mnuShowEdit.Checked = False
    End If
Else
    tblButtons.Buttons(2).Enabled = True
    ShellOn (False)
    End
End If
lDisplayHwnd = lblData.hwnd
Call CreateCaret(lblData.hwnd, 0, 0, 0) 'Erase the flashing Cursor
Call SendMessage(lblData.hwnd, EM_SETREADONLY, 1, 0) 'Set the Textbox to ReadOnly without Disabling the property (because that makes it font grayed out)
Call SendMessage(txtLoop.hwnd, EM_SETREADONLY, 1, 0) 'Set the Textbox to ReadOnly without Disabling the property (because that makes it font grayed out)
RegHotKeys
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim tmpLong As Single
 Dim apierror As Long
    tmpLong = x / Screen.TwipsPerPixelX
    Select Case tmpLong 'For system tray icon
        Case WM_LBUTTONUP
            apierror = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmEditor.PopupMenu mnuMain
        Case WM_RBUTTONUP
            apierror = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmEditor.PopupMenu mnuMain
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ShellOn(False)
Call Unhook(Me.hwnd)
End Sub

Private Sub lblData_GotFocus()
Call SetTheFocus(tblButtons.hwnd)
End Sub

Private Sub mnuExit_Click()
On Error GoTo fin
SelectOption (5)
Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Sub

Private Sub mnuHotKeys_Click()

    On Error GoTo fin

If frmHotKeys.Visible = False Then
    frmEditor.mnuHotKeys.Checked = True
    SelectOption (8)
Else
    frmEditor.mnuHotKeys.Checked = False
End If


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub mnuOpen_Click()

    On Error GoTo fin

SelectOption (4)


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub mnuPlay_Click()

    On Error GoTo fin

SelectOption (2)


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub mnuProp_Click()

    On Error GoTo fin


If frmProp.Visible = False Then
    frmEditor.mnuProp.Checked = True
    SelectOption (6)
Else
    frmEditor.mnuProp.Checked = False
    Unload frmProp
End If



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub mnuRecord_Click()

    On Error GoTo fin

SelectOption (1)


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub mnuShowEdit_Click()

    On Error GoTo fin


If frmEditor.Visible = False Then
    frmEditor.mnuShowEdit.Checked = True
    SelectOption (7)
Else
    frmEditor.mnuShowEdit.Checked = False
    frmEditor.Visible = False
End If


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub mnuStop_Click()

    On Error GoTo fin

SelectOption (3)


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub sldMaxLoop_Change()

    On Error GoTo fin

frmEditor.txtLoop.Text = frmEditor.sldMaxLoop.Value


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub tblButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo fin

SelectOption (Button.Index)


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub tblButtons_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo fin

Call Unhook(frmEditor.hwnd)
Select Case ButtonMenu.Index
    Case 1
        frmHotKeys.Show
    Case 2
        frmProp.Show
End Select



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Sub


