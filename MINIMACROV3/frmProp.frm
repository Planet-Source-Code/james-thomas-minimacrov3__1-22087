VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Macro Properties"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   Icon            =   "frmProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox mebLoop 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   3
      Format          =   "###"
      PromptChar      =   " "
   End
   Begin VB.CheckBox chkStartUp 
      Caption         =   "Start MiniMacro in the system tray?"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox cmbKeys 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Ctrl + Alt + S"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbExt 
      Height          =   315
      ItemData        =   "frmProp.frx":014A
      Left            =   1440
      List            =   "frmProp.frx":014C
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Maximun Loop Value"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "File Extention:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblStop 
      Caption         =   "Stop Keys:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddExt As Boolean
Dim WinHook As Long
Dim AddKey As Boolean
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" _
         Alias "SetWindowsHookExA" _
         (ByVal idHook As Long, _
         ByVal lpfn As Long, _
         ByVal hmod As Long, _
         ByVal dwThreadId As Long) As Long

      Private Declare Function PostMessage Lib "user32" _
         Alias "PostMessageA" _
         (ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Long) As Long
Private Const WM_KEYDOWN = &H100
Private Const WH_KEYBOARD = 2
Private Const KBH_MASK = &H20000000
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202



Private Sub cmbExt_Change()
AddExt = True
End Sub

Private Sub cmbExt_KeyPress(KeyAscii As Integer)
AddExt = True
End Sub

Private Sub cmbExt_LostFocus()
Dim hold As String
If AddExt = True And cmbExt Like ".???" Then
    hold = cmbExt
    cmbExt.Clear
    cmbExt.AddItem (hold)
    cmbExt.Refresh
    cmbExt = hold
    SaveExt
Else
    cmbExt = ""
    MsgBox "You must enter a valid four character file extention!", vbCritical, "File Extention"
End If
LoadExt
End Sub

Private Sub cmbKeys_GotFocus()
WinHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, App.ThreadID)

End Sub

Private Sub cmbKeys_KeyDown(KeyCode As Integer, Shift As Integer)
Dim m As MSG
'this will display the key combination
If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF16 Then
    cmbKeys.Text = "Ctrl + Alt + " & GetFunctionKey(KeyCode)
Else
    Select Case KeyCode
        Case vbKeyCancel To vbKeyHelp
            cmbKeys.Text = "Ctrl + Alt + " & GetFormatKeys(KeyCode)
        Case Else
            cmbKeys.Text = "Ctrl + Alt + " & UCase(Chr(KeyCode))
    End Select
End If
sStopKey = Chr(KeyCode)
'Now remove the Key from the queue
With m
    .hwnd = cmbKeys.hwnd
    .message = WM_KEYDOWN
    .lParam = KeyCode
End With
ret = PeekMessage(m, cmbKeys.hwnd, 0, 0, PM_REMOVE)
AddKey = True
'set the stop key sequence in the DLL

End Sub

Private Sub cmbKeys_LostFocus()
Call UnhookWindowsHookEx(WinHook)
End Sub

Private Sub Command1_Click()
RegFileAssociation (cmbExt.Text)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
hwndActiveBox = frmProp.cmbKeys.hwnd
UnLoadHotKeys
lStopKeyDisplay = cmbKeys.hwnd
AddExt = False
AddKey = False
LoadExt
LoadKey
frmEditor.mnuProp.Checked = True
bStartUp = GetSetting(App.Title, "Settings", "SysTray", False)
lMaxLoop = GetSetting(App.Title, "Settings", "MaxLoop", 30)
mebLoop.Text = lMaxLoop
If bStartUp = True Then
    chkStartUp.Value = 1
Else
    chkStartUp.Value = 0
End If
cmbExt = cmbExt.List(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnhookWindowsHookEx(WinHook)
Dim hold As String
If AddExt = True Then
    sEXT = cmbExt
    hold = cmbExt
    cmbExt.Clear
    cmbExt.AddItem (hold)
    cmbExt.Refresh
End If
SaveExt
SaveKey
frmEditor.mnuProp.Checked = False
Call SaveSetting(App.Title, "Settings", "SysTray", chkStartUp)
Call SaveSetting(App.Title, "Settings", "MaxLoop", lMaxLoop)
frmEditor.mnuProp.Checked = False
RegHotKeys
End Sub

Public Function LoadExt()
Dim i As Integer
Dim x As Integer
Dim b As Boolean
Dim hold As String
Dim ret As String
For i = 0 To 4
    ret = GetSetting(App.Title, "Extentions", i, ".mac")
    If ret <> "" Then
        For x = 0 To cmbExt.ListCount
            If ret <> cmbExt.List(x) Then
                b = True
            Else
                b = False
                Exit For
            End If
            'hold = ret
        Next x
        If b = True Then
            cmbExt.AddItem (ret)
        End If
    Else
        Exit For
    End If
Next i

End Function

Public Function LoadKey()
Dim ret As String
ret = GetSetting(App.Title, "Key", "Key", "S")
If ret <> "" Then
    cmbKeys.Text = "Ctrl + Alt + " & ret
ElseIf ret = "" Then
    cmbKeys.Text = "Ctrl + Alt + S"
End If
End Function

Public Function SaveKey()
If AddKey = True Then
    SaveSetting App.Title, "Key", "Key", Right(cmbKeys, 1)
End If
End Function

Public Function SaveExt()
Dim i As Integer
Dim j As Integer
Dim jret As String
Dim ret As String
Dim bNew As Boolean
bNew = True
    ret = GetSetting(App.Title, "Extentions", 0)
    If cmbExt.List(0) = "" Or ret <> cmbExt.List(0) Then
        j = 0
        jret = GetSetting(App.Title, "Extentions", j)
        Do Until jret = ""
            jret = GetSetting(App.Title, "Extentions", j)
            If jret <> "" Then
                If cmbExt.List(0) = jret Then
                    bNew = False
                    'Means that this extention has been previously
                    'entered into the Extention list. All other entries
                    'must be deleted before continuing.
                    Call DeleteSetting(App.Title, "Extentions", j)
                    'Now that its deleted, we have to collapse the
                    'Extentions that follow
                    'for entries before the find
                    i = j
                    jret = GetSetting(App.Title, "Extentions", 0)
                    Do Until jret = ""
                        jret = GetSetting(App.Title, "Extentions", i - 1)
                        If jret <> "" And i > -1 Then
                            Call SaveSetting(App.Title, "Extentions", i, jret)
                            i = i - 1
                        End If
                    Loop
                End If
                If j >= 3 Then
                    j = 3
                    Exit Do
                Else
                    j = j + 1
                End If
                'returns the first blank extention number
            End If
        Loop
        If bNew = True Then
            'if this is a new file extention, drop the last
            For i = j To 1 Step -1
                If i = 0 Then
                    Exit For
                End If
                jret = GetSetting(App.Title, "Extentions", i - 1)
                Call SaveSetting(App.Title, "Extentions", i, jret)
            Next i
        End If
        Call SaveSetting(App.Title, "Extentions", 0, cmbExt.List(0))
   End If

End Function


Private Sub mebLoop_Change()
lMaxLoop = Val(mebLoop.Text)
End Sub

