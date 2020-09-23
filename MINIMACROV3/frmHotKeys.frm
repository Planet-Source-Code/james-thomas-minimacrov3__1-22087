VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmHotKeys 
   BorderStyle     =   0  'None
   Caption         =   "HotKey Registration"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   6705
   Icon            =   "frmHotKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHotKey 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3240
      TabIndex        =   3
      Text            =   "CTRL + ALT + A"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton End 
      Caption         =   "End"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid FlxGd 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      ScrollTrack     =   -1  'True
      GridLines       =   3
      AllowUserResizing=   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmHotKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Const WM_KEYUP = &H101
Private Const WH_KEYBOARD = 2
Private Const KBH_MASK = &H20000000
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Sub cmdOpen_Click()
Dim sfilter As String
Dim strRet As String
sfilter = "Macro Files (*" & sEXT & ")" & vbNullChar & "*" & sEXT & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*"
strRet = ShowOpen(Me.hwnd, sfilter)
If strRet <> "" Then
    FlxGd.Text = strRet
End If
'Resize the cell
'FlxGd.ColWidth(1) = TextWidth(FlxGd.Text) * 1.44
End Sub

Private Sub End_Click()
On Error GoTo fin
Dim iIndx As Integer
Dim keyindx As Long
Dim retReg As String
Dim bRegkeyOn As Boolean
Dim iShiftState As Integer
sAtoms(0, 0) = ""
bRegkeyOn = False
With FlxGd
    For iIndx = 0 To .Rows - 1
        If .TextMatrix(iIndx, 1) = "" Then
            Exit For
        End If
        bRegkeyOn = True
        If InStr(1, .TextMatrix(iIndx, 2), "CTRL") > 0 Then
            iShiftState = MOD_CONTROL
        ElseIf InStr(1, .TextMatrix(iIndx, 2), "ALT") > 0 Then
            iShiftState = iShiftState + MOD_ALT
        ElseIf InStr(1, .TextMatrix(iIndx, 2), "SHIFT") > 0 Then
            iShiftState = iShiftState + MOD_SHIFT
        End If
        Call SaveHotKey(.TextMatrix(iIndx, 1), .TextMatrix(iIndx, 2), iIndx)
        'Call SetHotKeys(GetFunctionKey(, Trim(Right(.TextMatrix(iIndx, 2), RightInstr(.TextMatrix(iIndx, 2), "+") - 1))), .TextMatrix(iIndx, 1), iShiftState)
    Next iIndx
    If iIndx = 0 Then
        keyindx = 0
        Do
            'Return the number of entries so I can Dimension the array
            retReg = GetSetting(App.Title, "HotKeys", keyindx, "")
            keyindx = keyindx + 1
        Loop Until retReg = ""
        keyindx = keyindx - 2
        For iIndx = 0 To keyindx
            Call DeleteSetting(App.Title, "HotKeys", iIndx)
        Next iIndx
    End If
End With
If bRegkeyOn = True Then
    RegHotKeys
End If
Unload Me
Exit Sub
fin:
Select Case Err.Number
    Case 9 'subscript out of range
        Resume Next
    Case Else
        MsgBox Err.Description & Err.Number
End Select
End Sub

Private Sub FlxGd_Scroll()
txtHotKey.Visible = False
cmdOpen.Visible = False
End Sub

Private Sub Form_Load()
Dim x As Integer
Dim sKeys() As String
hwndActiveBox = frmHotKeys.txtHotKey.hwnd
Call UnLoadHotKeys
'Combo1_Load
With FlxGd
    .ColAlignment(-1) = 1       'all Left alligned
    'Get the hotkeys from the registry
    sKeys() = LoadHotKeys()
    'Apply them to the grid
    For x = 0 To UBound(sKeys)
        .ColWidth(0) = TextWidth(Str(x)) * 2
        If sKeys(x, 0) = "" Then
            Exit For
        End If
        .TextMatrix(x, 0) = Str(x)
        .TextMatrix(x, 1) = sKeys(x, 0)
        If TextWidth(sKeys(x, 0)) * 1.44 > .ColWidth(1) Then
                .ColWidth(1) = TextWidth(sKeys(x, 0)) * 1.44
        End If
        If .ColWidth(1) = 0 Then
            .ColWidth(1) = 1000
        End If
        .TextMatrix(x, 2) = sKeys(x, 1)
        .ColWidth(2) = TextWidth(sKeys(x, 1)) * 2.5
        If x <> UBound(sKeys) Then
            AddGridRow
        End If
    Next
    .ColWidth(2) = TextWidth(txtHotKey) * 1.44
    .Row = 0
    .Col = 0
    
End With

End Sub


Private Sub FlxGd_EnterCell()
If FlxGd.Col <> 0 Then
    FlxGd.CellBackColor = &HC0FFFF    'lt. yellow
End If
FlxGd.Tag = ""                    'clear temp storage
End Sub

Private Sub FlxGd_LeaveCell()
If FlxGd.Col <> 0 Then
    FlxGd.CellBackColor = &H80000005 'white
End If
txtHotKey.Visible = False
cmdOpen.Visible = False
End Sub

Private Sub FlxGd_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 46                 '<Del>, clear cell
        FlxGd.Tag = FlxGd   'assign to temp storage
        FlxGd = ""
  End Select
End Sub

Private Sub FlxGd_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13            'ENTER key
            Advance_Cell   'advance new cell
        Case 8             'Backspace
            If Len(FlxGd) Then
              FlxGd = left$(FlxGd, Len(FlxGd) - 1)
            End If
        Case 27                      'ESC
            If FlxGd.Tag > "" Then   'only if not NULL
              FlxGd = FlxGd.Tag      'restore original text
            End If
        Case Else
            FlxGd = FlxGd + Chr(KeyAscii)
    End Select
End Sub

Private Sub FlxGd_Click()
    If txtHotKey.Visible = True Then
      txtHotKey.Visible = False
      Call UnhookWindowsHookEx(WinHook)
      FlxGd.CellBackColor = &H80000005  'white
    End If
    
    If FlxGd.Col = 2 Then     ' Position and size the ComboBox, then show it.
        txtHotKey.BackColor = &HC0FFFF
        txtHotKey.Font = FlxGd.Font
        txtHotKey.left = FlxGd.CellLeft + FlxGd.left + 10
        txtHotKey.Height = FlxGd.CellHeight
        txtHotKey.Width = FlxGd.CellWidth - 10
        txtHotKey.top = FlxGd.CellTop + FlxGd.top + 10
        txtHotKey.Text = FlxGd.Text
        txtHotKey.Visible = True
        WinHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, App.ThreadID)
    End If
    If FlxGd.Col = 1 Then
        cmdOpen.left = (FlxGd.CellLeft + FlxGd.CellWidth + FlxGd.left) - cmdOpen.Width
        cmdOpen.top = FlxGd.CellTop + FlxGd.top
        cmdOpen.Height = FlxGd.CellHeight
        cmdOpen.Visible = True
    Else
        cmdOpen.Visible = False
    End If
        
End Sub

Private Sub txtHotKey_Change()
If FlxGd.Col = 2 Then
      FlxGd.Text = txtHotKey.Text
End If
End Sub

Private Sub txtHotKey_Click()  ' Place the selected item into the Cell and hide the ComboBox.
    If FlxGd.Col = 2 Then
      FlxGd.Text = txtHotKey.Text
      txtHotKey.Visible = False
      Call UnhookWindowsHookEx(WinHook)
    End If
End Sub

Private Sub Combo1_Load()   ' Load the ComboBox's list.
'Dim iIndx As Integer
'FlxGd.RowHeightMin = txtHotKey.Height
'txtHotKey.Visible = False
'txtHotKey.Clear
'txtHotKey.Refresh
'For iIndx = 65 To 90
'    txtHotKey.AddItem "CTRL + ALT + " & Chr(iIndx)
'Next iIndx
End Sub
   
Private Sub FlxGd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Row As Integer, Col As Integer
    Row = FlxGd.MouseRow
    Col = FlxGd.MouseCol
    If Button = 2 And (Col = 0 Or Row = 0) Then
      FlxGd.Col = IIf(Col = 0, 1, Col)
      FlxGd.Row = IIf(Row = 0, 1, Row)
      'PopupMenu MnuFGridRows
    End If
End Sub

Private Sub AddGridRow()
    With FlxGd
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, 0) = Str(.Row)
    End With
txtHotKey.Visible = False
cmdOpen.Visible = False
End Sub


Private Sub DeleteGridRow()
    Dim Row As Integer, n As Integer, x As Integer, iIndx As Integer, retReg As String
    With FlxGd
          Row = .Row
          For n = 1 To .Cols - 1
             If .TextMatrix(Row, n) > "" Then
               x = 1
               Exit For
             End If
          Next
          If x Then
            n = MsgBox("Data in Row" + Str$(Row) + ".  Delete anyway?", vbYesNo, "Delete Row...")
          End If
          If x = 0 Or n = 6 Then 'no exist. data or YES
            'refresh the grid
             iIndx = 0
            Do Until GetSetting(App.Title, "HotKeys", iIndx, "") = ""
                'Delete everthing in the HotKey registry
                Call DeleteSetting(App.Title, "HotKeys", iIndx)
                'retReg = GetSetting(App.Title, "HotKeys", iIndx, "")
                iIndx = iIndx + 1
            Loop
            For n = .Row To .Rows - 2      'move exist data up 1 row
               For x = 1 To FlxGd.Cols - 1
                  .TextMatrix(n, x) = .TextMatrix(n + 1, x)
               Next
            Next
            'If Row = .Rows - 1 Then     'set new cursor row
            '  .Row = .Rows - 2
            'End If
            .Rows = .Rows - 1           'delete last row
          End If
    End With
txtHotKey.Visible = False
cmdOpen.Visible = False
End Sub

Private Sub Advance_Cell()                  'advance to next cell
    With FlxGd
        .HighLight = flexHighlightNever     'turn off hi-lite
        If .Col < .Cols - 1 Then
          .Col = .Col + 1
        Else
          If .Row < .Rows - 1 Then
            .Row = .Row + 1                 'down 1 row
            .Col = 1                        'first column
          Else
            .Row = 1
            .Col = 1
          End If
        End If
        If .CellTop + .CellHeight > .top + .Height Then
          .TopRow = .TopRow + 1             'make sure row is visible
        End If
        .HighLight = flexHighlightAlways    'turn on hi-lite
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnhookWindowsHookEx(WinHook)
frmEditor.mnuHotKeys.Checked = False
End Sub

Private Sub mnuAdd_Click()
AddGridRow
End Sub


Private Sub mnuClose_Click()
End_Click
End Sub

Private Sub mnuDelete_Click()
DeleteGridRow
End Sub

Private Sub txtHotKey_KeyDown(KeyCode As Integer, Shift As Integer)
Dim m As MSG
Dim ret As Long
Dim strShift As String
strShift = GetShiftState(Shift)
If strShift <> "" Then
    strShift = strShift & " + "
End If
'this will display the key combination
If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF16 Then
    txtHotKey.Text = strShift & GetFunctionKey(KeyCode)
Else
    Select Case KeyCode
        Case vbKeyEscape To vbKeyHelp
            txtHotKey.Text = strShift & GetFormatKeys(KeyCode)
        Case Else
            If KeyCode >= vbKeyA Then
                txtHotKey.Text = strShift & UCase(Chr(KeyCode))
            End If
    End Select
End If
'Now remove the Key from the queue
KeyCode = 0
With m
    .hwnd = txtHotKey.hwnd
    .message = WM_KEYDOWN
    .lParam = KeyCode
End With
ret = PeekMessage(m, txtHotKey.hwnd, 0, 0, PM_REMOVE)
'set the stop key sequence in the DLL
End Sub

Private Sub txtHotKey_KeyUp(KeyCode As Integer, Shift As Integer)
Dim m As MSG
Dim ret As Long
With m
    .hwnd = txtHotKey.hwnd
    .message = WM_KEYUP
    .lParam = KeyCode
End With
ret = PeekMessage(m, txtHotKey.hwnd, 0, 0, PM_REMOVE)
End Sub

Private Sub txtHotKey_LostFocus()
Call UnhookWindowsHookEx(WinHook)
End Sub
