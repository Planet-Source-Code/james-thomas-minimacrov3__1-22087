Attribute VB_Name = "basHotKeys"
Option Explicit
'hotkey constants
Public Const WM_HOTKEY = &H312
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33

'used by the RegisterHotKey method
Public Enum RegisterHotKeyModifiers
   MOD_ALT = &H1
   MOD_CONTROL = &H2
   MOD_SHIFT = &H4
End Enum

Public Declare Function RegisterHotKey Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal id As Long, _
   ByVal fsModifiers As RegisterHotKeyModifiers, _
   ByVal vk As KeyCodeConstants) As Long
   
Public Declare Function UnregisterHotKey Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal id As Long) As Long

Public Declare Function GlobalAddAtom Lib "kernel32" _
   Alias "GlobalAddAtomA" _
  (ByVal lpString As String) As Long
   
Public Declare Function GlobalDeleteAtom Lib "kernel32" _
   (ByVal nAtom As Long) As Long

Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal MSG As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   
Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC As Long = (-4)

Public lpPrevWndProc As Long

'used by the PrintScreen method
Public Declare Function BitBlt Lib "gdi32" _
  (ByVal hDCDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long

Public Declare Function GetDesktopWindow Lib _
   "user32" () As Long

Public Declare Function GetWindowDC Lib _
   "user32" (ByVal hwnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32" _
   (ByVal hwnd As Long, ByVal hDC As Long) As Long
   
Public lMainhWnd As Long
Public sAtoms() As String
   
Public Sub Hook(ByVal gHW As Long)

    On Error GoTo fin


  'Establish a hook
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Public Sub SaveHotKey(sFileName As String, sKey As String, Optional iIndex As Integer)

    On Error GoTo fin

Dim keyindx As Integer
Dim retReg As String
If IsMissing(iIndex) Then
    keyindx = 0
    Do
        'This will give us the last used index
        retReg = GetSetting(App.ProductName, "HotKeys", keyindx, "")
        If retReg = "" Then
            Exit Do
        End If
        keyindx = keyindx + 1
    Loop Until retReg = ""
Else
    keyindx = iIndex
End If
Call SaveSetting(App.ProductName, "HotKeys", keyindx, sFileName & "," & sKey)



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Public Sub SetHotKeys(iKey As Integer, sFileName As String, iShiftState As Integer)

    On Error GoTo fin

On Error GoTo fin
Dim RetArray() As String
Dim iArrayIndx As Integer
Dim iIndx As Integer
Dim ret As Long
iIndx = UBound(sAtoms) + 1
If sAtoms(0, 0) = "" Then
    iIndx = iArrayIndx
    ReDim RetArray(iIndx, 1)
Else
    ReDim RetArray(iIndx, 1)
    For iArrayIndx = 0 To iIndx - 1
        RetArray(iArrayIndx, 0) = sAtoms(iArrayIndx, 0)
        RetArray(iArrayIndx, 1) = sAtoms(iArrayIndx, 1)
    Next iArrayIndx
End If
RetArray(iIndx, 0) = GlobalAddAtom(UCase(Chr(iKey)))
RetArray(iIndx, 1) = sFileName
ret = RegisterHotKey(frmEditor.hwnd, CLng(RetArray(UBound(RetArray), 0)), iShiftState, iKey)
ReDim sAtoms(iIndx, 1)
sAtoms() = RetArray()
Exit Sub
fin:
Select Case Err
    Case 9 'subscript error
        iIndx = 0
        Resume Next
    Case Else
        MsgBox Err.Description & Err.Number
End Select

End Sub

Public Sub Unhook(ByVal gHW As Long)

    On Error GoTo fin

'Reset the message handler
Call SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub


Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
                    ByVal wParam As Long, ByVal lParam As Long) As Long
Dim iIndx As Integer
   Select Case hwnd
      Case frmEditor.hwnd
         If uMsg = WM_HOTKEY Then
           'add code to process the Hotkey
            'wParam returns the Atom Code
            For iIndx = 0 To UBound(sAtoms)
                If wParam = sAtoms(iIndx, 0) Then
                    'Put Play code here but first unhook the subclass and hotkeys
                    Call UnLoadHotKeys
                    sFileName = sAtoms(iIndx, 1)
                    WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
                    Call frmEditor.SelectOption(2)
                    Exit Function
                End If
            Next iIndx
         End If
      Case Else
   End Select
  'Pass message to the original window message handler
   WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
   
End Function


Private Sub PrintScreen()

  Dim hWndDesk As Long
  Dim hDCDesk As Long

  Dim LeftDesk As Long
  Dim TopDesk As Long
  Dim WidthDesk As Long
  Dim HeightDesk As Long
   
 'define the screen coordinates (upper
 'corner (0,0) and lower corner (Width, Height)
  LeftDesk = 0
  TopDesk = 0
  WidthDesk = Screen.Width \ Screen.TwipsPerPixelX
  HeightDesk = Screen.Height \ Screen.TwipsPerPixelY
   
 'get the desktop handle and display context
  hWndDesk = GetDesktopWindow()
  hDCDesk = GetWindowDC(hWndDesk)
   
 'copy the desktop to the picture box
  'Call BitBlt(frmMain.Picture1.hdc, 0, 0, _
  '           WidthDesk, HeightDesk, hDCDesk, _
  '           LeftDesk, TopDesk, vbSrcCopy)

  Call ReleaseDC(hWndDesk, hDCDesk)

End Sub

Public Function LoadHotKeys() As String()
Dim retHotKeys() As String
Dim retReg As String
Dim keyindx As Integer
'the registry will hold a string formated like this:
' "c:\macrofile.mac,H"
' The first part is the file name of the macro to be run
' the Second is the HotKey reference all Hot keys will
' be CTRL+ALT+ whatever hot key.
keyindx = 0
Do
    'Return the number of entries so I can Dimension the array
    retReg = GetSetting(App.ProductName, "HotKeys", keyindx, "")
    keyindx = keyindx + 1
Loop Until retReg = ""
keyindx = keyindx - 1
ReDim Preserve retHotKeys(keyindx, 1)
keyindx = 0
Do
    retReg = GetSetting(App.ProductName, "HotKeys", keyindx, "")
    If retReg <> "" Then
        retHotKeys(keyindx, 0) = Mid(retReg, 1, InStr(1, retReg, ",") - 1)
        retHotKeys(keyindx, 1) = Mid(retReg, InStr(1, retReg, ",") + 1, Len(retReg) - InStr(1, retReg, ","))
    End If
    keyindx = keyindx + 1
Loop Until retReg = ""
LoadHotKeys = retHotKeys()
End Function


Public Sub UnLoadHotKeys()
On Error GoTo fin
    'unregister hotkey
Dim iIndx As Long
For iIndx = 0 To UBound(sAtoms)
   Call UnregisterHotKey(frmEditor.hwnd, Val(sAtoms(iIndx, 0)))
Next iIndx
Unhook (frmEditor.hwnd)
Exit Sub
fin:
Select Case Err
    Case 9 'Subscript out of rang
        Exit Sub
    Case Else
        MsgBox Err.Number & Err.Description
        Resume Next
End Select
End Sub
   
Public Sub RegHotKeys()
Dim keyindx As Long
Dim sKey() As String
Dim iShiftState As Integer
sKey() = LoadHotKeys()
keyindx = 0

For keyindx = 0 To UBound(sKey) - 1
    iShiftState = 0
    If InStr(1, sKey(keyindx, 1), "CTRL") > 0 Then
        iShiftState = MOD_CONTROL
    ElseIf InStr(1, sKey(keyindx, 1), "ALT") > 0 Then
        iShiftState = iShiftState + MOD_ALT
    ElseIf InStr(1, sKey(keyindx, 1), "SHIFT") > 0 Then
        iShiftState = iShiftState + MOD_SHIFT
    End If
    'Return the number of entries so I can Dimension the array
    Call SetHotKeys(GetFunctionKey(, Trim(Right(sKey(keyindx, 1), RightInstr(sKey(keyindx, 1), "+") - 1))), sKey(keyindx, 0), iShiftState)
Next keyindx
Call Hook(frmEditor.hwnd)
End Sub

