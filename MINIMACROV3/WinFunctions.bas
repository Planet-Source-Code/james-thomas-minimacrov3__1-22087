Attribute VB_Name = "WinFunctions"
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const conHwndTopmost = -1
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const flags = SWP_NOMOVE + SWP_NOSIZE
'**************************************
'Windows API/Global Declarations for :Re
'     gister File Association
'**************************************
Option Explicit
'BGS 10.23.2000 Constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
'Public Const KEY_ALL_ACCESS = ((&H1F00
'     00 Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 O
'     r &H20) And (Not &H100000))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const SHCNE_ASSOCCHANGED = &H8000000
Public Const SHCNF_IDLIST = &H0&
Public Const APP_PATH_EXE = "App.Path & ""\"" & App.EXEName"
'BGS 10.23.2000 API


Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Public Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Const SPI_GETWORKAREA = 48


Private Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub RestOnBar(F As Form)
    Dim RC As RECT
    Dim x As Long
    x = SystemParametersInfo(SPI_GETWORKAREA, vbNull, RC, 0)
    F.Move RC.left * _
    Screen.TwipsPerPixelX, RC.top * _
    Screen.TwipsPerPixelY, RC.Right * _
    Screen.TwipsPerPixelX, RC.Bottom * _
    Screen.TwipsPerPixelY
End Sub

        

Public Function SaveRegSetting(ByVal plHKEY As Long, ByVal psSection As String, ByVal psKey As String, ByVal psSetting As String) As Boolean

    On Error GoTo fin

    Dim lRet As Long
    Dim lhKey As Long
    Dim lResult As Long
    lRet = RegCreateKey(plHKEY, psSection, lhKey)


    If lRet = ERROR_SUCCESS Then
        psSetting = psSetting & vbNullChar
        lRet = RegSetValueEx(lhKey, psKey, 0&, REG_SZ, ByVal psSetting, Len(psSetting))
        lRet = RegCloseKey(lhKey)
    End If
    SaveRegSetting = (lRet = ERROR_SUCCESS)


    If Not SaveRegSetting Then
        Err.Raise -9999
    End If

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function



Public Function RegFileAssociation(psEXT As String, _
    Optional psEXEPath As String = APP_PATH_EXE, _
    Optional pbUseBigIcon As Boolean = True) As Boolean
    On Error GoTo EH
    Dim sEXT As String 'BGS the ETX without the dot.
    Dim sEXEPathIcon As String
    RegFileAssociation = True
    'BGS 10.27.2000 Allow the ext to be pass
    '     ed with dot or
    'without a dot.
    psEXT = Replace(psEXT, ".", vbNullString, , vbTextCompare)
    psEXT = LCase(psEXT)
    sEXT = psEXT
    psEXT = "." & psEXT
    'BGS Allow the exe Path to be passed wit
    '     h .exe or with out it.
    'As well, Concatinate proper strings to
    '     be passed to the Registry


    If psEXEPath = APP_PATH_EXE Then
        psEXEPath = App.Path & "\" & App.EXEName
    End If
    psEXEPath = Replace(psEXEPath, ".exe", vbNullString, , vbTextCompare)
    sEXEPathIcon = psEXEPath & ".exe,0"
    psEXEPath = """" & psEXEPath & ".exe"" " & "%1"
    'BGS update the registry to Auto Open th
    '     e parameter specified Extentions
    'with the parameter specified exe applic
    '     ation.
    'First set up HKEY_CLASSES_ROOT
    SaveRegSetting HKEY_CLASSES_ROOT, psEXT, vbNullString, sEXT & "_auto_file"
    SaveRegSetting HKEY_CLASSES_ROOT, sEXT & "_auto_file", vbNullString, UCase(sEXT) & " File"
    SaveRegSetting HKEY_CLASSES_ROOT, sEXT & "_auto_file\Shell\Open", vbNullString, vbNullString
    'BGS The Command line string sent to the
    '     registry has to look something like this
    '     ...
    ' "C:\Program Files\LaunchARViewer\MyApp
    '     licationName.exe" %1
    SaveRegSetting HKEY_CLASSES_ROOT, sEXT & "_auto_file\Shell\Open\Command", vbNullString, psEXEPath
    'BGS Now do HKEY_LOCAL_MACHINE
    SaveRegSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\" & psEXT, vbNullString, sEXT & "_auto_file"
    SaveRegSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\" & sEXT & "_auto_file", vbNullString, UCase(sEXT) & " File"
    SaveRegSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\" & sEXT & "_auto_file\Shell\Open", vbNullString, vbNullString
    'BGS The Command line string sent to the
    '     registry has to look something like this
    '     ...
    ' "C:\Program Files\LaunchARViewer\MyApp
    '     licationName.exe" %1
    SaveRegSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\" & sEXT & "_auto_file\Shell\Open\Command", vbNullString, psEXEPath
    'BGS Set the Icon to be the EXE ICON if
    '     pbUseBigIcon is true.


    If pbUseBigIcon Then
        SaveRegSetting HKEY_CLASSES_ROOT, sEXT & "_auto_file\DefaultIcon", vbNullString, sEXEPathIcon
        SaveRegSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\" & sEXT & "_auto_file\DefaultIcon", vbNullString, sEXEPathIcon
    Else
        SaveRegSetting HKEY_CLASSES_ROOT, sEXT & "_auto_file\DefaultIcon", vbNullString, vbNullString
        SaveRegSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\" & sEXT & "_auto_file\DefaultIcon", vbNullString, vbNullString
    End If
    'BGS Refresh the Icons
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    Exit Function
EH:
    Err.Clear
    RegFileAssociation = False
End Function

Public Function OnTop(frmForm As Form)
On Error GoTo fin
Dim hwnd As Long
hwnd = frmForm.hwnd
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function


Public Sub frmPosition(vForm As Form, vPosition As Integer)
 
Select Case vPosition
    Case 0
    ' TopRight
        vForm.top = Screen.Height - Screen.Height
            vForm.left = Screen.Width - vForm.Width
    Case 1
    ' TopLeft
        vForm.top = Screen.Height - Screen.Height
            vForm.left = Screen.Width - Screen.Width
    Case 2
    ' BottomLeft
        vForm.top = Screen.Height - vForm.Height
            vForm.left = Screen.Width - Screen.Width
    Case 3
    ' BottomRight
        vForm.top = Screen.Height - vForm.Height
            vForm.left = Screen.Width - vForm.Width
    Case 4
        'TOP Center
        vForm.top = Screen.Height - Screen.Height
        vForm.left = (Screen.Width / 2) - (vForm.Width / 2)
End Select

'NoOffScreen vForm
'Exit Sub

End Sub

Public Sub NoOffScreen(vForm As Form)

    On Error GoTo fin

If vForm.WindowState = 1 Then Exit Sub ' form cannot be moved if minimized
    If vForm.top < Screen.Height - Screen.Height + 10 Then ' Top Less than 10
         vForm.top = Screen.Height - Screen.Height + 10
    End If
            If vForm.left < Screen.Width - Screen.Width + 10 Then ' Left Less than 10
            vForm.left = Screen.Width - Screen.Width + 10
    End If
         If vForm.top > Screen.Height - vForm.Height - 10 Then ' Left more than 10
             vForm.top = Screen.Height - vForm.Height - 10
    End If
            If vForm.left > Screen.Width - vForm.Width - 10 Then ' Bottom more than 10
            vForm.left = Screen.Width - vForm.Width - 10
    End If
        Exit Sub



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub
Public Function TrimNulls(ByVal StrValue As String) As String

    On Error GoTo fin

Dim intPos As Integer
intPos = InStr(StrValue, vbNullChar)
Select Case intPos
    Case 0
        
    Case 1
        StrValue = ""
    Case Is > 1
        StrValue = left$(StrValue, intPos - 1)
End Select
TrimNulls = StrValue

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


TrimNulls = StrValue
End Function

Public Function AddtoTaskbar()
Dim hTaskbar As Long
Dim hStartButton As Long
Dim l As Long
Dim r As RECT
hTaskbar = FindWindowEx(0, 0, "Shell_TrayWnd", 0&)
hStartButton = FindWindowEx(hTaskbar, 0, "BUTTON", 0&)
l = CreateWindowEx(0, "BUTTON", "TEXT", WS_VISIBLE + WS_CHILD, 0, 0, 0, 0, hTaskbar, vbNull, 1, vbNull)
End Function

Public Function RightInstr(String1 As String, String2 As String) As Integer

    On Error GoTo fin

Dim strHold As String
Dim a As String
Dim i As Integer

For i = 1 To Len(String1)
    a = Right(String1, i)
    strHold = strHold & left(a, 1)
Next i
i = InStr(1, strHold, String2)
RightInstr = Len(String1) - i - 1

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Function
