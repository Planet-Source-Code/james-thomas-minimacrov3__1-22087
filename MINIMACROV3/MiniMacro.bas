Attribute VB_Name = "Macro"
'***************************************************
'Name: MiniMacro.bas
'Created: Jan 2000 By James Thomas
'Description: This is the primary module for
'   the MiniMacro application. It is used for
'   retreiving and setting values that are
'   essential the proper operation the application.
'****************************************************
'Procedures:
'4)GetSysDir
'5)ShellOn
'****************************************************
'Dependencies:
'1)
'****************************************************
'Modified History:
'1) 23-Jan-2001: Altered the mMacro Data type
'   to remove the AsciiHold element. This was no
'   longer needed after a change to the MiniMacro.dll.
'****************************************************
Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000
'Dialog Constants
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256
Public Const LF_FACESIZE = 32
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
Public Const DN_DEFAULTPRN = &H1

'Hook Constants
Public Const WH_KEYBOARD = 2
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_JOURNALRECORD = 0
Public Const WH_CBT = 5
'Hook Call Constants
Public Const HC_SYSMODALOFF = 5
Public Const HC_SYSMODALON = 4
Public Const HC_SKIP = 2
Public Const HC_GETNEXT = 1
Public Const HC_ACTION = 0

'Systray Constants
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
'Message Queue Constants
Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_SETTEXT = &HC
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_CANCELJOURNAL = &H4B


'Operating System Version
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'Windows Properties
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10

Public Const GWL_HINSTANCE = (-6)
Public Const HCBT_ACTIVATE = 5

'Virtual Key Contants
Public Const VK_SHIFT = &H10
Public Const VK_PAUSE = &H13
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const KEYEVENTF_KEYUP = &H2

'MiniMacro Constants



Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Type PRINTDLGS
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hDC As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sfilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Public Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Public Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hDC As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String          '  return the style field here
    nFontType As Integer          '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type
Public Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

'Systray Type
Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Public Type EVENTMSG
        message As Long
        paramL As Long
        paramH As Long
        time As Long
        hwnd As Long
End Type


Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    ' Maintenance string for PSS usage.
End Type

'Kernal Declares
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Sub FreeLibraryAndExitThread Lib "kernel32" (ByVal hLibModule As Long, ByVal dwExitCode As Long)
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub CopyMemoryT2H Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Long, Source As EVENTMSG, ByVal Length As Long)
Public Declare Sub CopyMemoryH2T Lib "kernel32" Alias "RtlMoveMemory" (Dest As EVENTMSG, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer


'Shell Declares
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'User32 Declares
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long


'Common Dialog Declares
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Public Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Public Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long

'Public Properties
Public pos As POINTAPI
Public o As OSVERSIONINFO
Public CloseWhenDone As Boolean 'When true the program ends after completing the macro
Public starttime As Long 'Hold the Time the record or play is
Public MiniMacroHWnd As Long
Public NOTIFY As NOTIFYICONDATA
Public lX As Long
Public lY As Long
Public strCmdLine As String
Public EventArr() As EVENTMSG
Public EventLog As Long
Public PlayLog As Long
Public hHook As Long
Public hPlay As Long
Public recOK As Long
Public canPlay As Boolean
Public bDelay As Boolean
Public sFileName As String 'Current Open File
Public iFileNumber As String
Public lDisplayHwnd As Long
Public pt As POINTAPI
Public bRecording As Boolean
Public bPlaying As Boolean
Public sEXT As String 'File extension
Public sStopKey As String 'Stop Key for key combination
Public bStartUp As Boolean 'if true then starts up in the systray
Public lMaxLoop As Long 'Controls that maximum number of times a Macro will repeat.
Public lStartTime As Long
Public lStopKeyDisplay As Long
Public PC_SPEED As Integer
'Dialog Properties
Public ColorDialog As CHOOSECOLORS
Public FontDialog As CHOOSEFONTS
Public PrintDialog As PRINTDLGS
Public ParenthWnd As Long
Public bCancel As Boolean
Public bKeyStop As Boolean
Public strOSVER As String
Public hwndActiveBox As Long


Function GetOSVer() As String

    On Error GoTo fin

    Dim osv As OSVERSIONINFO
    osv.dwOSVersionInfoSize = Len(osv)
    If GetVersionEx(osv) = 1 Then
        Select Case osv.dwPlatformId
            Case VER_PLATFORM_WIN32s
                GetOSVer = "Windows 3.x"
            Case VER_PLATFORM_WIN32_WINDOWS
                Select Case osv.dwMinorVersion
                    Case 0
                        If InStr(osv.szCSDVersion, "C") Then
                            GetOSVer = "Windows 95 OSR2"
                        Else
                            GetOSVer = "Windows 95"
                        End If
                    Case 10
                        If InStr(osv.szCSDVersion, "A") Then
                            GetOSVer = "Windows 98 SE"
                        Else
                            GetOSVer = "Windows 98"
                        End If
                    Case 90
                        GetOSVer = "Windows Me"
                End Select
            Case VER_PLATFORM_WIN32_NT
                Select Case osv.dwMajorVersion
                    Case 3
                        Select Case osv.dwMinorVersion
                            Case 0
                                GetOSVer = "Windows NT 3"
                            Case 1
                                GetOSVer = "Windows NT 3.1"
                            Case 5
                                GetOSVer = "Windows NT 3.5"
                            Case 51
                                GetOSVer = "Windows NT 3.51"
                        End Select
                    Case 4
                        GetOSVer = "Windows NT 4"
                    Case 5
                        Select Case osv.dwMinorVersion
                            Case 0
                                GetOSVer = "Windows 2000"
                            Case 1
                                GetOSVer = "Whistler"
                        End Select
            End Select
        End Select
    End If



Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function

Public Function GetShiftState(lShift As Integer) As String

    On Error GoTo fin

Select Case lShift
      Case 1 ' or vbShiftMask
         GetShiftState = "SHIFT"
      Case 2 ' or vbCtrlMask
         GetShiftState = "CTRL"
      Case 4 ' or vbAltMask
         GetShiftState = "ALT"
      Case 3
         GetShiftState = "SHIFT + CTRL"
      Case 5
         GetShiftState = "SHIFT + ALT"
      Case 6
         GetShiftState = "CTRL + ALT"
      Case 7
         GetShiftState = "SHIFT + CTRL + ALT"
      End Select





Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function

Private Sub ProcessMessages()

    On Error GoTo fin

    Dim message As MSG
    'loop until bCancel is set to True
    Do While Not bCancel
        'wait for a message
        WaitMessage
        'check if it's a HOTKEY-message
        If PeekMessage(message, frmEditor.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            'minimize the form
            'WindowState = vbMinimized
        End If
        'let the operating system process other events
        DoEvents
    Loop

Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Sub

Public Sub OpenMacro(Optional bHotKey As Boolean = False)

    On Error GoTo fin

'bHotKey means that a filename has been provided
'Yes it would have made more sence to call it
'bOpenFile but OH Well
Dim FileLength As Long
Dim sfilter As String
sfilter = "Macro Files (*" & sEXT & ")" & vbNullChar & "*" & sEXT & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*"
iFileNumber = FreeFile()
If bHotKey = False Then
    sFileName = ShowOpen(frmEditor.hwnd, sfilter)
End If
If sFileName <> vbNullChar Then
    ReDim EventArr(0)
    Open sFileName For Random Access Read As iFileNumber Len = Len(EventArr(0))
        FileLength = 1
        Do
            ReDim Preserve EventArr(FileLength - 1)
            Get #iFileNumber, FileLength, EventArr(FileLength - 1)
            FileLength = FileLength + 1
        Loop Until EOF(iFileNumber)
        EventLog = FileLength - 2
    Close iFileNumber
    frmEditor.tblButtons.Buttons(2).Enabled = True
    frmEditor.mnuPlay.Enabled = True
    frmEditor.Caption = sFileName
Else
    frmEditor.tblButtons.Buttons(2).Enabled = False
    frmEditor.mnuPlay.Enabled = False
    frmEditor.Caption = App.Title
End If

Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Sub
Public Sub SaveMacro()

    On Error GoTo fin

Dim iIndex As Long
Dim sfilter As String
sfilter = "Macro Files (*" & sEXT & ")" & vbNullChar & "*" & sEXT & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*"
iFileNumber = FreeFile()
DoEvents
sFileName = ShowSave(frmEditor.hwnd, True, sfilter)
If sFileName <> vbNullChar Then
    If Mid(Right(sFileName, 4), 1, 1) <> "." Then
        sFileName = sFileName & sEXT
    End If
    If Len(Dir(sFileName)) > 0 Then
        Kill sFileName
    End If
    Open sFileName For Random Access Write As iFileNumber Len = Len(EventArr(EventLog))
        For iIndex = 1 To EventLog
            Put iFileNumber, iIndex, EventArr(iIndex - 1)
        Next iIndex
    Close iFileNumber
End If
Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Sub

Public Function ShowOpen(Optional ByVal hwnd As Long = -1, Optional sfilter As String = "Macro Files (*.mac)" & vbNullChar & "*.mac" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*") As String

    On Error GoTo fin

Dim ret As Long
Dim Count As Integer
Dim fileNameHolder As String
Dim LastCharacter As Integer
Dim NewCharacter As Integer
Dim tempFiles(1 To 200) As String
Dim hInst As Long
Dim Thread As Long
Dim FileDialog As OPENFILENAME
    FileDialog.nStructSize = Len(FileDialog)
    If IsMissing(hwnd) Then
        FileDialog.hwndOwner = GetActiveWindow()
    Else
        FileDialog.hwndOwner = hwnd
    End If
    ParenthWnd = FileDialog.hwndOwner
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = FileDialog.sFile & Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    FileDialog.sfilter = sfilter
    FileDialog.flags = OFS_FILE_OPEN_FLAGS
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    'If centerForm = True Then
    '    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    'Else
    '    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    'End If
    ret = GetOpenFileName(FileDialog)
    If ret Then
       ShowOpen = TrimNulls(FileDialog.sFile)
    Else
        ShowOpen = vbNullChar
    End If
Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function
Public Function ShowSave(Optional ByVal hwnd As Long = -1, Optional ByVal centerForm As Boolean, Optional sfilter As String = "Macro Files (*.mac)" & vbNullChar & "*.mac" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*") As String

    On Error GoTo fin

Dim ret As Long
Dim hInst As Long
Dim Thread As Long
Dim FileDialog As OPENFILENAME
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    If IsMissing(hwnd) Then
        FileDialog.hwndOwner = GetActiveWindow()
    Else
        FileDialog.hwndOwner = hwnd
    End If
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    FileDialog.sfilter = sfilter
    
    If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_SAVE_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    'If centerForm = True Then
    '    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    'Else
    '    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    'End If
    ret = GetSaveFileName(FileDialog)
    If ret Then
       ShowSave = TrimNulls(FileDialog.sFile)
    Else
        ShowSave = vbNullChar
    End If



Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function


Public Function ShowColor(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedColor

    On Error GoTo fin

Dim customcolors() As Byte  ' dynamic (resizable) array
Dim i As Integer
Dim ret As Long
Dim hInst As Long
Dim Thread As Long

    ParenthWnd = hwnd
    If ColorDialog.lpCustColors = "" Then
        ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
    
        For i = LBound(customcolors) To UBound(customcolors)
          customcolors(i) = 254 ' sets all custom colors to white
        Next i
        
        ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
    End If
    
    ColorDialog.hwndOwner = hwnd
    ColorDialog.lStructSize = Len(ColorDialog)
    ColorDialog.flags = COLOR_FLAGS
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = ChooseColor(ColorDialog)
    If ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function

Public Function ShowFont(ByVal hwnd As Long, ByVal startingFontName As String, Optional ByVal centerForm As Boolean = True) As SelectedFont

    On Error GoTo fin

Dim ret As Long
Dim lfLogFont As LOGFONT
Dim hInst As Long
Dim Thread As Long
Dim i As Integer
    
    ParenthWnd = hwnd
    FontDialog.nSizeMax = 0
    FontDialog.nSizeMin = 0
    FontDialog.nFontType = Screen.FontCount
    FontDialog.hwndOwner = hwnd
    FontDialog.hDC = 0
    FontDialog.lpfnHook = 0
    FontDialog.lCustData = 0
    FontDialog.lpLogFont = VarPtr(lfLogFont)
    If FontDialog.iPointSize = 0 Then
        FontDialog.iPointSize = 10 * 10
    End If
    FontDialog.lpTemplateName = Space$(2048)
    FontDialog.rgbColors = RGB(0, 255, 255)
    FontDialog.lStructSize = Len(FontDialog)
    
    If FontDialog.flags = 0 Then
        FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
    End If
    
    For i = 0 To Len(startingFontName) - 1
        lfLogFont.lfFaceName(i) = Asc(Mid(startingFontName, i + 1, 1))
    Next
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = ChooseFont(FontDialog)
        
    If ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
        Next
    
        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
        Exit Function
    Else
        ShowFont.bCanceled = True
        Exit Function
    End If

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function
Public Function ShowPrinter(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As Long

    On Error GoTo fin

Dim hInst As Long
Dim Thread As Long
    
    ParenthWnd = hwnd
    PrintDialog.hwndOwner = hwnd
    PrintDialog.lStructSize = Len(PrintDialog)
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ShowPrinter = PrintDlg(PrintDialog)



Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function

Public Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo fin

    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        x = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.left) / 2
        y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.top) / 2
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False



Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function

Public Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo fin

    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        x = (rectForm.left + (rectForm.Right - rectForm.left) / 2) - ((rectMsg.Right - rectMsg.left) / 2)
        y = (rectForm.top + (rectForm.Bottom - rectForm.top) / 2) - ((rectMsg.Bottom - rectMsg.top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False




Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function

Public Function HookProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo fin
Dim Result As Long
Dim sMsg As String
Dim m As MSG
ReDim Preserve EventArr(EventLog)
recOK = 1
Result = 0
If iCode < 0 Then
    Result = CallNextHookEx(hHook, iCode, wParam, lParam)
ElseIf iCode = HC_SYSMODALON Then
    recOK = 0
ElseIf iCode = HC_SYSMODALOFF Then
    recOK = 1
ElseIf ((recOK > 0) And (iCode = HC_ACTION)) Then
    If CheckStopKey = True Then
        StopMacro
        bKeyStop = True
        HookProc = CallNextHookEx(hHook, iCode, wParam, lParam)
        Exit Function
    End If
    'Insert Caption for Display here
    If lDisplayHwnd <> 0 Then
        'Something to insert the display data
        Call GetCursorPos(pt)
        sMsg = "X:" & pt.x & " Y:" & pt.y & " Event:" & EventLog
        Call SendMessage(lDisplayHwnd, WM_SETTEXT, 0, ByVal sMsg)
    End If
    CopyMemoryH2T EventArr(EventLog), lParam, Len(EventArr(EventLog))
    EventLog = EventLog + 1
    Result = CallNextHookEx(hHook, iCode, wParam, lParam)
End If
HookProc = Result
Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Function
Public Function PlaybackProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo fin
Dim Result As Long
Dim sMsg As String
Dim lPause As Long
Dim evtMsg As EVENTMSG
canPlay = True
Result = 0
If iCode < 0 Then
    Result = CallNextHookEx(hPlay, iCode, wParam, lParam)
ElseIf iCode = HC_SYSMODALON Then
    canPlay = False
ElseIf iCode = HC_SYSMODALOFF Then
    canPlay = True
ElseIf ((canPlay = True) And (iCode = HC_GETNEXT)) Then
'This code controls the timing between system messages
    If bDelay Then
        bDelay = False
        'This result should be based on how fast your PC is
        If PlayLog > 0 And PlayLog < EventLog Then
            lPause = (EventArr(PlayLog).time - lStartTime) + 7
        Else
            lPause = 0
        End If
        lStartTime = EventArr(PlayLog).time
        Result = lPause
    End If
    If lDisplayHwnd <> 0 Then
        'Something to insert the display data
        Call GetCursorPos(pt)
        sMsg = "X:" & pt.x & " Y:" & pt.y & " Event:" & PlayLog
        Call SendMessage(lDisplayHwnd, WM_SETTEXT, 0, ByVal sMsg)
    End If
    CopyMemoryT2H lParam, EventArr(PlayLog), Len(EventArr(PlayLog))
   
ElseIf ((canPlay = True) And (iCode = HC_SKIP)) Then
    If PlayLog >= EventLog Or CheckStopKey = True Then
        StopMacro
        bPlaying = False
        PlaybackProc = Result
        Exit Function
    End If
    bDelay = True
    Result = CallNextHookEx(hPlay, iCode, wParam, lParam)
    PlayLog = PlayLog + 1
End If
PlaybackProc = Result
Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function
Public Function GetSysDir() As String

    On Error GoTo fin

Dim SysDir As String
Dim SD As String
SysDir = Space(144)
SD = GetSystemDirectory(SysDir, 144)
GetSysDir = TrimNulls(SysDir)
Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function


Public Function ShellOn(bOn As Boolean)

    On Error GoTo fin

'Set the Systray Data type
With NOTIFY
    .cbSize = Len(NOTIFY)
    .hwnd = frmEditor.hwnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = 512
    .hIcon = frmEditor.Icon 'Icon shown on the systray
    .szTip = frmEditor.Caption + Chr(0) 'Tooltip text
End With
'Only show the Icon in the Systray if
'the Editor is On
If bOn = True Then
    Call Shell_NotifyIcon(NIM_ADD, NOTIFY) 'Adds the Icon
Else
    Call Shell_NotifyIcon(NIM_DELETE, NOTIFY) ' Deletes the Icon
End If

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function

Public Function GetExt() As String

    On Error GoTo fin

GetExt = GetSetting(App.Title, "Extentions", 0, ".mac")

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function
Public Sub StartRecord()

    On Error GoTo fin

UnLoadHotKeys
frmEditor.mnuRecord.Enabled = False
frmEditor.tblButtons.Buttons(3).Enabled = True
frmEditor.mnuStop.Enabled = True
frmEditor.tblButtons.Buttons(4).Enabled = False
frmEditor.mnuStop.Enabled = False
EventLog = 0
hHook = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf HookProc, App.hInstance, 0)
If hHook <> 0 Then bRecording = True

Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Sub



Public Sub StartPlay()
frmEditor.tblButtons.Buttons(2).Enabled = False
frmEditor.mnuPlay.Enabled = False
frmEditor.tblButtons.Buttons(1).Enabled = False
frmEditor.mnuRecord.Enabled = False
frmEditor.tblButtons.Buttons(3).Enabled = True
frmEditor.mnuStop.Enabled = True
PlayLog = 0
lStartTime = EventArr(0).time
hPlay = SetWindowsHookEx(WH_JOURNALPLAYBACK, AddressOf PlaybackProc, App.hInstance, 0)
bKeyStop = False
If hPlay <> 0 Then
    bPlaying = True
Else
    bPlaying = False
End If

End Sub

Public Sub StopMacro()

On Error GoTo fin

Dim retval As Long
If bPlaying Then
    retval = UnhookWindowsHookEx(hPlay)
    DoEvents
    If retval Then
        bPalying = False
    End If
ElseIf bRecording Then
    retval = UnhookWindowsHookEx(hHook)
    DoEvents
    If retval Then
       bRecording = False
       Call SaveMacro
    End If
End If
frmEditor.tblButtons.Buttons(1).Enabled = True
frmEditor.mnuRecord.Enabled = True
frmEditor.tblButtons.Buttons(2).Enabled = True
frmEditor.mnuPlay.Enabled = True
frmEditor.tblButtons.Buttons(3).Enabled = False
frmEditor.mnuStop.Enabled = False
frmEditor.tblButtons.Buttons(4).Enabled = True
frmEditor.mnuOpen.Enabled = True
frmEditor.Caption = sFileName
Call OnTop(frmEditor)
Call frmPosition(frmEditor, 4)
PlayLog = 0
EventLog = 0
RegHotKeys


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Public Function CheckStopKey() As Boolean

    On Error GoTo fin

If GetAsyncKeyState(VK_CONTROL) And GetAsyncKeyState(VK_MENU) And GetAsyncKeyState(Asc(sStopKey)) Then
    CheckStopKey = True
    bKeyStop = True
    Call keybd_event(VK_CONTROL, &H45, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_MENU, &H45, KEYEVENTF_KEYUP, 0)
Else
    CheckStopKey = False
    bKeyStop = False
End If

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function
Public Function KeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo fin

If nCode >= 0 And GotFocus = lStopKeyDisplay Then
    Call PostMessage(hwndActiveBox, WM_KEYDOWN, wParam, lParam)
    KeyboardProc = 1
    Exit Function
End If


Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


KeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)
End Function

Public Function GetFunctionKey(Optional iKeyCode As Integer, Optional strKeyCode As String) As String

    On Error GoTo fin

If Not IsMissing(iKeyCode) And iKeyCode > 0 Then
    Select Case iKeyCode
        Case vbKeyF1
            GetFunctionKey = "F1"
        Case vbKeyF2
            GetFunctionKey = "F2"
        Case vbKeyF3
            GetFunctionKey = "F3"
        Case vbKeyF4
            GetFunctionKey = "F4"
        Case vbKeyF5
            GetFunctionKey = "F5"
        Case vbKeyF6
            GetFunctionKey = "F6"
        Case vbKeyF7
            GetFunctionKey = "F7"
        Case vbKeyF8
            GetFunctionKey = "F8"
        Case vbKeyF9
            GetFunctionKey = "F9"
        Case vbKeyF10
            GetFunctionKey = "F10"
        Case vbKeyF11
            GetFunctionKey = "F11"
        Case vbKeyF12
            GetFunctionKey = "F12"
        Case vbKeyF13
            GetFunctionKey = "F13"
        Case vbKeyF14
            GetFunctionKey = "F14"
        Case vbKeyF15
            GetFunctionKey = "F15"
        Case vbKeyF16
            GetFunctionKey = "F16"
        Case Else
            GetFunctionKey = Chr(iKeyCode)
    End Select
Else
    Select Case strKeyCode
        Case "F1"
            GetFunctionKey = vbKeyF1
        Case "F2"
            GetFunctionKey = vbKeyF2
        Case "F3"
            GetFunctionKey = vbKeyF3
        Case "F4"
            GetFunctionKey = vbKeyF4
        Case "F5"
            GetFunctionKey = vbKeyF5
        Case "F6"
            GetFunctionKey = vbKeyF6
        Case "F7"
            GetFunctionKey = vbKeyF7
        Case "F8"
            GetFunctionKey = vbKeyF8
        Case "F9"
            GetFunctionKey = vbKeyF9
        Case "F10"
            GetFunctionKey = vbKeyF10
        Case "F11"
            GetFunctionKey = vbKeyF11
        Case "F12"
            GetFunctionKey = vbKeyF12
        Case "F13"
            GetFunctionKey = vbKeyF13
        Case "F14"
            GetFunctionKey = vbKeyF14
        Case "F15"
            GetFunctionKey = vbKeyF15
        Case "F16"
            GetFunctionKey = vbKeyF16
        Case Else
            GetFunctionKey = Asc(strKeyCode)
    End Select
End If
    
Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Function

Public Function GetFormatKeys(iKeyCode As Integer) As String

    On Error GoTo fin

Select Case iKeyCode
    Case vbKeyLButton '1 Left mouse button
    Case vbKeyRButton '2 Right mouse button
    Case vbKeyCancel '3 CANCEL key
        GetFormatKeys = "CANCEL"
    Case vbKeyMButton '4 Middle mouse button
        GetFormatKeys = ""
    Case vbKeyBack '8 BACKSPACE key
        GetFormatKeys = "BACKSPACE"
    Case vbKeyTab '9 TAB key
        GetFormatKeys = "TAB"
    Case vbKeyClear '12 CLEAR key
        GetFormatKeys = "CLEAR"
    Case vbKeyReturn '13 ENTER key
        GetFormatKeys = "ENTER"
    Case vbKeyShift '16 SHIFT key
        GetFormatKeys = "SHIFT"
    Case vbKeyPause '19 PAUSE key
        GetFormatKeys = "PAUSE"
    Case vbKeyEscape '27 ESC key
        GetFormatKeys = "ESC"
    Case vbKeySpace '32 SPACEBAR key
        GetFormatKeys = "SPACEBAR"
    Case vbKeyPageUp '33 PAGE UP key
        GetFormatKeys = "PAGE UP"
    Case vbKeyPageDown '34 PAGE DOWN key
        GetFormatKeys = "PAGE DOWN"
    Case vbKeyEnd '35 END key
        GetFormatKeys = "END"
    Case vbKeyHome '36 HOME key
        GetFormatKeys = "HOME"
    Case vbKeyLeft '37 LEFT ARROW key
        GetFormatKeys = "LEFT ARROW"
    Case vbKeyUp '38 UP ARROW key
        GetFormatKeys = "UP ARROW"
    Case vbKeyRight '39 RIGHT ARROW key
        GetFormatKeys = "RIGHT ARROW"
    Case vbKeyDown '40 DOWN ARROW key
        GetFormatKeys = "DOWN ARROW"
    Case vbKeySelect '41 SELECT key
        GetFormatKeys = "SELECT"
    Case vbKeyPrint '42 PRINT SCREEN key
        GetFormatKeys = "PRINT SCREEN"
    Case vbKeyExecute '43 EXECUTE key
        GetFormatKeys = "EXECUTE"
    Case vbKeySnapshot '44 SNAPSHOT key
        GetFormatKeys = "SNAPSHOT"
    Case vbKeyInsert '45 INS key
        GetFormatKeys = "INSERT"
    Case vbKeyDelete '46 DEL key
        GetFormatKeys = "DELETE"
    Case vbKeyHelp '47 HELP key
        GetFormatKeys = "HELP"
End Select

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function

Public Sub WatchForCancel()

    On Error GoTo fin

Dim message As MSG
Do Until bKeyStop = True Or bPlaying = False
   WaitMessage
    'check if it's a HOTKEY-message
    If PeekMessage(message, 0, WM_CANCELJOURNAL, WM_CANCELJOURNAL, PM_REMOVE) Then
        bKeyStop = True
        bPlaying = False
        StopMacro
    End If
    'let the operating system process other events
    DoEvents
Loop



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub
