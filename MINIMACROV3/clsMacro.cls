VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***************************************************
'Name: clsMacro
'Created: Jan 2000 By James Thomas
'Description: This is the primary module for
'   the MiniMacr application. It is used for
'   retreiving and setting values that are
'   essential the proper operation the application.
'****************************************************
'Procedures:
'4)GetSysDir
'5)ShellOn
'****************************************************
'Dependencies:

'****************************************************
'Modified History:

'****************************************************
Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_ANYCOLOR = &H100

Private Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_SHOWHELP = &H4&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_USESTYLE = &H80&
Private Const CF_EFFECTS = &H100&
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_TTONLY = &H40000
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOVERTFONTS = &H1000000

Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400

Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Private Const SHAREVISTRING = "commdlg_ShareViolation"
Private Const FILEOKSTRING = "commdlg_FileNameOK"
Private Const COLOROKSTRING = "commdlg_ColorOK"
Private Const SETRGBSTRING = "commdlg_SetRGBColor"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const FINDMSGSTRING = "commdlg_FindReplace"

Private Const CD_LBSELNOITEMS = -1
Private Const CD_LBSELCHANGE = 0
Private Const CD_LBSELSUB = 1
Private Const CD_LBSELADD = 2

Private Const PD_ALLPAGES = &H0
Private Const PD_SELECTION = &H1
Private Const PD_PAGENUMS = &H2
Private Const PD_NOSELECTION = &H4
Private Const PD_NOPAGENUMS = &H8
Private Const PD_COLLATE = &H10
Private Const PD_PRINTTOFILE = &H20
Private Const PD_PRINTSETUP = &H40
Private Const PD_NOWARNING = &H80
Private Const PD_RETURNDC = &H100
Private Const PD_RETURNIC = &H200
Private Const PD_RETURNDEFAULT = &H400
Private Const PD_SHOWHELP = &H800
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_DISABLEPRINTTOFILE = &H80000
Private Const PD_HIDEPRINTTOFILE = &H100000
Private Const PD_NONETWORKBUTTON = &H200000
'Dialog Constants
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 256
Private Const LF_FACESIZE = 32
Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
Private Const DN_DEFAULTPRN = &H1

'Hook Constants
Private Const WH_KEYBOARD = 2
Private Const WH_JOURNALPLAYBACK = 1
Private Const WH_JOURNALRECORD = 0
Private Const WH_CBT = 5
'Hook Call Constants
Private Const HC_SYSMODALOFF = 5
Private Const HC_SYSMODALON = 4
Private Const HC_SKIP = 2
Private Const HC_GETNEXT = 1
Private Const HC_ACTION = 0

'Systray Constants
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
'Message Queue Constants
Private Const PM_NOREMOVE = &H0
Private Const PM_REMOVE = &H1
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_SETTEXT = &HC

'Operating System Version
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

'Windows Properties
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10

Private Const GWL_HINSTANCE = (-6)
Private Const HCBT_ACTIVATE = 5

'Virtual Key Contants
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12

Private Type PRINTDLGS
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

Private Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
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

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Private Type CHOOSECOLORS
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

Private Type LOGFONT
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

Private Type CHOOSEFONTS
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
Private Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Private Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Private Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Private Type SelectedFont
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
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type EVENTMSG
        message As Long
        paramL As Long
        paramH As Long
        time As Long
        hwnd As Long
End Type

Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    ' Maintenance string for PSS usage.
End Type

'Kernal Declares
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Sub FreeLibraryAndExitThread Lib "kernel32" (ByVal hLibModule As Long, ByVal dwExitCode As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemoryT2H Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Long, Source As EVENTMSG, ByVal Length As Long)
Private Declare Sub CopyMemoryH2T Lib "kernel32" Alias "RtlMoveMemory" (Dest As EVENTMSG, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Shell Declares
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'User32 Declares
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

'Common Dialog Declares
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long

'Private Properties
Private pos As POINTAPI
Private o As OSVERSIONINFO
Private NOTIFY As NOTIFYICONDATA
Private lX As Long
Private lY As Long
Private strCmdLine As String
Private EventArr() As EVENTMSG
Private EventLog As Long
Private PlayLog As Long
Private hHook As Long
Private hPlay As Long
Private recOK As Long
Private canPlay As Long
Private bDelay As Boolean
Private sFileName As String
Private iFileNumber As String
Private lDisplayHwnd As Long
Private pt As POINTAPI
Private bRecording As Boolean
Private bPlaying As Boolean
Private sExt As String 'File extension
Private sStopKey As String 'Stop Key for key combination
Private lStartTime As Long
'Dialog Properties
Private ColorDialog As CHOOSECOLORS
Private FontDialog As CHOOSEFONTS
Private PrintDialog As PRINTDLGS
Private ParenthWnd As Long

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent MacroStopped[(arg1, arg2, ... , argn)]
Public Event MacroStopped()
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent StartedRecording[(arg1, arg2, ... , argn)]
Public Event StartedRecording()
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent StartedPlaying[(arg1, arg2, ... , argn)]
Public Event StartedPlaying()

Public Property Let MacroFileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MacroFileName = 5
    sFileName = vData
End Property


Public Property Get MacroFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MacroFileName
    MacroFileName = sFileName
End Property




Public Property Let DisplayHwnd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DisplayHwnd = 5
    lDisplayHwnd = vData
End Property


Public Property Get DisplayHwnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DisplayHwnd
    DisplayHwnd = lDisplayHwnd
End Property




Public Property Let FileExt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileExt = 5
    sExt = vData
End Property


Public Property Get FileExt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileExt
    FileExt = sExt
End Property

'local variable(s) to hold property value(s)
Public Property Let StopKey(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StopKey = 5
    sStopKey = vData
End Property

Public Property Get StopKey() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StopKey
    StopKey = sStopKey
End Property

Public Sub OpenMacro()
Dim FileLength As Long
iFileNumber = FreeFile()
sFileName = ShowOpen()
If sFileName <> vbNullChar Then
    ReDim EventArr(0)
    Open sFileName For Random Access Read As iFileNumber Len = Len(EventArr(0))
        FileLength = 1
        Do Until EOF(iFileNumber)
            ReDim Preserve EventArr(FileLength)
            Get #iFileNumber, FileLength, EventArr(FileLength - 1)
            FileLength = FileLength + 1
        Loop
        EventLog = FileLength - 1
    Close iFileNumber
End If
End Sub

Public Sub SaveMacro()
Dim iIndex As Long
iFileNumber = FreeFile()
sFileName = ShowSave()
If sFileName <> vbNullChar Then
    If Mid(Right(sFileName, 4), 1, 1) <> "." Then
        sFileName = sFileName & sExt
    End If
    Open sFileName For Random Access Write As iFileNumber Len = Len(EventArr(EventLog))
        For iIndex = 1 To EventLog
            Put iFileNumber, iIndex, EventArr(iIndex - 1)
        Next iIndex
    Close iFileNumber
End If
End Sub

Public Function ShowOpen(Optional ByVal hwnd As Long, Optional ByVal centerForm As Boolean, Optional sFilter As String = "Macro Files (*.mac)" & vbNullChar & "*.mac" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*") As String
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
    FileDialog.sFilter = sFilter
    FileDialog.flags = OFS_FILE_OPEN_FLAGS
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    ret = GetOpenFileName(FileDialog)
    If ret Then
       ShowOpen = TrimNulls(FileDialog.sFile)
    Else
        ShowOpen = vbNullChar
    End If
End Function

Public Function ShowSave(Optional ByVal hwnd As Long, Optional ByVal centerForm As Boolean, Optional sFilter As String = "Macro Files (*.mac)" & vbNullChar & "*.mac" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*") As String
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
    FileDialog.sFilter = sFilter
    
    If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_SAVE_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    ret = GetSaveFileName(FileDialog)
    If ret Then
       ShowSave = TrimNulls(FileDialog.sFile)
    Else
        ShowSave = vbNullChar
    End If
End Function

Public Function ShowColor(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As Long
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
End Function

Public Function ShowFont(ByVal hwnd As Long, ByVal startingFontName As String, Optional ByVal centerForm As Boolean = True) As Long
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
End Function
Public Function ShowPrinter(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As Long
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
End Function
Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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
End Function

Public Function HookProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Result As Long
Dim sMsg As String
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
    CopyMemoryH2T EventArr(EventLog), lParam, Len(EventArr(EventLog))
    If CheckStopKey = True Then
        StopMacro
    End If
    EventLog = EventLog + 1
    'Insert Caption for Display here
    If lDisplayHwnd <> 0 Then
        'Something to insert the display data
        Call GetCursorPos(pt)
        sMsg = "X:" & pt.x & " Y:" & pt.y & " Event:" & EventLog
        Call SendMessage(lDisplayHwnd, WM_SETTEXT, 0, ByVal sMsg)
    End If
End If
HookProc = Result
End Function

Public Function PlaybackProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Result As Long
    Dim sMsg As String
    Dim lPause As Long
    canPlay = 1
    Result = 0
    If iCode < 0 Then
        Result = CallNextHookEx(hPlay, iCode, wParam, lParam)
    ElseIf iCode = HC_SYSMODALON Then
        canPlay = 0
    ElseIf iCode = HC_SYSMODALOFF Then
        canPlay = 1
    ElseIf ((canPlay = 1) And (iCode = HC_GETNEXT)) Then
        If bDelay Then
            bDelay = False
            Result = 50
        End If
        CopyMemoryT2H lParam, EventArr(PlayLog), Len(EventArr(PlayLog))
        If lDisplayHwnd <> 0 Then
            'Something to insert the display data
            Call GetCursorPos(pt)
            sMsg = "X:" & pt.x & " Y:" & pt.y & " Event:" & PlayLog
            Call SendMessage(lDisplayHwnd, WM_SETTEXT, 0, ByVal sMsg)
        End If
    ElseIf ((canPlay = 1) And (iCode = HC_SKIP)) Then
        bDelay = True
        lPause = EventArr(PlayLog).time - lStartTime
        If EventArr(PlayLog).time > 0 Then
            Sleep (lPause)
        End If
        lStartTime = EventArr(PlayLog).time
        PlayLog = PlayLog + 1
    End If
    If PlayLog >= EventLog Then
        UnhookWindowsHookEx hPlay
        bPlaying = False
    End If
    PlaybackProc = Result
End Function

Public Function GetSysDir() As String
Dim SysDir As String
Dim SD As String
SysDir = Space(144)
SD = GetSystemDirectory(SysDir, 144)
GetSysDir = TrimNulls(SysDir)
End Function

Public Function ShellOn(bOn As Boolean)
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
End Function

Public Function GetExt() As String
GetExt = GetSetting(App.Title, "Extentions", 0, ".mac")
End Function
Public Sub StartRecord()
EventLog = 0
hHook = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf HookProc, App.hInstance, 0)
If hHook <> 0 Then
    bRecording = True
    RaiseEvent StartedRecording
End If
End Sub


Public Sub Startplay()
PlayLog = 0
lStartTime = EventArr(0).time
hPlay = SetWindowsHookEx(WH_JOURNALPLAYBACK, AddressOf PlaybackProc, App.hInstance, 0)
If hPlay <> 0 Then bPlaying = True
End Sub

Public Sub StopMacro()
If bPalying Then
    UnhookWindowsHookEx hPlay
    bPalying = False
ElseIf bRecording Then
    UnhookWindowsHookEx hHook
    bRecording = False
    SaveMacro
End If
RaiseEvent MacroStopped
End Sub

Public Function CheckStopKey() As Boolean
If GetAsyncKeyState(VK_CONTROL) And GetAsyncKeyState(VK_MENU) And GetAsyncKeyState(Asc(sStopKey)) Then
    CheckStopKey = True
Else
    CheckStopKey = False
End If

End Function
