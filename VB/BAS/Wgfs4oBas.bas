Attribute VB_Name = "Wgfs4oBas"
'Hey everybody. This bas file I made in hope of seeing beter progs
'For AOL4.o. I also hope that all you lazey fools start writing your own code
'And not leavin it up to the .bas makers.
'This bas has a TON of faders on it so use them
'I belive its also the best bas so far for AOL4.o, So have fun
'Also it was used to make Blur ToolZ v3.0 Fo AOL4.o(which isn't out yet)
'L8ter, WGF
'������������������������������������������������ '
'Copyright 1998 Any modifications with out the consult of WGF
'Are not premited.

' Windows 95 API Public Function Declarations '
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function ExitWindows Lib "kernel32" (ByVal dwReturnCode&, ByVal wReserved%) As Integer
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Declare Function LZCopy Lib "lz32.dll" (ByVal hfSource As Long, ByVal hfDest As Long) As Long
Declare Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, lpOf As OFSTRUCT, ByVal style As Long) As Long
Declare Function LZSeek Lib "lz32.dll" (ByVal hfFile As Long, ByVal lOffset As Long, ByVal nOrigin As Long) As Long
Declare Function LZRead Lib "lz32.dll" (ByVal hfFile As Long, ByVal lpvBuf As String, ByVal cbread As Long) As Long
Declare Function LZClose Lib "lz32.dll" (ByVal hfFile As Long)

' Windows 95 API Public Functions Substitutes '
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Windows 95 API Private Function & Sub Declarations'
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function ShellUse Lib "shell32.dll Alias (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long" ()
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Private Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function ReleaseCapture Lib "user32" ()
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)

' Public Windows 95 API Constant Functions '
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CLEAR = &H303
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203

Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_GETITEMDATA = &H199
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Global Const CB_ADDSTRING = (WM_USER + 3)
Global Const CB_DELETESTRING = (WM_USER + 4)
Global Const CB_DIR = (WM_USER + 5)
Global Const CB_ERR = (-1)
Global Const CB_ERRSPACE = (-2)
Global Const CB_FINDSTRING = (WM_USER + 12)
Global Const CB_FINDSTRINGEXACT = (WM_USER + 24)
Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_GETCURSEL = (WM_USER + 7)
Global Const CB_GETDROPPEDCONTROLRECT = (WM_USER + 18)
Global Const CB_GETDROPPEDSTATE = (WM_USER + 23)
Global Const CB_GETEDITSEL = (WM_USER + 0)
Global Const CB_GETEXTENDEDUI = (WM_USER + 22)
Global Const CB_GETITEMDATA = (WM_USER + 16)
Global Const CB_GETITEMHEIGHT = (WM_USER + 20)
Global Const CB_GETLBTEXT = (WM_USER + 8)
Global Const CB_GETLBTEXTLEN = (WM_USER + 9)
Global Const CB_INSERTSTRING = (WM_USER + 10)
Global Const CB_LIMITTEXT = (WM_USER + 1)
Global Const CB_MSGMAX = (WM_USER + 19)
Global Const CB_OKAY = 0
Global Const CB_RESETCONTENT = (WM_USER + 11)
Global Const CB_SELECTSTRING = (WM_USER + 13)
Global Const CB_SETCURSEL = (WM_USER + 14)
Global Const CB_SETEDITSEL = (WM_USER + 2)
Global Const CB_SETEXTENDEDUI = (WM_USER + 21)
Global Const CB_SETITEMDATA = (WM_USER + 17)
Global Const CB_SETITEMHEIGHT = (WM_USER + 19)
Global Const CB_SHOWDROPDOWN = (WM_USER + 15)

Public Const VK_DELETE = &H2E
Public Const VK_RIGHT = &H27
Public Const VK_HOME = &H24
Public Const VK_CONTROL = &H11
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_CREATE = &H3

Public Const MB_OK = 0
Public Const MB_OKCANCEL = 1
Public Const MB_ABORTRETRYIGNORE = 2
Public Const MB_YESNOCANCEL = 3
Public Const MB_YESNO = 4
Public Const MB_RETRYCANCEL = 5

Public Const MB_ICONSTOP = 16
Public Const MB_ICONQUESTION = 32
Public Const MB_ICONEXCLAMATION = 48
Public Const MB_ICONINFORMATION = 64

Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Const EW_RESTARTWINDOWS = &H42
Public Const EW_REBOOTSYSTEM = &H43

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_APPEND = &H100&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_UNCHECKED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_INVALID_FUNCTION = 1&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_OUTOFMEMORY = 14&
Public Const ERROR_BAD_NETPATH = 53&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_INVALID_PASSWORDNAME = 1216&

Public Const GWL_STYLE = (-16)

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)

Public Const PROCESS_VM_READ = &H10

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const SPI_SCREENSAVERRUNNING = 97

Private Const OF_READ = &H0
Private Const OF_WRITE = &H1
Private Const OF_READWRITE = &H2
Private Const OF_SHARE_COMBAT = &H0
Private Const OF_SHARE_EXCLUSIVE = &H10
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const OF_SHARE_DENY_READ = &H30
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OF_PARSE = &H100
Private Const OF_DELETE = &H200
Private Const OF_VERIFY = &H400
Private Const OF_CANCEL = &H800
Private Const OF_CREATE = &H1000
Private Const OF_PROMPT = &H2000
Private Const OF_EXIST = &H4000
Private Const OF_REOPEN = &H8000

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYVTHUMB = 9
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14
Private Const SM_CYMENU = 15
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYKANJIWINDOW = 18
Private Const SM_MOUSEPRESENT = 19
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21
Private Const SM_DEBUG = 22
Private Const SM_SWAPBUTTON = 23
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXMINTRACK = 34
Private Const SM_CYMINTRACK = 35
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CXICONSPACING = 38
Private Const SM_CYICONSPACING = 39
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_PENWINDOWS = 41
Private Const SM_DBCSENABLED = 42
Private Const SM_CMOUSEBUTTONS = 43
Private Const SM_CMENTRICS = 44

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2


Private Type OFSTRUCT
        cBytes As Byte
        fFixedByte As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(128) As Byte
End Type

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
    
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type

Private inpOFS As OFSTRUCT
Private outOFS As OFSTRUCT
Private mybuf As String
Private size As Long
Private expandedName As String
Private compressedName As String

Global giBeepBox As Integer
Global r&
Global entry$
Global iniPath$

Sub AOLHideWelcome()
'Hides the welcome screen that I know time after time you sat thier and
'and pushed the A button trying to close it
X = FindChildByTitle(AOLMDI(), "Welcome, " & AOLUserSN & "!")
Call ShowWindow(X, SW_HIDE)
End Sub


Sub AOLShowWelcome()
'Shows the welcome menu
X = FindChildByTitle(AOLMDI(), "Welcome, " & AOLUserSN & "!")
Call ShowWindow(X, SW_SHOW)
End Sub

Sub AOLGetMemberProfile(sn)
'Will get a  member profile
AppActivate "America  Online"
SendKeys "^g"
pause 0.9
prof% = FindChildByTitle(AOLMDI(), "Get a Member's Profile")
pause 0.7
Edit% = FindChildByClass(prof%, "_AOL_Edit")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, sn)
Call Enter(Edit%)
End Sub
Function Chat_RoomName()
Call GetCaption(AOLFindChatRoom)
End Function
Function AOLActivate()
cap% = GetCaption(AOLWindow())
AppActivate cap%
End Function
Function AOLVersion()
hMenu% = GetMenu(AOLWindow())
submenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")
FindString% = GetMenuString(submenu%, subitem%, MenuString$, 100, 1)
If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 4
End If
End Function




Function FindParentByTitle(parenthand)
'Finds parent by title
End Function

Sub Mail_Forward(sn, message)
'Forwards a mail
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDI Client")
FwdWin% = FindChildByTitle(AOL%, "Fwd: ")

Edit% = FindChildByClass(FwdWin%, "_AOL_Edit")
Rich% = FindChildByClass(FwdWin%, "RICHCNTL")
icon% = FindChildByClass(FwdWin%, "_AOL_ICON")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, sn)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, message)
Call AOLClickIcon(icon%)
End Sub

Sub NotOnTop(frm As Form)
'Make the form of your choice not on top
notop% = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub


Public Sub AOLButton(but%)
'This clicks a aol button (that are so rare you probally won't use this)
ClickIcon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
ClickIcon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub


Public Function AOLRoomFull()
'Closes the message box that AOL say "The room you requested is full"
'Nice feture to add to your room buster
Msg% = FindWindow("#32770", "America Online")
WinTxt% = GetWinText(Msg%)
If InStr(1, WinTxt%, "is full") Then
Call Enter(full%)
Else
End If
End Function
Sub AOLFileSearch(File)
'This goes to aol's file serch and searches for a phile
'This is pointless but I made it at 1:30 in the morning so I don't care
Call AOL40_Keyword("File Search")
First% = FindChildByTitle(AOLMDI(), "Filesearch")
icon% = FindChildByClass(First%, "_AOL_Icon")
icon% = GetWindow(icon%, 2)
Call AOLClickIcon(icon%)

Secnd% = FindChildByTitle(AOLMDI(), "Software Search")
Edit% = FindChildByClass(Secnd%, "_AOL_Edit")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, File)
Call SendMessageByNum(Rich%, WM_CHAR, 0, 13)
End Sub


Function LastChatLine()
'Gets last chat lines text
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Function LastChatLineWithSN()
'Gets last chat lines text with the SN
chattext$ = GetchatText

For FindChar = 1 To Len(chattext$)

thechar$ = Mid(chattext$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(chattext$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function
Sub AOLlocateMember(sn)
'This will locate where a member is
Call AOL40_Keyword("aol://3548:" & sn)
End Sub

Sub Chat_Attending()
'This cheks if the person usin your prog is in a chat
'I don't think its very usefull
Room% = AOL40_FindChatRoom()
If Room% = 0 Then
MsgBox "You must be in a chat room to use this feature", 64, "Must Be In Chat!"
Else
End If
End Sub

Sub Chat_ChangeCaption(newcap)
'Changes the Chat Rooms caption
Room% = AOL40_FindChatRoom()
Call Window_ChangeCaption(Room%, newcap)
End Sub

Sub Chat_EnterPR(PR)
'Manualy goes to a private room
'Good for a room buster unless you want to do it
'your own way.
Call AOL40_Keyword("aol://2719:2-2-" + PR)
End Sub

Sub Chat_Ignore(sn%)
Room% = AOL40_FindChatRoom
List% = FindChildByClass(Room%, "_AOL_Listbox")
End Sub

Function Chat_Text()
'This stores all text from da chat
Rich% = FindChildByClass(AOL40_FindChatRoom(), "RICHCNTL")
txt% = GetWinText(Rich%)
AOL40_GetChatText = txt%
End Function

Function Mail_Count()
'This counts the # of mail(s) the user has
mail% = FindChildByClass(AOLMDI(), "AOL Child")
tree% = FindChildByClass(mail%, "_AOL_Tree")
Mail_Count = SendMessage(tree%, LB_GETCOUNT, 0, 0)
End Function



Sub AOLRunMenuByString(stringer As String)
'This searches all the popup menus for the item you
'want to find on AOL
Call RunMenuByString(AOLWindow(), stringer)
End Sub





Sub IMBuddy(Recipiant, message)
'Does the same as SendIM4 exept sends the IM thourgh the buddy list
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub SendIM4(sn, Text)
'Sends an aol4.0 IM
AOLActivate
SendKeys "^i"
Do: DoEvents
IM% = FindChildByTitle(AOLMDI(), "Send Instant Message")
Edit% = FindChildByClass(IM%, "_AOL_Edit")
Rich% = FindChildByClass(IM%, "RICHCNTL")
icon% = FindChildByClass(IM%, "_AOL_Icon")
Loop Until Edit% <> 0 And Rich% <> 0 And icon% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, sn)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, Text)
For X = 1 To 9
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next X
Call pause(0.01)
AOLClickIcon (icon%)
End Sub

Sub SetCheckBoxToFalse(win%)
'This will set any checkbox's value to equal false
Check% = SendMessageByNum(win%, BM_SETCHECK, False, 0&)
End Sub

Sub SetCheckBoxToTrue(win%)
'This will set any checkbox's value to equal true
Check% = SendMessageByNum(win%, BM_SETCHECK, True, 0&)
End Sub
Sub WavyFader(txt)
'Sub Created By Bman
'Messed around with By WGF
a = Len(txt)
For w = 1 To a Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    T$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    B$ = Mid$(txt, w + 7, 1)
    C$ = Mid$(txt, w + 8, 1)
    D$ = Mid$(txt, w + 9, 1)
    H$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    Q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    z$ = Mid$(txt, w + 17, 1)
    P$ = P$ & "<b><FONT COLOR=#000019><sup>" & ab$ & "<FONT COLOR=#000026></sup>" & u$ & "<FONT COLOR=#00003F><sub>" & S$ & "<FONT COLOR=#000058></sub>" & T$ & "<FONT COLOR=#000072><sup>" & Y$ & "<FONT COLOR=#00008B></sup>" & L$ & "<FONT COLOR=#0000A5><sub>" & F$ & "<FONT COLOR=#0000BE></sub>" & B$ & "<FONT COLOR=#0000D7><sup>" & C$ & "<FONT COLOR=#0000F1></sup>" & D$ & "<FONT COLOR=#0000D7><sub>" & H$ & "<FONT COLOR=#0000BE></sub>" & j$ & "<FONT COLOR=#0000A5><sup>" & k$ & "<FONT COLOR=#00008B></sup>" & m$ & "<FONT COLOR=#000072><sub>" & n$ & "<FONT COLOR=#000058></sub>" & Q$ & "<FONT COLOR=#00003F><sup>" & V$ & "<FONT COLOR=#000026></sup>" & z$
Next w
GeM3 = P$
SendChat4 (P$)
End Sub


Sub MouseScroll()
'Scrolls the mouse
X = Val(X) - 1
Y = Val(Y) - 1
whatever = SetCursorPos(X, Y)
End Sub

Function AOLFindChatRoom()
'This sets focus on a chat room
childfocus% = GetWindow(AOLMDI(), 5)
While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")
If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindChatRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, GW_HWNDNEXT)
Wend
End Function

Sub AOL40_SignOff()
'Sign Off very quick
Call AOLRunMenuByString("&Sign Off")
End Sub
Sub AOLSNReset(sn$, aoldir$, replace$)
'Resets your SN to a new User
l0036 = Len(sn$)
Select Case l0036
Case 3
i = sn$ + "       "
Case 4
i = sn$ + "      "
Case 5
i = sn$ + "     "
Case 6
i = sn$ + "    "
Case 7
i = sn$ + "   "
Case 8
i = sn$ + "  "
Case 9
i = sn$ + " "
Case 10
i = sn$
End Select
l0036 = Len(replace$)
Select Case l0036
Case 3
replace$ = replace$ + "       "
Case 4
replace$ = replace$ + "      "
Case 5
replace$ = replace$ + "     "
Case 6
replace$ = replace$ + "    "
Case 7
replace$ = replace$ + "   "
Case 8
replace$ = replace$ + "  "
Case 9
replace$ = replace$ + " "
Case 10
replace$ = replace$
End Select
X = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, X, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = replace$
ReplaceX$ = replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = replace$
Put #2, X + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
X = X + 32000
LF2 = LOF(2)
Close #2
If X > LF2 Then GoTo 301
Loop
301:
End Sub


Public Function GetListIndex(LB As ListBox, txt As String) As Integer
Dim Index As Integer
With LB
For Index = 0 To .ListCount - 1
If .List(iIndex) = txt Then
GetListIndex = Index
Exit Function
End If
Next Index
End With
GetListIndex = -2
End Function
Function GetWinText(hwnd As Integer) As String
'This gets the text from any window
TextLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(TextLength)
GetTheText = SendMessageByString(hwnd, WM_GETTEXT, LengthOfText + 1, Buffer$)
GetWinText = Buffer$
End Function
Sub KillWait()
'Killz the hour galsss in aol.ONLY in aol
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
AOLClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Sub KillUser()
'This is pretty cool. You could put it in a punter and then the   user thinks
'there punting some one else  but insstead it punts them with the 1 IM error.
'It keeps looping, so if you don't want to anoy them to bad add a stop button
Do:
DoEvents
Call SendIM4(UserSN, "<font= 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")
Loop
End Sub
Sub KillModal()
'Kills modal
modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(modal%, WM_CLOSE, 0, 0)
End Sub
Sub Form_Move(frm As Form)
'This will allow you to move the form to a different
'part of your screen
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(frm.hwnd, &HA1, 2, 0)
End Sub
Sub Mail_Open()
'This will open your Online Mailbox
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Call AOLClickIcon(icon%)
End Sub

Sub AOLClose()
'This closes AOL very quick
Call Window_Close(AOLWindow())
End Sub
Sub AOLChangeCaption(newcaption)
'It  change's AOL's caption
Call AOLSetText(AOLWindow(), newcaption)
End Sub
Sub AOLSetText(win, txt)
'This allows you to change the text from a window.

Text% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Function Window_ChangeCaption(win, txt)
'This will change the caption of any window that you
'tell it to as long as it is a valid window
Text% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Function

Function AOLIsOnline() As Integer
'This returns if the user is online
welcome% = FindChildByTitle(AOLMDI(), "Welcome, ")
If welcome% = 0 Then
MsgBox "You Must Sign On Before Using This Feature.", 64, "Must Be Online"
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Function Chat_RoomCount()
'Tells you how many people are in a chat room
'Hince the name "Chat room count
Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Sub SendchatBold(BoldChat)
'This makes the chat text bold.
SendChat4 ("<b>" & BoldChat & "</b>")
End Sub
Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
SendChat4 ("<i>" & ItalicChat & "</i>")
End Sub
Sub SendChatStrike(StrikeOutChat)
'This strikes out text in a chat
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub
Sub WierdAttention(Text)
'This is going to look really meesed up
SendChat ("<b>�</b><i> �</i><u> �</u><s> �</s> " & Text & " <s>�</s><u> �</u><i> �</i><b> �</b>")
BoldSendChat (Text)
ItalicSendChat (Text)
UnderLineSendChat (Text)
StrikeOutSendChat (Text)
SendChat ("<b>�</b><i> �</i><u> �</u><s> �</s> " & Text & " <s>�</s><u> �</u><i> �</i><b> �</b>")
End Sub
Sub AddRoomCombo(ListBox As ListBox, ComboBox As ComboBox)
'Adds A room to a comno box
Call AOL40_AddRoomList(ListBox)
For Q = 0 To ListBox.ListCount
ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
Sub AntiIdle()
'Clicks ok on the aol messgae box that says You have been on for an idle while
modal% = FindWindow("_AOL_Modal", vbNullString)
icon% = FindChildByClass(modal%, "_AOL_Icon")
AOLClickIcon (icon%)
End Sub

Sub KillGlyph()
'This will close that little  AOL spinning
'thingy on the top corner of AOL 4.0
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Glyph% = FindChildByClass(Toolbar%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub


Sub ADD_AOL_LB(itm As String, lst As ListBox)
'Add a list of names to a VB ListBox
If lst.ListCount = 0 Then
lst.AddItem itm
Exit Sub
End If
Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub
Sub AddRoomList(lst As ListBox)
'This adds all  the peeps in a chat room
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
names$ = String$(256, " ")
ret = AOLGetList(Index, names$)
names$ = Left$(Trim$(names$), Len(Trim(names$)))
ADD_AOL_LB names$, lst
Next Index
endaddroom:
lst.RemoveItem lst.ListCount - 1
i = GetListIndex(lst, AOLUserSN())
If i <> -2 Then lst.RemoveItem i
End Sub
Public Function AOLGetList(Index As Long, Buffer As String)
'This gets the list you request
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = AOL40_FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If
Buffer$ = Person$
End Function
Sub WAVStop()
'This will stop a WAV that is playing
Call WAVPlay(" ")
End Sub

Sub WAVLoop(File)
'This will play the WAV you want over and over
SoundName$ = File
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Function AOLMessageFromIM()
'This gets the full text from a IM(good to use for a IM recorder)
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo txt
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo txt
Exit Function
txt:
IMtxt% = FindChildByClass(IM%, "RICHCNTL")
txt% = AOLGetText(IMtxt%)
snlen = Len(AOLSNFromIM()) + 3
Blah% = Mid(txt%, InStr(txt%, AOLSNFromIM()) + snlen)
AOLMessageFromIM = Left(Blah%, Len(Blah%) - 1)
End Function

Sub WAVPlay(File)
'This will play a WAV file
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Function SendCharNum(win, chars)
'This sends any character number that you enter to
'your desired  window
e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Function

Sub AOLRespondIM(message)
'This finds the Instant Message window and responds
'it with the message you want, than closes
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call AOLSetText(e2, message)
AOLClickIcon (e)
pause 0.8
e = FindChildByClass(IM%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
AOLClickIcon (e)
End Sub

Function Enter(win)
'This presses enter
Call SendCharNum(win, 13)
End Function
Sub EliteTalker(word$)
'Its an elite talker that when used sends text auotmaticly
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "�"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "�"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "s" Then Leet$ = "�"
    If letter$ = "t" Then Leet$ = "�"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "�"
    If letter$ = "0" Then Leet$ = "�"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "B" Then Leet$ = "�"
    If letter$ = "C" Then Leet$ = "�"
    If letter$ = "D" Then Leet$ = "�"
    If letter$ = "E" Then Leet$ = "�"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "�"
    If X = 2 Then Leet$ = "�"
    If X = 3 Then Leet$ = "�"
    End If
    If letter$ = "N" Then Leet$ = "�"
    If letter$ = "O" Then Leet$ = "�"
    If letter$ = "S" Then Leet$ = "�"
    If letter$ = "U" Then Leet$ = "�"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "�"
    If letter$ = "`" Then Leet$ = "�"
    If letter$ = "!" Then Leet$ = "�"
    If letter$ = "?" Then Leet$ = "�"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
SendChat4 (Made$)
End Sub
Sub File_Delete(File)
'This will delete a file from the users HardDrive
Kill (File)
End Sub
Sub File_Open(File)
'This will open a file. The whole dir and file name needed
Shell (File)
End Sub
Sub File_ReName(sFromLoc As String, sToLoc As String)
'This immediately renames a file for you
Name sOldLoc As sNewLoc
End Sub
Sub Directory_Create(dir)
'This adds a directory to your system
'Example of what it should look like:
'Call Directory_Create("C:\WgfRules\ReportOnhowgoodhisbasis")
MkDir dir
End Sub
Sub Directory_Delete(dir)
'This deletes a directory from your HD
RmDir (dir)
End Sub

Function GetCaption(hwnd)
'Gets a caption
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function

Function AOLUserSN()
'The person usin your prog(good to use if you want to say loaded by)
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLUserSN = User
End Function
Function AOLSNFromIM()
'This returns the SN from an IM
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IM%)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
AOLSNFromIM = Naw$
End Function
Public Sub Disable_Ctrl_Alt_Del()
'Disables the Crtl+Alt+Del
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub Enable_Ctrl_Alt_Del()
'Enables the Crtl+Alt+Del
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub SendMail4(sn, subject, message)
'This will send mail from AOL4.0
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
icon% = GetWindow(icon%, GW_HWNDNEXT)
Call AOLClickIcon(icon%)
Do: DoEvents
mail% = FindChildByTitle(AOLMDI(), "Write Mail")
Edit% = FindChildByClass(mail%, "_AOL_Edit")
Rich% = FindChildByClass(mail%, "RICHCNTL")
icon% = FindChildByClass(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And Edit% <> 0 And Rich% <> 0 And icon% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, sn)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Call SendMessageByString(Edit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, message)
For GetIcon = 1 To 18
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next GetIcon
Call AOLClickIcon(icon%)
End Sub

Function GetClass(child)
'gets class
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function
Sub AOLAnti45MinTimer()
'Killz the 45 message box
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AOL40_Load()
'This loads  AOL4.0
X% = Shell("C:\America Online 4.0\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\America Online 4.0A\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\America Online 4.0B", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Function AOLMDI()
'This function sets focus on AOL's parent window
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Sub SendChat4(txt)
'This will send text to the chat to AOL 4.0
Rich% = FindChildByClass(AOL40_FindChatRoom, "RICHCNTL")
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Call SetFocusAPI(Rich%)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, txt)
DoEvents
Call Enter(Rich%)
End Sub


Sub AOL40_Keyword(KeyWord As String)
'Sends a AOL4.0 Keyword
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Tool2%, "_AOL_Icon")
For GetIcon = 1 To 20
icon% = GetWindow(icon%, 2)
Next GetIcon
Call pause(0.05)
Call AOLClickIcon(icon%)
Do: DoEvents
MDI% = FindChildByClass(AOLWindow(), "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
Edit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
Icon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And Edit% <> 0 And Icon2% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, KeyWord)
Call pause(0.05)
Call AOLClickIcon(Icon2%)
Call AOLClickIcon(Icon2%)
End Sub
Sub Mail_WaitForLoad()
Do
Box% = FindChildByTitle(AOLMDI(), AOLUserSN & "'s Online Mailbox")
Loop Until Box% <> 0
List = FindChildByClass(Box%, "_AOL_Tree")
Do
DoEvents
M1% = SendMessage(List%, LB_GETCOUNT, 0, 0&)
M2% = SendMessage(List%, LB_GETCOUNT, 0, 0&)
M3% = SendMessage(List%, LB_GETCOUNT, 0, 0&)
Loop Until M1% = M2% And M2% = M3%
M1% = SendMessage(List%, LB_GETCOUNT, 0, 0&)
End Sub


Function FreeProcess()
'This feature will allow you to be in your program while its playing and
'not freeze or have too many errors
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function Mail_FindOpen() As Long
'Finds the mail and opens it
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
child% = FindChildByClass(MDI%, "AOL Child")
Rich% = FindChildByClass(child%, "RICHCNTL")
icond% = FindChildByClass(child%, "_AOL_Icon")
Stat% = FindChildByClass(child%, "_AOL_Static")
If InStr(1, AOLGetText(Rich%), "Subj:") <> 0 And icond% <> 0 And Stat% <> 0 Then GoTo Found
For Y = 1 To 10
child% = GetWindow(child%, 2)
Rich% = FindChildByClass(child%, "RICHCNTL")
icond% = FindChildByClass(child%, "_AOL_Icon")
Stat% = FindChildByClass(child%, "_AOL_Static")
If InStr(1, AOLGetText(Rich%), "Subj:") <> 0 And icond% <> 0 And Stat% <> 0 Then GoTo Found
Next Y
AOLFindOpenMail = 0
Exit Function
Found:
Mail_FindOpen = child%
End Function

Function AOLGetText(child)
'This get's the text of a child window
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
AOLGetText = TrimSpace$
End Function

Function FindChatRoom()
'Maybe it finds......A CHAT ROOM!
Room% = FindChildByClass(AOLMDI(), "AOL Child")
roomlst% = FindChildByClass(Room%, "_AOL_Listbox")
roomtxt% = FindChildByClass(Room%, "RICHCNTL")
If roomlst% <> 0 And roomtxt% <> 0 Then
AOL40_FindChatRoom = Room%
Else
AOL40_FindChatRoom = 0
End If
End Function
Function FindChildByClass(parentw, childhand)
'This will find an MDI Child by the childhand's class
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function

Function FindChildByTitle(parent, childhand)
'Finds a child by title
firs% = GetWindow(parent, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo last
firs% = GetWindow(parent, GW_CHILD)
While firs%
firss% = GetWindow(parent, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo last
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo last
Wend
FindChildByTitle = 0
last:
Room% = firs%
FindChildByTitle = Room%
End Function
Sub AOLHide()
'This hides aol, and if your an idiot i'll explian it deeper.
'It HIDES dosn't CLOSE
X = FindWindow("AOL Frame25", 0&)
Window_Hide (X)
End Sub

Sub AOLShow()
'This Shows aol.
'Again I will explian deeper on the hide/show.
'It SHOWS it if  HIDDEN it. NOT OPEN
X = FindWindow("AOL Frame25", 0&)
Window_Show (X)
End Sub

Function AOLClickList(List)
'Clicks a list
Click% = SendMessage(List, WM_LBUTTONDBLCLK, 0, 0)
End Function
Sub AOLClickIcon(icon%)
'Clicks an icon like one on a IM
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Function Window_Close(win)
'This closes any window of your choice
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Function
Sub ColorIM(Person)
'This sends someone blank IMs with a different colors
'in each one. It sends 5 IMs but then it loops so
'add a stop button
Do:
DoEvents
Call SendIM4(Person, "<body bgcolor=#000000>")
Call SendIM4(Person, "<body bgcolor=#0000FF>")
Call SendIM4(Person, "<body bgcolor=#FF0000>")
Call SendIM4(Person, "<body bgcolor=#00FF00>")
Call ISendIM4(SPerson, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub

Sub IMPUNT(LaMeR, subject)
'Make a stop button or It will punt forever!
'Example:
'IMPUNT("A WgfPoser, Freekin poser learn how to program!")
Do:
DoEvents
Call SendMail4(LaMeR, subject, "<font =999999999999999999999999999999999999999999999999999999999999999999999999")
Loop
End Sub
Function Stopbutton()
'Put this in a command button labled Stop
Do: DoEvents
Loop
End Function
Sub StayOnTop(frm As Form)
'Allows your form to stay on top of all other windows
ontop% = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Window_Minimize(win)
'This will minimize any window of your choice
X = ShowWindow(win, SW_MINIMIZE)
End Sub

Sub Window_Maximize(win)
'This maximizes the window of your choice
X = ShowWindow(win, SW_MAXIMIZE)
End Sub


Sub Window_Hide(hwnd)
'This hides any window of your choice
X = ShowWindow(hwnd, SW_HIDE)
End Sub



'2 color fade combinations begin here

Function BlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlue = Msg
SendChat4 (Msg)
End Function

Function BlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreen = Msg
SendChat4 (Msg)
End Function

Function BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 220 / a
        F = e * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
SendChat4 (Msg)
End Function

Function BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurple = Msg
SendChat4 (Msg)
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRed = Msg
SendChat4 (Msg)
End Function

Function BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellow = Msg
SendChat4 (Msg)
End Function

Function BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlack = Msg
SendChat4 (Msg)
End Function

Function BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreen = Msg
SendChat4 (Msg)
End Function

Function BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurple = Msg
SendChat4 (Msg)
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRed = Msg
SendChat4 (Msg)
End Function

Function BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellow = Msg
SendChat4 (Msg)
End Function

Function GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlack = Msg
SendChat4 (Msg)
End Function

Function GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlue = Msg
SendChat4 (Msg)
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurple = Msg
SendChat4 (Msg)
End Function

Function GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRed = Msg
SendChat4 (Msg)
End Function

Function GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellow = Msg
SendChat4 (Msg)
End Function

Function GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 220 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlack = Msg
SendChat4 (Msg)
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlue = Msg
SendChat4 (Msg)
End Function

Function GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreen = Msg
SendChat4 (Msg)
End Function

Function GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurple = Msg
SendChat4 (Msg)
End Function

Function GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRed = Msg
SendChat4 (Msg)
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellow = Msg
SendChat4 (Msg)
End Function

Function PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlack = Msg
SendChat4 (Msg)
End Function

Function PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlue = Msg
SendChat4 (Msg)
End Function

Function PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreen = Msg
SendChat4 (Msg)
End Function

Function PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRed = Msg
SendChat4 (Msg)
End Function

Function PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellow = Msg
SendChat4 (Msg)
End Function

Function RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlack = Msg
SendChat4 (Msg)
End Function

Function RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlue = Msg
SendChat4 (Msg)
End Function

Function RedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreen = Msg
SendChat4 (Msg)
End Function

Function RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
RedPurple = (Msg)
SendChat (Msg)
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellow = Msg
SendChat4 (Msg)
End Function

Function YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlack = Msg
SendChat4 (Msg)
End Function

Function YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlue = Msg
SendChat4 (Msg)
End Function

Function YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreen = Msg
SendChat4 (Msg)
End Function

Function YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurple = Msg
SendChat4 (Msg)
End Function

Function YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRed = Msg
SendChat4 (Msg)
End Function


' 3 Color fade combinations begin here


Function BlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlueBlack = Msg
SendChat4 (Msg)
End Function

Function BlackGreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreenBlack = Msg
SendChat4 (Msg)
End Function

Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreyBlack = Msg
SendChat4 (Msg)
End Function

Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurpleBlack = Msg
SendChat4 (Msg)
End Function

Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRedBlack = Msg
SendChat4 (Msg)
End Function

Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellowBlack = Msg
SendChat4 (Msg)
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlackBlue = Msg
SendChat4 (Msg)
End Function

Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreenBlue = Msg
SendChat4 (Msg)
End Function

Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurpleBlue = Msg
SendChat4 (Msg)
End Function

Function BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRedBlue = Msg
SendChat4 (Msg)
End Function

Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellowBlue = Msg
SendChat4 (Msg)
End Function

Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlackGreen = Msg
SendChat4 (Msg)
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlueGreen = Msg
SendChat4 (Msg)
End Function

Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurpleGreen = Msg
SendChat4 (Msg)
End Function

Function GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRedGreen = Msg
SendChat4 (Msg)
End Function

Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellowGreen = Msg
SendChat4 (Msg)
End Function

Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlackGrey = Msg
SendChat4 (Msg)
End Function

Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlueGrey = Msg
SendChat4 (Msg)
End Function

Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreenGrey = Msg
SendChat4 (Msg)
End Function

Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurpleGrey = Msg
SendChat4 (Msg)
End Function

Function GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRedGrey = Msg
SendChat4 (Msg)
End Function

Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellowGrey = Msg
SendChat4 (Msg)
End Function

Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlackPurple = Msg
SendChat4 (Msg)
End Function

Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBluePurple = Msg
SendChat4 (Msg)
End Function

Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreenPurple = Msg
SendChat4 (Msg)
End Function

Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRedPurple = Msg
SendChat4 (Msg)
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellowPurple = Msg
SendChat4 (Msg)
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlackRed = Msg
SendChat4 (Msg)
End Function

Function RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlueRed = Msg
SendChat4 (Msg)
End Function

Function RedGreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreenRed = Msg
SendChat4 (Msg)
End Function

Function RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurpleRed = Msg
SendChat4 (Msg)
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellowRed = Msg
SendChat4 (Msg)
End Function

Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlackYellow = Msg
SendChat4 (Msg)
End Function

Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlueYellow = Msg
SendChat4 (Msg)
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreenYellow = Msg
SendChat4 (Msg)
End Function

Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurpleYellow = Msg
SendChat4 (Msg)
End Function

Function YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRedYellow = Msg
SendChat4 (Msg)
End Function


'2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function

'Form back color fade codes are here
'Works best when used in the Form_Paint() sub

Sub FadeFormBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FadeFormGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormPurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub Window_Show(hwnd)
'This shows the window of your choice
X = ShowWindow(hwnd, SW_SHOW)
End Sub
Sub TimeOut(Duration)
'Makes a timeout of the intervel you wish
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub
Sub Upchat()
'Upload and chat or do other stuff at same time
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub pause(interval)
'This pauses all activity in the program for the intervel you enter
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub

Sub Form_Center(frm As Form)
'Centers a form
frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Sub Form_Maximize(frm As Form)
'This will maximize the form of your choice
frm.WindowState = 2
End Sub
Sub Form_Minimize(frm As Form)
'This will minimize the form of your choice
frm.WindowState = 1
End Sub
Sub waitforok()
'Waits for the AOL OK messages that pops up then gets
'rid of it
Do
DoEvents
okw = FindWindow("#32770", "America Online")
DoEvents
Loop Until okw <> 0
okb = FindChildByTitle(okw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub
Sub AOLIMsOn()
'This turns your IMz on
Call SendIM4("$IM_ON", "�")
End Sub
Sub AOLIMsOff()
'This turns on your IMz
Call SendIM4("$IM_OFF", "�")
End Sub
Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)
For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)
For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If
Next GetString
Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub

Function AOLWindow()
'This sets focus on the AOL window
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function


