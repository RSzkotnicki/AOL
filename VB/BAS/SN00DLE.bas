Attribute VB_Name = "Sn00dle95"
'           ;¯¯¯¯¯¯;   ;¯¯¯¯¯¯¯¯;   ;¯¯¯¯¯¯;  ;¯¯¯¯¯¯;     ;¯¯;   ;¯¯;    ;¯¯¯¯¯¯¯¯;    ;¯¯¯¯¯¯¯¯¯¯;    ;¯¯¯¯¯¯¯¯¯¯¯¯;
'          ;  ;¯¯¯¯   ; ;¯¯¯¯; ;   ; ;¯¯; ;  ; ;¯¯; ;     ;  ;   ;  ;    ;  ;¯¯¯¯¯¯    ;  ;¯¯;     ;   ;  ;¯¯¯¯¯¯¯¯¯¯
'      :¯¯¯¯ ;       ; ;    ; ;   ;  ¯¯¯ ;  ;  ¯¯¯ ;  ;¯¯¯  ;   ;  ;    ;   ¯¯¯¯;       ;  ¯¯      ;  ;  ;
'      ¯¯¯¯¯¯        ¯¯     ¯¯     ¯¯¯¯¯     ¯¯¯¯¯   ; ;¯; ;   ;  ;    ;  ;¯¯¯¯¯         ¯¯¯¯¯¯;  ;  ; ¯¯¯¯¯¯¯¯¯¯;
'                                                     ; ¯  ;  ;  ;    ;   ¯¯¯¯¯¯¯;            ;  ;   ¯¯¯¯¯¯¯¯¯;  ;
'      ~By: FiNGr                                      ¯¯¯   ;  ;      ¯¯¯¯¯¯¯¯¯¯             ;  ;           ;  ;
'                                                           ;   ¯¯¯¯¯¯¯¯;                     ;  ;    ;¯¯¯¯¯¯  ;
'                                                            ¯¯¯¯¯¯¯¯¯¯¯                       ¯¯     ¯¯¯¯¯¯¯¯¯
'
'This bas is for Visual Basics 5.  It works for America Online 95 - 32 bit.
'Theres alot of features in this bas file.
'Released: 5-26-98
'Instant Messenge me at ilfingrll.
'My new Sn00dle40.bas for AOL 4.0 will be out pretty soon.




Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)

Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function getnextwindow Lib "user32" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFilename As String) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal Dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (Object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal Dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function getparent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMEssageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function MciSendString& Lib "Winmm" Alias "mciSendStringA" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen&, ByVal hcallback&)


Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_ALIAS = &H10000
Public Const SND_FILENAME = &H20000
Public Const SND_RESOURCE = &H40004
Public Const SND_ALIAID = &H110000
Public Const SND_ALIASTART = 0
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_VALID = &H1F
Public Const SND_NOWAIT = &H2000
Public Const SND_VALIDFLAGS = &H17201F
 
 
Public Const SND_RESERVED = &HFF000000
Public Const SND_TYPE_MASK = &H170007
 
 
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const EM_GETLINE = &HC4


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
Public Const HWND_TOPMOST = -1
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

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESVM_READ = &H10
Public Const STANDARD_RIGHTREQUIRED = &HF0000
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2

Public Const HTERROR = (-2)
Public Const HTTRANSPARENT = (-1)
Public Const HTNOWHERE = 0
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTSIZE = HTGROWBOX
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const HTREDUCE = HTMINBUTTON
Public Const HTZOOM = HTMAXBUTTON
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT

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

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   x As Long
   Y As Long
End Type

Global r&
Global entry$
Global iniPath$

Global sockettype As Integer
Global Const FD_SETSIZE = 64

Type fd_set_type
  fd_count As Integer
  fd_array(FD_SETSIZE) As Integer
End Type
Global FD_SET As fd_set_type

Declare Function FD_ISSET Lib "winsock.dll" Alias "__WSAFDIsSet" (ByVal S As Integer, passed_set As fd_set_type) As Integer



Type timeval
  tv_sec As Long
  tv_usec As Long
End Type

Type SockAddr_in
  sin_family As Integer
  sin_port As Integer
  sin_addr As Long
  sin_zero(7) As String * 1
End Type

Type in_addr
 temp As Long
End Type

Type sockaddr
  sa_family As Integer
  sa_data(14) As Integer
End Type

Global Const WSADESCRIPTION_LEN = 256
Global Const WSASYS_STATUS_LEN = 128

Type WSAdata_type
   wVersion As Integer
   wHighVersion As Integer
   szDescription As String * 257
   szSystemStatus As String * 129
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpVendorInfo As String * 200
End Type

Global WSAdata As WSAdata_type


Type sockproto
  sp_family As Integer
  sp_protocol As Integer
End Type

Type linger
  l_onoff As Integer
  l_linger As Integer
End Type

Declare Function Accept Lib "winsock.dll" (ByVal S As Integer, addr As SockAddr_in, addrlen As Integer) As Integer
Declare Function Bind Lib "winsock.dll" (ByVal S As Integer, addr As SockAddr_in, ByVal namelen As Integer) As Integer
Declare Function closesocket Lib "winsock.dll" (ByVal S As Integer) As Integer
Declare Function htonl Lib "winsock.dll" (ByVal A As Long) As Long
Declare Function inet_addr Lib "winsock.dll" (ByVal S As String) As Long
Declare Function ntohl Lib "winsock.dll" (ByVal A As Long) As Long
Declare Function socket Lib "winsock.dll" (ByVal af As Integer, ByVal typesock As Integer, ByVal protocol As Integer) As Integer
Declare Function htons Lib "winsock.dll" (ByVal A As Integer) As Integer
Declare Function ntohs Lib "winsock.dll" (ByVal A As Integer) As Integer
Declare Function Connect Lib "winsock.dll" (ByVal sock As Integer, sockstruct As SockAddr_in, ByVal structlen As Integer) As Integer
Declare Function send Lib "winsock.dll" (ByVal sock As Integer, ByVal Msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function Recv Lib "winsock.dll" (ByVal sock As Integer, ByVal Msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function Listen Lib "winsock.dll" (ByVal S As Integer, ByVal backlog As Integer) As Integer


Declare Function WSaStartup Lib "winsock.dll" (ByVal A As Integer, b As WSAdata_type) As Integer
Declare Function WSACleanup Lib "winsock.dll" () As Integer



Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2
Global Const AF_INET = 2

Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
Declare Function CreateCompatibleDC% Lib "GDI" (ByVal hDC%)
Declare Function CreateFont% Lib "GDI" (ByVal H%, ByVal W%, ByVal E%, ByVal O%, ByVal W%, ByVal i%, ByVal U%, ByVal S%, ByVal c%, ByVal OP%, ByVal CP%, ByVal Q%, ByVal PAF%, ByVal f$)

Declare Function CreateWindow% Lib "User" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hWndParent%, ByVal hMenu%, ByVal hInstance%, ByVal lpParam$)
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function DrawText Lib "User" (ByVal hDC As Integer, ByVal lpStr As String, ByVal nCount As Integer, lpRect As RECT, ByVal wFormat As Integer) As Integer
Declare Function EnableHardwareInput Lib "User" (ByVal bEnableInput As Integer) As Integer

Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer

Declare Function FlashWindow Lib "User" (ByVal hWnd As Integer, ByVal bInvert As Integer) As Integer


Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer




Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer


Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer

Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetTempDrive Lib "Kernel" (ByVal cDriveLetter As Integer) As Integer
Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer





Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function IsWindow Lib "User" (ByVal hWnd As Integer) As Integer


Declare Function LoadBitmap Lib "User" (ByVal hInstance%, ByVal lpBitMapName As Any) As Integer
Declare Function lstrcpy Lib "Kernel" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long







Declare Function SetMenu Lib "User" (ByVal hWnd As Integer, ByVal hMenu As Integer) As Integer
Declare Function SetMenuItemBitmaps Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal hBitmapUnchecked As Integer, ByVal hBitmapChecked As Integer) As Integer

Declare Function SetPixel Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long) As Long



Declare Function StretchBlt% Lib "GDI" (ByVal hDestDC%, ByVal x%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal xSrc%, ByVal ySrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop As Long)
Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function WindowFromPoint Lib "User" (ByVal ptScreen As Any) As Integer

Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Integer

'API Sub's




Declare Sub hmemcpy Lib "Kernel" (hpvDest As Any, hpvSource As Any, ByVal cbCopy&)
Declare Sub InvertRect Lib "User" (ByVal hDC As Integer, lpRect As RECT)
Declare Sub ModifyMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpString As Long)



 
Declare Sub Yield Lib "Kernel" ()


'Important Global's

Global Const SW_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'API Global's






Global Const EM_GETLINECOUNT = WM_USER + 10
Global Const EM_GETSEL = WM_USER + 0
Global Const EM_REPLACESEL = WM_USER + 18
Global Const EM_SCROLL = WM_USER + 5
Global Const EM_SETFONT = WM_USER + 19
Global Const EM_SETREADONLY = (WM_USER + 31)
Global Const EW_REBOOTSYSTEM = &H43








Global Const KEY_DELETE = &H2E
Global Const LB_32GETCOUNT = &H18B
Global Const LB_32GETCURSEL = &H188
Global Const LB_32GETITEMDATA = &H199
Global Const LB_32GETTEXT = &H189
Global Const LB_32GETTEXTLEN = &H18A
Global Const LB_32SETCURSEL = &H186

Global Const LB_GETITEMRECT = (WM_USER + 25)
Global Const LBN_DBLCLK = 2
Global Const MB_TASKMODAL = &H2000
Global Const MF_BITMAP = &H4
Global Const SRCCOPY = &HCC0020
Global Const SW_NORMAL = 1
Global Const SW_SHOWNA = 8



Global Const WM_COPY = &H301
Global Const WM_GETFONT = &H31





Global Const WM_MOVE = &H3

Global Const WM_SETCURSOR = &H20
Global Const WM_SETFONT = &H30

Global Const WS_BORDER = &H800000
Global Const WS_THICKFRAME = &H40000

'Other Globals
Global Abort As Integer
Global AscBord(2) As String
Global FindChild As Integer
Global HoldText As String
Global IntMin As Integer
Global IntSec As Integer
Global OldText As String
Global OldTextLength As Integer
Global Pause As Integer

Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd%, ByVal hDC%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer

Sub Hide(waht As String)
A = showwindow(waht, SW_HIDE)
End Sub







Function GetText(child)
GathTrim = sendmessagebynum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GathTrim)
GathString = SendMEssageByString(child, 13, GathTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, GW_CHILD)
   

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

Function Getwindtext(read As Integer)
BufLen% = sendmessagebynum(read%, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String(BufLen%, 0)
Q% = SendMEssageByString(read%, WM_GETTEXT, BufLen% + 1, Buffer$)
DoEvents
Getwindtext = Trim(Buffer$)
End Function


Sub FileCopy(file$, DestFile$)
If Not FileExists(file$) Then Exit Sub
FileCopy file$, DestFile$
End Sub

Sub Settext(Where%, thestring$)
Dim SendTheText%
SendTheText% = SendMEssageByString(Where%, WM_SETTEXT, 0, thestring$)
End Sub
Sub LoopClickFav(name, timer As timer)
rm% = FindWindow("AOLChild", name)
prefer% = FindChildByTitle(AOLMDI(), "Favorite Places")
but% = FindChildByTitle(prefer%, "Connect")
Do
Timeout (0.2)
Clickbut but%
A% = FindWindow("#32770", "America Online")
z% = FindChildByClass(A%, "Button")
If A% <> 0 Then
Clickbut z%
End If
Loop Until rm% <> 0

del% = FindChildByTitle(prefer%, "Delete")
Clickbut del%
Q% = FindWindow("#32770", "America Online")
V% = FindChildByClass(A%, "Button")
Clickbut V%
S = sendmessagebynum(prefer%, WM_CLOSE, 0, 0)
End Sub


Sub Favplace(name As String, address As String)
Call RunAOLmenu("Favorite Places")
Do ': DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Favorite Places")
maillab% = FindChildByTitle(prefer%, "Add Favorite Place")
Tree% = FindChildByClass(prefer%, "_AOL_Tree")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop
Timeout (1)
Clickbut maillab%
Timeout (0.5)

Do: DoEvents
Add% = FindChildByTitle(AOLMDI(), "Add Favorite Place")
but% = FindChildByClass(Add%, "_AOL_Edit")
but2% = GetWindow(but%, GW_HWNDNEXT)
but3% = GetWindow(but2%, GW_HWNDNEXT)
If Add% <> 0 And but% <> 0 Then Exit Do
Loop
Timeout (0.4)
Call Settext(but%, name)
Call Settext(but3%, address)
But4% = FindChildByTitle(Add%, "OK")
Clickbut But4%
clikr% = FindWindow("#32770", "America Online")
clikbut% = FindChildByClass(clikr%, "Button")
Timeout (0.2)
If clikr% <> 0 Then
Clickbut clikbut%
Timeout (0.1)
S = sendmessagebynum(Add%, WM_CLOSE, 0, 0)
End If
For D = 0 To sendmessagebynum(Tree%, LB_GETCOUNT, 0, 0&) - 1
S = sendmessagebynum(Tree%, LB_SETCURSEL, D, 0&)
Next D
End Sub




Sub GetMemberProfile(name As String)

RunAOLmenu ("Get a Member's Profile")
Timeout 0.3
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
prof% = FindChildByTitle(MDI%, "Get a Member's Profile")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call Settext(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
Clickbut okbutton%
End Sub

Sub Add_list_with_comma(List1 As ListBox, List2 As ListBox)
x = InputBox("Enter the screen name and password seperated with a ,   Example: SteveCase,AOLRules")
A = Trim(x)
z = InStr(A, ",")
If z Then
List1.AddItem Left(A, z - 1)
List2.AddItem Mid(A, z + 1)
End If

End Sub

Sub Fade_DitherForm(vForm As Form)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub addroom(Lst As ListBox)

'This Adds The Room To A List Box
'Example: Call Addroom(List1)

LB = FindChildByClass(FindChatRoom(), "_AOL_Listbox")
Coun = sendmessagebynum(LB, LB_GETCOUNT, 0, 0)
For G = 0 To Coun - 1
x = sendmessagebynum(LB, LB_SETCURSEL, G, 0)
Q = sendmessagebynum(LB, WM_LBUTTONDBLCLK, 0, 0)
z = RoomName
chil% = FindChildByClass(AOLMDI(), "AOL Child")
Mess = FindChildByTitle(chil%, "Message")

Do
x = getwintext(chil%)
Loop Until z <> x

Lst.AddItem x
S = sendmessagebynum(chil%, WM_CLOSE, 0, 0)
Next G
For i = 0 To Lst.ListCount - 1
If Lst.List(i) = "" & aolUser & "" Then
Lst.RemoveItem i
End If
Next i
End Sub

Sub AddroomCombo(Cmb As ComboBox)

'This Adds The Room To A List Box
'Example: Call AddroomCombo(combo1)

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = FindChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESVM_READ Or STANDARD_RIGHTREQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Cmb.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
For i = 0 To Cmb.ListCount - 1
If Cmb.List(i) = "" & CaPRe_User & "" Then
Cmb.RemoveItem i
End If
Next i
End Sub

Sub AddroomList(Lst As ListBox)

'This Adds The Room To A List Box
'Example: Call AddroomList(list1)

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = FindChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESVM_READ Or STANDARD_RIGHTREQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Lst.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
For i = 0 To Lst.ListCount - 1
If Lst.List(i) = "" & GetUser() & "" Then
Lst.RemoveItem i
End If
Next i
End Sub

Sub AddtoAOL(frm As Form, xpos, ypos)

'This Adds Whatever You Want Into AOL.
'Example: Call AddtoAOL(Form1, ,34, 34)

frm.Top = ypos
frm.Left = xpos
AOL% = FindWindow("AOL FRAME25", vbNullString)
TL% = FindChildByClass(AOL%, "AOL TOOLBAR")
sett = SetParent(frm.hWnd, AOL%)
ack = showwindow(AOL%, 2)
ack = showwindow(AOL%, 3)

End Sub

Sub AntiPunt()

'This Is An Anti-Punter.
'It Locks The IM Box.

AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = SnFromIM
x = Getwindtext(rch2%)
If InStr(x, "    ") Then
Do
S = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
sendtext "" & nme & " Is trying to punt me"
End If

If InStr(x, "") Then
Do
S = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
sendtext "" & nme & " Is trying to punt me"
End If
End Sub

Function AOLGetListString(Parent, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

aolhandle = Parent

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESVM_READ Or STANDARD_RIGHTREQUIRED, False, AOLProcess)

If AOLProcessThread Then
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = person$
End Function

Function AOLGetTopWindow()
AOLGetTopWindow = GetTopWindow(AOLMDI())
End Function

Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Function AOLRoomCount()
chld% = FindChatRoom()
ListBox% = FindChildByClass(chld%, "_AOL_Listbox")

countem = SendMessage(ListBox%, LB_GETCOUNT, 0, 0)
AOLRoomCount = countem
End Function

Sub AOLSignoff()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunAOLmenu("&Sign Off")

End Sub

Function AOLToolbar()
AOL% = FindWindow("AOL Frame25", vbNullString)
tob% = FindChildByClass(AOL%, "AOL Toolbar")
AOLToolbar = tob%
End Function

Function AOLVersion()
RunAOLmenu ("&About America Online")
Do
bout% = FindWindow("_AOL_Modal", vbNullString)
Loop Until bout% <> 0
stc% = FindChildByClass(bout%, "_AOL_Static")
red = GetText(stc%)
If InStr(1, red, "Windows 95") Then
S = sendmessagebynum(bout%, WM_CLOSE, 0, 0)
AOLVersion = "32 bit"
Else
S = sendmessagebynum(bout%, WM_CLOSE, 0, 0)
AOLVersion = "16 bit"
End If
End Function

Function AOLWin()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWin = AOL%
End Function

Sub ButtonDblClk(Button%)
Dim DoubleClick
DoubleClick = sendmessagebynum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub

Sub Center(frm As Form)

'This Centers The Form
'Example: CenterForm(Form1)

abc = (Screen.Width - frm.Width) / 2
xyz = (Screen.Height - frm.Height) / 2
frm.Move abc, xyz
End Sub

Sub Change_AOL_Caption(newcaption As String)

'This Changes AOLs Caption.
'Example: Call Change_AOL_Caption(Sn00dle95 Bas Rules)

Call Settext(AOLMDI, newcaption)
End Sub

Sub Clickbut(Wut%)

clicks = SendMessage(Wut%, WM_KEYDOWN, VK_SPACE, 0)
clicks = SendMessage(Wut%, WM_KEYUP, VK_SPACE, 0)
End Sub

Sub Clickicon(icon%)
c = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
c = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function ClickList(hWnd)
cliklst% = sendmessagebynum(hWnd, &H203, 0, 0&)
End Function

Sub CloseAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
closes = SendMessage(AOL%, WM_CLOSE, 0, 0)
End Sub

Sub CloseGoodbye()
AOL% = FindChildByTitle(AOLWin(), "Goodbye from America Online!")
D = sendmessagebynum(AOL%, WM_CLOSE, 0, 0)
End Sub

Sub Closewin(wind)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)

End Sub

Sub Common_Load(List1 As ListBox)
Dim A As Variant
Dim b As Variant
CMDialog1.DialogTitle = "Load List File" ' set title
CMDialog1.Filter = "Tee (*.txt)|*.txt|All Files (*.*)|*.*|"
CMDialog1.FileName = "*.txt"
CMDialog1.FLAGS = &H1000&
CMDialog1.Action = 1
A = 1
If (CMDialog1.FileTitle <> "") Then
List1.Clear ' clear the list
Open CMDialog1.FileTitle For Input As A
While (EOF(A) = False)
Line Input #A, b
List1.AddItem b
Wend
Close A
End If
End Sub

Sub Common_Save(List1 As ListBox)
Dim b As Variant
CMDialog1.DialogTitle = "Save List File" ' set CMDialog's title bar
CMDialog1.Filter = "Tee (*.txt)|*.txt|All Files (*.*)|*.*|"
CMDialog1.FLAGS = &H1000&
CMDialog1.FileName = "*.txt"
CMDialog1.Action = 2
If (CMDialog1.FileTitle <> "") Then
A = 2
Open CMDialog1.FileName For Output As A
b = 0
Do While b < List1.ListCount
Print #A, List1.List(b)
b = b + 1
Loop
Close A
End If
End Sub

Sub CopyFile(FileName$, CopyTo$)

If FileName$ = "" Then Exit Sub
If CopyTo$ = "" Then Exit Sub
If Not FileExists(FileName$) Then Exit Sub
On Error GoTo AnErrOccured
If InStr(Right$(FileName$, 4), ".") = 0 Then Exit Sub
If InStr(Right$(CopyTo$, 4), ".") = 0 Then Exit Sub
FileCopy FileName$, CopyTo$
Exit Sub
AnErrOccured:
MsgBox "An Unexpected Error Occured!", 16, "Error"
End Sub

Function EnterKey()
EnterKey = CStr(Chr(13) & Chr(10))
End Function

Sub Exit_prog(Form As Form)
    
    'Put In Query Unload
    
    Dim Msg ' Declare variable.

If UnloadMode > 0 Then
        ' If exiting the application.
        Msg = "Are you sure you want to leave?"
    Else
        ' If just closing the form.
        Msg = "Dont Leave"
    End If
    ' If user clicks the No button, stop QueryUnload.
    If MsgBox(Msg, vbQuestion + vbYesNo + vbSystemModal, Form.Caption) = vbNo Then Cancel = True
End Sub

Sub Change_Welc_Caption(blah As String)
'when u have the sign on window open this makes it look like yer signed on
'so u can open certain features like fate when yer not signed on

'Example: Fakesignedon "Layzie"

x = SendMEssageByString(Sign_On_Window(), WM_SETTEXT, 0, "Welcome, " + blah + "!")


End Sub

Sub FileAdd(file$, WhatToAdd$)

If FileExists(file$) = False Then Exit Sub
Open file$ For Append As #1
    Print #1, WhatToAdd$
    Close #1
End Sub

Sub FileAddto(file$, WhatToAdd$)

If FileExists(file$) = False Then Exit Sub
Open file$ For Append As #1
    Print #1, WhatToAdd$
    Close #1
End Sub

Sub FileDelete(FileName$)
If Not FileExists(FileName$) Then MsgBox FileName$ & Chr(13) & "Bad File Name!", 16, "Error": Exit Sub
On Error GoTo ErrorInDeletion
Kill FileName$
Exit Sub
ErrorInDeletion:
MsgBox Error$
Resume Exitinga
Exitinga:
Exit Sub
End Sub

Function FileExists(ByVal sFileName As String) As Integer
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        FileExists = False
        Else
            FileExists = True
    End If

End Function

Sub FileMakDirectory(DirName$)
MkDir DirName$
End Sub

Sub FileOpen(file$)
x = Shell(file$, 1): NoFreeze% = DoEvents()
End Sub

Sub FileProperties(Choice, file$)

If Not FileExists(file$) Then Exit Sub
If LCase$(Choice) = "normal" Then
SetAttr file$, vbNormal
ElseIf LCase$(Choice) = "readonly" Then: SetAttr file$, vbReadOnly
ElseIf LCase$(Choice) = "hidden" Then: SetAttr file$, vbHidden
ElseIf LCase$(Choice) = "system" Then: SetAttr file$, vbSystem
ElseIf LCase$(Choice) = "archive" Then: SetAttr file$, vbArchive
End If
NoFreeze% = DoEvents()
End Sub

Sub FileRename(file$, NewName$)
Name file$ As NewName$
NoFreeze% = DoEvents()
End Sub

Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindChatRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function

Function FindChildByClass(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
Buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, Buffer$, 250)

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByClass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function FindChildByTitle(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
hwndlength% = GetWindowTextLength(childfocus%)
Buffer$ = String$(hwndlength%, 0)
WindowText% = GetWindowText(childfocus%, Buffer$, (hwndlength% + 1))

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function FindFWDWindow()

wind = FindChildByClass(AOLMDI, "AOL CHILD")
SN = FindChildByTitle(wind, "Send Now")
SL = FindChildByTitle(wind, "Send Later")

Toz = FindChildByClass(wind, "_AOL_ICON")

If SL <> 0 Then
FindFWDWindow = 1
Else
FindFWDWindow = 0
End If
End Function

Function FindKeyword()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
keyw% = FindChildByTitle(MDI%, "Keyword")
KEdit% = FindChildByClass(keyw%, "_AOL_Edit")
FindKeyword = KEdit%
End Function

Function FindRoom()

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfcs% = GetWindow(MDI%, 5)

While childfcs%
listers% = FindChildByClass(childfcs%, "_AOL_Edit")
listere% = FindChildByClass(childfcs%, "_AOL_View")
listerb% = FindChildByClass(childfcs%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindRoom = childfcs%: Exit Function
childfcs% = GetWindow(childfcs%, GW_HWNDNEXT)
Wend
End Function

Sub First_Time_Load()
If Len(Dir(App.Path + "\" + "first.txt")) = 0 Then
MsgBox "This is the first time your running this program"
Open (App.Path + "\" + "first.txt") For Append As #1
Print #1, "Hey!"
Close #1
Else
Open (App.Path + "\" + "first.txt") For Append As #1
Print #1, "Hey!"
Close #1
End If
End Sub

Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop

End Function

Function GetCaption(hWnd)
hwndlength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndlength%, 0)
A% = GetWindowText(hWnd, hwndTitle$, (hwndlength% + 1))

GetCaption = hwndTitle$
End Function

Function GetAOL() As Integer
AOL% = FindWindow("AOL Frame25", 0&)
Menu% = GetMenu(AOL%)
aol2% = SearchMenu(Menu%, "Download Manager...")
If (aol2% <> 0) Then
    GetAOL = 2
    Exit Function
End If
aol3% = SearchMenu(Menu%, "&Log Manager")
If (aol3% <> 0) Then
    AOL95% = SearchMenu(Menu%, "America Online Help &Topics")
    If (AOL95% <> 0) Then
        GetAOL = 95
    Else:
        GetAOL = 3
    End If
    Exit Function
End If
AOL4% = SearchMenu(Menu%, "Open Picture &Gallery...")
If (AOL4% <> 0) Then
    GetAOL = 4
    Exit Function
End If
    
'menu% = GetMenu(FindWindow("AOL Frame25", 0&))
'If menu% = 0 Then Exit Function
'MenuName$ = Space(500)
'x = GetMenuString(menu%, 0, MenuName$, 500, WM_USER)
'If InStr(1, FixAPIString(MenuName$), "*", 1) Then subMenu% = GetSubMenu(menu%, 1)
'If InStr(1, FixAPIString(MenuName$), "&File", 1) Then subMenu% = GetSubMenu(menu%, 0)
'mnuCount% = GetMenuItemCount(subMenu%)
'Count = 0
'Do
'    MenuName$ = String(500, Chr(32))
'    x = GetMenuString(subMenu%, Count, MenuName$, 500, WM_USER)
'    Count = Count + 1
'    If InStr(FixAPIString(MenuName$), "&Log Manager") Then Exit Do
'    If InStr(FixAPIString(MenuName$), "Download Manager...") Then Exit Do
'Loop Until Count = mnuCount%
'If InStr(FixAPIString(MenuName$), "&Log Manager") Then
'    mnuCount% = GetMenuItemCount(menu%)
'    MenuName$ = String(500, Chr(32))
'    If online() = True Then mnuCount% = mnuCount% - 1
'    x = GetMenuString(menu%, mnuCount% - 1, MenuName$, 500, WM_USER)
'    If InStr(1, MenuName$, "&Help", 1) Then Found = True
'    If Found = True Then
'        Found = False
'        SubMnu% = GetSubMenu(menu%, mnuCount% - 1)
'        SubMnuCount% = GetMenuItemCount(SubMnu%)
'        Count = 0
'        Do
'            MenuName$ = String(500, Chr(32))
'            x = GetMenuString(SubMnu%, Count, MenuName$, 500, WM_USER)
'            Count = Count + 1
'            MenuName$ = FixAPIString(MenuName$)
'            If InStr(1, MenuName$, "America Online Help &Topics", 1) Then
'                Found = True
'                Exit Do
'            End If
'        Loop Until Count = SubMnuCount%
'    End If
'    If Found = True Then
'        GetAOL = 95
'    Else :
'        GetAOL = 3
'    End If
'ElseIf InStr(FixAPIString(MenuName$), "Download Manager...") Then
'    GetAOL = 2
'Else :
'    GetAOL = 4
'End If
End Function

Function GetChatText() As String
On Error Resume Next
If FindChatRoom() = 0 Then Exit Function
ChatText$ = getapitext(ChatView())
For x = Len(ChatText$) To 1 Step -1
If Mid(ChatText$, x, 1) = Chr(13) Then Exit For
Next x
ChatText$ = Mid(ChatText$, x, Len(ChatText$))
GetChatText = ChatText$
End Function

Function ChatView() As Integer
Chat% = FindChatRoom()
view% = getwindowbyclass(Chat%, "_AOL_View")
ChatView = view%
End Function

Function GetChildWin(Parent As Integer, Caption As String, Class As String) As Integer
win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text1$ = getapitext(win%)
text2$ = GetClass(win%)
If InStr(1, Text1$, Caption$, 1) And InStr(1, text2$, Class$, 1) Then Exit Do
win% = GetWindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
GetChildWin = win%
End Function

Function GetClass(hWnd As Integer) As String
Text$ = Space(1000)
x = GetClassName(hWnd%, Text$, 1000)
Text$ = FixAPIString(Text$)
GetClass = Text$
End Function

Function GetGuest() As Integer
MODAL% = FindWindow("_AOL_Modal", 0&)
ScreenName% = getnextwindow(getwindowbytitle(MODAL%, "Screen Name:"), 1)
Password% = getnextwindow(getwindowbytitle(MODAL%, "Password:"), 1)
If (ScreenName% <> 0) And (Password% <> 0) Then GetGuest = MODAL%
End Function


Sub getnum(og%, A)
Do
If A = 0 Then Exit Sub
b = 1 + b
og% = GetWindow(og%, GW_HWNDNEXT)
Loop Until b >= A - 1
End Sub

Function GetRoomName() As String
On Error Resume Next
GetRoomName = getapitext(FindChatRoom())
End Function

Function GetRoomURL() As String
x = SetFocusAPI(FindChatRoom())
RunMenu "Add to Favorite Places"
If GetAOL() = 2 Then
    AOL% = FindWindow("AOL Frame25", 0&)
    MDI% = getwindowbyclass(AOL%, "MDIClient")
    AOLToobar% = getwindowbyclass(AOL%, "AOL Toolbar")
    AOLIcon% = getnextwindow(GetWindow(AOLToobar%, GW_CHILD), 19)
    Clickicon (AOLIcon%)
End If
Go:
Do
    DoEvents
    MODAL% = FindWindow("#32770", "Name Conflict")
    FavoritePlaces% = getwindowbytitle(MDI%, "Favorite Places")
Loop Until (MODAL% <> 0) Or (FavoritePlaces% <> 0)
If (MODAL% <> 0) Then
    Call Closewin(MODAL%)
    Clickicon (AOLIcon%)
    GoTo Go
End If
If (FavoritePlaces% <> 0) Then
    AOLTree% = getwindowbyclass(FavoritePlaces%, "_AOL_Tree")
    num = sendmessagebynum(AOLTree%, LB_GETCOUNT, 0, 0)
    For x = 0 To num - 1
        length = sendmessagebynum(AOLTree%, LB_GETTEXTLEN, x, 0)
        Text$ = Space(length)
        i = SendMEssageByString(AOLTree%, LB_GETTEXT, x, Text$)
        Text$ = FixAPIString(Text$)
        If SpaceCase(Text$) = SpaceCase(GetRoomName()) Then
            i = sendmessagebynum(AOLTree%, LB_SETCURSEL, x, 0)
            Exit For
        End If
    Next x
    Modify% = getwindowbytitle(FavoritePlaces%, "Modify")
    Clickicon (Modify%)
    Do
    DoEvents
    Modi% = getwindowbytitle(MDI%, Text$)
    EnterURL% = GetWindow(getwindowbytitle(Modi%, "Enter the Internet Address:"), GW_HWNDNEXT)
    If EnterURL% <> 0 Then
        GetRoomURL = getapitext(EnterURL%)
        Call Closewin(Modi%)
        Call Closewin(FavoritePlaces%)
        Exit Do
    End If
    Loop
End If
End Function

Function GetSignOn() As Integer
If online() = True Then Exit Function
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDI Client")
Goodbye% = getwindowbytitle(MDI%, "Goodbye From America Online!")
Welcome% = getwindowbytitle(MDI%, "Welcome")
If Goodbye% <> 0 Then SignOn% = Goodbye%
If Welcome% <> 0 Then SignOn% = Welcome%
If IsWindowVisible(SignOn%) <> 0 Then
  GetSignOn = SignOn%
Else:
  GetSignOn = 0
End If
End Function

Function GetTextLen(hWnd As Integer) As Integer
GetTextLen = sendmessagebynum(hWnd%, WM_GETTEXTLENGTH, 0, 0)
End Function

Function getusersn() As String
On Error Resume Next
If online() = False Then Exit Function
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDI Client")
Welcome% = getwindowbytitle(MDI%, "Welcome, ")
Text$ = getapitext(Welcome%)
If Len(Text$) = 0 Then Exit Function
Text$ = Mid(Text$, InStr(Text$, ",") + 2)
Text$ = Mid(Text$, 1, InStr(Text$, "!") - 1)
getusersn = Trim(Text$)
End Function

Function getwindowbyclass(Parent As Integer, ByVal Class As String) As Integer
win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = GetClass(win%)
If SpaceCase(Text$) = SpaceCase(Class$) Then Exit Do
If FindChild = True Then
    If GetWindow(win%, GW_CHILD) Then
        ChildWin% = getwindowbyclass(win%, Class$)
        If ChildWin% <> 0 Then
            win% = ChildWin%
            Exit Do
        End If
    End If
End If
win% = GetWindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
getwindowbyclass = win%
End Function

Function getwindowbytitle(Parent As Integer, ByVal Title As String) As Integer
win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = FixAPIString(getapitext(win%))
If InStr(1, Text$, Title$, 1) Then Exit Do
If FindChild = True Then
    If GetWindow(win%, GW_CHILD) Then
        ChildWin% = getwindowbytitle(win%, Title$)
        If ChildWin% <> 0 Then
            win% = ChildWin%
            Exit Do
        End If
    End If
End If
win% = GetWindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
getwindowbytitle = win%
End Function

Function getwintext(hWnd As Integer) As String
lentos = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
x = SendMEssageByString(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
getwintext = Buffer$
End Function

Sub KillAd()
Advertisement% = FindWindow("_AOL_MODAL", 0&)
Cancel1% = getwindowbytitle(Advertisement%, "Cancel")
Cancel2% = getwindowbytitle(Advertisement%, "No Thanks")
If Cancel1% <> 0 Then CancelButton% = Cancel1%
If Cancel2% <> 0 Then CancelButton% = Cancel2%
If (CancelButton% <> 0) And (Advertisement% <> GetGuest()) Then Clickicon (CancelButton%)
End Sub

Sub killwait()
AOL% = FindWindow("AOL Frame25", 0&)
x = EnableWindow(AOL%, True)
If GetAOL() = 2 Then RunMenu "Exit Free Area"
If GetAOL() = 3 Or GetAOL() = 95 Then RunMenu "Exit Unlimited Use area"
End Sub

Sub killwin(Windo)
x = sendmessagebynum(Windo, WM_CLOSE, 0, 0)
End Sub

Function Locate(who As String) As String
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
RunMenu "Send an Instant Message"
Do
    DoEvents
    IMs% = getwindowbytitle(MDI%, "Send Instant Message")
Loop Until IMs% <> 0
AOLEdit% = getwindowbyclass(IMs%, "_AOL_Edit")
If GetAOL() = 2 Then Message% = getnextwindow(AOLEdit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then Message% = getwindowbyclass(IMs%, "RICHCNTL")
Call setedit(AOLEdit%, who$)
Call setedit(Message%, " ")
Avail% = getwindowbytitle(IMs%, "Available")
If Avail% = 0 Then Avail% = getnextwindow(Message%, 1)
Clickicon (Avail%)
Do
    DoEvents
    Off% = FindWindow("#32770", "America Online")
    MsgStatic% = getnextwindow(getwindowbyclass(Off%, "Static"), 1)
    If (Off% <> 0) Then
        txt$ = getapitext(MsgStatic%)
        Locate = txt$
        Call Closewin(Off%)
        Call Closewin(IMs%)
        Exit Do
    End If
Loop
Call Closewin(IMs%)

End Function

Function Mail_ListToText2(Lst As ListBox) As String
For i = 0 To Lst.ListCount - 1
    Final$ = Final$ & "," & Lst.List(i)
    Next i
MM_createmaillist$ = "( " & Final$ & " )"
End Function

Function online() As Integer
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
Welcome% = getwindowbytitle(MDI%, "Welcome, ")
Cap$ = getapitext(Welcome%)
If InStr(Cap$, ",") <> 0 Then online = True
If InStr(Cap$, ",") = 0 Then online = False
End Function

Sub PercentBar(Shape As Control, Done As Integer, total As Variant)
On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False
x = Done / total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
Shape.Line (0, 0)-(x - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(255, 0, 0)
Shape.Print Percent(Done, total, 100) & "%"
End Sub

Sub ReEnter()
Call Keyword("aol://2719:2-2-" + GetRoomName())
StartTime = timer
Do While (timer - StartTime < 5)
    DoEvents
    Text$ = GetChatText()
    If InStr(Text$, "OnlineHost:") Then Exit Do
    Full% = FindWindow("#32770", "America Online")
    If Full% <> 0 Then
        Call Closewin(Full%)
        Exit Do
    End If
Loop
Call killwait
End Sub

Function RemoveDeadNames(List As ListBox) As Integer
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
ErrorWin% = getwindowbytitle(MDI%, "Error")
If ErrorWin% <> 0 Then
    AOLView% = getwindowbyclass(ErrorWin%, "_AOL_View")
    Text$ = getapitext(AOLView%)
Search_For_Dead:
    For x = 0 To List.ListCount - 1
        If InStr(1, RemoveSpace(Text$), RemoveSpace((List.List(x))), 1) Then List.RemoveItem x: GoTo Search_For_Dead
    Next x
    RemoveDeadNames = True
    Call Closewin(ErrorWin%)
End If
End Function

Function RemoveSpace(txt$) As String
NoSpace$ = txt$
While InStr(NoSpace$, " ") <> 0
Where = InStr(NoSpace$, " ")
NoSpace$ = Mid(NoSpace$, 1, Where - 1) + Mid(NoSpace$, Where + 1)
Wend
RemoveSpace = NoSpace$
End Function

Function RemoveString(txt As String, Char As String) As String
NoChar$ = txt$
While InStr(NoChar$, Char$) <> 0
Where = InStr(NoChar$, Char$)
NoChar$ = Mid(NoChar$, 1, Where - 1) + Mid(NoChar$, Where + Len(Char$))
Wend
RemoveString = NoChar$
End Function

Function Get_Class(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

Get_Class = Buffer$
End Function

Function GetLineCount(Text)

theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Sub ResetSN(SN As String, aoldir As String, Replace As String)
SN$ = SN$ + String(10 - Len(SN$), Chr(32))
Replace$ = Replace$ + String(10 - Len(Replace$), Chr(32))
Free = FreeFile
Open aoldir$ + "\idb\main.idx" For Binary As #Free
For x = 1 To LOF(Free) Step 32000
    Text$ = Space(32000)
    Get #Free, x, Text$
Search:
    If InStr(1, Text$, SN$, 1) Then
        Where = InStr(1, Text$, SN$, 1)
        Put #Free, (x + Where) - 1, Replace$
        Mid$(Text$, Where, 10) = String(10, " ")
        GoTo Search
    End If
    DoEvents
Next x
Close #Free
End Sub

Sub RunMenu(MenuCaption As String)
AOL% = FindWindow("AOL Frame25", 0&)
Menu% = GetMenu(AOL%)
ID% = SearchMenu(Menu%, MenuCaption$)
x = sendmessagebynum(AOL%, WM_COMMAND, ID%, 0)

'AOL% = FindWindow("AOL Frame25", 0&)
'menu% = GetMenu(AOL%)
'mnuCount% = GetMenuItemCount(menu%)
'For Mnu = 0 To mnuCount%
'    subMenu% = GetSubMenu(menu%, Mnu)
'    SubMnuCount% = GetMenuItemCount(subMenu%)
'    For SubMnu = 0 To SubMnuCount%
'        SubSubMenu% = GetSubMenu(subMenu%, SubMnu)
'        'for menu's with double sub menu's
'        'If SubSubMenu% <> 0 Then
'        '    SubSubMnuCount% = GetMenuItemCount(SubSubMenu%)
'        '    For SubSubMnu = 0 To SubSubMnuCount%
'        '        txt$ = Space(256)
'        '        x = GetMenuString(SubSubMenu%, SubSubMnu, txt$, 256, True)
'        '        txt$ = FixAPIString(txt$)
'        '        If InStr(UCase(txt$), UCase(MenuCaption$)) Then
'        '            ID% = GetMenuItemID(SubSubMenu%, SubSubMnu)
'        '            Found = True
'        '        End If
'        '        If Found = True Then Exit For
'        '    Next SubSubMnu
'        '    If Found = True Then Exit For
'        'End If
'        Txt$ = Space(256)
'        x = GetMenuString(subMenu%, SubMnu, Txt$, 256, WM_USER)
'        Txt$ = FixAPIString(Txt$)
'        If InStr(UCase(Txt$), UCase(MenuCaption$)) Then
'            ID% = GetMenuItemID(subMenu%, SubMnu)
'            Found = True
'        End If
'        If Found = True Then Exit For
'    Next SubMnu
'    If Found = True Then Exit For
'Next Mnu
'x = SendMessageByNum(AOL%, WM_COMMAND, ID%, 0)
End Sub

Function ScanFile(FileName As String, SearchString As String) As Long
Free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #Free
For x = 1 To LOF(Free) Step 32000
    Text$ = Space(32000)
    Get #Free, x, Text$
    Debug.Print x
    If InStr(1, Text$, SearchString$, 1) Then
        Where = InStr(1, Text$, SearchString$, 1)
        ScanFile = (Where + x) - 1
        Close #Free
        Exit For
    End If
    Next x
Close #Free
End Function

Function Scramble(Word$) As String
Static Words(256) As String
Static Dick(256) As String
Word$ = Word$
Word$ = Word$ + " "
Do
    Where = InStr(UCase$(Word$), UCase$(" "))
    If Where = False Then Exit Do
    Sep$ = (Mid$(Word$, 1, Where - 1))
    x = x + 1
    Words$(x) = (Sep$)
    Word$ = Mid$(Word$, Where + 1)
Loop
For i = 1 To x
    Dick$(i) = ScrambleWord(Words(i))
Next i
For i = 1 To x
    shit$ = shit$ + Dick$(i) + " "
Next i
shit$ = (shit$)
Scramble = Trim$(shit$)
End Function

Function ScrambleWord(Word$)
Static letter(1000)
Static KK(1000)
S = Word$
tak = 0

For Y = 1 To Len(S)
p = p + 1
letter(p) = Mid(S, Y, 1)
Next Y

For Q = 1 To Len(S) '- 1
againdood:
Randomize timer
f = Int(Rnd * Len(S) + 1)
If f = 0 Then GoTo againdood
For W = 0 To tak
If f = KK(W) Then GoTo againdood
Next W
tak = tak + 1
KK(tak) = f
tempt = tempt & letter(f)
Next Q
ScrambleWord = tempt '& Right(S, 1)
End Function

Function SearchMenu(mnuWnd As Integer, MenuCaption As String) As Integer
mnuCount = GetMenuItemCount(mnuWnd%)
For num = 0 To mnuCount - 1
    Text$ = Space(100)
    x = GetMenuString(mnuWnd%, num, Text$, 100, WM_USER)
    Text$ = FixAPIString(Text$)
    SubMenu% = GetSubMenu(mnuWnd%, num)
    If InStr(1, Text$, MenuCaption$, 1) Then
        SubMenu% = GetSubMenu(mnuWnd%, num)
        Menu% = SubMenu%
        MenuID% = GetMenuItemID(mnuWnd%, num)
    ElseIf (SubMenu% <> 0) Then
        MenuID% = SearchMenu(SubMenu%, MenuCaption$)
    End If
    If (MenuID% <> 0) Then
        Exit For
    End If
Next num
SearchMenu = MenuID%
End Function


Sub Sendd(chatedit, sill$)
sndtext = SendMEssageByString(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub SendEMail(Names As String, Subject As String, Message As String, SendMail As Integer, WaitForSend As Integer)
On Error Resume Next
If online() = False Then Exit Sub
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
RunMenu "&Compose Mail"
Do
DoEvents
Compose% = getwindowbytitle(MDI%, "Compose Mail")
Loop Until Compose% <> 0
NamesEdit% = GetWindow(getwindowbytitle(Compose%, "To:"), GW_HWNDNEXT)
SubjectEdit% = GetWindow(getwindowbytitle(Compose%, "Subject:"), GW_HWNDNEXT)
If GetAOL() = 2 Then MessageEdit% = getnextwindow(SubjectEdit%, 3)
If GetAOL() = 3 Or GetAOL() = 95 Then MessageEdit% = getwindowbyclass(Compose%, "RICHCNTL")
SendNow% = getwindowbyclass(Compose%, "_AOL_Icon")
Send_The_Mail:
Call setedit(NamesEdit%, Names$)
Call setedit(SubjectEdit%, Subject$)
Call setedit(MessageEdit%, Message$)
If SendMail = True Then Clickicon (SendNow%)
If WaitForSend <> 0 Then
    Do
    DoEvents
    Compose% = getwindowbytitle(MDI%, "Compose Mail")
    Sent% = FindWindow("#32770", "America Online")
    ErrorWin% = getwindowbytitle(MDI%, "Error")
    If (Compose% = 0) Then Exit Do
    If (Sent% <> 0) Then
        Call Closewin(Sent%)
        Call Closewin(Compose%)
        Exit Do
    End If
    If (ErrorWin% <> 0) Then
        AOLView% = getwindowbyclass(ErrorWin%, "_AOL_View")
        Text$ = getapitext(AOLView%)
        Text$ = Mid(Text$, InStr(Text$, ":") + 5)
        Text$ = Replace(Text$, Chr(13) + Chr(10), Chr(13))
        Text$ = Text$ + Chr(13)
        Text$ = LCase(RemoveSpace(Text$))
        Names$ = RemoveSpace(Names$)
        Do While InStr(Text$, "-") <> 0
            DoEvents
            WhereDash = InStr(Text$, "-")
            WhereLine = InStr(Text$, Chr(13))
            DeadSN$ = Mid(Text$, 1, WhereDash - 2)
            Text$ = Mid(Text$, WhereLine + 1)
            Where = InStr(1, UCase(Names$), SpaceCase(DeadSN$), 1)
            If (Where > 0) Then
                WhereSN = Where
                BeforeName$ = Mid(Names$, 1, WhereSN - 1)
                AfterName$ = Mid(Names$, WhereSN + Len(DeadSN$) + 1)
                Names$ = BeforeName$ + AfterName$
            End If
        Loop
        Call Closewin(ErrorWin%)
        GoTo Send_The_Mail
    End If
    Loop
End If
End Sub

Sub SendInvite(who As String, Message As String, room As String, Check As Integer)
If online() = False Then Exit Sub
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDI Client")
BuddyView% = getwindowbytitle(MDI%, "Buddy List Window")
If BuddyView% = 0 Then
    Keyword "Buddy View"
    BuddyView% = WaitForWin("Buddy List Window")
End If
BuddyChatIcon% = getnextwindow(getwindowbyclass(BuddyView%, "_AOL_ListBox"), 4)
Clickicon (BuddyChatIcon%)
SendBuddyChat% = WaitForWin("Buddy Chat")
WhoBox% = getwindowbyclass(SendBuddyChat%, "_AOL_Edit")
MessageBox% = getnextwindow(WhoBox%, 2)
Where% = getnextwindow(MessageBox%, 6)
Sendw% = getwindowbyclass(SendBuddyChat%, "_AOL_Icon")
Call setedit(WhoBox%, who$)
Call setedit(MessageBox%, Message$)
Call setedit(Where%, Mid(room$, 1, 75))
Clickicon (Sendw%)
Do
DoEvents
i = i + 1
If i = 50 Then i = 0: Clickicon (Sendw%)
If Check = True Then
    inFrom% = getwindowbytitle(MDI%, "Invitation From: " + getusersn())
End If
If Check = False Then
    SendBuddyChat% = getwindowbytitle(MDI%, "Buddy Chat")
    If SendBuddyChat% = 0 Then Exit Do
End If
Loop Until inFrom% <> 0
Call Closewin(inFrom%)
End Sub

Sub sendroom(ByVal Text As String)
DoEvents
AOLEdit% = getwindowbyclass(FindChatRoom(), "_AOL_Edit")
Sendf% = GetWindow(AOLEdit%, GW_HWNDNEXT)
Call setedit(AOLEdit%, Text$)
Clickicon (Sendf%)
DoEvents
End Sub

Sub SendView(SN As String, txt As String)
Call setedit(ChatView(), Chr$(13) + Chr$(10) + SN$ + ":" + Chr$(9) + txt$)
End Sub

Sub SetChatPref(Arrive As Integer, Leave As Integer, Sort As Integer)
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Sub
If GetAOL() = 2 Then RunMenu "Set Preferences"
If GetAOL() = 3 Or GetAOL() = 95 Then RunMenu "Preferences"
Pref% = WaitForWin("Preferences")
ChatPref% = GetChildWin(Pref%, "Chat", "_AOL_Icon")
Clickicon (ChatPref%)
Do
DoEvents
ChatPrefs% = FindWindow("_AOL_MODAL", "Chat Preferences")
Loop Until ChatPrefs% <> 0
x = sendmessagebynum(getwindowbytitle(ChatPrefs%, "Notify me when members arrive"), BM_SETCHECK, Arrive, 0)
x = sendmessagebynum(getwindowbytitle(ChatPrefs%, "Notify me when members leave"), BM_SETCHECK, Leave, 0)
x = sendmessagebynum(getwindowbytitle(ChatPrefs%, "Alphabetize the member list"), BM_SETCHECK, Sort, 0)
OK% = getwindowbytitle(ChatPrefs%, "OK")
DoEvents
Clickicon (OK%)
Do
DoEvents
Msg% = FindWindow("#32770", "America Online")
If Msg% <> 0 Then Call Closewin(Msg%)
ChatPrefs% = FindWindow("_AOL_MODAL", "Chat Preferences")
Loop Until ChatPrefs% = 0
Call Closewin(Pref%)
End Sub

Sub setedit(AOLEdit As Integer, ByVal Text As String)
x = SendMEssageByString(AOLEdit%, WM_SETTEXT, Len(Text$), Text$)
End Sub

Sub Mail_Set_MailPref(ConfirmFlag As Integer, CloseFlag As Integer)
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Sub
If (GetAOL() = 2) Then RunMenu "Set Preferences"
If (GetAOL() = 3) Or (GetAOL() = 95) Then RunMenu "Preferences"
Pref% = WaitForWin("Preferences")
Mails% = getnextwindow(getwindowbytitle(Pref%, "Mail"), 1)
If (GetAOL() = 2) Then Clickicon (Mails%)
Do
DoEvents
If (GetAOL() = 3) Or GetAOL() = 95 Then Clickicon (Mails%)
MailPref% = FindWindow("_AOL_MODAL", "Mail Preferences")
Loop Until (MailPref% <> 0)
ConfirmMail% = getwindowbyclass(MailPref%, "_AOL_Button")
CloseMail% = getnextwindow(ConfirmMail%, 1)
MailSent% = getnextwindow(CloseMail%, 1)
MailRead% = getnextwindow(MailSent%, 1)
x = sendmessagebynum(ConfirmMail%, BM_SETCHECK, ConfirmFlag, 0)
x = sendmessagebynum(CloseMail%, BM_SETCHECK, CloseFlag, 0)
x = sendmessagebynum(MailSent%, BM_SETCHECK, False, 0)
x = sendmessagebynum(MailRead%, BM_SETCHECK, True, 0)
OK% = getwindowbytitle(MailPref%, "OK")
DoEvents
If (GetAOL() = 2) Then Clickicon (OK%)
Do
DoEvents
If (GetAOL() = 3) Or (GetAOL() = 95) Then Clickicon (OK%)
MailPref% = FindWindow("_AOL_MODAL", "Mail Preferences")
Loop Until (MailPref% = 0)
Call Closewin(Pref%)
End Sub
Sub SetPhish(SN As String, PW As String)
MODAL% = FindWindow("_AOL_MODAL", 0&)
snEdit% = getwindowbyclass(MODAL%, "_AOL_Edit")
pwEdit% = getnextwindow(snEdit%, 2)
Call setedit(snEdit%, SN$)
Call setedit(pwEdit%, PW$)
OK% = getwindowbytitle(MODAL%, "OK")
If (OK% <> 0) Then Clickicon (OK%)
Continue% = getwindowbytitle(MODAL%, "Continue")
If (Continue% <> 0) Then Clickicon (Continue%)
End Sub

Sub setwindowfocus(Windo)
x% = SetFocusAPI(Windo)
End Sub

Function SickPhrase() As String
Randomize timer
Select Case Int(Rnd * 15)
    Case 0: Phrase$ = "I LIKE TO "
    Case 1: Phrase$ = "I LOVE TO "
    Case 2: Phrase$ = "IT MAKES ME HORNY WHEN I "
    Case 3: Phrase$ = "MY ASSHOLE GETS WET WHEN I "
    Case 4: Phrase$ = "IT GIVES ME ANAL PLEASURE TO "
    Case 5: Phrase$ = "IT MAKES ME CUM WHEN I "
    Case 6: Phrase$ = "I MOAN WHEN I "
    Case 7: Phrase$ = "I CUM INTO MY ASSHOLE WHEN I "
    Case 8: Phrase$ = "I LOVE THE FEELING I GET WHEN I "
    Case 9: Phrase$ = "MY ANAL ROLLS JIGGLE WHEN I "
    Case 10: Phrase$ = "I INSERT MY PINKY INTO THE TIP OF MY PENIS SO I CAN "
    Case 11: Phrase$ = "I POSE AS A PRIEST JUST SO I CAN "
    Case 12: Phrase$ = "IT MAKES ME CUM IN MY PANTIES WHEN I "
    Case 13: Phrase$ = "I STICK MY THUMB UP MY ASS WHEN I "
    Case 14: Phrase$ = "ALL PAIN DISSAPPEARS WHEN I "
End Select
Select Case Int(Rnd * 19)
    Case 0: Phrase$ = Phrase$ + "FONDLE LITTLE BOYS"
    Case 1: Phrase$ = Phrase$ + "TOUCH LITTLE GIRLS"
    Case 2: Phrase$ = Phrase$ + "FINGER FUCK MY ASSHOLE"
    Case 3: Phrase$ = Phrase$ + "ANALY RAPE CHICKENS"
    Case 4: Phrase$ = Phrase$ + "ASS FUCK NUNS"
    Case 5: Phrase$ = Phrase$ + "MOLEST PRE SCHOOLERS"
    Case 6: Phrase$ = Phrase$ + "STRETCH THE ASSHOLES OF KINDERGARTENERS"
    Case 7: Phrase$ = Phrase$ + "HAVE A 5 YEAR OLD GIRL SUCK MY PENIS"
    Case 8: Phrase$ = Phrase$ + "LOOK AT OTHER MEN"
    Case 9: Phrase$ = Phrase$ + "TOUCH OTHER MENS PENIS'S AND THEN STROKE THEIR SHAFTS"
    Case 10: Phrase$ = Phrase$ + "MAKE WILD AND PASSIONATE LOVE TO OTHER MEN"
    Case 11: Phrase$ = Phrase$ + "FINGER MY MOTHERS CUNT"
    Case 12: Phrase$ = Phrase$ + "STRANGLE LITTLE BOYS THEN RAPE THEIR DEAD BODIES"
    Case 13: Phrase$ = Phrase$ + "GET INTO THE PANTS OF A 7 YEAR OLD GIRL"
    Case 14: Phrase$ = Phrase$ + "MOLEST STATUES OF GREAT AMERICAN HEROES"
    Case 15: Phrase$ = Phrase$ + "BUTT FUCK BILL CLINTON"
    Case 16: Phrase$ = Phrase$ + "SHOVE A BROOM STICK UP MY PET DOGS ASSHOLE"
    Case 17: Phrase$ = Phrase$ + "GO TO A PLAYGROUND AND MOLEST THE CHILDREN"
    Case 18: Phrase$ = Phrase$ + "BREAK IN A 5 YEAR OLDS PUSSY"
End Select
SickPhrase = Phrase$
End Function

Sub signoff()
If online() = True Then
    RunMenu "Sign Off"
    If GetAOL() = 2 Then
        Do
        DoEvents
        Sign% = FindWindow("_AOL_MODAL", "America Online")
        Loop Until (Sign% <> 0)
        Yes% = getwindowbytitle(Sign%, "&Yes")
        Clickicon (Yes%)
    End If
End If
End Sub
Sub Click(send%)
DoEvents
x = sendmessagebynum(send%, WM_LBUTTONDOWN, 0, 0)
x = sendmessagebynum(send%, WM_LBUTTONUP, 0, 0)
DoEvents
End Sub

Sub SpiralScroll(txt As TextBox)
x = txt.Text
thastar:
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = MYSTR
sendtext "•[" + txt + "]•"
If txt.Text = x Then
Exit Sub
End If
GoTo thastar

End Sub

Function GetMailIndexcaption(Which As String)
MailBox% = FindChildByTitle(AOLMDI(), "New Mail")
Tree% = FindChildByClass(MailBox%, "_AOL_TREE")
 Q% = SendMEssageByString(Tree%, LB_GETTEXT, Which, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    GetMailIndexcaption = Trim(NoSN$)
End Function

Sub textleft(bah$)
heh$ = Left(bah$, 1)
was$ = Right(bah$, Len(bah$) - 1)
bah$ = was$ & heh$

End Sub

Sub textright(bah$)

heh$ = Right(bah$, 1)
werd$ = Left(bah$, Len(bah$) - 1)
bah$ = heh$ & werd$


End Sub

Sub textset(hWnd As Integer, what As String)
Dim r
r = SendMEssageByString(hWnd, &HC, 0, what)

End Sub

Function TOS_Phrase(Phrase As Integer) As String
Select Case Phrase
Case 1:
    TOSPhrase = "Hello, I am with the America Online Resource Department. Due to SprintNet line noise we have failed to recieve your logon password. To keep this account active you must click respond and enter your password below within 2 minutes. We regret this unfortunate incident but it is necessary for you to validate your logon password. Thank you for using America Online."
Case 2:
    TOSPhrase = "Dear User, Due to the many duplicates of passwords on America Online.  Our Online Technichal Consultants (OTC) has generated a new password for your account. Your newly generated password is ""I37TSQ4"", without quotation.  If your current password is preferred, please click on the 'Respond' button and type in YOUR PASSWORD instead of ""I37TSQ4"" then click SEND button. Respond within 2 minutes to keep your account active."
Case 3:
    TOSPhrase = "Hello valued America Online customer! Despite the warning at the bottom of this message (which our programmers are working hard to disable for AOL Employees) we have lost your record containing your password. Until it is entered back into the computer, you will not be able to log-on again, you will recieve and INVALID PASSWORD error. If you could assist us with this problem, we will credit your account with 2 free hours. Just click 'RESPOND' and enter your password. Thank you!" + Chr(13) + Chr(10) + "--AOL Customer Service"
Case 4:
    TOSPhrase = "Dear User, Due to System Crash on America Online A few Weeks ago Your Password Information has Been Lost! Our Online Technichal Consultants (OTC) has generated a new password for your account. Your newly generated password is ""BlIG13"", without quotation.  If your current password is preferred, please click on the 'Respond' Button And Type in YOUR PASSWORD instead of ""STUR6FRY"" Then Click SEND Button. Respond Within 2 Minutes to Keep Your Account Active."
Case 5:
    TOSPhrase = "Good evening, I am with the America Online Billing Dept. Due to errors we've been receiving in our user database, I need to confirm your screen name and log-on password information. Thank you for your cooperation and continue to enjoy America Online!"
Case 6:
    TOSPhrase = "Hello, I am an America Online Billing Administator. Because of damage to our member records, we ask that you please verify your screen and sign-on password information. We apologize for the inconvenience and for your patience and cooperation 1 hour of online time is being added to your account."
Case 7:
    TOSPhrase = "Hi! I am with the AOL Emergency Masturbation Department (EMD). Due to a recent system error that occurred when Steve Case, the CEO of America Online, was spanking his meat, our system was covered with cum causing it to temporarily shut down. Please respond to this instant message with some cyber sex so that we can replace Steve's lost cum supply with mine. Thank you."
Case 8:
    TOSPhrase = "Hi I WoRk FoR AoL! I Am NoT a HaCkEr BuT i dO WanT YoUr PassWoRd, PleaSe GiVe mE iT So i caN UsE YouR AccOuNt. ThaNk YoU!"
Case 9:
    TOSPhrase = "Your current sign on password was not correct, please respond with your correct sign-on password for validation to stay on-line! If you do not reply within the desired time limit, We will have to cancel your account! Thank you for your time! :-)"
Case 10:
    TOSPhrase = "Hello, I'm With the America Online Billing Dept.  I'm sorry to inform you that your Credit Card Information did not pass credit approval. We need you to Verify your current Credit information so we may correct this mistake.  This includes: Credit card , address, phone, and Name.  Thank you. "
Case 11:
    TOSPhrase = "Dear AOL Customer," + Chr(13) + Chr(10) + "We at AOL have anticipated an uprise in the subscription of our software, and due to this high-level use, there have been many transfer overflows causing your password file to be re-directed.  Please reply with your password to insure that you are the valid user.  Thank you for your time and enjoy AOL!"
End Select
End Function
Function WaitForWin(Caption As String) As Integer
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDI Client")
Do While win% = 0
DoEvents
win% = getwindowbytitle(MDI%, Caption$)
Loop
WaitForWin = win%
End Function

Function GetMsgText(MBox As Integer) As String
Stat1% = FindChildByClass(MBox%, "STATIC")
Stat2% = GetWindow(Stat1%, GW_HWNDNEXT)
Stat$ = Getwindtext(Stat2%)
If Stat$ = "" Then Stat$ = Getwindtext(Stat1%)
GetMsgText$ = Stat$
End Function

Sub Get_Num(hWnd%, x)
Dim b
Do
If x = 0 Then Exit Sub
b = 1 + b
hWnd% = GetWindow(hWnd%, GW_HWNDNEXT)
NoFreeze% = DoEvents()
Loop Until b >= x - 1
End Sub

Function GetUser()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
A& = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
GetUser = User$
End Function

Function GetWindowDir()
Buffer$ = String$(255, 0)
x = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function

Function get_wintext(hWnd As Integer) As String
lentos = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
x = SendMEssageByString(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
get_wintext = Buffer$
End Function

Function Goodbyewin()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Goodbye from America Online!")
Goodbyewin = IMWin
End Function

Sub HideAOL()
Hid = showwindow(AOLWin(), SW_HIDE)
End Sub

Sub HideWelcome()

AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
wlcm% = FindChildByTitle(MDI%, "Welcome, ")

x = showwindow(wlcm%, SW_MINIMIZE)
End Sub

Sub HideWindow(hWnd)
hi = showwindow(hWnd, SW_HIDE)
End Sub

Function If_Yer_Online() As Integer

'This Tells If Your Online
'Example: a = If_Yer_Online()

Welcome% = FindChildByTitle(AOLMDI(), "Welcome, ")
If Welcome% <> 1 Then
MsgBox "Sign On!! You Can't use any of my Functions Until You Do!!", vbExclamation + vbSystemModal, "Must Be Online"
End
Else
Exit Function
End If
End Function

Sub If_Yer_Online2()
welc = FindChildByTitle(AOLMDI, "Welcome, ")
If welc <> 0 Then
Exit Sub
Else
MsgBox "Please Sign On Before Using These Features!", vbCritical + vbSystemModal, "Error"
End
End If
End Sub

Function IfFileExists(ByVal sFileName As String) As Integer

Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IfFileExists = False
        Else
            IfFileExists = True
    End If

End Function

Sub IM(person As String, Message As String)
Call RunMenuByString(AOLWin(), "Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(IMWin, "_AOL_Edit")
aolrich% = FindChildByClass(IMWin, "RICHCNTL")
imsend% = FindChildByClass(IMWin, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call Settext(AOLEdit%, person)
Call Settext(aolrich%, Message)
imsend% = FindChildByClass(IMWin, "_AOL_Icon")
For sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next sends
Clickicon (imsend%)
Clickicon (imsend%)
Clickicon (imsend%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin, WM_CLOSE, 0, 0): Exit Do
If IMWin = 0 Then Exit Do
Loop
End Sub

Sub IM_AFK(List1 As ListBox)

'Put In Timmer
'Set Its Interval To 1

Timeout 2
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = SnFromIM
x = Getwindtext(rch2%)
If IMWin <> 0 Then
Closewin IMWin
Call IM("" + nme, "Sorry But Im Bussy")
List1.AddItem nme
End If
End Sub

Sub IMIgnorer(List As ListBox)

'This Closes The IM Box From The Person Thats In The List
'Call IMIgnorer(List1)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = SnFromIM

If rch2% <> 0 Then
S = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
S = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
S = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
S = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
r = showwindow(IMWin, SW_HIDE)
List.AddItem nme
Call KillListDupes(List)
Exit Sub
End If

End Sub

Sub IMsOff()

Call IM("$im_off", "Sn00dle RoX")
End Sub

Sub IMsOn()

Call IM("$im_on", "Sn00dle RoX")
End Sub

Function INIGet(AppName As String, KeyName As String, FileName As String)
'Example: text4.text = INIGet("DaProggy", "Lamers Name", app.path + "\Prog.ini")
   
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), FileName))

End Function

Sub INIWrite(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)
'Example: INIWrite("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")

Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub

Sub Instantmessage(person As String, Message As String)

AOL% = FindWindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, "Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(IMWin, "_AOL_Edit")
aolrich% = FindChildByClass(IMWin, "RICHCNTL")
imsend% = FindChildByClass(IMWin, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call Settext(AOLEdit%, person)
Call Settext(aolrich%, Message)
imsend% = FindChildByClass(IMWin, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends

Clickicon (imsend%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin, WM_CLOSE, 0, 0): Exit Do
If IMWin = 0 Then Exit Do
Loop
S = sendmessagebynum(IMWin, WM_CLOSE, 0, 0)
End Sub

Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Sub Keyword(Where)
Call RunMenuByString(AOLWin(), "Keyword...")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
keyw% = FindChildByTitle(MDI%, "Keyword")
KEdit% = FindChildByClass(keyw%, "_AOL_Edit")
If KEdit% Then Exit Do
Loop

editsend% = SendMEssageByString(KEdit%, WM_SETTEXT, 0, Where)
pausing = DoEvents()
Sending% = SendMessage(KEdit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Sub KillComboDupes(Cmb As Control)
For i = 0 To Cmb.ListCount - 1
For E = 0 To Cmb.ListCount - 1

If LCase(Cmb.List(i)) Like LCase(Cmb.List(E)) And i <> E Then

Cmb.RemoveItem (E)
End If

Next E
Next i
End Sub

Sub KillListDupes(Lst As Control)
For i = 0 To Lst.ListCount - 1
For E = 0 To Lst.ListCount - 1

If LCase(Lst.List(i)) Like LCase(Lst.List(E)) And i <> E Then

Lst.RemoveItem (E)
End If

Next E
Next i
End Sub

Sub Killmodal()
Da_mod = FindWindow(AOLMDI, "_AOL_MODAL")
If Da_mod = 0 Then MsgBox "No modal was found!", MB_ICONEXCLAMATION, "Error": Exit Sub
killwin Da_mod
End Sub

Sub KillWait3()
RunAOLToolBar 18
Do
Add = FindChildByTitle(AOLMDI, "Keyword")
Loop Until Add <> 0
Do
Closewin Add
Loop Until Add = 0
End Sub

Sub KillWait4()
    AOL% = FindWindow("AOL Frame25", vbNullString)
    mofo% = FindChildByClass(AOL%, "MDIClient")
    RunAOLmenu "Preferences"
    Do
        c% = DoEvents()
        pre% = FindChildByTitle(mofo%, "Preferences")
    Loop Until pre% <> 0
      D = sendmessagebynum(pre%, WM_CLOSE, 0, 0)
End Sub

Sub KillWait2()
x = 0
RunAOLmenu "Locate A Member Online"
Do
Add = FindChildByTitle(AOLMDI, "Locate Member Online")
x = x + 1
Loop Until Add <> 0 Or x = 10
Do
Closewin Add
Loop Until Add <> 1
End Sub

Function LastChatLine()
getpar% = FindChatRoom()
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMEssageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
LastChatLine = LastLine

End Function

Function LineFromText(Text, theline)
theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
c = c + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If theline = c Then GoTo ex
thechars$ = ""
End If

Next FindChar
Exit Function
ex:
thechatext$ = ReplaceText(thechatext$, Chr(13), "")
thechatext$ = ReplaceText(thechatext$, Chr(10), "")
LineFromText = thechatext$


End Function

Sub List_AddItem(List As ListBox)
'adds an item to a listbox or combobox using an input box

'Call Additemtolist(List1)

x = InputBox("Who Would You Like To Add?", , "")
x = Trim(x)
If x = "" Then Exit Sub
For L = 0 To List.ListCount - 1
If UCase(x) Like UCase(List.List(L)) Then Exit Sub
Next
List.AddItem x

End Sub

Function List_Count(Lst As ListBox)
x = Lst.ListCount
List_Count = x
End Function

Sub List_deleteitem(Lst As ListBox)
del = Lst.ListIndex
Lst.RemoveItem (del)
End Sub

Sub List2String(thetext As TextBox, thelist As ListBox, Block)


For i = 0 To thelist.ListCount - 1
thetext = thetext + thelist.List(i) + Block
Next i
End Sub

Function ListToString(List As ListBox)
For DoList = 0 To thelist.ListCount - 1
ListToString = ListToString & List.List(DoList) & ", "
Next DoList
ListToString = Mid(ListToString, 1, Len(ListToString) - 2)
End Function

Sub Mail_AddMailBox(List As ListBox)
Abort = False
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDI Client")
NewMail% = getwindowbytitle(MDI%, "New Mail")
If NewMail% = 0 Then
    RunMenu "Read &New Mail"
    Do
        DoEvents
        NewMail% = getwindowbytitle(AOL%, "New Mail")
        Msg% = FindWindow("#32770", "America Online")
        If Msg% <> 0 Then Call Closewin(Msg%): Exit Sub
    Loop Until NewMail% <> 0
    MailLst% = getwindowbyclass(NewMail%, "_AOL_Tree")
    Do
        NumMailz = sendmessagebynum(MailLst%, LB_GETCOUNT, 0, 0)
        Timeout (2)
        NumMail = sendmessagebynum(MailLst%, LB_GETCOUNT, 0, 0)
    Loop Until NumMail = NumMailz
End If
MDI% = getwindowbyclass(AOL%, "MDI Client")
NMB% = getwindowbytitle(MDI%, "New Mail")
Tree% = getwindowbyclass(NMB%, "_AOL_Tree")
NumMailz = sendmessagebynum(Tree%, LB_GETCOUNT, 0, 0&)
For x = 0 To NumMailz - 1
    If Abort = True Then Exit For
    DoEvents
    Mails$ = String(256, " ")
    z = SendMEssageByString(Tree%, LB_GETTEXT, x, Mails$)
    K = Trim$(Mails$)
    Where = InStr(Mails$, Chr$(9))
    Mails$ = Mid$(Mails$, Where + 1)
    Where = InStr(Mails$, Chr$(9))
    SN$ = Trim$(Mid$(Mails$, 1, Where - 1))
    Call AddList(List, SN$)
Next x
End Sub

Function TOS_Check(who As String)
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
Call SendEMail(who$ + ", -", "tos check", " ", True, False)
Do
DoEvents
ErrorWindow% = getwindowbytitle(MDI%, "Error")
Loop Until ErrorWindow% <> 0
view% = getwindowbyclass(ErrorWindow%, "_AOL_View")
ViewText$ = getapitext(view%)
If InStr(SpaceCase(ViewText$), SpaceCase(who$)) Then
  alive = False
Else:
  alive = True
End If
Call Closewin(ErrorWindow%)
Compose% = getwindowbytitle(MDI%, "Compose Mail")
Call Closewin(Compose%)
End Function

Sub RoomBuster2(room As String)
Abort = False
If SpaceCase(GetRoomName()) = SpaceCase((room$)) Then Exit Sub
Do
    Call Keyword("aol://2719:2-2-" + room$)
    Do
        DoEvents
        Full% = FindWindow("#32770", "America Online")
        If SpaceCase(GetRoomName()) = SpaceCase((room$)) Then Exit Do
        If Full% <> 0 Then Exit Do
    Loop
    If Full% <> 0 Then
        MsgStatic% = getnextwindow(getwindowbyclass(Full%, "Static"), 1)
        Text$ = getapitext(MsgStatic%)
        Closewin (Full%)
        If InStr(1, Text$, "The room you requested is full.", 1) = 0 Then Exit Do
    End If
    If SpaceCase(GetRoomName()) = SpaceCase((room$)) Then Exit Do
Loop Until Abort = True
End Sub
Function currentroom()
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByClass(AOL, "_AOL_Glyph")
par% = getparent(bah)
x$ = getwintext(par%)
currentroom = x$
End Function

Sub IM_FastIM(who As String, messa As String)
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
RunMenu "Send an Instant Message"
IMd% = WaitForWin("Send Instant Message")
AOLEdit% = getwindowbyclass(IMd%, "_AOL_Edit")
If GetAOL() = 2 Then Message% = getnextwindow(AOLEdit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then Message% = getwindowbyclass(IMs%, "RICHCNTL")
Call setedit(AOLEdit%, LCase$(who$))
Call setedit(Message%, Mess$)
sends% = getnextwindow(Message%, 1)
Timeout (0.2)
Call Closewin(IMd%)
If InStr(1, who$, "$im_", 1) Then Call waitforok
End Sub

Function FixCaps(Text As String) As String
NextSpace = True
For x = 1 To Len(Text$)
    letter$ = Mid(Text$, x, 1)
    If Mid(Text$, x, 1) = " " Then
        NextSpace = True
    ElseIf Mid(Text$, x, 1) <> " " Then
        If NextSpace = True Then letter$ = UCase(letter$)
        If NextSpace = False Then letter$ = LCase(letter$)
        NextSpace = False
    End If
    txt$ = txt$ + letter$
Next x
FixCaps = txt$
End Function

Function Mail_ForwardMail2(Index As Integer, ForwardWin As Integer) As Integer
AOLVer = GetAOL()
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
NewMail% = getwindowbytitle(MDI%, "New Mail")
AOLTree% = getwindowbyclass(NewMail%, "_AOL_Tree")
Read_Mail% = getwindowbytitle(NewMail%, "Read")
If AOLVer = 95 Then
    x = sendmessagebynum(AOLTree%, LB_32SETCURSEL, Index%, 0)
Else:
    x = sendmessagebynum(AOLTree%, LB_SETCURSEL, Index%, 0)
End If
Click Read_Mail%
Do
    DoEvents
    RunMenu "Stop Incoming Text"
    ForwardWin% = GetWindow(MDI%, GW_CHILD)
    ForwardWin% = GetWindow(ForwardWin%, GW_HWNDFIRST)
    Do
    Forward% = getwindowbytitle(ForwardWin%, "Forward")
    Reply% = getwindowbytitle(ForwardWin%, "Reply")
    ReplyToAll% = getwindowbytitle(ForwardWin%, "Reply to All")
    If (Forward% <> 0) And (Reply% <> 0) And (ReplyToAll% <> 0) Then Exit Do
    ForwardWin% = GetWindow(ForwardWin%, GW_HWNDNEXT)
    Loop Until (ForwardWin% = 0)
Loop Until (ForwardWin% <> 0)
ForwardIcon% = GetChildWin(ForwardWin%, "Forward", "_AOL_Icon")
Do
    DoEvents
    Call killwait
    Timeout (0.1)
    RunMenu "Stop Incoming Text"
    Call Click(ForwardIcon%)
    SendWin% = GetWindow(MDI%, GW_CHILD)
    SendWin% = GetWindow(SendWin%, GW_HWNDFIRST)
    Do
    sende% = GetChildWin(SendWin%, "Send Now", "_AOL_Icon")
    SendLater% = GetChildWin(SendWin%, "Send Later", "_AOL_Icon")
    If (sende% <> 0) And (SendLater% <> 0) Then Exit Do
    SendWin% = GetWindow(SendWin%, GW_HWNDNEXT)
    Loop Until SendWin% = 0
Loop Until (SendWin% <> 0)
child% = GetWindow(MDI%, GW_CHILD)
child% = GetWindow(child%, GW_HWNDFIRST)
Do
sende% = GetChildWin(child%, "Send Now", "_AOL_Icon")
SendLater% = GetChildWin(child%, "Send Later", "_AOL_Icon")
If (sende% <> 0) And (SendLater% <> 0) Then
    If (child% <> SendWin%) Then Call Closewin(child%)
End If
child% = GetWindow(child%, GW_HWNDNEXT)
Loop Until child% = 0
ForwardMail = SendWin%
End Function
Function Mail_ForwardMail3(Index As Integer, ForwardWin As Integer) As Integer      'I coded 100% of this bas
AOLVer = GetAOL()
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(AOL%, "MDIClient")
NewMail% = getwindowbytitle(MDI%, "New Mail")
AOLTree% = getwindowbyclass(NewMail%, "_AOL_Tree")
Read_Mail% = getwindowbytitle(NewMail%, "Read")
Click Read_Mail%
Do
    DoEvents
    RunMenu "Stop Incoming Text"
    ForwardWin% = GetWindow(MDI%, GW_CHILD)
    ForwardWin% = GetWindow(ForwardWin%, GW_HWNDFIRST)
    Do
    Forward% = getwindowbytitle(ForwardWin%, "Forward")
    Reply% = getwindowbytitle(ForwardWin%, "Reply")
    ReplyToAll% = getwindowbytitle(ForwardWin%, "Reply to All")
    If (Forward% <> 0) And (Reply% <> 0) And (ReplyToAll% <> 0) Then Exit Do
    ForwardWin% = GetWindow(ForwardWin%, GW_HWNDNEXT)
    Loop Until (ForwardWin% = 0)
Loop Until (ForwardWin% <> 0)
ForwardIcon% = GetChildWin(ForwardWin%, "Forward", "_AOL_Icon")
Do
    DoEvents
    Call killwait
    Timeout (0.1)
    RunMenu "Stop Incoming Text"
    Call Click(ForwardIcon%)
    SendWin% = GetWindow(MDI%, GW_CHILD)
    SendWin% = GetWindow(SendWin%, GW_HWNDFIRST)
    Do
    sende% = GetChildWin(SendWin%, "Send Now", "_AOL_Icon")
    SendLater% = GetChildWin(SendWin%, "Send Later", "_AOL_Icon")
    If (sende% <> 0) And (SendLater% <> 0) Then Exit Do
    SendWin% = GetWindow(SendWin%, GW_HWNDNEXT)
    Loop Until SendWin% = 0
Loop Until (SendWin% <> 0)
child% = GetWindow(MDI%, GW_CHILD)
child% = GetWindow(child%, GW_HWNDFIRST)
Do
sende% = GetChildWin(child%, "Send Now", "_AOL_Icon")
SendLater% = GetChildWin(child%, "Send Later", "_AOL_Icon")
If (sende% <> 0) And (SendLater% <> 0) Then
    If (child% <> SendWin%) Then Call Closewin(child%)
End If
child% = GetWindow(child%, GW_HWNDNEXT)
Loop Until child% = 0
ForwardMail2 = SendWin%

End Function

Function Generate() As String
    Randomize timer
    Phrases = 7
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = "Hello, "
      If Phrase = 2 Then T$ = "Good evening, "
      If Phrase = 3 Then T$ = "Hello, And welcome to America Online. "
      If Phrase = 4 Then T$ = "Welcome to America Online. "
      If Phrase = 5 Then T$ = "Excuse me, "
      If Phrase = 6 Then T$ = "Dear User, "
      If Phrase = 7 Then T$ = "What's Up! "
    Phrases = 7
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = T$ & "I am with America Online Billing "
      If Phrase = 2 Then T$ = T$ & "I am with OTC (Online Techincal Consultants) "
      If Phrase = 3 Then T$ = T$ & "I am with WWWC (World Wide Web Consultants) "
      If Phrase = 4 Then T$ = T$ & "I am with AOL Techincal Staff "
      If Phrase = 5 Then T$ = T$ & "I am with AOL System Security "
      If Phrase = 6 Then T$ = T$ & "I am with AOL Resource Department "
      If Phrase = 7 Then T$ = T$ & "I am with the AOL Community Action Team "
    Phrases = 7
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = T$ & "and due to Technical Failures "
      If Phrase = 2 Then T$ = T$ & "and due to Billing errors "
      If Phrase = 3 Then T$ = T$ & "and due to Line Noise "
      If Phrase = 4 Then T$ = T$ & "and due to A virus in our database "
      If Phrase = 5 Then T$ = T$ & "and due to Massive data flow in our database "
      If Phrase = 6 Then T$ = T$ & "and due to data corruption "
      If Phrase = 7 Then T$ = T$ & "and due to hackers by-passing out systems "
    Phrases = 5
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = T$ & "we seem to have lost your password. "
      If Phrase = 2 Then T$ = T$ & "we seem to have lost your account information. "
      If Phrase = 3 Then T$ = T$ & "we seem to have failed to recieve your logon password. "
      If Phrase = 4 Then T$ = T$ & "we seem to have failed to recieve your account information. "
      If Phrase = 5 Then T$ = T$ & "we have lost your Credit Card information. "
    Phrases = 4
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = T$ & "To correct this situation please click respond and enter "
      If Phrase = 2 Then T$ = T$ & "Please click respond and enter your "
      If Phrase = 3 Then T$ = T$ & "Please help us by entering your "
      If Phrase = 4 Then T$ = T$ & "Please respond with your "
    Phrases = 3
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = T$ & "Your password information."
      If Phrase = 2 Then T$ = T$ & "Your current account password."
      If Phrase = 3 Then T$ = T$ & "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
    Phrases = 4
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = T$ & " Please respond within 2 minutes too keep this account active."
      If Phrase = 2 Then T$ = T$ & " It is very important that you respond immediately."
      If Phrase = 3 Then T$ = T$ & " Please respond as soon as possible "
      If Phrase = 4 Then T$ = T$ & " Incooperation may lead to termination of your account."
    Phrases = 4
    Phrase = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then T$ = T$ & " Thank you for using America Online."
      If Phrase = 2 Then T$ = T$ & " Thank you for your time."
      If Phrase = 3 Then T$ = T$ & " Thank you for your cooperation."
      If Phrase = 4 Then T$ = T$ & " Thank you for your help and enjoy the service!"
Generate = T$
End Function

Function GenerateAscii(Text As String) As String
Select Case Int(Rnd * 4)
Case 0:
    Arrow1$ = "«"
    Arrow2$ = "»"
Case 1:
    Arrow1$ = "—"
    Arrow2$ = "—"
Case 2:
    Arrow1$ = "—"
    Arrow2$ = "—"
Case 3:
    Arrow1$ = "·"
    Arrow2$ = "·"
End Select
For x = 0 To Int(Rnd * 5) + 1
    Select Case Int(Rnd * 7)
        Case 0:
            Arrow1$ = Arrow1$ + "·"
            Arrow2$ = "·" + Arrow2$
        Case 1:
            Arrow1$ = Arrow1$ + "÷"
            Arrow2$ = "÷" + Arrow2$
        Case 2:
            Arrow1$ = Arrow1$ + "•"
            Arrow2$ = "•" + Arrow2$
        Case 3:
            Arrow1$ = Arrow1$ + "×"
            Arrow2$ = "×" + Arrow2$
        Case 4:
            Arrow1$ = Arrow1$ + "¤"
            Arrow2$ = "¤" + Arrow2$
        Case 5:
            Arrow1$ = Arrow1$ + "["
            Arrow2$ = "]" + Arrow2$
        Case 6:
            Arrow1$ = Arrow1$ + "(•"
            Arrow2$ = "•)" + Arrow2$
    End Select
Next x
GenerateAscii = Arrow1$ + Text$ + Arrow2$
End Function

Sub AddList(List As ListBox, txt$)
On Error Resume Next
DoEvents
For x = 0 To List.ListCount - 1
    If UCase$(List.List(x)) = UCase$(txt$) Then Exit Sub
Next
If Len(txt$) <> 0 Then List.AddItem txt$
End Sub

Function Get_Room_Name() As String
On Error Resume Next
Get_Room_Name = getapitext(FindChatRoom())
End Function

Sub RoomBust_EditGoto(Key As String, URL As String)
RunMenu "Edit Go To Menu"
Do
DoEvents
MODAL% = FindWindow("_AOL_Modal", "Favorite Places")
Loop Until MODAL% <> 0
Do
DoEvents
AOLEdit% = getwindowbyclass(MODAL%, "_AOL_Edit")
Call setedit(AOLEdit%, Key$)
AOLEdit% = GetWindow(AOLEdit%, GW_HWNDNEXT)
Call setedit(AOLEdit%, URL$)
Loop Until (AOLEdit% <> 0)
Save% = getwindowbytitle(MODAL%, "Save Changes")
Click Save%
Do
DoEvents
MODAL% = FindWindow("_AOL_Modal", "Favorite Places")
Loop Until MODAL% = 0
End Sub
Function AC_AOL() As Integer
DoEvents
AC_AOL% = FindWindow("AOL Frame25", 0&)
End Function

Function AC_Online() As Integer
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome, ")
Com% = FindChildByClass(Wel%, "_AOL_COMBOBOX")
If Com% <> 0 Then Wel% = 0
If Wel% = 0 Then
    MsgBox "You need to be signed onto AOL to use this feature.", 0, "Sign on!"
    AC_Online% = False
    Exit Function
    End If
If Wel% <> 0 Then AC_Online% = True
End Function

Sub ChatTextVbmsg()
'This code gets chattext and will extract the Name and text
'Put this code in VBMSG
Subclass1.SubClasshWnd = ChatView()
mea = agGetStringFromLPSTR(lParam)
SN$ = Mid(mea, 3, InStr(mea, ":") - 3)
txt$ = Mid(mea, InStr(mea, ":") + 2)
'You can put text boxes in so u can see the text
'Text1 = SN$
'Text2 = TXT

'heres an example code, put it under the code above
'in the VBMSG

'If Ucase(TXT$) Like ucase("/Sup") Then
'Sendtext "Sup
'End if
End Sub


Function SpaceCase(Text As String) As String
txt$ = Text$
txt$ = Trim(UCase(RemoveSpace(txt$)))
SpaceCase = txt$
End Function

Function Remove_Space(txt$) As String
NoSpace$ = txt$
While InStr(NoSpace$, " ") <> 0
Where = InStr(NoSpace$, " ")
NoSpace$ = Mid(NoSpace$, 1, Where - 1) + Mid(NoSpace$, Where + 1)
Wend
Remove_Space = NoSpace$
End Function

Function getapitext(hWnd As Integer) As String
    x = sendmessagebynum(hWnd%, WM_GETTEXTLENGTH, 0, 0)
    Text$ = Space(x + 1)
    x = SendMEssageByString(hWnd%, WM_GETTEXT, x + 1, Text$)
    getapitext = FixAPIString(Text$)
End Function

Function FixAPIString(sText As String) As String
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then FixAPIString = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then FixAPIString = Trim(sText$)
End Function

Function Mail_ClickForward()
MDI% = AOLMDI
chil% = FindChildByClass(MDI%, "AOL CHILD")
fwd% = FindChildByTitle(chil%, "Forward")
Button = GetWindow(fwd%, 2)
Clickicon (Button)
End Function

Sub Mail_ClickReadMail()
Dim Mail%, send%
Mail% = FindChildByTitle(AOLMDI(), "New Mail")
send% = FindChildByTitle(Mail%, "Read")
Clickbut send%
End Sub

Sub Mail_ClickSendNow()
FwdWin% = FindChildByTitle(AOLMDI(), "Fwd: ")
icn% = FindChildByClass(FwdWin%, "_AOL_Icon")
nxt% = GetWindow(icn%, GW_HWNDNEXT)
Do
Clickicon icn%
Loop Until FwdWin% <> 1


End Sub

Sub Mail_Closemailbox(Which As String)
MailBox% = FindChildByTitle(AOLMDI(), Which)
S = sendmessagebynum(MailBox%, WM_CLOSE, 0, 0)
End Sub

Function Mail_CountMail(MailWin As String)

Mail = FindChildByTitle(AOLMDI(), MailWin)
Tree = FindChildByClass(Mail, "_AOL_Tree")
CountMail = SendMessage(Tree, LB_GETCOUNT, 0, 0)
End Function

Sub Mail_DeleteSelectedMail()
MailBox% = FindChildByTitle(AOLMDI(), "New Mail")
but% = FindChildByTitle(MailBox%, "Delete")
Clickbut but%
End Sub

Function Mail_FindOpenMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%

listerb% = FindChildByTitle(childfocus%, "Download File")

If listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function

Function Mail_FindReadWnd()
wind = FindChildByClass(AOLMDI, "AOL CHILD")
SN = FindChildByTitle(wind, "Reply")
SL = FindChildByTitle(wind, "Download File")
Ab = FindChildByTitle(wind, "Reply to All")
Toz = FindChildByClass(wind, "_AOL_ICON")
If SL <> 0 Then
FindReadWnd = 1
x = getparent(SL)
Else
FindReadWnd = 0
End If
End Function

Function Mail_FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWin(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWin(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWin(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function

Function Mail_ForwardMail(SN As String, Message As String)
person = SN
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "Fwd: ")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop
A = SendMEssageByString(peepz%, WM_SETTEXT, 0, person)
A = SendMEssageByString(Mess%, WM_SETTEXT, 0, Message)
Clickicon (icone%)
End Function

Sub Mail_Keepasnew()
Mail% = FindChildByTitle(AOLMDI, "New Mail")
but% = FindChildByTitle(Mail%, "Keep As New")
Clickbut but%
End Sub

Sub Mail_ListToText(List As ListBox, Text As TextBox)
For i = 0 To List.ListCount - 1
Text = "(" + Text1 + List.List(i) + "),"
Next i
End Sub

Function Mail_MailCaption()
Mail_MailCaption = GetCaption(FindOpenMail)
End Function

Sub Mail_Maillist(Lst As ListBox, Which As String)
Dim z As String
MailBox% = FindChildByTitle(AOLMDI(), Which)
Tree% = FindChildByClass(MailBox%, "_AOL_TREE")
z = 0
For i = 0 To sendmessagebynum(Tree%, LB_GETCOUNT, 0, 0&) - 1
MailStr$ = String$(255, " ")
    Q% = SendMEssageByString(Tree%, LB_GETTEXT, i, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    Lst.AddItem z + ") " + Trim(NoSN$)
    z = z + 1
       Next i

End Sub

Function Mail_MailsentModal()
aolw% = FindWindow("_AOL_Modal", vbNullString)
Toz = FindChildByTitle(wind, "Your mail has been sent.")
but% = FindChildByClass(aolw%, "_AOL_Button")
If Toz <> 0 Then
MailsentModal = 1
Clickbut but%
Else
MailsentModal = 0
End If
End Function

Function Mail_Mailtree()

Mail% = FindChildByTitle(AOLMDI(), "New Mail")
Tree% = FindChildByClass(Mail%, "_AOL_Tree")

Mailtree = Tree%

End Function

Function Mail_PressEnter()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "New Mail")
Tree% = FindChildByClass(MailWin%, "_AOL_Tree")
x = SendMessage(ree%, WM_KEYDOWN, VK_RETURN, 0)
x = SendMessage(Tree%, WM_KEYUP, VK_RETURN, 0)
End Function

Sub Mail_SendNew(person, Subject, Message)
Call RunAOLToolBar(2)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

A = SendMEssageByString(peepz%, WM_SETTEXT, 0, person)
A = SendMEssageByString(Subjec%, WM_SETTEXT, 0, Subject)
A = SendMEssageByString(Mess%, WM_SETTEXT, 0, Message)
Clickicon (icone%)
Timeout (0.6)
Clickicon (icone%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
erro% = FindChildByTitle(MDI%, "Error")
view% = FindChildByClass(erro%, "_AOL_View")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
A = SendMessage(aolw%, WM_CLOSE, 0, 0)
Clickbut (FindChildByTitle(aolw%, "OK"))
A = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
MsgBox "This is not a known AOL Member", 16, "CaPRe"
A = SendMessage(erro%, WM_CLOSE, 0, 0)
A = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub

Sub Mail_Set_Pref()
Call RunMenuByString(AOLWin(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Timeout (0.2)
Clickicon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

Clickbut (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function Mail_SetCursor(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "New Mail")
Tree% = FindChildByClass(MailWin%, "_AOL_Tree")
A6000% = SendMEssageByString(Tree%, LB_SETCURSEL, mailIndex, 0)
End Function

Function Mail_SetMailIndex(num)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "New Mail")
Tree% = FindChildByClass(MailWin%, "_AOL_Tree")
x = SendMEssageByString(Tree%, LB_SETCURSEL, num, 0)
End Function

Sub Mail_TrimFwdoffMail()
MailWin% = FindChildByTitle(AOLMDI, "Fwd: ")
x = getwintext(MailWin%)
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
cclabl% = GetWindow(peepz%, GW_HWNDNEXT)
cctext% = GetWindow(cclabl%, GW_HWNDNEXT)
SubjLabl% = GetWindow(cctext%, GW_HWNDNEXT)
SubjText% = GetWindow(SubjLabl%, GW_HWNDNEXT)
nofwd$ = Mid$(x, InStr(x, ":") + 2)
A = SendMEssageByString(SubjText%, WM_SETTEXT, 0, nofwd$)
End Sub

Sub Mail_Wait4Mail(MailBox As String)
'This will wait until yer mail fully loads
'Example: Waitmail "New Mail"
Mail = FindChildByTitle(AOLMDI, MailBox)
Tree = FindChildByClass(Mail, "_AOL_TREE")
DoEvents
x = SendMessage(Tree, 395, 0, 0)
Timeout 6
x = SendMessage(Tree, 395, 0, 0)
DoEvents
End Sub

Sub Mass_IMer(List As ListBox, Text As TextBox)
For i = 0 To List.ListCount - 1
who = List.List(i)
Call IM("" + who, "" + Text)
A = List.List(0)
Next i
End Sub

Sub MaximizeWindow(hWnd)
ma = showwindow(hWnd, SW_MAXIMIZE)
End Sub

Sub MinimizeWindow(hWnd)
mi = showwindow(hWnd, SW_MINIMIZE)
End Sub

Sub MoveFormNoCaption(frm As Form)


ReleaseCapture
G% = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, 2, 0)
End Sub

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)


End Function

Function Onlinecheck(person As String)
AOL% = FindWindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, "Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(IMWin, "_AOL_Edit")
aolrich% = FindChildByClass(IMWin, "RICHCNTL")
imsend% = FindChildByClass(IMWin, "_AOL_Icon")
Msg% = FindWindow("#32770", "America Online")

If AOLEdit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call Settext(AOLEdit%, person)
IMSEND2% = GetWindow(imsend%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
IMSEND2% = GetWindow(IMSEND2%, 2)
clik:
Clickicon (IMSEND2%)
Msg% = FindWindow("#32770", "America Online")
If Msg% = 0 Then
GoTo clik:
End If
stc = FindChildByClass(Msg%, "_AOL_Static")
x = GetMsgText(Msg%)
If InStr(1, x, "able") Then
text2 = "on"
If InStr(1, x, "cannot") Then
text2 = "off"
S = sendmessagebynum(Msg%, WM_CLOSE, 0, 0)
S = sendmessagebynum(Msg%, WM_CLOSE, 0, 0)
S = sendmessagebynum(Msg%, WM_CLOSE, 0, 0)
S = sendmessagebynum(IMWin, WM_CLOSE, 0, 0)
End If
End If
End Function

Sub ParentChange(Parent%, Location%)
doparent% = SetParent(Parent%, Location%)
End Sub

Function Percent()
'This will get the percent out of numbers
'On #            Out of        Leave 100
'Example: Text1 = Percent(List1.listcount,List2.Listcount,100)
Percent = Int(Complete / total * TotalOutput)
End Function

Sub PlayWav(file)
SoundName$ = file
SoundFlags& = &H20000 Or &H1
snd& = sndPlaySound(SoundName$, SoundFlags&)
End Sub

Sub PuntIM(person As String, Message As String)
AOL% = FindWindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, "Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(IMWin, "_AOL_Edit")
aolrich% = FindChildByClass(IMWin, "RICHCNTL")
imsend% = FindChildByClass(IMWin, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call Settext(AOLEdit%, person)
Call Settext(aolrich%, Message)
imsend% = FindChildByClass(IMWin, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends

Clickicon (imsend%)
Clickicon (imsend%)
Msg% = FindWindow("#32770", "America Online")
If Msg% <> 0 Then
stc = FindChildByClass(Msg%, "_AOL_Static")
x = GetMsgText(Msg%)
If InStr(1, x, "currently") Then z = person & "-  is not available"
If InStr(1, x, "able") Then z = person & "-  has IM's on"
If InStr(1, x, "cannot") Then z = person & "-  has IM's off"
MsgBox "" & x & ""
Exit Sub
End If
End Sub

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function READFILE(Where As String)
Filenum = FreeFile
Open (Where) For Input As Filenum
info = Input(LOF(Filenum), Filenum)
info = READFILE
End Function

Function ReadFwdWindCapt()
x% = FindForwardWindow
txt = Getwindtext(x%)
ReadFwdWindCapt = txt
End Function

Function ReadOpenMailCaption()
x% = FindOpenMail
tet = Getwindtext(x%)
ReadOpenMailCaption = tet
End Function

Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For Replacew = 1 To Len(Text)
thechar$ = Mid(Text, Replacew, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replacew

ReplaceText = thechars$

End Function

Function FiNGr()
'This is used for me, so peeps cant hex my programs and
'put their name in

FiNGr = Chr(70) & Chr(105) & Chr(78) & Chr(71) & Chr(114)
End Function

Function Replace(Text As String, what As String, withhis As String)

Do While (InStr(1, Text, what, 1) > 0)
    Where = InStr(1, Text, what, 1)
    If (Where > 0) Then
        LeftSide$ = Mid(Text, 1, Where - 1)
        RightSide$ = Mid(Text, Where + Len(what))
        Text = LeftSide$ + withhis + RightSide$
        Replace = Text
    End If
Loop
Replace = Text
End Function

Sub ResetName(aoldir, Chang, Create)
Dim BufSize As Long
Dim FilePosition As Long
Dim FileSize As Long
Dim FileSizeLeft As Long
Dim SN$
Dim Buffer$
Dim blah!
SN$ = Chang
Open aoldir + "\IDB\MAIN.IDX" For Binary Access Read Write As #2
FileSize = LOF(2)
FileSizeLeft = FileSize
FilePosition = 1
While FileSizeLeft >= 0
    If FileSizeLeft > 32000 Then
        BufSize = 32000
    ElseIf FileSizeLeft = 0 Then
        BufSize = 1
    Else
        BufSize = FileSizeLeft
    End If
    Buffer$ = String$(BufSize, " ")
    Get #2, FilePosition, Buffer$
    blah! = InStr(1, Buffer$, SN$, 1)
    If blah! Then Mid$(Buffer$, blah!) = Create
    Put #2, FilePosition, Buffer$
    FilePosition = FilePosition + BufSize
    FileSizeLeft = FileSize - FilePosition
    Wend
    Close #2
Call CloseAOL
x = Shell(aoldir + "\waol.exe")
MsgBox "AOL has been restarted to finish the Settings....Please wait", 64
End Sub

Function ReverseText(Text)
For Words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, Words, 1)
Next Words


End Function

Sub Roombust(roomcount As TextBox, room As ComboBox)
roomcount = "0"
Do
Keyword "aol://2719:2-2-" + room
roomcount = roomcount + 1
Timeout 2
AOL% = FindWindow("#32770", "America Online")
x = SendMEssageByString(AOL%, WM_CLOSE, 0, 0)
Loop Until FindChatRoom() <> 0
sendtext "I broke Into " + room
sendtext "In " + roomcount + " Trie(s)"
Exit Sub
End Sub

Function RoomName() As String
room% = FindChatRoom()
If room% = 0 Then RoomName$ = "No Room Detected": Exit Function
nontrimmed$ = Getwindtext(room%)
trimmed$ = TrimSpaces(nontrimmed$)
small = LCase(trimmed$)
RoomName$ = small
End Function

Sub RunAOLmenu(stringer As String)

AOL% = FindWindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, stringer)
End Sub

Sub RunAOLToolBar(tool)
Tbr = FindChildByClass(AOLWin(), "AOL Toolbar")
iconz% = FindChildByClass(Tbr, "_AOL_Icon")
For x = 1 To tool - 1
iconz% = GetWindow(iconz%, 2)
Next x
isen% = iswindowenabled(iconz%)
If isen% = 0 Then Exit Sub
Clickicon (iconz%)
End Sub

Sub Run_Menu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = sendmessagebynum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub

Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getstring = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getstring

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub

Sub ScrollLine8(txt As String)
lonh = String(116, Chr(32))
D = 116 - Len(Text1)
c$ = Left(lonh, D)
sendtext ("" & txt & c$ & txt)
sendtext ("" & txt & c$ & txt)
lonh = String(116, Chr(32))
D = 116 - Len(Text1)
c$ = Left(lonh, D)
sendtext ("" & txt & c$ & txt)
sendtext ("" & txt & c$ & txt)
End Sub

Sub SendCharNum(win, chars)
E = sendmessagebynum(win, WM_CHAR, chars, 0)

End Sub

Sub SendMailkillerror(who, Subject, Message)
Call RunAOLToolBar(2)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

A = SendMEssageByString(peepz%, WM_SETTEXT, 0, who)
A = SendMEssageByString(Subjec%, WM_SETTEXT, 0, Subject)
A = SendMEssageByString(Mess%, WM_SETTEXT, 0, Message)
Clickicon (icone%)
Timeout (0.6)
Clickicon (icone%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
erro% = FindChildByTitle(MDI%, "Error")
view% = FindChildByClass(erro%, "_AOL_View")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
A = SendMessage(aolw%, WM_CLOSE, 0, 0)
Clickbut (FindChildByTitle(aolw%, "OK"))
A = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
A = SendMessage(erro%, WM_CLOSE, 0, 0)
A = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub

Sub SendMailwitherror(who, Subject, Message)
Call RunAOLToolBar(2)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

A = SendMEssageByString(peepz%, WM_SETTEXT, 0, who)
A = SendMEssageByString(Subjec%, WM_SETTEXT, 0, Subject)
A = SendMEssageByString(Mess%, WM_SETTEXT, 0, Message)
Clickicon (icone%)
Timeout (0.6)
Clickicon (icone%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
erro% = FindChildByTitle(MDI%, "Error")
view% = FindChildByClass(erro%, "_AOL_View")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
A = SendMessage(aolw%, WM_CLOSE, 0, 0)
Clickbut (FindChildByTitle(aolw%, "OK"))
A = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
MsgBox "An error has occured ", 16, "CAprE"

Exit Do
End If
Loop
last:
End Sub

Sub sendtext(txt As String)

room% = FindChatRoom()
Call Settext(FindChildByClass(room%, "_AOL_Edit"), txt)
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)

End Sub

Sub Sendtext2(sendthis As String)


room% = FindChatRoom()
Call Settext(FindChildByClass(room%, "_AOL_Edit"), sendthis)
DoEvents
A% = FindChildByClass(room%, "_AOL_Icon")

Do
Call Clickicon(A%)
Loop Until InStr(1, LastChatLine, sendthis)


End Sub

Sub SetBackPre()
Call RunMenuByString(AOLWin(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Timeout (0.2)
Clickicon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 0, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

Clickbut (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function

Sub SetMsginFwdWindow(whattosay As String)
MailWin% = FindChildByTitle(AOLMDI, "Fwd: ")
Do
body% = FindChildByClass(MailWin%, "RICHCNTL")
Loop Until body% <> 0
A = SendMEssageByString(body%, WM_SETTEXT, 0, whattosay)
End Sub

Sub SetSNinFwdWindow(who As String)
MailWin% = FindChildByTitle(AOLMDI, "Fwd: ")
Do
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
Loop Until peepz% <> 0
A = SendMEssageByString(peepz%, WM_SETTEXT, 0, who)
End Sub

Sub SetToGuest()
'This sets the cursor to Guest in the sign on window
'good when making crackers and Multi <>< tos

'Example: SetToGuest

welc = FindChildByTitle(AOLMDI, "Welcome")
CB = FindChildByClass(Sign_On_Window(), "_AOL_COMBOBOX")
XX = sendmessagebynum(CB, CB_GETCOUNT, 0, 0&)
x = sendmessagebynum(CB, CB_SETCURSEL, (XX - 2), 0&)

End Sub

Sub Shell_to_file(Wut As String)
'This will shell to a file

'Example: ShellToFile("C:\VB\VB.EXE")
On Error GoTo Fawk
x = Shell(Wut, 1)
Exit Sub
Fawk:
MsgBox "That File Does Not Exist!", 16, "Error"
Exit Sub


End Sub

Function Sign_On_Window()
'This finds which sign on window is there
'"Welcome" Or "Goodbye from America Online!"
'useful when makin multi <>< tos, redialers, PW crackers

'Example: Comb = Findchildbyclass(Signonwnd(),"_AOL_COMBOBOX")


welc = FindChildByTitle(AOLMDI, "Welcome")
gb = FindChildByTitle(AOLMDI, "Goodbye from America Online!")
If welc <> 0 Then
Sign_On_Window = welc
ElseIf gb <> 0 Then
Sign_On_Window = gb
ElseIf gb = 0 & welc = 0 Then
Sign_On_Window = 0
End If
End Function

Function SignonWin()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Welcome")
SignonWin = IMWin
End Function

Function SnFromIM()

IMWin = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IMWin Then GoTo Greed
IMWin = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IMWin Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IMWin)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
SnFromIM = Naw$
End Function

Function StayOnline()
hwndz% = FindWindow(AOLWin(), "America Online")
childhwnd% = FindChildByTitle(hwndz%, "OK")
Clickbut (childhwnd%)
End Function

Sub stayontop(the As Form)

SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub String2List(items, List As ListBox)
If Not Mid(items, Len(items), 1) = "," Then
items = items & ","
End If

For DoList = 1 To Len(items)
thechars$ = thechars$ & Mid(items, DoList, 1)

If Mid(items, DoList, 1) = "," Then
List.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(items, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub

Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function

Function Text_Search(SearchFor, SearchThis)
x = InStr(1, SearchThis, SearchFor)
Text_Search = x
End Function

Sub Timeout(HowLong)
Beginning = timer
Do While timer - Beginning < HowLong
x = DoEvents()
Loop
End Sub

Function Toolbar() As Integer
AOL% = FindWindow("AOL FRAME25", vbNullString)
tol% = FindChildByClass(AOL%, "AOL TOOLBAR")
icon = FindChildByClass(tol%, "_AOL_ICON")
ico = GetWindow(icon, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
ico = GetWindow(ico, 2)
If ico > 0 Then
Toolbar = 5
End If
If ico = 0 Then
Toolbar = 8
End If
End Function

Function TrimCharacter(thetext, chars)
TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function Trimmail(thetext)
takechr13 = ReplaceText(thetext, "    ", "")

Trimmail = takechr13
End Function

Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For TrimSpace = 1 To Len(Text)
thechar$ = Mid(Text, TrimSpace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next TrimSpace

TrimSpaces = thechars$
End Function

Sub Wait_For_Child(child As String)
'This waits for a window to come up
'example: Call Waitforwnd (GetMDI(),"Caption of window to wait for")
Do
DoEvents
Na = FindChildByTitle(AOLMDI, child)
Loop Until Na <> 0
End Sub

Sub waitforok()
Do
AOL% = FindWindow("#32770", "America Online")

If AOL% Then
x = SendMEssageByString(AOL%, WM_CLOSE, 0, 0)
Exit Do
End If

aolw% = FindWindow("_AOL_Modal", vbNullString)

If aolw% Then
Clickbut (FindChildByTitle(aolw%, "OK"))
Exit Do
End If
Loop

End Sub

Sub waitmail()

AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
LB = FindChildByClass(A3000%, "_AOL_Tree")

Do
fir = sendmessagebynum(LB, LB_GETCOUNT, 0, 0)
Call Timeout(0.21)
sec = sendmessagebynum(LB, LB_GETCOUNT, 0, 0)
Call Timeout(0.21)
thir = sendmessagebynum(LB, LB_GETCOUNT, 0, 0)
Call Timeout(0.21)
forth = sendmessagebynum(LB, LB_GETCOUNT, 0, 0)

Loop Until fir = sec And sec = thir And thir = forth

End Sub

Sub Window_Disable(Windo%)
x = EnableWindow(Windo%, 0)
End Sub

Sub Window_Enable(Windo%)
x = EnableWindow(Windo%, 1)
End Sub

Sub WinHide(what As Integer)

Q% = showwindow(what, SW_HIDE)
DoEvents
End Sub

Sub WinShow(what As Integer)

Q% = showwindow(what, SW_SHOW)
DoEvents
End Sub

