Attribute VB_Name = "Snow32"
'      __
'  ___|¯¯|___  __  _____  ___  _  ___
' (__(|__|¯¯|\|¯¯|/¯¯/\¯¯\\¯¯\/¯\/¯¯¯/|
'|¯¯|)__)|__|\|__|\__\/__/;\___/\___/ -32
'|__||;;||;;|\|;;|\|;;;;|/\|;;|/\|;;|/
'|;;| ¯¯  ¯¯   ¯¯   ¯¯¯¯   ¯¯¯¯   ¯¯
' ¯¯  -  - zodi -
'      *¯¯¯¯¯¯¯¯¯*

Declare Function Getnextwindow Lib "User32" Alias "getnextwindow" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "User32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function WindowFromPointXY Lib "User32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "User32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "User32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function IsWindowEnabled Lib "User32" (ByVal hWnd As Long) As Long
Declare Function IsWindowVisible Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "User32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function Movewindow Lib "User32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (Object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "User32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "User32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function SendMessagebyNum& Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Declare Function Drawmenubar Lib "User32" Alias "DrawMenuBar" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Declare Function findwindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function sendmessagebystring Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage2 Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CreatePopupMenu Lib "User32" () As Long
Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function Getmenu Lib "User32" Alias "GetMenu" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function Gettopwindow Lib "User32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "User32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "User32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "User32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "User32" (ByVal hMenu%) As Integer
Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function MciSendString& Lib "Winmm" Alias "mciSendStringA" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen&, ByVal hcallback&)

'sndPlaySound  flag values for uFlags parameter
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_ALIAID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIASTART = 0  '  must be > 4096 to keep strings in same section of resource file
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside
 
 '  this range will raise an error
Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Public Const SND_TYPE_MASK = &H170007
 
 '  this range sets type
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

 '  Listbox Uses
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

 '  Keyboard Data
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

 '  Window postion junk
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

 '  getwindow and setwindow positions and yadayada
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
   y As Long
End Type

Global r&       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Global sockettype As Integer
Global Const FD_SETSIZE = 64

Type fd_set_type
  fd_count As Integer
  fd_array(FD_SETSIZE) As Integer
End Type
Global FD_SET As fd_set_type

Declare Function FD_ISSET Lib "winsock.dll" Alias "__WSAFDIsSet" (ByVal S As Integer, passed_set As fd_set_type) As Integer


'   Structure used in select() call, taken from the BSD file sys/time.h.

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

'   Socket function prototypes
Declare Function Accept Lib "winsock.dll" (ByVal S As Integer, addr As SockAddr_in, addrlen As Integer) As Integer
Declare Function Bind Lib "winsock.dll" (ByVal S As Integer, addr As SockAddr_in, ByVal namelen As Integer) As Integer
Declare Function closesocket Lib "winsock.dll" (ByVal S As Integer) As Integer
Declare Function htonl Lib "winsock.dll" (ByVal a As Long) As Long
Declare Function inet_addr Lib "winsock.dll" (ByVal S As String) As Long
Declare Function ntohl Lib "winsock.dll" (ByVal a As Long) As Long
Declare Function socket Lib "winsock.dll" (ByVal af As Integer, ByVal typesock As Integer, ByVal protocol As Integer) As Integer
Declare Function htons Lib "winsock.dll" (ByVal a As Integer) As Integer
Declare Function ntohs Lib "winsock.dll" (ByVal a As Integer) As Integer
Declare Function Connect Lib "winsock.dll" (ByVal sock As Integer, sockstruct As SockAddr_in, ByVal structlen As Integer) As Integer
Declare Function Send Lib "winsock.dll" Alias "send" (ByVal sock As Integer, ByVal Msg As String, ByVal msglen As Integer, ByVal Flag As Integer) As Integer
Declare Function Recv Lib "winsock.dll" (ByVal sock As Integer, ByVal Msg As String, ByVal msglen As Integer, ByVal Flag As Integer) As Integer
Declare Function Listen Lib "winsock.dll" (ByVal S As Integer, ByVal backlog As Integer) As Integer

'   Microsoft Windows Extension function prototypes

Declare Function WSaStartup Lib "winsock.dll" (ByVal a As Integer, b As WSAdata_type) As Integer
Declare Function WSACleanup Lib "winsock.dll" () As Integer

'  WINSOCK constants

Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2
Global Const AF_INET = 2
Sub Keepasnew()
mail% = findchildbytitle(AOLMDI, "New Mail")
but% = findchildbytitle(mail%, "Keep As New")
Clickbut but%
End Sub


Sub Keepmailasnew()

End Sub


Sub killwait()
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |Kill's wait using AOL's Timer                       | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim TheMod%, CloseNow%
RunMenu 2, 8
Do
TheMod% = findwindow("_AOL_Modal", "America Online")
NoFreeze% = DoEvents()
Loop Until TheMod% <> 0
Do
CloseNow% = SendMessagebyNum(TheMod%, WM_CLOSE, 0, 0)
NoFreeze% = DoEvents()
Loop Until TheMod% = 0
NoFreeze% = DoEvents()
End Sub


Sub AddtoAOL(FRM As Form, xpos, ypos)
FRM.Top = ypos
FRM.Left = xpos
AOL% = findwindow("AOL FRAME25", vbNullString)
TL% = findchildbyclass(AOL%, "AOL TOOLBAR")
sett = SetParent(FRM.hWnd, AOL%)
ack = ShowWindow(AOL%, 2)
ack = ShowWindow(AOL%, 3)

End Sub


Function AOLVersion()
RunAOLmenu ("&About America Online")
Do
bout% = findwindow("_AOL_Modal", vbNullString)
Loop Until bout% <> 0
stc% = findchildbyclass(bout%, "_AOL_Static")
red = GetText(stc%)
If InStr(1, red, "Windows 95") Then
S = SendMessagebyNum(bout%, WM_CLOSE, 0, 0)
AOLVersion = "32 bit"
Else
S = SendMessagebyNum(bout%, WM_CLOSE, 0, 0)
AOLVersion = "16 bit"
End If
End Function


Function FindReadWnd()
wind = findchildbyclass(AOLMDI, "AOL CHILD")
SN = findchildbytitle(wind, "Reply")
SL = findchildbytitle(wind, "Download File")
Ab = findchildbytitle(wind, "Reply to All")
Toz = findchildbyclass(wind, "_AOL_ICON")
If SL <> 0 Then
FindReadWnd = 1
x = GetParent(SL)
Else
FindReadWnd = 0
End If
End Function
Function FindFWDWindow()

wind = findchildbyclass(AOLMDI, "AOL CHILD")
SN = findchildbytitle(wind, "Send Now")
SL = findchildbytitle(wind, "Send Later")
'Ab = findchildbytitle(wind, "Reply to All")
Toz = findchildbyclass(wind, "_AOL_ICON")
'If SN <> 0 & SL <> 0 & Ab <> 0 & Toz <> 0 Then
If SL <> 0 Then
FindFWDWindow = 1
Else
FindFWDWindow = 0
End If
End Function


Function List_Count(Lst As ListBox)
x = Lst.ListCount
List_Count = x
End Function
Sub List_deleteitem(Lst As ListBox)
del = Lst.ListIndex
Lst.RemoveItem (del)
End Sub

Function MailsentModal()
aolw% = findwindow("_AOL_Modal", vbNullString)
Toz = findchildbytitle(wind, "Your mail has been sent.")
but% = findchildbyclass(aolw%, "_AOL_Button")
If Toz <> 0 Then
MailsentModal = 1
Clickbut but%
Else
MailsentModal = 0
End If
End Function

Function ReadFwdWindCapt()
x% = FindForwardWindow
TXT = Getwindtext(x%)
ReadFwdWindCapt = TXT
End Function

Function ReadOpenMailCaption()
x% = FindOpenMail
tet = Getwindtext(x%)
ReadOpenMailCaption = tet
End Function

Sub Sendtext2(sendthis As String)
'sends "txt" to the chat room actually clicks the send button
'and will click the button until the last chat line contains
'the line you sent
'Thanx Goob for the idea
' Fate uses this also    but i didnt get it from Fate

Room% = FindChatRoom()
Call Settext(findchildbyclass(Room%, "_AOL_Edit"), sendthis)
DoEvents
a% = findchildbyclass(Room%, "_AOL_Icon")

Do
Call Clickicon(a%)
Loop Until InStr(1, LastChatLine, sendthis)


End Sub







Function GetText(child)
GathTrim = SendMessagebyNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GathTrim)
GathString = sendmessagebystring(child, 13, GathTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function FindOpenMail()
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
'listers% = FindChildByTitle(childfocus%, "Reply")
'listere% = findchildbyclass(childfocus%, "_AOL_Icon")
listerb% = findchildbytitle(childfocus%, "Download File")

If listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


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










Sub Mail_SendNew(person, subject, message)
Call RunAOLToolBar(2)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(MailWin%, "_AOL_Icon")
peepz% = findchildbyclass(MailWin%, "_AOL_Edit")
subjt% = findchildbytitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, person)
a = sendmessagebystring(Subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, message)
Clickicon (icone%)
Timeout (0.6)
Clickicon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
View% = findchildbyclass(erro%, "_AOL_View")
aolw% = findwindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
Clickbut (findchildbytitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
MsgBox "This is not a known AOL Member", 16, "SnoW"
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub




Function FindSendWin(dosloop)
firs% = GetWindow(findchildbyclass(AOLWin(), "MDIClient"), 5)
forw% = findchildbytitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(findchildbyclass(AOLWin(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(findchildbyclass(AOLWin(), "MDIClient"), 5)
forw% = findchildbytitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = findchildbytitle(firs%, "Send Now")
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









Function AOLRoomCount()
chld% = FindChatRoom()
ListBox% = findchildbyclass(chld%, "_AOL_Listbox")

countem = SendMessage(ListBox%, LB_GETCOUNT, 0, 0)
AOLRoomCount = countem
End Function







Sub SetSNinFwdWindow(Who As String)
MailWin% = findchildbytitle(AOLMDI, "Fwd: ")
Do
peepz% = findchildbyclass(MailWin%, "_AOL_Edit")
Loop Until peepz% <> 0
a = sendmessagebystring(peepz%, WM_SETTEXT, 0, Who)
End Sub
Sub SetMsginFwdWindow(whattosay As String)
MailWin% = findchildbytitle(AOLMDI, "Fwd: ")
Do
body% = findchildbyclass(MailWin%, "RICHCNTL")
Loop Until body% <> 0
a = sendmessagebystring(body%, WM_SETTEXT, 0, whattosay)
End Sub
Sub TrimFwdoffMail()
MailWin% = findchildbytitle(AOLMDI, "Fwd: ")
x = getwintext(MailWin%)
peepz% = findchildbyclass(MailWin%, "_AOL_Edit")
cclabl% = GetWindow(peepz%, GW_HWNDNEXT)
cctext% = GetWindow(cclabl%, GW_HWNDNEXT)
SubjLabl% = GetWindow(cctext%, GW_HWNDNEXT)
Subjtext% = GetWindow(SubjLabl%, GW_HWNDNEXT)
nofwd$ = Mid$(x, InStr(x, ":") + 2)
a = sendmessagebystring(Subjtext%, WM_SETTEXT, 0, nofwd$)
End Sub

Function Text_Search(SearchFor, SearchThis)
x = InStr(1, SearchThis, SearchFor)
Text_Search = x
End Function

Sub Waitmail()

AOL% = findwindow("AOL Frame25", vbNullString)
A2000% = findchildbyclass(AOL%, "MDIClient")
A3000% = findchildbytitle(A2000%, "New Mail")
LB = findchildbyclass(A3000%, "_AOL_Tree")

Do
fir = SendMessagebyNum(LB, LB_GETCOUNT, 0, 0)
Call Timeout(0.21)
sec = SendMessagebyNum(LB, LB_GETCOUNT, 0, 0)
Call Timeout(0.21)
thir = SendMessagebyNum(LB, LB_GETCOUNT, 0, 0)
Call Timeout(0.21)
forth = SendMessagebyNum(LB, LB_GETCOUNT, 0, 0)

Loop Until fir = sec And sec = thir And thir = forth

End Sub


Function TrimSpaces(text)
If InStr(text, " ") = 0 Then
TrimSpaces = text
Exit Function
End If

For TrimSpace = 1 To Len(text)
thechar$ = Mid(text, TrimSpace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next TrimSpace

TrimSpaces = thechars$
End Function
Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function Trimmail(thetext)
takechr13 = ReplaceText(thetext, "    ", "")
'takechr10 = ReplaceText(takechr13, Chr$(10), "")
Trimmail = takechr13
End Function
Function TrimCharacter(thetext, chars)
TrimCharacter = ReplaceText(thetext, chars, "")

End Function









Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function


Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function


Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function StayOnline()
hwndz% = findwindow(AOLWin(), "America Online")
childhwnd% = findchildbytitle(hwndz%, "OK")
Clickbut (childhwnd%)
End Function



Function GetWindowDir()
Buffer$ = String$(255, 0)
x = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub




Sub SendCharNum(win, chars)
e = SendMessagebyNum(win, WM_CHAR, chars, 0)

End Sub


Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function


Sub setpreference()
Call RunMenuByString(AOLWin(), "Preferences")

Do: DoEvents
prefer% = findchildbytitle(AOLMDI(), "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Timeout (0.2)
Clickicon (mailbut%)

Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
aolcloses% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

Clickbut (aolOK%)
Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Sub SetBackPre()
Call RunMenuByString(AOLWin(), "Preferences")

Do: DoEvents
prefer% = findchildbytitle(AOLMDI(), "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Timeout (0.2)
Clickicon (mailbut%)

Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
aolcloses% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 0, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

Clickbut (aolOK%)
Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = Getmenu(findwindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
clickaolmenu = SendMessagebyNum(findwindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub


Function ReplaceText(text, charfind, charchange)
If InStr(text, charfind) = 0 Then
ReplaceText = text
Exit Function
End If

For Replace = 1 To Len(text)
thechar$ = Mid(text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function

Sub RunMenuByString(Application, StringSearch)
ToSearch% = Getmenu(Application)
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



Function ReverseText(text)
For Words = Len(text) To 1 Step -1
ReverseText = ReverseText & Mid(text, Words, 1)
Next Words


End Function

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Sub ParentChange(Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
'turns the "number" so vb recognizes it for
'addition, subtraction, ect.

End Function




Sub MinimizeWindow(hWnd)
mi = ShowWindow(hWnd, SW_MINIMIZE)
End Sub



Sub MaximizeWindow(hWnd)
ma = ShowWindow(hWnd, SW_MAXIMIZE)
End Sub

Sub WaitForOk()
Do
AOL% = findwindow("#32770", "America Online")

If AOL% Then
x = sendmessagebystring(AOL%, WM_CLOSE, 0, 0)
Exit Do
End If

aolw% = findwindow("_AOL_Modal", vbNullString)

If aolw% Then
Clickbut (findchildbytitle(aolw%, "OK"))
Exit Do
End If
Loop

End Sub


Function LineFromText(text, theline)
theview$ = text


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
Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Sub HideWindow(hWnd)
hi = ShowWindow(hWnd, SW_HIDE)
End Sub






Function GetLineCount(text)

theview$ = text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(text, Len(text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function


Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop

End Function


Function findchildbyclass(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
Buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, Buffer$, 250)

If InStr(UCase(Buffer$), UCase(child)) Then findchildbyclass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function



Function findchildbytitle(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
Buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, Buffer$, (hwndLength% + 1))

If InStr(UCase(Buffer$), UCase(child)) Then findchildbytitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function







Sub playwav(File)
SoundName$ = File
SoundFlags& = &H20000 Or &H1
snd& = SndPlaySound(SoundName$, SoundFlags&)
End Sub


Function READFILE(Where As String)
Filenum = FreeFile
Open (Where) For Input As Filenum
Info = Input(LOF(Filenum), Filenum)
Info = READFILE
End Function


Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function


Function Mail_SetCursor(mailIndex As String)
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "New Mail")
Tree% = findchildbyclass(MailWin%, "_AOL_Tree")
A6000% = sendmessagebystring(Tree%, LB_SETCURSEL, mailIndex, 0)
End Function


Function Mail_PressEnter()
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "New Mail")
Tree% = findchildbyclass(MailWin%, "_AOL_Tree")
x = SendMessage(ree%, WM_KEYDOWN, VK_RETURN, 0)
x = SendMessage(Tree%, WM_KEYUP, VK_RETURN, 0)
End Function

Function Mail_ClickForward()
MDI% = AOLMDI
chil% = findchildbyclass(MDI%, "AOL CHILD")
fwd% = findchildbytitle(chil%, "Forward")
button = GetWindow(fwd%, 2)
Clickicon (button)
End Function


Function Mail_ForwardMail(SN As String, message As String)
ForwardWindow
person = SN
Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Fwd: ")
icone% = findchildbyclass(MailWin%, "_AOL_Icon")
peepz% = findchildbyclass(MailWin%, "_AOL_Edit")
subjt% = findchildbytitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop
a = sendmessagebystring(peepz%, WM_SETTEXT, 0, person)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, message)
Clickicon (icone%)
End Function



Function FindKeyword()
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
keyw% = findchildbytitle(MDI%, "Keyword")
KEdit% = findchildbyclass(keyw%, "_AOL_Edit")
FindKeyword = KEdit%
End Function


Sub File_DeleteDirectory(DirName$)
'If Not File_IfDirectoryExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub


Function getwintext(hWnd As Integer) As String
lentos = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
x = sendmessagebystring(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
getwintext = Buffer$
End Function





Sub Addroom95(Lst As ListBox)
LB = findchildbyclass(FindChatRoom(), "_AOL_Listbox")
Coun = SendMessagebyNum(LB, LB_GETCOUNT, 0, 0)
For g = 0 To Coun - 1
x = SendMessagebyNum(LB, LB_SETCURSEL, g, 0)
Q = SendMessagebyNum(LB, WM_LBUTTONDBLCLK, 0, 0)
z = RoomName
chil% = findchildbyclass(AOLMDI(), "AOL Child")
Mess = findchildbytitle(chil%, "Message")

Do
x = getwintext(chil%)
Loop Until z <> x

Lst.AddItem x
S = SendMessagebyNum(chil%, WM_CLOSE, 0, 0)
Next g
For i = 0 To Lst.ListCount - 1
If Lst.List(i) = "" & aolUser & "" Then
Lst.RemoveItem i
End If
Next i
End Sub




















Function Toolbar() As Integer
AOL% = findwindow("AOL FRAME25", vbNullString)
tol% = findchildbyclass(AOL%, "AOL TOOLBAR")
icon = findchildbyclass(tol%, "_AOL_ICON")
Ico = GetWindow(icon, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
Ico = GetWindow(Ico, 2)
If Ico > 0 Then
Toolbar = 5
End If
If Ico = 0 Then
Toolbar = 8
End If
End Function










Sub MoveFormNoCaption(FRM As Form)
'this goes in Object_mousedown
'ie. Form1_mousedown

ReleaseCapture
g% = SendMessage(FRM.hWnd, WM_NCLBUTTONDOWN, 2, 0)
End Sub


Sub Scroll8Line(TXT As String)
lonh = String(116, Chr(32))
d = 116 - Len(text1)
c$ = Left(lonh, d)
Sendtext ("" & TXT & c$ & TXT)
Sendtext ("" & TXT & c$ & TXT)
lonh = String(116, Chr(32))
d = 116 - Len(text1)
c$ = Left(lonh, d)
Sendtext ("" & TXT & c$ & TXT)
Sendtext ("" & TXT & c$ & TXT)
End Sub

Function LastChatLine()
getpar% = FindChatRoom()
child = findchildbyclass(getpar%, "_AOL_View")
GetTrim = SendMessagebyNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = sendmessagebystring(child, 13, GetTrim + 1, TrimSpace$)

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
lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
LastChatLine = lastline

End Function


Sub Keyword(Where)
Call RunMenuByString(AOLWin(), "Keyword...")

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
keyw% = findchildbytitle(MDI%, "Keyword")
KEdit% = findchildbyclass(keyw%, "_AOL_Edit")
If KEdit% Then Exit Do
Loop

editsend% = sendmessagebystring(KEdit%, WM_SETTEXT, 0, Where)
pausing = DoEvents()
Sending% = SendMessage(KEdit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub


Sub Clickicon(icon%)
c = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
c = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub HideAOL()
Hid = ShowWindow(AOLWin(), SW_HIDE)
End Sub
Function AOLWin()
AOL% = findwindow("AOL Frame25", vbNullString)
AOLWin = AOL%
End Function
Function AOLToolbar()
AOL% = findwindow("AOL Frame25", vbNullString)
tob% = findchildbyclass(AOL%, "AOL Toolbar")
AOLToolbar = tob%
End Function

Function ListToString(List As ListBox)
For DoList = 0 To thelist.ListCount - 1
ListToString = ListToString & List.List(DoList) & ", "
Next DoList
ListToString = Mid(ListToString, 1, Len(ListToString) - 2)
End Function

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


Function RoomName() As String
Room% = FindChatRoom()
If Room% = 0 Then RoomName$ = "No Room Detected": Exit Function
nontrimmed$ = Getwindtext(Room%)
trimmed$ = TrimSpaces(nontrimmed$)
small = LCase(trimmed$)
RoomName$ = small
End Function
Function GetMsgText(MBox As Integer) As String
Stat1% = findchildbyclass(MBox%, "STATIC")
Stat2% = GetWindow(Stat1%, GW_HWNDNEXT)
Stat$ = Getwindtext(Stat2%)
If Stat$ = "" Then Stat$ = Getwindtext(Stat1%)
GetMsgText$ = Stat$
End Function

Function Onlinecheck(person As String)
AOL% = findwindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, "Send an Instant Message")

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Send Instant Message")
aoledit% = findchildbyclass(IMWin, "_AOL_Edit")
aolrich% = findchildbyclass(IMWin, "RICHCNTL")
imsend% = findchildbyclass(IMWin, "_AOL_Icon")
Msg% = findwindow("#32770", "America Online")

If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call Settext(aoledit%, person)
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
Msg% = findwindow("#32770", "America Online")
If Msg% = 0 Then
GoTo clik:
End If
stc = findchildbyclass(Msg%, "_AOL_Static")
x = GetMsgText(Msg%)
If InStr(1, x, "currently") Then z = person & "-  is not available" 'put what to say here if they are offline
If InStr(1, x, "able") Then z = person & "-  has IM's on"    'Put what to say in here for IMs on
If InStr(1, x, "cannot") Then z = person & "-  has IM's off" 'Put what to say here for IMs off
Onlinecheck = z
S = SendMessagebyNum(Msg%, WM_CLOSE, 0, 0)
S = SendMessagebyNum(Msg%, WM_CLOSE, 0, 0)
S = SendMessagebyNum(Msg%, WM_CLOSE, 0, 0)
S = SendMessagebyNum(IMWin, WM_CLOSE, 0, 0)
End Function

Function Online2()
AOL% = findwindow("AOL Frame25", "America  Online")
MDI% = findchildbyclass(AOL%, "MDIClient")
wlcm% = findchildbytitle(MDI%, "Welcome, " & SnoW_User & "!")
If wlcm% = 0 Then
Online2 = False
Exit Function
Else
Online2 = True
End If

End Function
Function Online()
'Checks if the User is Online
' If Snow_Online = False then
' Exit sub

AOL% = findwindow("AOL Frame25", "America  Online")
MDI% = findchildbyclass(AOL%, "MDIClient")
wlcm% = findchildbytitle(MDI%, "Welcome, " & SnoW_User & "!")
If wlcm% = 0 Then
MsgBox "Please Sign on before using this feature", 48, "SnoW"
Online = False
Exit Function
Else
Online = True
End If

End Function

Sub AddroomCombo(Cmb As ComboBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindChatRoom()
aolhandle = findchildbyclass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESVM_READ Or STANDARD_RIGHTREQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Cmb.AddItem person$
Next index
Call CloseHandle(AOLProcessThread)
End If
For i = 0 To Cmb.ListCount - 1
If Cmb.List(i) = "" & SnoW_User & "" Then
Cmb.RemoveItem i
End If
Next i
End Sub
Sub AddroomList(Lst As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindChatRoom()
aolhandle = findchildbyclass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESVM_READ Or STANDARD_RIGHTREQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Lst.AddItem person$
Next index
Call CloseHandle(AOLProcessThread)
End If
For i = 0 To Lst.ListCount - 1
If Lst.List(i) = "" & SnoW_User & "" Then
Lst.RemoveItem i
End If
Next i
End Sub






Sub Window_Disable(Windo%)
x = EnableWindow(Windo%, 0)
End Sub

Sub Window_Enable(Windo%)
x = EnableWindow(Windo%, 1)
End Sub

Function Getwindtext(read As Integer)
BufLen% = SendMessagebyNum(read%, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String(BufLen%, 0)
Q% = sendmessagebystring(read%, WM_GETTEXT, BufLen% + 1, Buffer$)
DoEvents
Getwindtext = Trim(Buffer$)
End Function


Sub FileCopy(File$, DestFile$)
If Not FileExists(File$) Then Exit Sub
FileCopy File$, DestFile$
End Sub
Sub IM(person As String, message As String)
Call RunMenuByString(AOLWin(), "Send an Instant Message")

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Send Instant Message")
aoledit% = findchildbyclass(IMWin, "_AOL_Edit")
aolrich% = findchildbyclass(IMWin, "RICHCNTL")
imsend% = findchildbyclass(IMWin, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call Settext(aoledit%, person)
Call Settext(aolrich%, message)
imsend% = findchildbyclass(IMWin, "_AOL_Icon")
For sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next sends
Clickicon (imsend%)
Clickicon (imsend%)
Clickicon (imsend%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Send Instant Message")
aolcl% = findwindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin, WM_CLOSE, 0, 0): Exit Do
If IMWin = 0 Then Exit Do
Loop
End Sub

Sub CenterForm(FRM As Form)
'This Will center a form

'Example: Center Me

abc = (Screen.Width - FRM.Width) / 2
xyz = (Screen.Height - FRM.Height) / 2
FRM.Move abc, xyz
End Sub
Sub RunAOLmenu(stringer As String)
'This will Look through each Menu on the AOL
'menu bar and clicks the one you want
'if you wanna Click the READ NEW MAIL one
'you must put Call runaolmenu("Read &New Mail")
'where & = the underline.

AOL% = findwindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, stringer)
End Sub

Sub FileRename(File$, NewName$)
Name File$ As NewName$
NoFreeze% = DoEvents()
End Sub
Sub FileOpen(File$)
x = Shell(File$, 1): NoFreeze% = DoEvents()
End Sub
Sub FileMakDirectory(DirName$)
MkDir DirName$
End Sub
Sub ClickReadMail()
Dim mail%, Send%
mail% = findchildbytitle(AOLMDI(), "New Mail")
Send% = findchildbytitle(mail%, "Read")
Clickbut Send%
End Sub

Sub FileAddto(File$, WhatToAdd$)
'this adds Whatever you put as Whattoadd$ into any file
'Example: Call Snow_AddtoFile("C:\Windows\win.ini", "SnoW Owns me")

If FileExists(File$) = False Then Exit Sub
Open File$ For Append As #1
    Print #1, WhatToAdd$
    Close #1
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
Sub ButtonDblClk(button%)
Dim DoubleClick
DoubleClick = SendMessagebyNum(button%, WM_LBUTTONDBLCLK, &HD, 0)
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

Sub RunAOLToolBar(tool)
Tbr = findchildbyclass(AOLWin(), "AOL Toolbar")
iconz% = findchildbyclass(Tbr, "_AOL_Icon")
For x = 1 To tool - 1
iconz% = GetWindow(iconz%, 2)
Next x
isen% = IsWindowEnabled(iconz%)
If isen% = 0 Then Exit Sub
Clickicon (iconz%)
End Sub

Function ClickList(hWnd)
cliklst% = SendMessagebyNum(hWnd, &H203, 0, 0&)
End Function
Sub CloseWin(wind)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)
closeit = SendMessage(wind, WM_CLOSE, 0, 0)

End Sub
Function CountMail(MailWin As String)
mail = findchildbytitle(AOLMDI(), MailWin)
Tree = findchildbyclass(mail, "_AOL_Tree")
CountMail = SendMessage(Tree, LB_GETCOUNT, 0, 0)
End Function
Function FindChatRoom()
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = findchildbyclass(childfocus%, "_AOL_Edit")
listere% = findchildbyclass(childfocus%, "_AOL_View")
listerb% = findchildbyclass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindChatRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function


Function AOLGetListString(Parent, index, Buffer As String)
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
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
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
AOLGetTopWindow = Gettopwindow(AOLMDI())
End Function

Function GetUser()
On Error Resume Next
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
welcome% = findchildbytitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
GetUser = User$
End Function

Function AOLMDI()
AOL% = findwindow("AOL Frame25", vbNullString)
AOLMDI = findchildbyclass(AOL%, "MDIClient")
End Function
Sub SendMailwitherror(Who, subject, message)
Call RunAOLToolBar(2)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(MailWin%, "_AOL_Icon")
peepz% = findchildbyclass(MailWin%, "_AOL_Edit")
subjt% = findchildbytitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, Who)
a = sendmessagebystring(Subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, message)
Clickicon (icone%)
Timeout (0.6)
Clickicon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
View% = findchildbyclass(erro%, "_AOL_View")
aolw% = findwindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
Clickbut (findchildbytitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
MsgBox "An error has occured during the mailing process.  Please review it and try again", 16, "Prohecies by Snow and Lunik"
'a = sendmessage(erro%, WM_CLOSE, 0, 0)
'a = sendmessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub

Sub SendMailkillerror(Who, subject, message)
Call RunAOLToolBar(2)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(MailWin%, "_AOL_Icon")
peepz% = findchildbyclass(MailWin%, "_AOL_Edit")
subjt% = findchildbytitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, Who)
a = sendmessagebystring(Subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, message)
Clickicon (icone%)
Timeout (0.6)
Clickicon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
View% = findchildbyclass(erro%, "_AOL_View")
aolw% = findwindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
Clickbut (findchildbytitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
'MsgBox "An error has occured during the mailing process.  Please review it and try again", 16, "Prohecies by Snow and Lunik"
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub



Sub AOLSignoff()
AOL% = findwindow("AOL Frame25", vbNullString)
If AOL% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunAOLmenu("&Sign Off")

End Sub


Sub Killwait2()
    AOL% = findwindow("AOL Frame25", vbNullString)
    mofo% = findchildbyclass(AOL%, "MDIClient")
    RunAOLmenu "Preferences"
    Do
        c% = DoEvents()
        pre% = findchildbytitle(mofo%, "Preferences")
    Loop Until pre% <> 0
      d = SendMessagebyNum(pre%, WM_CLOSE, 0, 0)
End Sub

Sub KillWait1()
RunAOLToolBar 18
Do
Add = findchildbytitle(AOLMDI, "Keyword")
Loop Until Add <> 0
'Do
CloseWin Add
'Loop Until Add = 0
End Sub
Sub KillWait3()
x = 0
RunAOLmenu "Locate A Member Online"
Do
Add = findchildbytitle(AOLMDI, "Locate Member Online")
x = x + 1
Loop Until Add <> 0 Or x = 10
Do
CloseWin Add
Loop Until Add <> 1
End Sub

Sub Settext(Where%, thestring$)
Dim SendTheText%
SendTheText% = sendmessagebystring(Where%, WM_SETTEXT, 0, thestring$)
End Sub
Sub GetNum(hWnd%, x)
Dim b
Do
If x = 0 Then Exit Sub
b = 1 + b
hWnd% = GetWindow(hWnd%, GW_HWNDNEXT)
NoFreeze% = DoEvents()
Loop Until b >= x - 1
End Sub




Sub ClickSendNow()
fwdwin% = findchildbytitle(AOLMDI(), "Fwd: ")
icn% = findchildbyclass(fwdwin%, "_AOL_Icon")
nxt% = GetWindow(icn%, GW_HWNDNEXT)
Do
Clickicon icn%
Loop Until fwdwin% <> 1


End Sub




Function ForwardMail(SN As String, Msg As String)
FindForwardWindow
person = SN
message = Msg
Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "Fwd: ")
icone% = findchildbyclass(MailWin%, "_AOL_Icon")
peepz% = findchildbyclass(MailWin%, "_AOL_Edit")
subjt% = findchildbytitle(MailWin%, "Subject:")
Subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And Subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop
a = sendmessagebystring(peepz%, WM_SETTEXT, 0, person)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, message)
Clickicon (icone%)
End Function
Sub KillListDupes(Lst As Control)
For i = 0 To Lst.ListCount - 1
For e = 0 To Lst.ListCount - 1

If LCase(Lst.List(i)) Like LCase(Lst.List(e)) And i <> e Then

Lst.RemoveItem (e)
End If

Next e
Next i
End Sub
Sub KillComboDupes(Cmb As Control)
For i = 0 To Cmb.ListCount - 1
For e = 0 To Cmb.ListCount - 1

If LCase(Cmb.List(i)) Like LCase(Cmb.List(e)) And i <> e Then

Cmb.RemoveItem (e)
End If

Next e
Next i
End Sub
Function EnterKey()
EnterKey = CStr(Chr(13) & Chr(10))
End Function

Sub DeleteSelectedMail()
Mailbox% = findchildbytitle(AOLMDI(), "New Mail")
but% = findchildbytitle(Mailbox%, "Delete")
Clickbut but%
End Sub


Function Mailtree()

mail% = findchildbytitle(AOLMDI(), "New Mail")
Tree% = findchildbyclass(mail%, "_AOL_Tree")

Mailtree = Tree%

End Function


Sub Maillist(Lst As ListBox, Which As String)
Dim z As String
Mailbox% = findchildbytitle(AOLMDI(), Which)
Tree% = findchildbyclass(Mailbox%, "_AOL_TREE")
z = 0
For i = 0 To SendMessagebyNum(Tree%, LB_GETCOUNT, 0, 0&) - 1
MailStr$ = String$(255, " ")
    Q% = sendmessagebystring(Tree%, LB_GETTEXT, i, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    Lst.AddItem z + ") " + Trim(NoSN$)
    z = z + 1
       Next i

End Sub







Sub FileProperties(Choice, File$)
'Choices are: Normal, Readonly, Hidden, System, and archive
If Not FileExists(File$) Then Exit Sub
If LCase$(Choice) = "normal" Then
SetAttr File$, vbNormal
ElseIf LCase$(Choice) = "readonly" Then: SetAttr File$, vbReadOnly
ElseIf LCase$(Choice) = "hidden" Then: SetAttr File$, vbHidden
ElseIf LCase$(Choice) = "system" Then: SetAttr File$, vbSystem
ElseIf LCase$(Choice) = "archive" Then: SetAttr File$, vbArchive
End If
NoFreeze% = DoEvents()
End Sub



Sub Closemailbox(Which As String)
Mailbox% = findchildbytitle(AOLMDI(), Which)
S = SendMessagebyNum(Mailbox%, WM_CLOSE, 0, 0)
End Sub

Function Goodbyewin()
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Goodbye from America Online!")
Goodbyewin = IMWin
End Function



Function SignonWin()
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Welcome")
SignonWin = IMWin
End Function

Sub CloseAOL()
AOL% = findwindow("AOL Frame25", vbNullString)
closes = SendMessage(AOL%, WM_CLOSE, 0, 0)
End Sub

Function GetFromINI(Appname$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(Appname$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function
Sub ResetName(aoldir, Chang, Create)
'                AOL DIR           Name to change       create...
Dim BufSize As Long
Dim FilePosition As Long
Dim FileSize As Long
Dim FileSizeLeft As Long
Dim SN$
Dim Buffer$
Dim Blah!
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
    Blah! = InStr(1, Buffer$, SN$, 1)
    If Blah! Then Mid$(Buffer$, Blah!) = Create
    Put #2, FilePosition, Buffer$
    FilePosition = FilePosition + BufSize
    FileSizeLeft = FileSize - FilePosition
    Wend
    Close #2
Call CloseAOL
x = Shell(aoldir + "\waol.exe")
MsgBox "AOL has been restarted to finish the Settings....Please wait", 64
End Sub
Sub Timeout(HowLong)
Beginning = timer
Do While timer - Beginning < HowLong
x = DoEvents()
Loop
End Sub
Sub Stayontop(the As Form)
'sets your form to be the topmost window all the
'time. Example:  StayOnTop Me
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Sub LoopClickFav(name, timer As timer)
rm% = findwindow("AOLChild", name)
prefer% = findchildbytitle(AOLMDI(), "Favorite Places")
but% = findchildbytitle(prefer%, "Connect")
Do
Timeout (0.2)
Clickbut but%
a% = findwindow("#32770", "America Online")
z% = findchildbyclass(a%, "Button")
If a% <> 0 Then
Clickbut z%
End If
Loop Until rm% <> 0

del% = findchildbytitle(prefer%, "Delete")
Clickbut del%
Q% = findwindow("#32770", "America Online")
V% = findchildbyclass(a%, "Button")
Clickbut V%
S = SendMessagebyNum(prefer%, WM_CLOSE, 0, 0)
End Sub


Function GetMailIndexcaption(Which As String)
Mailbox% = findchildbytitle(AOLMDI(), "New Mail")
Tree% = findchildbyclass(Mailbox%, "_AOL_TREE")
 Q% = sendmessagebystring(Tree%, LB_GETTEXT, Which, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    GetMailIndexcaption = Trim(NoSN$)
End Function








Function SetMailIndex(Num)
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
MailWin% = findchildbytitle(MDI%, "New Mail")
Tree% = findchildbyclass(MailWin%, "_AOL_Tree")
x = sendmessagebystring(Tree%, LB_SETCURSEL, Num, 0)
End Function


Sub Favplace(name As String, address As String)
Call RunAOLmenu("Favorite Places")
Do ': DoEvents
prefer% = findchildbytitle(AOLMDI(), "Favorite Places")
maillab% = findchildbytitle(prefer%, "Add Favorite Place")
Tree% = findchildbyclass(prefer%, "_AOL_Tree")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop
Timeout (1)
Clickbut maillab%
Timeout (0.5)

Do: DoEvents
Add% = findchildbytitle(AOLMDI(), "Add Favorite Place")
but% = findchildbyclass(Add%, "_AOL_Edit")
but2% = GetWindow(but%, GW_HWNDNEXT)
but3% = GetWindow(but2%, GW_HWNDNEXT)
If Add% <> 0 And but% <> 0 Then Exit Do
Loop
Timeout (0.4)
Call Settext(but%, name)
Call Settext(but3%, address)
But4% = findchildbytitle(Add%, "OK")
Clickbut But4%
clikr% = findwindow("#32770", "America Online")
clikbut% = findchildbyclass(clikr%, "Button")
Timeout (0.2)
If clikr% <> 0 Then
Clickbut clikbut%
Timeout (0.1)
S = SendMessagebyNum(Add%, WM_CLOSE, 0, 0)
End If
For d = 0 To SendMessagebyNum(Tree%, LB_GETCOUNT, 0, 0&) - 1
S = SendMessagebyNum(Tree%, LB_SETCURSEL, d, 0&)
Next d
End Sub

Sub CloseGoodbye()
AOL% = findchildbytitle(AOLWin(), "Goodbye from America Online!")
d = SendMessagebyNum(AOL%, WM_CLOSE, 0, 0)
End Sub
Sub instantmessage(person As String, message As String)
'sends an Instant Message to "Person" with the
'message of "message"
AOL% = findwindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, "Send an Instant Message")

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Send Instant Message")
aoledit% = findchildbyclass(IMWin, "_AOL_Edit")
aolrich% = findchildbyclass(IMWin, "RICHCNTL")
imsend% = findchildbyclass(IMWin, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call Settext(aoledit%, person)
Call Settext(aolrich%, message)
imsend% = findchildbyclass(IMWin, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends

Clickicon (imsend%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Send Instant Message")
aolcl% = findwindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin, WM_CLOSE, 0, 0): Exit Do
If IMWin = 0 Then Exit Do
Loop
S = SendMessagebyNum(IMWin, WM_CLOSE, 0, 0)
End Sub

Sub PuntIM(person As String, message As String)
AOL% = findwindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, "Send an Instant Message")

Do: DoEvents
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, "Send Instant Message")
aoledit% = findchildbyclass(IMWin, "_AOL_Edit")
aolrich% = findchildbyclass(IMWin, "RICHCNTL")
imsend% = findchildbyclass(IMWin, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call Settext(aoledit%, person)
Call Settext(aolrich%, message)
imsend% = findchildbyclass(IMWin, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends

Clickicon (imsend%)
Clickicon (imsend%)
Msg% = findwindow("#32770", "America Online")
If Msg% <> 0 Then
stc = findchildbyclass(Msg%, "_AOL_Static")
x = GetMsgText(Msg%)
If InStr(1, x, "currently") Then z = person & "-  is not available" 'put what to say here if they are offline
If InStr(1, x, "able") Then z = person & "-  has IM's on"    'Put what to say in here for IMs on
If InStr(1, x, "cannot") Then z = person & "-  has IM's off" 'Put what to say here for IMs off
MsgBox "" & x & ""
Exit Sub
End If
End Sub



Sub IMsOn()
'turns IM's on
Call IM("$im_on", "SnoW     IMs On")
End Sub
Sub IMsOff()
'turns IM's off
Call IM("$im_off", "SnoW     IMs Off")
End Sub




Sub GetMemberProfile(name As String)
'this will get the profile of name
RunAOLmenu ("Get a Member's Profile")
Timeout 0.3
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
prof% = findchildbytitle(MDI%, "Get a Member's Profile")
putname% = findchildbyclass(prof%, "_AOL_Edit")
Call Settext(putname%, name)
okbutton% = findchildbyclass(prof%, "_AOL_Button")
Clickbut okbutton%
End Sub
Function FindRoom()
'This will set the the focus to the Chat room window
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
childfcs% = GetWindow(MDI%, 5)

While childfcs%
listers% = findchildbyclass(childfcs%, "_AOL_Edit")
listere% = findchildbyclass(childfcs%, "_AOL_View")
listerb% = findchildbyclass(childfcs%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindRoom = childfcs%: Exit Function
childfcs% = GetWindow(childfcs%, GW_HWNDNEXT)
Wend
End Function

Sub AOLCaption(newcaption As String)
'This changes the AMerica Online Caption to whatever
'you change newcaption to
Call Settext(AOL%, newcaption)
End Sub
Sub Clickbut(wut%)

Click = SendMessage(wut%, WM_KEYDOWN, VK_SPACE, 0)
Click = SendMessage(wut%, WM_KEYUP, VK_SPACE, 0)
End Sub
Sub List2String(thetext As TextBox, thelist As ListBox, Block)
'This will take a listbox you have and read through
'it and add a what ever you have as BLOCK
' like... Call List2string (text1, list1, "@aol.com,")

For i = 0 To thelist.ListCount - 1
thetext = thetext + thelist.List(i) + Block
Next i
End Sub

Sub AntipuntALL(List As ListBox)
' this antipunt is good but it doesn't distinguish the IMs
' it kills them all
' put in a timer with an interval of about 50-100

AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, ">Instant Message From:")
rch2% = findchildbyclass(IMWin, "RICHCNTL")
nme = SNFROMIM

If rch2% <> 0 Then
S = SendMessagebyNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessagebyNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessagebyNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessagebyNum(rch2%, WM_CLOSE, 0, 0)
r = ShowWindow(IMWin, SW_HIDE)
List.AddItem nme
Call KillListDupes(List)
Exit Sub
End If

End Sub


Function SNFROMIM()

IMWin = findchildbytitle(AOLMDI(), ">Instant Message From:")
If IMWin Then GoTo Greed
IMWin = findchildbytitle(AOLMDI(), "  Instant Message From:")
If IMWin Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IMWin)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
SNFROMIM = naw$
End Function
Sub Sendtext(TXT As String)
'Send what txt = to the Chat room
'Example: Call Sendtext ("This Bas Rox")
Room% = FindChatRoom()
Call Settext(findchildbyclass(Room%, "_AOL_Edit"), TXT)
Call SendCharNum(findchildbyclass(Room%, "_AOL_Edit"), 13)
Call SendCharNum(findchildbyclass(Room%, "_AOL_Edit"), 13)

End Sub



Sub AntiPuntDis()
' this anti punt goes in a timer with an
' interval of about 50-100
' this will also distinguish whether the IM contains
' the h3 or the CTRL Backspace punt codes
'just type  Call AntiPuntDis in the timer code

AOL% = findwindow("AOL Frame25", "America  Online")
MDI% = findchildbyclass(AOL%, "MDIClient")
IMWin = findchildbytitle(MDI%, ">Instant Message From:")
rch2% = findchildbyclass(IMWin, "RICHCNTL")
nme = SNFROMIM
x = Getwindtext(rch2%)
If InStr(x, "    ") Then
Do
S = SendMessagebyNum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
Sendtext "" & nme & " Is trying to punt me"
End If

If InStr(x, "") Then
Do
S = SendMessagebyNum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
Sendtext "" & nme & " Is trying to punt me"
End If
End Sub
Function IfFileExists(ByVal sFileName As String) As Integer
'aol://1391:43-16662
'This looks for a File.
'Example: If Not Snow_IFileExists("c:\snowowns.hah") then...
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IfFileExists = False
        Else
            IfFileExists = True
    End If

End Function
Sub HideWelcome()
'Umm this hides the Welcome window
'Call HideWelcome
AOL% = findwindow("AOL Frame25", "America  Online")
MDI% = findchildbyclass(AOL%, "MDIClient")
wlcm% = findchildbytitle(MDI%, "Welcome, ")
'wlcm% = FindChildByTitle(MDI%, "Welcome, " & SnoW_User & "!")
x = ShowWindow(wlcm%, SW_MINIMIZE)
End Sub
Sub WinHide(What As Integer)
'This Sub will hide a target window.
Q% = ShowWindow(What, SW_HIDE)
DoEvents
End Sub
Sub WinShow(What As Integer)
'This Sub will Show a target window.
Q% = ShowWindow(What, SW_SHOW)
DoEvents
End Sub

Sub WriteINI(Appname, KeyName, NewString, FileName)
r = WritePrivateProfileString(Appname, KeyName, NewString, FileName)
End Sub


Sub FileAdd(File$, WhatToAdd$)

If FileExists(File$) = False Then Exit Sub
Open File$ For Append As #1
    Print #1, WhatToAdd$
    Close #1
End Sub

