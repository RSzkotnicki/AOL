Attribute VB_Name = "Jaguar32"
'Jaguar32.Bas (For Visual Basic Versions 4, 5, 6)
'For use with Aol (95 and 4.0)
'Release Date: 9/16/98
'Please note disclaimer at bottom.

'Creators contact addresses:
'Jaguar     Jaguar32X@Juno.com (Can send and recieve files here)
'Jaguar     PuRRBaLL@aol.com (AIM Account)
'Flux        dr_flux@hotmail.com
'VSTD      VSTD COORD@aol.com

'Creators:
'Jaguar
'Flux
'VSTD COORD
'Baron
'Genghis

'Jaguar32 News and Updates:
'Flux has now been officially added as a creator in Jaguar32.
'His work has contributed to the next generation of XAOL4 subs.
'Myself and Flux, have decided to make a Gold Series of the
'Jaguar32 bas file. Thus, we will have bas files
'for AOL 2.5, 3.0, 95, and 4.0, but thats way down the road.
'I can now send and recieve files at Jaguar32X@Juno.com so I can
'send you the latest version or if you need help you can send your
'forms and I can fix them and send them back or whatever. Also I
'now have an AIM account which I can sign on at remote locations.
'If you want, you can add PuRRBaLL11 to your buddylist. Thats me.
'Thats about it we may update it in the future probably a month from
'now.

'A note from Jaguar about disclaimer:
'Programmers of the Jaguar32 team have decided that since we make
'and distribute Jaguar32 as freeware and that we do not have a charge
'for Jaguar32 that we strictly enforce the disclaimer. In other words,
'we spent a hell of a lot of time bringing you Jaguar32 and its a aid
'to you in your programming so we ask that you follow the disclaimer.

'Dislcaimer:
'Thank you for choosing Jaguar32.Bas. Before you start using this bas
'file there are a couple of things we have all agreed on. First off, we do
'not want anyone to add on to this bas file and call it theirs. Make your
'own fucking bas file from scratch like we did! Second, please do not
'tamper with the code that is in the bas file unless you have emailed
'Jaguar and I have said it was ok. Third, if someone wants Jaguar32
'please tell them to email Jaguar at Jaguar32X@Juno.com. Fourth,
'unless you have the expressed written permission by one of the
'creators, we do not want to see this bas file on servers or mms
'because it is constantly being updated, and we have the latest
'update. Basically cause we have to send the thing out again to
'people that we have sent it to a million times. Thats about it. Later.

'PS. The Addroom (not the one for AOL4) works only on AOL95
'not AOL3.0

'Special Thanks Given To:
'All of Jaguar, Genghis, Baron, VSTD, and Flux's friends (you know who you are)
'All of Teel
'Cougar & Panther & Leopard
'People who dont decompile and steal codes.
'People who prog their ass off and dont get any credit for what they do.

'Enjoy Jaguar32!
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
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
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ALIAS_START = 0
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

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
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
Global Title
Global Buff3
Global buff2
Global Buff
Global ct
Global RoomHits
Global Log
Global thelist
Global r&
Global entry$
Global iniPath$
Global mmlastline
Function AOLActivate()
x = GetCaption(AOLWindow)
AppActivate x
End Function
Sub AOLChatPunter(SN1 As TextBox, Bombs As TextBox)
'This will see if somebody types /Punt: in a chat
'room...then punt the SN they put.
On Error GoTo errhandler
GINA69 = AOLGetUser
GINA69 = UCase(GINA69)

heh$ = AOLLastChatLine
heh$ = UCase(heh$)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
Pause (0.3)
SN = Mid(naw$, InStr(naw$, ":") + 1)
SN = UCase(SN)
Pause (0.3)
pntstr = Mid$(naw$, 1, (InStr(naw$, ":") - 1))
Gina = pntstr
If Gina = "/PUNT" Then
SN1 = SN
If SN1 = GINA69 Or SN1 = " " + GINA69 Or SN1 = "  " + GINA69 Or SN1 = "   " + GINA69 Or SN1 = "     " + GINA69 Or SN1 = "      " + GINA69 Then
SN1 = AOLGetSNfromCHAT
    AOLChatSend "· ···•(\›•    SouthPark Punter Final"
    AOLChatSend "· ···•(\›•    I can't punt myself BITCH!"
    AOLChatSend "· ···•(\›•    Now U Get PUNTED!"
    GoTo JAKC
    Pause (1)
Exit Sub
End If
    GoTo SendITT
Else
    Exit Sub
End If
SendITT:
AOLChatSend "· ···•(\›•    SouthPark Punter Final"
AOLChatSend "· ···•(\›•    Request Noted"
AOLChatSend "· ···•(\›•    Now Punting - " + SN1
AOLChatSend "· ···•(\›•    Punting With - " + Bombs + " IMz"
JAKC:
Call AOLIMOff
Do
Call AOLInstantMessage(SN1, "       ")
Bombs = Str(Val(Bombs - 1))
If FindWindow("#32770", "America Online") <> 0 Then Exit Sub: MsgBox "This User is not currently signed on, or his/her IMz are Off."
Loop Until Bombs <= 0
Call AOLIMsOn
Bombs = "10"
errhandler:
    Exit Sub
End Sub
Public Sub AOLChatSend2(txt As TextBox)
'This scrolls a multilined textbox adding pauses where needed
'This is basically for macro shops and things like that.
AOLChatSend "· ···•(\›• INCOMMING TEXT"
Pause 4
Dim onelinetxt$, x$, start%, i%
start% = 1
fa = 1
For i% = start% To Len(txt.text)
x$ = Mid(txt.text, i%, 1)
onelinetxt$ = onelinetxt$ + x$
If Asc(x$) = 13 Then
AOLChatSend ": " + onelinetxt$
Pause (0.5)
J% = J% + 1
i% = InStr(start%, txt.text, x$)
If i% >= Len(txt.text) Then Exit For
start% = i% + 1
onelinetxt$ = ""
End If
Next i%
AOLChatSend ":" + onelinetxt$
End Sub


Public Sub AddRoom_SNs(Listboxes As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6
PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)
PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
Listboxes.AddItem PERSON$
Next index
Call CloseHandle(AOLProcessThread)
End If
End Sub


Sub AOLFakeOH()
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
Pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth ßy ·Labyrinth" & Chr(4) & "/-=!=-\"
End Sub
Function CutStringDown(ByVal pws As String) As String
On Error Resume Next
CutStringDown = Trim(Left$(pws, InStr(pws, Chr$(0)) - 1))
End Function
Function Juno_Activate()
x = GetCaption(JunoWindow)
AppActivate x
End Function
Function AOLGotoPrivateRoom(Room As String)
Theroomcode = "aol://2719:2-2-" & Room
AOLKeyword (Theroomcode)
End Function

Function AOLFindIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
AOLFindIM = IM%
End Function











Function AOLSupRoom()
AOLIsOnline
If AOLIsOnline = 0 Then GoTo last
AOLFindRoom
If AOLFindRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6

PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)

PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
Call AOLChatSend("~Genghis~ Sup, " & PERSON$)
Pause (0.4)
Next index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function




Function findaol()
aol% = FindWindow("AOL Frame25", vbNullString)
findaol = aol%
End Function

Function FindKeyword()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(mdi%, "Keyword")
kedit% = FindChildByClass(keyw%, "_AOL_Edit")
FindKeyword = kedit%
End Function

Function FindNewIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
 End Function

Function FindWelcome()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
FindWelcome = FindChildByTitle(mdi%, "Welcome, ")
End Function

Function AOLIMRoomIMer(mess As String)
AOLIsOnline
If AOLIsOnline = 0 Then GoTo last


On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6

PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)

PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
Call AOLInstantMessage(PERSON$, mess)
Next index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function

Function Juno_Tab()
JunoTab = FindChildByClass(JunoWindow, "#32770")
End Function

Function Mail_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
KillWin (A3000%)
End Function

Function Mail_DeleteSent()
Call AOLRunMenuByString("Check Mail You've &Sent")

aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
again:
Pause (1)
A3000% = FindChildByTitle(A2000%, "Outgoing Mail")
If A3000% = 0 Then GoTo again
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
Pause (6)
AOLButton (Delete%)
KillWin (A3000%)
End Function

Function Mail_KeepAsNew()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Keepasnew% = FindChildByTitle(A3000%, "Keep As New")
AOLButton (Keepasnew%)
End Function


Function Mail_DeleteSingle()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
AOLButton (Delete%)
End Function


Function Mail_FindComposed()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi = FindChildByClass(aol%, "MDIClient")
Mail_FindComposed = FindChildByTitle(mdi, "Compose Mail")
End Function

Function Mail_ForwardMail(SN As String, message As String)
FindForwardWindow
PERSON = SN
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Fwd: ")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)
End Function

Function GetFromINI(Appname$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(Appname$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function



Function Mail_ClickForward()
x = FindOpenMail
If x = 0 Then GoTo last
AOLActivate
SendKeys "{TAB}"
AG:
Pause (0.2)
SendKeys " "
x = FindSendWin(2)
If x = 0 Then GoTo AG
last:
End Function

Function Mail_KillComposed()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi = FindChildByClass(aol%, "MDIClient")
Composed = FindChildByTitle(mdi, "Compose Mail")
KillWin (Composed)
End Function
Function MAil_BuildList(lst As ListBox)
AOLMDI
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then GoTo Justamin
Pause (7)
End If

MailWin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(MailWin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    x = SendMessageByString(MailTree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    lst.AddItem buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo start
last:

End Function
Function Mail_ListMail(Box As ListBox)
Box.Clear
AOLMDI
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then GoTo Justamin
Pause (7)
End If

MailWin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(MailWin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    x = SendMessageByString(MailTree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    Box.AddItem buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo start
last:
End Function

Function Mail_Out_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
KillWin (A3000%)
End Function

Function Mail_Out_CursorSet(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function






Function Mail_Out_ListMail(Box As ListBox)
Box.Clear
AOLMDI
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then GoTo Justamin
Pause (7)
End If

MailWin = FindChildByTitle(AOLMDI, "Outgoing FlashMail")
AOLCountMail
start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(MailWin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    x = SendMessageByString(MailTree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    Box.AddItem buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo start
last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
theMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(theMail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
x = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
x = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
x = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
x = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function

Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function


Function READFILE(Where As String)
Filenum = FreeFile
Open (Where) For Input As Filenum
Info = Input(LOF(Filenum), Filenum)
Info = READFILE
End Function




Function Text_backwards(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let Newsent$ = NextChr$ & Newsent$
Loop
Text_backwards = Newsent$
End Function
Function Text_Elite(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed
If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "|-|"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "]V["
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "ö"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "<" Then Let NextChr$ = "«"
If NextChr$ = ">" Then Let NextChr$ = "»"
If NextChr$ = "*" Then Let NextChr$ = "¤"
If NextChr$ = "`" Then Let NextChr$ = "“"
If NextChr$ = "'" Then Let NextChr$ = "”"
If NextChr$ = "0" Then Let NextChr$ = "º"
Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Text_Elite = Newsent$
End Function
Function Text_Hacker(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Hacker = Newsent$
End Function
Function Text_Spaced(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Spaced = Newsent$
End Function
Function Text_Khangolian(txt As String)
'This translats text into my lang(Khangolian)
Dim Firstletter, LastLetter, Middle
txtlen = Len(txt)
Firstletter = Left$(txt, 1)
LastLetter = Right$(txt, 1)
Middle = NotSure
withnofirst = Right$(txt, txtlen - 1)
nofirstlen = Len(withnofirst)
Withnofirstorlast = Left$(withnofirst, nofirstlen - 1)
Text_Encode = LastLetter & Withnofirstorlast & Firstletter
End Function

Function Text_Looping(txt As String)
Dim thecaption, Captionlen, middlelen, Firstletter, Middle
If txt = "" Then GoTo dead
thecaption = txt
Captionlen = Len(thecaption)
middlelen = Captionlen - 1
Firstletter = Left$(thecaption, 1)
Middle = Right(thecaption, middlelen)
Text_Looping = Middle & Firstletter
GoTo last
dead:
Text_Looping = ""
last:
End Function

Function Text_StripLetter(txt As String, which As String)
'This takes out a certain letter
'Which is the letter you take out(its in number value)
'For example..in the work Khan if I wanted to
'take out the H I would use
'Text_StripLetter("Khan", 2)
txtlen = Len(txt)
before = Left$(txt, which - 1)
MsgBox before
beforelen = Len(before)
afterthat = txtlen - beforelen - 1
After = Right$(txt, afterthat)
MsgBox After
Text_StripLetter = before & After
End Function

Public Sub TextColor_Blue(txt As TextBox)
txt.ForeColor = &HFFFF00
Pause 0.1
txt.ForeColor = &HFF0000
Pause 0.1
txt.ForeColor = &HC00000
Pause 0.1
txt.ForeColor = &H800000
Pause 0.1
txt.ForeColor = &H400000
Pause 0.1
End Sub

Public Sub TextColor_Teal(txt As TextBox)
txt.ForeColor = &HFFFF00
Pause 0.1
txt.ForeColor = &HC0C000
Pause 0.1
txt.ForeColor = &H808000
Pause 0.1
txt.ForeColor = &H404000
Pause 0.1
End Sub

Public Sub TextColor_Green(txt As TextBox)
txt.ForeColor = &HFF00&
Pause 0.1
txt.ForeColor = &HC000&
Pause 0.1
txt.ForeColor = &H8000&
Pause 0.1
txt.ForeColor = &H4000&
Pause 0.1
End Sub

Public Sub TextColor_Yellow(txt As TextBox)
txt.ForeColor = &HFFFF&
Pause 0.1
txt.ForeColor = &HC0C0&
Pause 0.1
txt.ForeColor = &H8080&
Pause 0.1
txt.ForeColor = &H4040&
Pause 0.1
End Sub


Public Sub TextColor_Red(txt As TextBox)
txt.ForeColor = &HFF&
Pause 0.1
txt.ForeColor = &HC0&
Pause 0.1
txt.ForeColor = &H80&
Pause 0.1
txt.ForeColor = &H40&
Pause 0.1
End Sub

Function Text_TurnToUpperCase(txt As String)
Text_TurntoUCase = UCase(txt)
End Function

Function Text_TurnToLowerCase(txt As String)
Text_TurntoLCase = LCase(txt)
End Function
Sub ULAlign(Frm As Form)
    Dim x, Y                    ' New top, left for the form
    Y = 0
    x = 0
    Frm.Move x, Y             ' Change location of the form

End Sub

Sub playwav(File)
SoundName$ = File
SoundFlags& = &H20000 Or &H1
snd& = sndPlaySound(SoundName$, SoundFlags&)
End Sub


Sub AOLChangeCaption(newcaption As String)
Call AOLSetText(AOLWindow(), newcaption)
End Sub

Sub AOLBuddyBLOCK(SN As TextBox)
BUDLIST% = FindChildByTitle(AOLMDI(), "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_HWNDNEXT)
setup% = GetWindow(IM1%, GW_HWNDNEXT)
AOLIcon (setup%)
Pause (2)
STUPSCRN% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(Edit%, GW_HWNDNEXT)
view% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(view%, GW_HWNDNEXT)
AOLIcon PRCYPREF%
Pause (1.8)
Call AOLKillWindow(STUPSCRN%)
Pause (2)
PRYVCY% = FindChildByTitle(AOLMDI(), "Privacy Preferences")
DABUT% = FindChildByTitle(PRYVCY%, "Block only those people whose screen names I list")
AOLButton (DABUT%)
DaPERSON% = FindChildByClass(PRYVCY%, "_AOL_EDIT")
Call AOLSetText(DaPERSON%, SN)
Creat% = FindChildByClass(PRYVCY%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
AOLIcon Edit%
Pause (1)
Save% = GetWindow(Edit%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
AOLIcon Save%
End Sub
Public Sub AOLKillWindow(Windo)
x = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Sub XAOL4_15Liner(txt As String)
'Max of 14 chr or else u get Msg is too long
Call XAOL4_SetFocus
a = String(116, Chr(32))
d = 116 - Len(txt)
c$ = Left(a, d)
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.8
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.8
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.8
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.8
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.3
XAOL4_ChatSend "" + txt + "" & c$ & "" + txt + ""
Pause 0.8
End Sub
Sub XAOL4_AntiIdle()
'Sub contributed and written by ieet xero.
'If you would like to contact ieet xero,
'please email Jaguar at Jaguar32X@Juno.com
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
AOLIcon (AOIcon%)
End Sub

Public Sub XAOL4_AddRoom(Listboxes As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = XAOL4_FindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6
PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)
PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
Listboxes.AddItem PERSON$
Next index
Call CloseHandle(AOLProcessThread)
End If
End Sub


Sub XAOL4_BuddyVIEW()
Call XAOL4_Keyword("Buddy View")
End Sub
Sub XAOL4_BudList(lst As ListBox)
'This adds the AOL Buddy List to a VB listbox
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = FindChildByTitle(AOLMDI(), "Buddy List Window")
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6
PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)
PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
lst.AddItem PERSON$
Next index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub XAOL4_ChangeCaption(newcaption As String)
Call AOLSetText(AOLWindow(), newcaption)
End Sub
Sub XAOL4_ChatManipulator(Who$, What$)
'This makes the chat room text near the VERY TOP
'what u want
view% = FindChildByClass(XAOL4_FindRoom(), "RICHCNTL")
Buffy$ = Chr$(13) & Chr$(10) & "" & (Who$) & ":" & Chr$(9) & "" & (What$) & ""
x% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub XAOL4_ChatSend(txt)
    Room% = XAOL4_FindRoom()
    If Room% Then
        hChatEdit% = Find2ndChildByClass(Room%, "RICHCNTL")
        ret = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, txt)
        ret = SendMessageByNum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub
Function Find2ndChildByClass(parentw, childhand)
'DO NOT TAMPER WITH THIS CODE!
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    While firs%
        firs% = GetWindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Wend
    Find2ndChildByClass = 0
Found:
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = firs%
End Function

Sub XAOL4_ClearChat()
childs% = XAOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Function XAOL4_CountMail()
'This is an advanced mail counter. Don't change the code.
Dim PASS As Integer
PASS = 0
begin:
theMail% = FindChildByTitle(AOLMDI(), XAOL4_GetUser & "'s Online Mailbox")
If theMail% = 0 Then
Call XAOL4_MailReadNew
GoTo begin
End If
PASS = PASS + 1
If PASS <> 10 Then
GoTo begin
Else
tabcont% = FindChildByClass(theMail%, "_AOL_TabControl")
tabpage% = FindChildByClass(tabcont%, "_AOL_TabPage")
thetree% = FindChildByClass(tabpage%, "_AOL_Tree")
XAOL4_CountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
AOLClose (theMail%)
End If
End Function

Function XAOL4_FindRoom()
'Finds the chat room and sets focus on it
    aol% = FindWindow("AOL Frame25", vbNullString) '   MDI% = FindChildByClass(AOL%, "MDIClient")
    mdi% = FindChildByClass(aol%, "MDIClient")
    firs% = GetWindow(mdi%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    listere% = FindChildByClass(firs%, "RICHCNTL")
    listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (L <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            listers% = FindChildByClass(firs%, "RICHCNTL")
            listere% = FindChildByClass(firs%, "RICHCNTL")
            listerb% = FindChildByClass(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            L = L + 1
    Loop
    If (L < 100) Then
       XAOL4_FindRoom = firs%
       Exit Function
     End If
End Function

Function XAOL4_FindToolbar()
toolbar% = FindChildByClass(AOLWindow, "AOL Toolbar")
toolbar2% = FindChildByClass(toolbar%, "_AOL_Toolbar")
XAOL4_FindToolbar = toolbar2%
End Function
Function XAOL4_GetChat()
'This gets all the txt from chat room
childs% = XAOL4_FindRoom
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
XAOL4_GetChat = theview$
End Function
Function XAOL4_GetCurrentRoomName()
XAOL4_GetCurrentRoomName = GetCaption(XAOL4_FindRoom)
End Function
Function XAOL4_GetUser()
On Error Resume Next
Welcome% = FindChildByTitle(AOLMDI, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
XAOL4_GetUser = user
End Function
Sub XAOL4_Hide()
a = ShowWindow(AOLWindow(), SW_HIDE)
End Sub
Sub XAOL4_IMOff()
Call XAOL4_InstantMessage("$IM_OFF", "Jaguar32")
waitforok
End Sub
Sub XAOL4_IMOn()
Call XAOL4_InstantMessage("$IM_ON", "Jaguar32")
waitforok
End Sub
Sub XAOL4_InstantMessage(PERSON, message)
Call XAOL4_Keyword("aol://9293:" & PERSON)
Pause (2)
Do
DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
Loop Until (IM% <> 0 And aolrich% <> 0 And imsend% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, message)
For Sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next Sends
AOLIcon imsend%
If IM% Then Call AOLKillWindow(IM%)
End Sub

Sub XAOL4_Keyword(txt)
    aol% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(aol%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub
Sub XAOL4_LocateMember(name As String)
Call XAOL4_Keyword("aol://3548:" + name)
End Sub
Sub XAOL4_MailReadNew()
mailicon% = FindChildByClass(XAOL4_FindToolbar, "_AOL_Icon")
AOLIcon (mailicon%)
End Sub
Sub XAOL4_Mail(PERSON, SUBJECT, message)
Const LBUTTONDBLCLK = &H203
aol% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(tool2%, "_AOL_Icon")
icon2% = GetWindow(ico3n%, 2)
x = SendMessageByNum(icon2%, WM_LBUTTONDOWN, 0&, 0&)
x = SendMessageByNum(icon2%, WM_LBUTTONUP, 0&, 0&)
Pause (4)
    aol% = FindWindow("AOL Frame25", vbNullString)
    mdi% = FindChildByClass(aol%, "MDIClient")
    mail% = FindChildByTitle(mdi%, "Write Mail")
    aoledit% = FindChildByClass(mail%, "_AOL_Edit")
    aolrich% = FindChildByClass(mail%, "RICHCNTL")
    subjt% = FindChildByTitle(mail%, "Subject:")
    subjec% = GetWindow(subjt%, 2)
        Call AOLSetText(aoledit%, PERSON)
        Call AOLSetText(subjec%, SUBJECT)
        Call AOLSetText(aolrich%, message)
e = FindChildByClass(mail%, "_AOL_Icon")
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
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
AOLIcon (e)
End Sub

Public Sub XAOL4_MassIM(lst As ListBox, txt As TextBox)
lst.Enabled = False
i = lst.ListCount - 1
lst.ListIndex = 0
For x = 0 To i
lst.ListIndex = x
Call XAOL4_InstantMessage(lst.text, txt.text)
Pause (1)
Next x
lst.Enabled = True
End Sub
Sub XAOL4_OpenChat()
XAOL4_Keyword ("PC")
End Sub
Sub XAOL4_OpenPR(PRrm As String)
Call XAOL4_Keyword("aol://2719:2-2-" & PRrm)
End Sub
Sub XAOL4_Punter(SN As TextBox, Bombz As TextBox)
Call XAOL4_IMOff
waitforok
Do
DoEvents:
Call XAOL4_InstantMessage(SN, "<h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h3><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3>")
DAWIN% = FindWindow("#32770", "America Online")
If DAWIN% Then Exit Sub: MsgBox "Sorry person isn't online!", 48, "DAMMIT!"
Bombz = Str(Val(Bombz - 1))
Loop Until Bombz <= 0
Call XAOL4_IMOn
waitforok
End Sub
Sub XAOL4_Read1Mail()
'This will read the very first mail in the User's box
theMail% = FindChildByTitle(AOLMDI(), XAOL4_GetUser & "'s Online Mailbox")
If theMail% = 0 Then
Exit Sub
End If
e = FindChildByClass(theMail%, "_AOL_Icon")
AOLIcon (e)
End Sub
Function XAOL4_RoomCount()
thechild% = XAOL4_FindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")
getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
XAOL4_RoomCount = getcount
End Function
Sub XAOL4_SetFocus()
x = GetCaption(AOLWindow)
AppActivate x
End Sub
Sub XAOL4_SignOff()
Call RunMenuByString(AOLWindow(), "Sign Off")
End Sub
Function XAOL4_SpiralScroll(txt As String)
Dim AODCOUNTER, a, thetxtlen
AODCOUNTER = 1
thetxtlen = Len(txt)
start:
a = a + 1
If a = thetxtlen Then GoTo last
x = Text_Looping(txt)
txt = x
XAOL4_ChatSend x
Pause (0.5)
AODCOUNTER = AODCOUNTER + 1
If AODCOUNTER = 4 Then
   AODCOUNTER = 2
   End If
GoTo start
last:

End Function

Sub XAOL4_UnHide()
a = ShowWindow(AOLWindow(), SW_SHOW)
End Sub
Sub XAOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(aolmod%, SW_RESTORE)
Call XAOL4_SetFocus
End Sub
Function XAOL4_UpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(die%, SW_HIDE)
x = ShowWindow(die%, SW_MINIMIZE)
Call XAOL4_SetFocus
End Function
Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
Function AOLGetTopWindow()
AOLGetTopWindow = GetTopWindow(AOLMDI())
End Function

Sub AOLSetFocus()
'SetFocusAPI doesn't work AOL because AOL has added
'a safeguard against other programs calling certain
'API functions (like owner-drawn things and like.)
'This is the only way known for setting the focus
'to AOL.  This is a normal VB command!
x = GetCaption(AOLWindow)
AppActivate x
End Sub
Public Sub AOLMassIM(lst As ListBox, txt As TextBox)
lst.Enabled = False
i = lst.ListCount - 1
lst.ListIndex = 0
For x = 0 To i
lst.ListIndex = x
Call AOLInstantMessage(lst.text, txt.text)
Pause 0.5
Next x
lst.Enabled = True
End Sub
Public Sub AOLOnlineChecker(PERSON)
Call AOLInstantMessage4(PERSON, "Sup?")
Pause 2
AOLIMScan
End Sub
Public Sub AddRoom_ByBox(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindChildByTitle(AOLMDI, "Who's Chatting")
If Room = 0 Then MsgBox "Not Open"
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
OLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6

PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)

PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
LOP = Len(PERSON$)
PERSON$ = Right$(PERSON$, LOP - 2)
PERSON$ = PERSON$ & "@AOL.COM"
Listboxes.AddItem PERSON$
Next index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub AddRoom(lst As ListBox)
Dim index As Long
Dim i As Integer
For index = 0 To 25
namez$ = String$(256, " ")
ret = AOLGetList(index, namez$)
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)))
ADD_AOL_LB namez$, lst
Next index
end_addr:
lst.RemoveItem lst.ListCount - 1
i = GetListIndex(lst, AOLGetUser())
If i <> -2 Then lst.RemoveItem i
End Sub

Public Sub AddRoom_WithExt(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6

PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)

PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
PERSON$ = PERSON$ & "@AOL.COM"
Listboxes.AddItem PERSON$
Next index
Call CloseHandle(AOLProcessThread)
End If
End Sub


Function ListToList(Source, destination)
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, buffer$)
addstrings% = SendMessageByString(destination, LB_ADDSTRING, 0, buffer$)
Next Adding

End Function

Function MouseOverHwnd()
    ' Declares
      Dim pt32 As POINTAPI
      Dim ptx As Long
      Dim pty As Long
   
      Call GetCursorPos(pt32)               ' Get cursor position
      ptx = pt32.x
      pty = pt32.Y
      MouseOverHwnd = WindowFromPointXY(ptx, pty)    ' Get window cursor is over
End Function

Function UntilWindowClass(Parent, news$)
Do: DoEvents
e = FindChildByClass(Parent, news$)
Loop Until e
UntilWindowClass = e
End Function


Function UntilWindowTitle(Parent, news$)
Do: DoEvents
e = FindChildByTitle(Parent, news$)
Loop Until e
UntilWindowTitle = e
End Function
Public Function AOLGetList(index, buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6

PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)

PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

buffer$ = PERSON$
End Function





Function AddListToString(thelist As ListBox)
For DoList = 0 To thelist.ListCount - 1
AddListToString = AddListToString & thelist.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)

End Function

Function AddListToMailString(thelist As ListBox)
If thelist.List(0) = "" Then GoTo last
For DoList = 0 To thelist.ListCount - 1
AddListToMailString = AddListToMailString & "(" & thelist.List(DoList) & "), "
Next DoList
AddListToMailString = Mid(AddListToMailString, 1, Len(AddListToMailString) - 2)
last:
End Function
Function ScrambleText(thetext, lbl)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Scrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)

'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe

'adds the scrambled text to the full scrambled element
cityz:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = scrambled$
lbl.Caption = ScrambleText
Exit Function
End Function

Public Sub AOLGetCurrentRoomName()
x = GetCaption(AOLFindRoom())
MsgBox x
End Sub
Function SearchForSelected(lst As ListBox)
If lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

start:
counterf = counterf + 1
If lst.ListCount = counterf + 1 Then GoTo last
If lst.Selected(counterf) = True Then GoTo last
If couterf = lst.ListCount Then GoTo last
GoTo start

last:
SearchForSelected = counterf
End Function
Function Juno_Window()
jun% = FindWindow("Afx:b:152e:6:386f", vbNullString)
JunoWindow = jun%
End Function


Sub AddStringToList(theitems, thelist As ListBox)
If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
thelist.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub

Sub Scroll8Line(txt As TextBox)
lonh = String(116, Chr(32))
d = 116 - Len(text1)
c$ = Left(lonh, d)
AOLChatSend ("" & txt & c$ & text1)
AOLChatSend ("" & txt & c$ & text1)
lonh = String(116, Chr(32))
d = 116 - Len(text1)
c$ = Left(lonh, d)
AOLChatSend ("" & txt & c$ & text1)
AOLChatSend ("" & txt & c$ & text1)
End Sub


Function AOLClickList(hwnd)
clicklist% = SendMessageByNum(hwnd, &H203, 0, 0&)
End Function

Function AOLCountMail()
theMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(theMail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function AOLGetListString(Parent, index, buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

aolhandle = Parent

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6

PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)

PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

buffer$ = PERSON$
End Function

Sub AOLHide()
a = ShowWindow(AOLWindow(), SW_HIDE)
End Sub

Sub AOLOpenChat()
If AOLFindRoom() Then Exit Sub
AOLKeyword ("pc")
Do: DoEvents
Loop Until AOLFindRoom()

End Sub
Public Sub AOLOpenNewMail()
Call AOLRunMenuByString("Read &New Mail")
End Sub


Public Sub AOLOpenOLDMail()
Call AOLRunMenuByString("Check Mail You've &Read")
End Sub
Public Sub AOLOpenSentMail()
Call AOLRunMenuByString("Check Mail You've &Sent")
End Sub
Public Sub AOLSignOnCaption(newcaption As String)
setup% = FindChildByTitle(AOLMDI(), "Welcome")
Call AOLSetText(setup%, newcaption)
End Sub
Sub AOLRespondIM(message)
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Z
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Z
Exit Sub
Z:
e = FindChildByClass(IM%, "RICHCNTL")

e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e2 = GetWindow(e, 2) 'Send Text
e = GetWindow(e2, 2) 'Send Button
Call AOLSetText(e2, message)
AOLIcon (e)
Pause 4
KillWin (IM%)
End Sub

Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub


Sub AOLUnHide()
a = ShowWindow(AOLWindow(), SW_SHOW)
End Sub

Sub AOLWaitMail()
MailWin% = GetTopWindow(AOLMDI())
aoltree% = FindChildByClass(MailWin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
Pause (10)
secondcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub


Function EncryptType(text, types)
'to encrypt, example:
'encrypted$ = EncryptType("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = EncryptType("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(text)
If types = 0 Then
Current$ = Asc(Mid(text, God, 1)) - 1
Else
Current$ = Asc(Mid(text, God, 1)) + 1
End If
Process$ = Process$ & Chr(Current$)
Next God

EncryptType = Process$
End Function

Function FindChildByTitle(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, buffer$, (hwndLength% + 1))

If InStr(UCase(buffer$), UCase(child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function FindChildByClass(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, buffer$, 250)

If InStr(UCase(buffer$), UCase(child)) Then FindChildByClass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function


Function DescrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Descrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo city
lastchar$ = Mid(chars$, 2, 1)
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 3, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
city:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
scrambled$ = scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed buffers
sniff:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = scrambled$

End Function



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

Sub HideWindow(hwnd)
hi = ShowWindow(hwnd, SW_HIDE)
End Sub


Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

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



Sub MaxWindow(hwnd)
ma = ShowWindow(hwnd, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hwnd)
MI2 = ShowWindow(hwnd, SW_MINIMIZE)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
'turns the "number" so vb recognizes it for
'addition, subtraction, ect.
End Function

Sub ParentChange(Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub


Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function ReverseText(text)
For Words = Len(text) To 1 Step -1
ReverseText = ReverseText & Mid(text, Words, 1)
Next Words


End Function

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

Sub AOLRunTool(tool)
toolbar% = FindChildByClass(AOLWindow(), "AOL Toolbar")
iconz% = FindChildByClass(toolbar%, "_AOL_Icon")
For x = 1 To tool - 1
iconz% = GetWindow(iconz%, 2)
Next x
isen% = IsWindowEnabled(iconz%)
If isen% = 0 Then Exit Sub
AOLIcon (iconz%)
End Sub

Function ScrambleGame(thestring As Integer)
Dim bytestring As String
thestringcount = Len(thestring)
If Not Mid(thestring, thestringcount, 1) = " " Then thestring = thestring & " "
For Stringe = 1 To Len(thestring)
characters$ = Mid(thestring, Stringe, 1)
thestrings$ = thestrings$ & characters$

If characters$ = " " Then
smoked:

DoEvents
For Ensemble = 1 To Len(thestrings$) - 1
Randomize
randomstring = Int((Len(thestrings$) * Rnd) + 1)
If randomstring = Len(thestrings$) Then GoTo already
If bytesread Like "*" & randomstring & "*" Then GoTo already
stringrandom$ = Mid(thestrings$, randomstring, 1)
stringfound$ = stringfound$ & stringrandom$
bytesread = bytesread & randomstring
GoTo really
already:
Ensemble = Ensemble - 1
really:
Next Ensemble
If stringfound$ = thestrings$ Then stringfound$ = "": GoTo smoked
thestrings2$ = thestrings2$ & stringfound$ & " "
stringfound$ = ""
thestrings$ = ""
bytesread = ""
strngfound$ = ""
End If

Next Stringe
ScrambleGame = Mid(thestrings2$, 1, Len(thestring) - 1)
End Function


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


Sub SetBackPre()
Call RunMenuByString(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
AOLIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 0, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

AOLButton (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function AOLStayOnline()
hwndz% = FindWindow(AOLWindow(), "America Online")
childhwnd% = FindChildByTitle(hwndz%, "OK")
AOLButton (childhwnd%)
End Function

Public Sub CenterCorner(frmForm As Form)
'This will center you form in the upper right
'of the users screen
   With frmForm
      .Left = (Screen.Width - .Width) / 1
      .Top = (Screen.Height - .Height) / 2000
   End With
End Sub
Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function
Function TrimCharacter(thetext, chars)
TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

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


Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")
End Function



Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindFwdWin = firs%

Exit Function
begis:
FindFwdWin = firss%
End Function


Function FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
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


Function KTEncrypt(ByVal password, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(password) = 0 Then Error 31100
  
  'Is password too long
  If Len(password) > 255 Then Error 31100

  'Is there a strng$ to work with?
  If Len(strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
          strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
        End If
        look = look + 1
      Loop
      
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
    End If
  Else
    'force% flag set, ecrypt string regardless of tag
    EncryptFlag% = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(password)
  
  
  'Tack on leading characters to prevent repetative recognition
  password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
  password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
  password = password + Chr$(Asc(Right$(password, 1)) Xor PassMax)
  password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(tochange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, strng, Chr$(1))
      If look > 0 Then
        strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(strng, Chr$(0))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(strng, Chr$(10))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(strng, Chr$(13))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(strng, Chr$(26))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      strng = Mid$(strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function

Public Sub CenterForm(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub

Public Sub CenterFormTop(Frm As Form)
   With Frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub

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

Public Sub AOLButton(but%)
clickicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
clickicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function AOLIMSTATIC(newcaption As String)
ANTI1% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
STS% = FindChildByClass(ANTI1%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call ChangeCaption(st%, newcaption)
End Function

Function AOLGetUser()
On Error Resume Next
aol& = FindWindow("AOL Frame25", "America  Online")
mdi& = FindChildByClass(aol&, "MDIClient")
Welcome% = FindChildByTitle(mdi&, "Welcome, ")
WelcomeLength& = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength& + 1))
user$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = user$
End Function

Sub AOLIMOff()
Call AOLInstantMessage("$IM_OFF", "TheKhan")
End Sub

Sub AOLIMsOn()
Call AOLInstantMessage("$IM_ON", "TheKhan")

End Sub


Sub AOLChatSend(txt)
Room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(Room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(Room%, "_AOL_Edit"), 13)
'A1000% = FindChildByClass(Room%, "_AOL_Edit")
'A2000% = GetWindow(A1000%, 2)
'AOLIcon (A2000%)

End Sub


Sub AOLClose(winew)
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub
Public Sub AOLEightLine(txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(txt)
c$ = Left(a, d)
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.3
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.3
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.3
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.7
End Sub

Sub AOLCursor()
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Function AOLFindRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function
Function FindOpenMail()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "RICHCNTL")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function

Function FindForwardWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByTitle(childfocus%, "Send Now")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function



Function AOLGetChat()
childs% = AOLFindRoom()
child = FindChildByClass(childs%, "_AOL_View")


GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$
AOLGetChat = theview$
End Function

Sub ADD_AOL_LB(itm As String, lst As ListBox)
If lst.ListCount = 0 Then
lst.AddItem itm
Exit Sub
End If
Do Until XX = (lst.ListCount)
Let diss_itm$ = lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub
Sub XAOL4_PRBust(Room As String)
'Sub contributed and written by ieet xero.
'If you would like to contact ieet xero,
'please email Jaguar at Jaguar32X@Juno.com
chat = XAOL4_FindRoom
Do
Call XAOL4_OpenPR(Room$)
waitforok
If chat = Room$ Then GoTo xero
Loop
xero:
End Sub

Sub XAOL4_KillModal()
'Sub contributed and written by ieet xero.
'If you would like to contact ieet xero,
'please email Jaguar at Jaguar32X@Juno.com
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call AOLKillWindow(Modal%)
End Sub
Sub AOLAntiIdle()
aol% = FindWindow("_AOL_Modal", vbNullString)
xstuff% = FindChildByTitle(aol%, "Favorite Places")
If xstuff% Then Exit Sub
xstuff2% = FindChildByTitle(aol%, "File Transfer *")
If xstuff2% Then Exit Sub
yes% = FindChildByClass(aol%, "_AOL_Button")
AOLButton yes%
End Sub

Sub AOLGetMemberProfile(name As String)
AOLRunMenuByString ("Get a Member's Profile")
Pause 0.3
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(mdi%, "Get a Member's Profile")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
End Sub


Function FindIMTextwindow()
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMTextwindow = FindChildByClass(IM%, "RICHCNTL")
End Function
Function FindIMCaption()
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMCaption = FindChildByClass(IM%, "_AOL_Static")
End Function
Function AOLChangeIMCaption(txt As String)
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
IMtext% = FindChildByClass(IM%, "_AOL_Static")
Call ChangeCaption(IMtext%, txt)
End Function
Function Mail_GetErrorMessage()
Errors% = FindChildByTitle(AOLMDI(), "Error")
IMtext% = FindChildByClass(Errors%, "_AOL_VIEW")
Mail_GetErrorMessage = AOLGetText(IMtext%)
End Function

Function MakeSpaceInGoto(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If NextChr$ = " " Then Let NextChr$ = "%20"
Let Newsent$ = Newsent$ + NextChr$
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
MakeSpaceInGoto = Newsent$
End Function

Sub AOLAntiPunter()
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRich% = FindChildByClass(ANT%, "RICHCNTL")
STS% = FindChildByClass(ANT%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call AOLSetText(st%, "SouthPark FINAL - This IM Window Should Remain OPEN.")
mi = ShowWindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRich% <> 0 Then
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
End If
Loop
End Sub
Sub AOLCatWatch()
Do
    Y% = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo lol
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
W = InStr(LCase$(namez$), LCase$("catwatch"))
x = InStr(LCase$(namez$), LCase$("catid"))
If W <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Cat had entered the room."
End If
If x <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Cat had entered the room."
End If
Next index%
lol:
Loop
End Sub
Sub AOLChangeWavDirect(wav As String)
'change the directory of the wav's
'AOLChangeWavDirect("C:\aol30\download")
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub
Public Sub AOLChangeWelcome(newwelcome As String)
Welc% = FindChildByTitle(AOLMDI(), "Welcome, " & AOLGetUser & "!")
Call AOLSetText(Welc%, newwelcome)
End Sub
Public Sub AOLChatManipulator(Who$, What$)
view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "" & (Who$) & ":" & Chr$(9) & "" & (What$) & ""
x% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLClearChatRoom()
'clears the chat room
x$ = Format$(String$(100, Chr$(13)))
Call AOLChatManipulator(" ", x$)
a = String(116, Chr(32))
d = 116 - Len(txt)
c$ = Left(a, d)
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.3
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.3
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.3
AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
Pause 0.7
End Sub

Sub AOLGuideWatch()
Do
    Y = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
x = InStr(LCase$(namez$), LCase$("guide"))
If x <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Guide had entered the room."
End If
Next index%
end_ad:
Loop
End Sub
Sub WaitForLoadedMail()
Do
Box = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail")
Pause (0.1)
Loop Until Box <> 0
List = FindChildByClass(Box, "_AOL_Tree")
Do
DoEvents
MailNum = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call Pause(0.5)
MailNum2 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call Pause(0.5)
MailNum3 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until MailNum = MailNum2 And MailNum2 = MailNum3
    MailNum = SendMessage(List, LB_GETCOUNT, 0, 0&)

End Sub
Sub AOLHostManipulator(What$)
'AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (What$) & ""
x% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLHostNameChange(SN As String)
x = AOLVersion()
If x = "C:\aol30" Then
Open "C:\aol30\tool\aolchat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
ElseIf x = "C:\aol25" Then
Open "C:\aol25\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
End If
End Sub
Function AOLFindChatWindow() As Integer
  Dim genhWnd%
  Dim AOLChildhWnd%
  Dim ChildWnd As Integer
  Dim ControlWnd As Integer
  Dim ChatWnd As Integer
  Dim TargetsFound As Integer
  Dim RetClsName As String * 255
  Dim x%
genhWnd% = GetWindow(FindWindow("AOL Frame25", 0&), GW_CHILD)
Do
  x% = GetClassName(genhWnd%, RetClsName$, 254)
    If InStr(RetClsName$, "MDIClient") Then
      AOLChildhWnd% = genhWnd% 'Child window found!
    End If
  genhWnd% = GetWindow(genhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While genhWnd% <> 0
ChildWnd = GetWindow(AOLChildhWnd%, GW_CHILD)
Do
  ControlWnd = GetWindow(ChildWnd, GW_CHILD)
  Do
    x% = GetClassName(ControlWnd, RetClsName$, 254)

    
    If InStr(RetClsName$, "_AOL_Edit") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_View") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_Listbox") Then
      TargetsFound = TargetsFound + 1:
    End If
    ControlWnd = GetWindow(ControlWnd, GW_HWNDNEXT)
    DoEvents
  Loop While ControlWnd <> 0

  If TargetsFound = 3 Then ChatWnd = ChildWnd: Exit Do

  
  ChildWnd = GetWindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
Chat_FindTheWin = ChatWnd

End Function
Sub Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Function AOLKillDupes()
'Gets rid of excess mail
num = AOLCountMail()
DELNUM% = 0
aol% = FindWindow("AOL Frame25", 0&)
mld% = FindChildByClass(aol%, "MDIClient")
FNMB% = FindChildByTitle(mld%, "New Mail")
If FNMB% = 0 Then
AOLRunTool (1)
Pause (0.4)
Do: DoEvents
    NMB% = FindChildByTitle(mofo%, "New Mail")
    Pause (0.1)
Loop Until NMB% <> 0
WaitForLoadedMail
End If

Do: DoEvents
LSTTXT$ = ","
DELTXT$ = ","
btnDEL% = FindChildByTitle(FindChildByTitle(mofo%, "New Mail"), "Delete")
If AOLCountMail() = 0 Then MsgBox "You have no New Mail.", 12, "Dupe Killer": Exit Function
List% = FindChildByClass(FindChildByTitle(mofo%, "New Mail"), "_AOL_Tree")
For i = 0 To AOLCountMail() - 1
Ln = SendMessage(List%, LB_GETTEXTLEN, i, 0)
If Ln = -1 And i >= AOLCountMail() Then
    Exit For
ElseIf Ln = -1 And i <= AOLCountMail() Then
    MAILTXT$ = String$(60, 0)
Else
    MAILTXT$ = String$(Ln, 0)
End If
GTTXT = SendMessageByString(List%, LB_GETTEXT, i, MAILTXT$)
MAILTXT$ = Right$(MAILTXT$, Len(MAILTXT$) - InStr(InStr(MAILTXT$, Chr$(9)) + 1, MAILTXT$, Chr$(9)))
If InStr(LSTTXT$, "," & MAILTXT$ & ",") And InStr(DELTXT$, "," & MAILTXT$ & ",") = 0 Then
            x = SendMessage(List%, LB_SETCURSEL, i, 0)
            Call Click(btnDEL%)
            DELNUM% = DELNUM% + 1
            num = num - 1
            i = i - 1
            DELTXT$ = DELTXT$ + MAILTXT$ + ","
Else
LSTTXT$ = LSTTXT$ + MAILTXT$ + ","
End If
Next i
Loop Until Len(DELTXT$) = 1
MsgBox "There were " & DELNUM% & " duplicate mails deleted.", 12, "Dupe Count"
Mail_KillDupes = DELNUM%
End Function
Sub AOLMakeMeParent(Frm As Form)
'AOLMakeParent Me
'this makes the form an aol parent
aol% = FindChildByClass(FindWindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = SetParent(Frm.hwnd, aol%)
End Sub
Sub AOLMail(PERSON, SUBJECT, message)
Call RunMenuByString(AOLWindow(), "Compose Mail")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, SUBJECT)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub AOLLocateMember(name As String)
'locates, if possible, member "name"
AOLRunMenuByString ("Locate a Member Online")
Pause 0.3
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(mdi%, "Locate Member Online")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
closes = SendMessage(prof%, WM_CLOSE, 0, 0)
End Sub

Function WriteErrorNameToList(lst As ListBox)
'Not working yet 4/8/98
messa = Mail_GetErrorMessage
LC = GetLineCount(messa)

a = 1

start:
thetext = LineFromText(messa, a)
Stcount = Len(thetext)
SC = Stcount - 34
AC2 = Left$(thetext, SC)
lst.AddItem AC2
a = a + 1
GoTo start
last:
End Function

Function MessageFromIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = AOLGetText(IMtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(IMmessage, Len(IMmessage) - 1)
End Function

Sub SizeFormToWindow(Frm As Form, win%)
Dim wndRect As RECT, lRet As Long
lRet = GetWindowRect(win%, wndRect)
With Frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub
Sub StuffOff(Frm As Form, btn As Object)
With btn
   .FontItalic = False
   .FontStrikethru = False
   .FontUnderline = False
End With
End Sub
Function SNfromIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IM%)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
SNfromIM = naw$
End Function
Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

AOLGetText = TrimSpace$
End Function

Sub AOLIcon(icon%)
Click2% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click2% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLInstantMessage(PERSON, message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, PERSON)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(IM%, "_AOL_Icon")
For Sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next Sends
AOLIcon (imsend%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
End Sub
Sub AOLInstantMessage2(PERSON)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, ">Instant Message From: ")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, PERSON)
End Sub
Sub AOLInstantMessage3(PERSON, message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, PERSON)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(IM%, "_AOL_Icon")
For Sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next Sends
AOLIcon (imsend%)
Loop Until IM% = 0
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop

End Sub
Sub AOLInstantMessage4(PERSON, message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, PERSON)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(IM%, "_AOL_Icon")
For Sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next Sends
AOLIcon (imsend%)
End Sub

Function AOLChildIM()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
AOLChildIM = imsend%
End Function
Function AOLCreateIM(PERSON, message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, PERSON)
Call AOLSetText(aolrich%, message)
SendKeys "{TAB}"
End Function
Function AOLIsOnline()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Welcome% = FindChildByTitle(mdi%, "Welcome, ")
If Welcome% = 0 Then
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Public Sub AOLLoadAol()
On Error Resume Next
Dim x%
x% = Shell("C:\aol30\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x% = Shell("C:\aol30a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x% = Shell("C:\aol30b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x% = Shell("C:\aol25\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x% = Shell("C:\aol25a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x% = Shell("C:\aol25b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Function AOLIMScan()
aolcl% = FindWindow("#32770", "America Online")
If aolcl% > 0 Then
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs OFF and can't be punted."
End If
If aolcl% = 0 Then
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs ON and can be punted."
End If
End Function

Sub AOLKeyword(text)
Call RunMenuByString(AOLWindow(), "Keyword...")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(mdi%, "Keyword")
kedit% = FindChildByClass(keyw%, "_AOL_Edit")
If kedit% Then Exit Do
Loop

editsend% = SendMessageByString(kedit%, WM_SETTEXT, 0, text)
pausing = DoEvents()
Sending% = SendMessage(kedit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Function AOLLastChatLine()
getpar% = AOLFindRoom()
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

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
AOLLastChatLine = lastline
End Function

Sub Mail_SendNew(PERSON, SUBJECT, message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, SUBJECT)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub
Sub Mail_SendNew3(PERSON, SUBJECT, message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, SUBJECT)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)
HideWindow (MailWin%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub
Sub Mail_SendNew2(PERSON, SUBJECT, message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, SUBJECT)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
x = MsgBox("Please Attch File And Send", vbCritical, "BuM Auto Tagger l3y GenghisX")
last:
End Sub



Sub AOLResetNewUser(SN As String, tru_sn As String, pth As String)
'creates a new sn
'example : Call AOLResetNewUser("NewSN", "CurrentSN", "C:\aol30\Organize")
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already on new user!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The screen name has to be at least 7 characters long :)"): Exit Sub
tru_sn = tru_sn + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(16384, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(16384, " ")
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 16384
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next
End Sub

Sub File_BirthDay(File$)
If Not File_IfileExists(File$) Then Exit Sub
MsgBox FileDateTime(File$)
NoFreeze% = DoEvents()
End Sub
Sub File_Copy(File$, DestFile$)
If Not File_IfileExists(File$) Then Exit Sub
FileCopy File$, DestFile$
End Sub
Sub File_Delete(File$)
If Not File_IfileExists(File$) Then Exit Sub
Kill File$
NoFreeze% = DoEvents()
End Sub



Sub File_DeleteDir(DirName$)
If Not File_IfDirectoryExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub
Sub File_DeleteDirectory(DirName$)
If Not File_IfDirectoryExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub



Function File_IfDirectoryExists(TheDirectory)
Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(Dir$(TheDirectory))
If Err Or Check = 0 Then
    File_IfDirectoryExists = False
Else
    File_IfDirectoryExists = True
End If
End Function


Sub File_MakeDirectory(DirName$)
MkDir DirName$
End Sub
Function File_IfileExists(ByVal sFileName As String) As Integer
'Example: If Not File_ifileexists("win.com") then...
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        File_IfileExists = False
        Else
            File_IfileExists = True
    End If

End Function
Function File_LoadINI(look$, FileNamer$) As String
On Error GoTo Sla
Open FileNamer$ For Input As #1
Do While Not EOF(1)
    Input #1, CheckOut$
    If InStr(UCase$(CheckOut$), UCase$(look$)) Then
        Where = InStr(UCase$(CheckOut$), UCase$(look$))
        out$ = Mid$(CheckOut$, Where + Len(look$))
        File_LoadINI = out$
    End If
Loop
Sla:
Close #1
Resume nigger
nigger:
End Function
Sub File_OpenEXE(File$)
OpenEXE = Shell(File$, 1): NoFreeze% = DoEvents()
End Sub






Sub File_ReName(File$, NewName$)
Name File$ As NewName$
NoFreeze% = DoEvents()
End Sub
Sub FormFlash(Frm As Form)
Frm.Show
Frm.BackColor = &H0&
Pause (".1")
Frm.BackColor = &HFF&
Pause (".1")
Frm.BackColor = &HFF0000
Pause (".1")
Frm.BackColor = &HFF00&
Pause (".1")
Frm.BackColor = &H8080FF
Pause (".1")
Frm.BackColor = &HFFFF00
Pause (".1")
Frm.BackColor = &H80FF&
Pause (".1")
Frm.BackColor = &HC0C0C0
End Sub
Public Sub FortuneBot()
'steps...
'1. in Timer1 tye Call FortuneBot
'2. make 2 command buttons
'3. in command button 1 type-
'Timer1.enbled = True
'AOLChatSend "Type /fortune to get your fortune"
'4. in the command button 2 type-
'Timer1.enabled = false
'AOLChatSend "Fortune Bot Off!"
FreeProcess
Timer1.interval = 1
On Error Resume Next
Dim last As String
Dim name As String
Dim a As String
Dim n As Integer
Dim x As Integer
DoEvents
a = AOLLastChatLine
last = Len(a)
For x = 1 To last
name = Mid(a, x, 1)
final = final & name
If name = ":" Then Exit For
Next x
final = Left(final, Len(final) - 1)
If final = AOLGetUser Then
Exit Sub
Else
If InStr(a, "/fortune") Then
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call AOLChatSend("" & final & ", You will win the lottery and spend it all on BEER!")
If rand = 2 Then Call AOLChatSend("" & final & ", You will kill Steve Case and take over AoL!")
If rand = 3 Then Call AOLChatSend("" & final & ", You will marry Carmen Electra!")
If rand = 4 Then Call AOLChatSend("" & final & ", You will DL a PWS and get thousands of bucks charged on your account!")
If rand = 5 Then Call AOLChatSend("" & final & ", You will end up werking at McDonalds and die a lonely man")
If rand = 6 Then Call AOLChatSend("" & final & ", You will get a check for ONE MILLION $$ from me! Yeah right!")
If rand = 7 Then Call AOLChatSend("" & final & ", You will be OWNED by shlep")
If rand = 8 Then Call AOLChatSend("" & final & ", You will be OWNED by epa")
If rand = 9 Then Call AOLChatSend("" & final & ", You will get an OH and delete Steve Case's SN!")
If rand = 10 Then Call AOLChatSend("" & final & ", You will slip on a banana peel in Japan and land on some egg foo yung!")
Call Pause(0.6)
End If
End If
End Sub
Public Function GetListIndex(oListBox As ListBox, sText As String) As Integer

Dim iIndex As Integer

With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With

GetListIndex = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function
Sub RemoveItemFromListbox(lst As ListBox)
'this code works well in the double click part of your listbox
Jaguar% = lst.ListIndex
lst.RemoveItem (Jaguar%)
End Sub
Public Sub TransferListToTextBox(lst As ListBox, txt As TextBox)
'This moves the individual highlighted part of a
'listbox to a textbox
Ind = lst.ListIndex
daname$ = lst.List(Ind)
txt.text = ""
txt.text = daname$
End Sub
Function AOLUpChat()
Do
    x% = DoEvents()
aolmod = FindWindow("_AOL_Modal", 0&)
KillWin (aolmod)
Loop Until aolmod = 0
End Function
Sub Win95_StartButton()
wind% = FindWindow("Shell_TrayWnd", 0&)
btn% = FindChildByClass(wind%, "Button")
SendNow% = SendMessageByNum(btn%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(btn%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub AOLHidenMail(PERSON, SUBJECT, message)
Call RunMenuByString(AOLWindow(), "Compose Mail")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
HideWindow MailWin%
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, SUBJECT)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub


Sub AOLMainMenu()
Call RunMenu(2, 3)
End Sub

Function AOLRoomCount()
thechild% = AOLFindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")

getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOLRoomCount = getcount
End Function

Sub AOLSetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub AOLSignOff()
aol% = FindWindow("AOL Frame25", vbNullString)
If aol% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu(2, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(aol%, "Sign Off?")
If pfc% <> 0 Then
icon1% = FindChildByClass(pfc%, "_AOL_Icon")
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
clickicon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
clickicon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Function AOLVersion()
aol% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(aol%)

submenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(submenu%, subitem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 2.5
End If
End Function

Function AOLWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = aol%
End Function



Function GetCaption(hwnd)
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

Function GetWindowDir()
buffer$ = String$(255, 0)
x = GetWindowsDirectory(buffer$, 255)
If Right$(buffer$, 1) <> "\" Then buffer$ = buffer$ + "\"
GetWindowDir = buffer$
End Function

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Function


Sub SendCharNum(win, chars)
e = SendMessageByNum(win, WM_CHAR, chars, 0)

End Sub

Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function

Sub SetPreference()
Call RunMenuByString(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
AOLIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

AOLButton (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub


Public Sub AOLPuntCombo(PERSON$)
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
End Sub

Public Sub AOLPuntH1(PERSON$)
Call AOLInstantMessage(PERSON$, "<h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1>")
End Sub



Public Sub AOLPuntExtreme(PERSON$)
Call AOLInstantMessage(PERSON$, "<a hreh><a href></a>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(PERSON$, "</a>")
End Sub


Public Sub DisableCRTL_ALT_DEL()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub EnableCRTL_ALT_DEL()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
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


Sub UnHideWindow(hwnd)
un = ShowWindow(hwnd, SW_SHOW)
End Sub



Sub waitforok()
Do: DoEvents
aol% = FindWindow("#32770", "America Online")

If aol% Then
CloseAOL% = SendMessage(aol%, WM_CLOSE, 0, 0)
Exit Do
End If

aolw% = FindWindow("_AOL_Modal", vbNullString)

If aolw% Then
AOLButton (FindChildByTitle(aolw%, "OK"))
Exit Do
End If
Loop

End Sub

Sub WaitWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
topmdi% = GetWindow(mdi%, 5)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
topmdi2% = GetWindow(mdi%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop

End Sub


Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop

End Function


Function AOLKillChatRoom()
'This finds the chat room
Room = AOLFindRoom()
'This kills the chat room
Do
    KillWin (Room)
    Loop Until Room = 0
End Function
Sub KillWin(Windo)
x = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

