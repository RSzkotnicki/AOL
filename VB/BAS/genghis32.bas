Attribute VB_Name = "Genghis"
'|X-X| Genghis X
'|X-X| Made By Genghis HimSelf
'|X-X| Mail: UnitedNation@MailCity.Com
'|X-X|
'|X-X| Enjoy your good programming times.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
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
Declare Function mciSendString& Lib "Winmm" Alias "mciSendStringA" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen&, ByVal hCallBack&)
'sndPlaySound  flag values for uFlags parameter
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside
 '  this range will raise an error
Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
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

Global R&       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Function GetFromINI(AppName$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function



Sub ULAlign(FRM As Form)
    Dim x, Y                    ' New top, left for the form
    Y = 0
    x = 0
    FRM.Move x, Y             ' Change location of the form

End Sub

Sub PlayWav(File)
SoundName$ = File
SoundFlags& = &H20000 Or &H1
Snd& = sndPlaySound(SoundName$, SoundFlags&)
End Sub


Sub AOLChangeCaption(NewCaption As String)
Call AOLSetText(AOLWindow(), NewCaption)
End Sub
Sub AOLKHANCAPTION(NewCaption As String)
Call AOLSetText(AOLFindRoom(), NewCaption)
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

AppActivate "America  Online"
End Sub


Public Sub AddRoom(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

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
Listboxes.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Function ListToList(source, destination)
counts = SendMessage(source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(destination, LB_ADDSTRING, 0, Buffer$)
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

Function UntilWindowClass(parent, news$)
Do: DoEvents
e = FindChildByClass(parent, news$)
Loop Until e
UntilWindowClass = e
End Function


Function UntilWindowTitle(parent, news$)
Do: DoEvents
e = FindChildByTitle(parent, news$)
Loop Until e
UntilWindowTitle = e
End Function
Public Function AOLGetList(Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

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





Function AddListToString(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function
Sub Addit(Lst As ListBox)
AOL = FindWindow%("AOL Frame25", 0&)
Happy% = FindChildByTitle%(AOL, "Welcome,")
joy$ = GetWinText((Happy%))
If InStr(joy$, "!") = 0 Then Exit Sub
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
For Index% = 0 To 25
names$ = String$(256, " ")
Ret = AOLGetList(Index%, names$) & ErB$
If Len(Trim$(names$)) <= 1 Then Exit For
names$ = Left$(Trim$(names$), Len(Trim(names$)) - 1)
If names$ = yoursn Then GoTo boba2
Do
If Lst.List(a) = names$ Then GoTo boba2
a = 1 + a
Loop Until Lst.List(a) = ""
a = 0
Lst.AddItem (names$)
boba2:
a = 0
Next Index%
End Sub
Sub Addsta(Lst As ListBox)
AOL = FindWindow%("AOL Frame25", "America  Online")
Happy% = FindChildByTitle%(AOL, "Welcome,")
joy$ = GetWinText((Happy%))
If InStr(joy$, "!") = 0 Then Exit Sub
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
For Index% = 0 To 25
names$ = String$(256, " ")
Ret = AOLGetList(Index%, names$) & ErB$
If Len(Trim$(names$)) <= 1 Then Exit For
names$ = Left$(Trim$(names$), Len(Trim(names$)) - 1)
If names$ = yoursn Then GoTo bo2
If InStr(LCase$(names$), "tos") Then GoTo bo2
If InStr(LCase$(names$), "guide") Then GoTo bo2
If InStr(LCase$(names$), "host") Then GoTo bo2
Lst.AddItem (names$)
bo2:
Next Index%
End Sub
Function AF_Script(ByVal Strt As String, ByVal ReplaceMe As String, ByVal ReplaceWith As String) As String
start$ = Strt
Do While InStr(start$, ReplaceMe$) <> 0
    x% = DoEvents()
    pos% = InStr(start$, ReplaceMe$)
    start$ = Left(start$, pos% - 1) & ReplaceWith$ & Right(start$, Len(start$) - pos% - Len(ReplaceMe$) + 1)
    Loop
AF_Script$ = start$
End Function


Sub AddStringToList(theitems, TheList As ListBox)
If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
TheList.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub


Function AOLClickList(hWnd)
clicklist% = SendMessageByNum(hWnd, &H203, 0, 0&)
End Function


Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function AOLGetListString(parent, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

aolhandle = parent

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

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

Sub AOLHide()
a = ShowWindow(AOLWindow(), SW_HIDE)
End Sub



Sub ClickAOLMenu(Menu_String As String, Top_Position As String)
Dim Top_Position_Num As Integer
Dim Buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim AOL As Integer
Dim Menu_Handle As Integer


Top_Position_Num = -1
AOL% = FindWindow("AOL Frame25", 0&)
Menu_Handle = GetMenu(AOL%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    Buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, Buffer$, Len(Top_Position) + 1, &H400)
    Trim_Buffer = Trimnull(Buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    Buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, Buffer$, Len(Menu_String) + 1, &H400)
    Trim_Buffer = Trimnull(Buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
Loop
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = SendMessageByNum(AOL, &H111, Get_ID%, 0&)

End Sub
Function Trimnull (in$) As String
For x = 1 To Len(in$)
    If (Mid$(in$, x, 1) <> Chr$(0)) Then
        Total$ = Total$ + Mid$(in$, x, 1)
    Else
        GoTo NullDetect
    End If
Next
NullDetect:
Trimnull1 = Total$

End Function

Sub AOLOpenMail(which)
If which = 1 Then
Call AOLRunMenuByString("Read &New Mail")
End If

If which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Read")
End If

If Not which = 1 Or Not which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Sent")
End If

End Sub




Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub


Sub AOLUnHide()
a = ShowWindow(AOLWindow(), SW_SHOW)
End Sub

Sub AOLWaitMail()
mailwin% = GetTopWindow(AOLMDI())
aoltree% = FindChildByClass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
Pause (10)
secondcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub


Function EncryptType(TEXT, types)
'to encrypt, example:
'encrypted$ = EncryptType("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = EncryptType("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(TEXT)
If types = 0 Then
Current$ = Asc(Mid(TEXT, God, 1)) - 1
Else
Current$ = Asc(Mid(TEXT, God, 1)) + 1
End If
Process$ = Process$ & Chr(Current$)
Next God

EncryptType = Process$
End Function

Function FindChildByTitle(parent, child As String) As Integer
childfocus% = GetWindow(parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
Buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, Buffer$, (hwndLength% + 1))

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function FindChildByClass(parent, child As String) As Integer
childfocus% = GetWindow(parent, 5)

While childfocus%
Buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, Buffer$, 250)

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByClass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function



Sub ClickKW25()
AOL% = FindWindow("AOL Frame25", 0&)
tool% = FindChildByClass(AOL%, "Aol Toolbar")
Icon1% = FindChildByClass(tool%, "_Aol_Icon")
Icon2% = GetWindow(Icon1%, 2)
Icon3% = GetWindow(Icon2%, 2)
Icon4% = GetWindow(Icon3%, 2)
Icon5% = GetWindow(Icon4%, 2)
Icon6% = GetWindow(Icon5%, 2)
Icon7% = GetWindow(Icon6%, 2)
Icon8% = GetWindow(Icon7%, 2)
Icon9% = GetWindow(Icon8%, 2)
Icon10% = GetWindow(Icon9%, 2)
Icon11% = GetWindow(Icon10%, 2)
Icon12% = GetWindow(Icon11%, 2)
Icon13% = GetWindow(Icon12%, 2)
Z% = SendMessage(Icon13%, WM_LBUTTONDOWN, 0&, 0&)
q% = SendMessage(Icon13%, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub ClickKW30()
AOL% = FindWindow("AOL Frame25", 0&)
tool% = FindChildByClass(AOL%, "Aol Toolbar")
Icon1% = FindChildByClass(tool%, "_Aol_Icon")
Icon2% = GetWindow(Icon1%, 2)
Icon3% = GetWindow(Icon2%, 2)
Icon4% = GetWindow(Icon3%, 2)
Icon5% = GetWindow(Icon4%, 2)
Icon6% = GetWindow(Icon5%, 2)
Icon7% = GetWindow(Icon6%, 2)
Icon8% = GetWindow(Icon7%, 2)
Icon9% = GetWindow(Icon8%, 2)
Icon10% = GetWindow(Icon9%, 2)
Icon11% = GetWindow(Icon10%, 2)
Icon12% = GetWindow(Icon11%, 2)
Icon13% = GetWindow(Icon12%, 2)
Icon14% = GetWindow(Icon13%, 2)
Icon15% = GetWindow(Icon14%, 2)
Icon16% = GetWindow(Icon15%, 2)
Icon17% = GetWindow(Icon16%, 2)
Icon18% = GetWindow(Icon17%, 2)
Icon19% = GetWindow(Icon18%, 2)
Z% = SendMessage(Icon18%, WM_LBUTTONDOWN, 0&, 0&)
q% = SendMessage(Icon18%, WM_LBUTTONUP, 0&, 0&)
End Sub




Function GetLineCount(TEXT)

theview$ = TEXT


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(TEXT, Len(TEXT), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function
Sub delitem(Lst As ListBox, item$)
Do
If LCase$(Lst.List(a)) = LCase$(item$) Then Lst.RemoveItem (a)
a = 1 + a
Loop Until a >= Lst.ListCount

End Sub
Function Disc(ByVal Marbro As String) As String
On Error Resume Next
Disc = Left$(Marbro$, InStr(Marbro$, Chr$(0)) - 1)
End Function
Sub emailbomb(numofbomb As TextBox, whotobom, subj$, mesg$)
Do
Call mail((whotobom), (subj$), (mesg$))
numofbomb = Str(Val(numofbomb - 1))
Loop Until numofbomb = 0
AOL% = FindWindow("aol Frame25", 0&)
compoze% = FindChildByTitle(AOL%, "Compose Mail")
If compoze% <> 0 Then
Do
AOL% = FindWindow("aol Frame25", 0&)
compoze% = FindChildByTitle(AOL%, "Compose Mail")
x = SendMessageByNum(compoze%, WM_CLOSE, 0, 0)
Loop Until compoze% = 0
End If

End Sub
Sub fastestfakeoh(room1$)
' change Command3d1 to whatever the start button is named
'Command3d1.Enabled = False

word:
Call keywor("aol://2719:2-2-" & (room1$))
AOL% = FindWindow("aol Frame25", 0&)
mdi% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
rooma% = FindChildByTitle(mdi%, (room1$))
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
Timeout 0.2
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
Timeout 0.2
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
Timeout 0.2
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
Timeout 0.2
'If Command3d1.Enabled = True Then
'x = waitforwindow((room1$), "AOL Child")
'killwin rooma%
'Exit Sub
'End If
GoTo word

End Sub
Function FileErrors(errVal As Integer) As Integer
' Return Value  Meaning     Return Value    Meaning
' 0             Resume      2               Unrecoverable error
' 1             Resume Next 3               Unrecognized error
Dim MsgType As Integer
Dim Response As Integer
Dim Action As Integer
Dim Msg As String
MsgType = MB_EXCLAIM
Select Case errVal
    Case Err_DeviceUnavailable  ' Error #68
        Msg = "That device appears to be unavailable."
        MsgType = MB_EXCLAIM + 5
    Case Err_DiskNotReady       ' Error #71
        Msg = "The disk is not ready."
    Case Err_DeviceIO
        Msg = "The disk is full."
    Case Err_BadFileName, Err_BadFileNameOrNumber   ' Errors #64 & 52
        Msg = "That file name is illegal."
    Case Err_PathDoesNotExist                        ' Error #76
        Msg = "That path doesn't exist."
    Case Err_BadFileMode                            ' Error #54
        Msg = "Can't open your file for that type of access."
    Case Err_FileAlreadyOpen                        ' Error #55
        Msg = "That file is already open."
    Case Err_InputPastEndOfFile                     ' Error #62
    Msg = "This file has a nonstandard end-of-file marker,"
    Msg = Msg + "or an attempt was made to read beyond "
    Msg = Msg + "the end-of-file marker."
    Case Else
        FileErrors = 3
        Exit Function
    End Select
    Response = MsgBox(Msg, MsgType, "File Error")
    Select Case Response
        Case 4          ' Retry button.
            FileErrors = 0
        Case 5          ' Ignore button.
            FileErrors = 1
        Case 1, 2, 3    ' Ok and Cancel buttons.
            FileErrors = 2
        Case Else
            FileErrors = 3
    End Select
End Function
Function findaol()
Dim AOL
AOL = FindWindow("AOL Frame25", 0&)
findaol = AOL
End Function
Function FindAOLChildByTitle(TitleText As String) As Integer
Dim x%
Dim ChildWnd As Integer
Dim MDIhWnd%
Dim AOLChildhWnd%
Dim RetClsName As String * 255
  
MDIhWnd% = GetWindow(FindWindow("AOL Frame25", 0&), GW_CHILD)
Do
  x% = GetClassName(MDIhWnd%, RetClsName$, 254)
  If InStr(RetClsName$, "MDIClient") Then AOLChildhWnd% = MDIhWnd%
  MDIhWnd% = GetWindow(MDIhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While MDIhWnd% <> 0
If TitleText = "MDIClient" Then FindAOLChildByTitle = AOLChildhWnd%
ChildWnd = GetWindow(AOLChildhWnd%, GW_CHILD)
Do
  If InStr(WindowCaption(ChildWnd), TitleText) <> 0 Then
      FindAOLChildByTitle = ChildWnd
      Exit Do
  End If
  ChildWnd = GetWindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
End Function

Function FindChatWnd() As Integer
  Dim MDIhWnd%
  Dim AOLChildhWnd%
  Dim ChildWnd As Integer
  Dim ControlWnd As Integer
  Dim ChatWnd As Integer
  Dim TargetsFound As Integer
  Dim RetClsName As String * 255
  Dim x%
MDIhWnd% = GetWindow(FindWindow("AOL Frame25", 0&), GW_CHILD)
Do
  x% = GetClassName(MDIhWnd%, RetClsName$, 254)
    If InStr(RetClsName$, "MDIClient") Then
      AOLChildhWnd% = MDIhWnd% 'Child window found!
    End If
  MDIhWnd% = GetWindow(MDIhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While MDIhWnd% <> 0
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
FindChatWnd = ChatWnd

End Function

Function findcomposemail()
Dim bb As Integer
Dim dis_win As Integer

dis_win = FindChildByClass(aolhwnd(), "AOL Child")

begin_find_composemail:

bb = FindChildByTitle(dis_win, "Send")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "To:")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Subject:")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Send" & Chr(13) & "Later")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Attach")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Address" & Chr(13) & "Book")
    If bb <> 0 Then Let countt = countt + 1

If countt = 6 Then
  findcomposemail = dis_win
  Exit Function
End If
Let countt = 0
dis_win = GetNextWindow(dis_win, 2)
If dis_win = GetWindow(dis_win, GW_HWNDLAST) Then
   findtocomposemail = 0
   Exit Function
End If
GoTo begin_find_composemail
End Function
Function FindReadWnd() As Integer
Dim HWNDS() As Integer
Dim R, x%, i, T, g%
x = 0
DoEvents
g% = GetWindow(AOhWnd(), 5)
DoEvents
Do
If FindChildByTitle(g%, "Reply") <> 0 Then x = x + 1
DoEvents
If FindChildByTitle(g%, "Forward") <> 0 Then x = x + 1
DoEvents
If FindChildByTitle(g%, "Reply to All") <> 0 Then x = x + 1
DoEvents
If x = 3 Then GoTo Founddd
DoEvents
g% = GetWindow(g%, 2)
DoEvents
x = 0
Loop While g% <> 0


Exit Function
Founddd:
DoEvents
FindReadWnd = g%
Exit Function


End Function
Function FindSN()
Dim dis_win As Integer

dis_win = FindChildByClass(aolhwnd(), "AOL Child")

begin_find_SN:

bb$ = WindowCaption(dis_win)
    If Left(bb$, 9) = "Welcome, " Then Let countt = countt + 1
If countt = 1 Then
  val1 = InStr(bb$, " ")
  val2 = InStr(bb$, "!")
  Let sn$ = Mid$(bb$, val1 + 1, val2 - val1 - 1)
  FindSN = Trim(sn$)
  Exit Function
End If
Let countt = 0
dis_win = GetNextWindow(dis_win, 2)
If dis_win = GetWindow(dis_win, GW_HWNDLAST) Then
   FindSN = 0
   Exit Function
End If

GoTo begin_find_SN

End Function
Function findtoim()
Dim bb As Integer
Dim dis_win As Integer
AOL% = FindWindow("AOL Frame25", 0&)

dis_win = FindChildByClass(AOL%, "AOL Child")

begin_find_To_im:

bb = FindChildByTitle(dis_win, "Send")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "To:")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Available?")
    If bb <> 0 Then Let countt = countt + 1

If countt = 3 Then
  findtoim = dis_win
  Exit Function
End If
Let countt = 0
dis_win = GetNextWindow(dis_win, 2)
If dis_win = GetWindow(dis_win, GW_HWNDLAST) Then
   findtoim = 0
   Exit Function
End If

GoTo begin_find_To_im
End Function
Function getroomname() As String
On Error Resume Next
Chat% = findchatroom1()
x = GetWindowTextLength(Chat%)
Title$ = Space(x + 1)
x = GetWindowText(Chat%, Title$, x + 1)
Title$ = FixAPIString(Title$)
getroomname = Title$
End Function

Function frac(whatever)
bah = Int(whatever)
valfrac = (Val(whatever) - Val(bah))
frac = valfrac
End Function
Function gencc(prefix)
loogie:
a = 0
Randomize Timer
heh = 99 * Rnd
Do
If Val(heh) = 10 Then heh = heh - (Int(10 * Rnd))
If Val(heh) < 10 Then GoTo hhhf
If heh > 10 Then
heh = heh - Int(10 * Rnd)
End If
Loop Until heh < 10
hhhf:

g1 = 4
g2 = Int(Val(heh) * Rnd)
g3 = Int(Val(heh) * Rnd)
g4 = Int(Val(heh) * Rnd)
g5 = Int(Val(heh) * Rnd)
g6 = Int(Val(heh) * Rnd)
g7 = Int(Val(heh) * Rnd)
g8 = Int(Val(heh) * Rnd)
g9 = Int(Val(heh) * Rnd)
g10 = Int(Val(heh) * Rnd)
g11 = Int(Val(heh) * Rnd)
g12 = Int(Val(heh) * Rnd)
g13 = Int(Val(heh) * Rnd)
g14 = Int(Val(heh) * Rnd)
g15 = Int(Val(heh) * Rnd)
g16 = Int(Val(heh) * Rnd)

g13 = Int(Val(heh) * Rnd)
g14 = Int(Val(heh) * Rnd)
g15 = Int(Val(heh) * Rnd)
g16 = Int(Val(heh) * Rnd)

gee = 0
hehd:
Do
T1$ = g1
T2$ = g2
T3$ = g3
T4$ = g4
T5$ = g5
T6$ = g6
T7$ = g7
T8$ = g8
T9$ = g9
T10$ = g10
T11$ = g11
T12$ = g12
T13$ = g13
T14$ = g14
T15$ = g15
T16$ = g16
evens = Val(T2$) + Val(T4$) + Val(T6$) + Val(T8$) + Val(T10$) + Val(T12$) + Val(T14$) + Val(T16$)
C1 = T1$
C3 = T3$
C5 = T5$
C7 = T7$
C9 = T9$
C11 = T11$
C13 = T13$
C15 = T15$
C1 = Val(C1) + Val(C1)
C3 = Val(C3) + Val(C3)
C5 = Val(C5) + Val(C5)
C7 = Val(C7) + Val(C7)
C9 = Val(C9) + Val(C9)
C11 = Val(C11) + Val(C11)
C13 = Val(C13) + Val(C13)
C15 = Val(C15) + Val(C15)
If C1 > 9 Then C1 = C1 - 9
If C3 > 9 Then C3 = C3 - 9
If C5 > 9 Then C5 = C5 - 9
If C7 > 9 Then C7 = C7 - 9
If C9 > 9 Then C9 = C9 - 9
If C11 > 9 Then C11 = C11 - 9
If C13 > 9 Then C13 = C13 - 9
If C15 > 9 Then C15 = C15 - 9
a = 1 + a
If a = 10 Then GoTo loogie
odds = Val(C1) + Val(C3) + Val(C5) + Val(C7) + Val(C9) + Val(C11) + Val(C13) + Val(C15)
If Int((odds + evens) / 60) = (odds + evens) / 60 Then GoTo bob

hah = 60
If gee = 0 Then g16 = -1
g16 = g16 + 1
If g16 > 10 Then
g13 = Int(10 * Rnd)
g14 = Int(Val(heh) * Rnd)
g15 = Int(Val(heh) * Rnd)
g16 = Int(Val(heh) * Rnd)
gee = 0
GoTo hehd
End If
gee = 1
Loop
bob:
bahq$ = g1 & g2 & g3 & g4 & "-" & g5 & g6 & g7 & g8 & "-" & g9 & g10 & g11 & g12 & "-" & g13 & g14 & g15 & g16
gencc = bahq$
End Function
Function getaolwinb(parent As Integer, ClassToFind As String, Num As Integer)
''——-=-By:GenOziDe-=-——''
Dim DooD As Integer
DooD = 0
Count = (Num) 'Count Out The Number Of Windows Over You Want To Look For
'If num = 0 Then MsgBox "Are You An Asshole": Exit Function'Won't Let You Look For Nothing
DoEvents
a% = FindChildByClass(parent%, ClassToFind$) 'Begin Your Search For Da Window
DooD = DooD + 1 'If You Find One Add 1 To Your Counter
Do
DoEvents
If DooD = Num Then 'If Your Counter = Your Number Then Exit Function
getaolwinb = a% 'Declare The Function
Exit Function
End If
DoEvents
Do         'Begin a Do...Loop to look For The Class Name
DoEvents
a% = GetWindow(a%, GW_HWNDNEXT)
bb$ = String(255, 0)
CC% = GetClassName(a%, bb$, 254) 'Use This To Get The Class Name Of The Window You Found
bb$ = Disc(bb$)
If bb$ = ClassToFind$ Then DooD = DooD + 1 'Then Compare
If DooD = Num Then Exit Do
Loop Until bb$ = ClassToFind$ 'Loop Until You Find The Window With The Class Name You're Looking For
Loop Until DooD = Num 'Loop Until Your Counter Is = To Your Number
getaolwinb = a% 'Declare The Function

End Function
Function GetCPUType(x)
'Example: text9.text = "Your system's CPU type is: " & sGetCPUType
Dim lWinFlags As Long

    lWinFlags = GetWinFlags()

    If lWinFlags And WF_CPU486 Then

        x = "486"
        ElseIf lWinFlags And WF_CPU386 Then
            x = "386"
            ElseIf lWinFlags And WF_CPU286 Then
                x = "286"
                Else
                    x = "Other"
    End If

End Function
Function GetFileName(Prompt As String) As String
    GetFileName = LTrim$(RTrim$(UCase$(InputBox$(Prompt, "Enter File Name"))))
End Function
Function GetFreeGDI(x)
'Example: text5.text = "Free GDI Resources: " & sGetFreeGDI
    x = Format$(GetFreeSystemResources(GFSR_GDIRESOURCES)) + "%"

End Function
Function GetFreeSys(x)
'Example: text3.text = "Free System Resources: " & sGetFreeSys
    x = Format$(GetFreeSystemResources(GFSR_SYSTEMRESOURCES)) + "%"

End Function
Function GetFreeUser(x)
'Example: text4.text = "Free User Resources: " & sGetFreeUser
    x = Format$(GetFreeSystemResources(GFSR_USERRESOURCES)) + "%"

End Function

Sub sendtext(hWnd As Integer, What As String)
Dim R
R = SendMessageByString(hWnd, &HC, 0, What)

End Sub

Function FileOpener(NewFileName As String, Mode As Integer, RecordLen As Integer, Confirm As Integer) As Integer
     Dim NewFileNum As Integer
     Dim Action As Integer
     Dim FileExists As Integer
     Dim Msg As String
     On Error GoTo OpenerError
     If NewFileName Like "*[;-?[* ]*" Or NewFileName Like "*]*" Then Error Err_BadFileName
     If Confirm Then
        If Dir(NewFileName) = "" Then
            FileExists = False
        Else
            FileExists = True
        End If
        If Mode = REPLACEFILE And FileExists Then
            Msg = "Replace contents of " + NewFileName + "?"
            If MsgBox(Msg, 49, "Replace File?") = 2 Then
                FileOpener = 0
                Exit Function
            End If
        End If
        If Not FileExists Then
            Msg = "The file " + NewFileName + " does not exist. "
            Msg = Msg + "Do you want to create it?"
            If MsgBox(Msg, 1, "Create File?") = 2 Then
                FileOpener = 0
                Exit Function
            End If
        End If
     End If
     NewFileNum = FreeFile
     Select Case Mode
          Case REPLACEFILE
            Open NewFileName For Output As NewFileNum
          Case READFILE
            Open NewFileName For Input As NewFileNum
          Case ADDTOFILE
            Open NewFileName For Append As NewFileNum
          Case RANDOMFILE
            Open NewFileName For Random As NewFileNum Len = RecordLen
          Case BINARYFILE
            Open NewFileName For Binary As NewFileNum
          Case Else
            Exit Function
     End Select
     FileOpener = NewFileNum
Exit Function
OpenerError:
     Action = FileErrors(Err)
     Select Case Action
        Case 0
            Resume
        Case Else
            FileOpener = 0
            Exit Function
     End Select
End Function

Sub eightballgen(tixt As TextBox)
Randomize
tixt = Int((Val("11") * Rnd) + 1)
If tixt = "1" Then
tixt = "Looks doubtful."
ElseIf tixt = "2" Then tixt = "Definately YES!"
ElseIf tixt = "3" Then tixt = "Definately No!"
ElseIf tixt = "4" Then tixt = "Not a FuKin chance"
ElseIf tixt = "5" Then tixt = "HEEELLLLLLLLLLLLLLL nO"
ElseIf tixt = "6" Then tixt = "HeLL yeA!"
ElseIf tixt = "7" Then tixt = "Response HaZey try again."
ElseIf tixt = "8" Then tixt = "ProbabLee"
ElseIf tixt = "9" Then tixt = "yep yep"
ElseIf tixt = "10" Then tixt = "I'm not suRe"
ElseIf tixt = "11" Then tixt = "AbsolootLee yeZ"
End If
End Sub

Sub deletmenu()
    AOL% = FindWindow("AOL Frame25", 0&)
    aolmenu% = GetMenu(AOL%)
    a = 0
    Do
        x% = DoEvents()
        check% = GetSubMenu(aolmenu%, a)
        If check% <> 0 Then a = a + 1 Else Exit Do
    Loop
    For b = 0 To a - 1
        Buff$ = String(255, 0)
        x% = GetMenuString(aolmenu%, b, Buff$, 255, MF_BYPOSITION)
        Buff$ = Left(Buff$, x%)
        If Buff$ = "&FATE Infinity" Then
            x% = DeleteMenu(aolmenu%, b, MF_BYPOSITION)
            b = b - 1
            a = a - 1
        End If
    Next b

    DrawMenuBar (AOL%)
AO% = FindWindow("aol Frame25", 0&)
Y% = SendMessageByString(AO%, WM_SETTEXT, 0, "America  Online")
End
End Sub
Sub eMail(Lst As ListBox, subj$, lst2 As ListBox)

Dim R, x%, C1%, C2%, C3%, C4%, C5%, C6%, C7%, C8%, C9%, c10%, C11%, C12%, C13%, C14%, C15%, c16%, C17%
Call runmenua(3, 1)
x% = waitforwindow("Compose Mail", "AOL Child")
C1% = GetWindow(x%, 5)  '|Send|
C2% = GetWindow(C1%, 2) '<Send>
C3% = GetWindow(C2%, 2) '|Send Later|
C4% = GetWindow(C3%, 2) '<Send Later>
C5% = GetWindow(C4%, 2) '|Attach|
C6% = GetWindow(C5%, 2) '<Attach>
C7% = GetWindow(C6%, 2) '|Address Book|
C8% = GetWindow(C7%, 2) '<Address Book>
C9% = GetWindow(C8%, 2) '|To:|
c10% = GetWindow(C9%, 2) '<To:>
C11% = GetWindow(c10%, 2) '|CC:|
C12% = GetWindow(C11%, 2) '<CC:>
C13% = GetWindow(C12%, 2) '|Subject:|
C14% = GetWindow(C13%, 2) '<Subject:>
C15% = FindChildByClass(x%, "#32770")
c16% = FindChildByClass(gen%, "Edit")
C17% = FindChildByClass(x%, "RICHCNTL") '<Message>
Do
Loop Until c10% <> 0
who$ = ""
For i = 0 To Lst.ListCount - 1
who$ = who$ + Lst.List(i) + ", "
Y = SendMessageByString(c10%, WM_SETTEXT, 0, who$)
Next i
Timeout 1
textset C14%, CStr(subj$)
whoa$ = ""
For i = 0 To lst2.ListCount - 1
whoa$ = whoa$ + lst2.List(i) & Chr(13)
txt$ = "IXI—ÅñImÕ§iT¥—IXI"
Y = SendMessageByString(Dad%, WM_SETTEXT, 0, whoa$ & Chr(13) & Chr(13) & Chr(13) & txt$)
Z = SendMessageByString(rich, WM_SETTEXT, 0, whoa$ & Chr(13) & Chr(13) & Chr(13) & txt$)
Next i
textset C17%, CStr(whoa$)
Timeout 0.2

    Click (C2%)


End Sub

Sub HideWindow(hWnd)
hi = ShowWindow(hWnd, SW_HIDE)
End Sub

Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Function LineFromText(TEXT, theline)
theview$ = TEXT


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
Sub ListCheck(itm As String, Lst As ListBox)

If itm = Sccc Then Exit Sub
If Lst.ListCount = 0 Then Lst.AddItem itm: Exit Sub

Do Until xx = (Lst.ListCount)
Let diss_itm$ = Lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub
Sub mail(sill$, Subject$, TEXT$)
Do
AOL% = FindWindow("AOL Frame25", 0&)
heh = FindChildByTitle(AOL%, "Compose Mail")
SEND heh, "Compose Mail"
DoEvents
Loop Until heh = 0
AOL% = FindWindow("AOL Frame25", 0&)
run "Compose"
Do
x = DoEvents()
chatlist% = FindChildByTitle(AOL%, "Compose Mail")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
Loop Until chatedit% <> 0

chatwin% = GetParent(chatlist%)
button% = FindChildByClass(chatlist%, "_AOL_Icon")
ChatViw% = FindChildByClass(chatlist%, "_AOL_View")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, sill$)
blah% = GetWindow(chatedit%, GW_HWNDNEXT)
good% = GetWindow(blah%, GW_HWNDNEXT)
bad% = GetWindow(good%, GW_HWNDNEXT)
Sad% = GetWindow(bad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Sad%, WM_SETTEXT, 0, Subject$)
Nad% = GetWindow(Sad%, GW_HWNDNEXT)
Wad% = GetWindow(Nad%, GW_HWNDNEXT)
Dad% = GetWindow(Wad%, GW_HWNDNEXT)
fad% = GetWindow(Dad%, GW_HWNDNEXT)
Qad% = GetWindow(fad%, GW_HWNDNEXT)
Ead% = GetWindow(Qad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Dad%, WM_SETTEXT, 0, TEXT$ & " ")
rich = FindChildByClass(chatlist%, "RICHCNTL")
sndtext% = SendMessageByString(rich, WM_SETTEXT, 0, TEXT$ & " ")

Timeout (0.5)
SendNow% = SendMessageByNum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(button%, WM_LBUTTONUP, &HD, 0)
Timeout (0.1)
Do
AOL% = FindWindow("AOL Frame25", 0&)
heh = FindChildByTitle(AOL%, "Compose Mail")
SEND heh, "Compose Mail"
DoEvents
Loop Until heh = 0

End Sub

Sub LogOff_Quick()
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Sub
x = SetFocusAPI(AOL%)
x = SendMessageByNum(AOL%, &H11, 0, &H0)
Do
x = DoEvents()
Beav% = FindWindow("_AOL_Modal", "America Online")
Loop Until Beav% <> 0
oK% = FindChildByTitle%(Beav%, "&Yes")
Goob = SendMessageByNum(oK%, &H201, 0, 0&)
Goob = SendMessageByNum(oK%, &H202, 0, 0&)
End Sub
Sub maniptext(who$, wut$)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (who$) & ":" & Chr(9) & (wut$))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
End Sub
Sub massinstantmessage(Lst As ListBox, TEXT$)
Do
For i = 0 To Lst.ListCount - 1
If m001C% = 1 Then Exit Sub
who$ = Lst.List(0)
Lst.ListIndex = 0
Next i
Okw = FindWindow("#32770", "America Online")
okb = FindChildByTitle(Okw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
run "Send an Instant Message"
Do
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(AOL, "Send Instant Message")
txt% = FindChildByClass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
txt% = FindChildByClass(bah, "_AOL_Edit")
Do
rich% = FindChildByClass(bah, "RICHCNTL")
bahqw% = FindChildByTitle(bah, "Send")
DoEvents
Timeout (0.001)
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then
SEND txt%, who$
SEND rich%, ((TEXT$) & Chr(13) & "      <//=-——• FATE Infinity •——-=\\>" & Chr(13) & "        ƒ——by GenOziDe——ƒ")
Timeout (0.001)
getnum rich%, 1
Click rich%
Else
SEND txt%, person$
getnum txt%, 1
SEND txt%, tt$
getnum txt%, 1
Click txt%
End If
Timeout (0.001)
x = SendMessageByNum(bah, WM_CLOSE, 0, 0)
a = Lst.List(0)
Call delitem(Lst, (a))
Loop Until Lst.ListCount = 0
If Lst.ListCount = 0 Then
Exit Sub
End If

End Sub
Function MenuTrim1(Ands) As String
Dim x, a, b
     For x = 1 To Len(Ands)
          If Mid$(Ands, x, 1) = "&" Then
               a = Mid$(Ands, 1, x - 1)
               b = Mid$(Ands, x + 1, Len(Ands))
               Ands = a & b
          End If
          
          If InStr(Ands, "Ctrl") <> 0 Then
               Ands = Mid$(Ands, 1, InStr(Ands, "Ctrl") - 1)
          End If
          
          If InStr(Ands, " Shift") Then
               Ands = Mid$(Ands, 1, InStr(Ands, " Shift") - 1)
          End If
          
          If InStr(Ands, Chr$(9)) <> 0 Then
               Ands = Mid$(Ands, 1, InStr(Ands, Chr$(9)) - 1)
          End If
          
     Next
     
     MenuTrim = Ands


End Function

Sub ohscroll()
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "FATE Infinity by GenOziDe" & Chr(4) & "~~>GenOziDe "

End Sub


Sub macrokill()
Dim x
Randomize   ' Seed random number generator.
    x = Int(3 * Rnd + 1)    ' Generate first die value.
    If x = 1 Then   ' Generate second die value.
    
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")
Call Timeout(4)
send1 ("§——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§ §————————————‹^›‹{ `;´ }›‹^›——")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›———————————— ———————‹^›‹{ `;´ }›‹^›———————§")
send1 ("§————————————‹^›‹{ `;´ }›‹^›—— ———————‹^›‹{ `;´ }›‹^›———————§ §——‹^›‹{ `;´ }›‹^›————————————")
send1 ("———————‹^›‹{ `;´ }›‹^›———————§ §—————FucK uR MaCRo!—————— ——————•§•FATE Infinity•§•———————§")

ElseIf x = 2 Then
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
Call Timeout(0.000000001)
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe"
Call Timeout(0.000000001)
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "FATE Infinity by GenOziDe" & Chr(4) & "~~>GenOziDe "

ElseIf x = 3 Then
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("•FATE Infinity Ðiçe Hë££•")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("•FATE Infinity Ðiçe Hë££•")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("•FATE Infinity Ðiçe Hë££•")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call send1("//Roll -dice999 -sides999")
Call Timeout(4)
Call send1("•FATE Infinity Ðiçe Hë££•")
End If
End Sub



Sub MaxWindow(hWnd)
ma = ShowWindow(hWnd, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hWnd)
mi = ShowWindow(hWnd, SW_MINIMIZE)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
'turns the "number" so vb recognizes it for
'addition, subtraction, ect.

End Function

Sub ParentChange(parent%, location%)
doparent% = SetParent(parent%, location%)
End Sub


Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
Sub resetname(oldsn, newsn, pathh)
Static moocow As String * 10000
Dim Trident As Long
Dim fish As Long
Dim Tribal As Integer
Dim werd As Integer
Dim qwerty As Variant
Dim meee As Integer
On Error GoTo err0r
tru_sn = newsn + String$(Len(oldsn) - Len(newsn), " ")
Let paath$ = (pathh & "\idb\main.idx")
Open paath$ For Binary As #1 'Len = 50000
Trident& = 1
fish& = LOF(1)
While Trident& < fish&
moocow = String$(40000, Chr$(0))
Get #1, Trident&, moocow
While InStr(UCase$(moocow), UCase$(oldsn)) <> 0
Mid$(moocow, InStr(UCase$(moocow), UCase$(oldsn))) = tru_sn
Wend
    
Put #1, Trident&, moocow
Trident& = Trident& + 40000
Wend

Seek #1, Len(oldsn)
Trident& = Len(oldsn)
While Trident& < fish&
moocow = String$(255, Chr$(0))
Get #1, Trident&, moocow
While InStr(UCase$(moocow), UCase$(oldsn)) <> 0
Mid$(moocow, InStr(UCase$(moocow), UCase$(oldsn))) = tru_sn
Wend
Put #1, Trident&, moocow
Trident& = Trident& + 11900000
Wend
Close #1
Screen.MousePointer = 0
err0r:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub

Function r_backwards(strin As TextBox)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
r_backwards = newsent$

End Function
Sub Stafftos(who$, phrase$)
run "Keyword"
AOL = FindWindow("AOL Frame25", 0&)
Do
keyword = FindChildByTitle(AOL, "Keyword")
Timeout (0.001)
editbox = FindChildByClass(keyword, "_AOL_Edit")
Loop Until editbox <> 0
editbox = FindChildByClass(keyword, "_AOL_Edit")
SEND editbox, "staffpager"
GO% = FindChildByClass(keyword, "_AOL_Icon")
Click GO%
Do
AOL = FindWindow("AOL Frame25", 0&)
mow = FindChildByTitle(AOL, "I NEED HELP")
ikqn% = FindChildByClass(mow, "_AOL_Icon")
DoEvents
PassWord = FindChildByTitle(AOL, "Report Password Solicitations")
Click ikqn%
Timeout (0.001)
Loop Until PassWord <> 0

Do
PassWord = FindChildByTitle(AOL, "Report Password Solicitations")
Timeout (0.001)
Loop Until PassWord <> 0
whototosbox% = FindChildByClass(PassWord, "_AOL_Edit")
getnum whototosbox%, 7
SEND whototosbox%, who$
getnum whototosbox%, 3
SEND whototosbox%, phrase$
Sendbutton% = FindChildByTitle(PassWord, "Send")
Click Sendbutton%
WaitForOk
x = SendMessageByNum(guidepager, WM_CLOSE, 0, 0)
DoEvents
x = SendMessageByNum(mow, WM_CLOSE, 0, 0)

End Sub
Sub Timeout(duration)
starttime = Timer

x = DoEvents()

End Sub
Sub upchat()
AOM = FindWindow("_AOL_MODAL", 0&)
x = ShowWindow(AOM, SW_MINIMIZE)
x = ShowWindow(AOM, SW_HIDE)
End Sub

Function AOLNEWMAIL()
Call AOLRunMenuByString("Read &New Mail")
End Function
Sub AOLOpenChat()
If AOLFindRoom() Then Exit Sub
AOLKeyword ("pc")
Do: DoEvents
Loop Until AOLFindRoom()

End Sub
Sub AOLRespondIM(message)
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Z
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Z
Exit Sub
Z:
e = FindChildByClass(im%, "RICHCNTL")

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
End Sub
Sub addroomwithoutme(Lst As ListBox)
AOL% = FindWindow("AOL Frame25", 0&)
Debug.Print AOL%
RoomList% = FindChildByTitle(AOL%, "List Rooms")
Debug.Print RoomList%
Chatroom% = GetParent(RoomList%)
Debug.Print Chatroom%
List% = FindChildByClass(Chatroom%, "_AOL_Listbox")
If List% = 0 Then Exit Sub
If List% = 0 Then GoTo 19
thatlb = SendMessage(List%, LB_GETCOUNT, 0&, 0&)
For RoomNames = 0 To thatlb - 1
Buffer$ = String$(64, 0)
BuddyName% = AOLGetList%(RoomNames, Buffer$)
FinalBuddyname$ = Left$(Buffer$, BuddyName%)
AOL% = FindWindow("AOL Frame25", 0&)
Dim mdi As Integer
mdi% = FindChildByClass(AOL%, "MDIClient")
Dim welcome As Integer
welcome = FindChildByTitle(mdi%, "Welcome, ")
Dim MyName As String
MyName = String$(22, 0)
x% = GetWindowText(welcome%, MyName$, 22)
MyName$ = Left$(MyName$, x%)
MyName$ = Mid$(MyName$, InStr(MyName$, ", "), InStr(MyName$, "!") - 1)
MyName$ = Left$(MyName$, Len(MyName$) - 1)
MyName$ = Right$(MyName$, Len(MyName$) - 2)
If MyName$ = FinalBuddyname$ Then GoTo 18
For names = 0 To List1_ListCount - 1
 If FinalBuddyname$ = Lst.List(names) Then GoTo 18
Next names
Lst.AddItem FinalBuddyname$
18:
Next RoomNames
19:
End Sub
Sub Changefocus()
Dim aa, ab As Integer
aa = getfocus()
Do
  DoEvents
  ab = getfocus()
Loop While aa = ab
End Sub
Function dupekill()
Num = countmail()
DELNUM% = 0
AOL% = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL%, "MDIClient")
FNMB% = FindChildByTitle(mdi%, "New Mail")

If FNMB% = 0 Then
RunToolBar (1)
Timeout (0.4)
Do: DoEvents
    NMB% = FindChildByTitle(mdi%, "New Mail")
    Timeout (0.1)
Loop Until NMB% <> 0
WaitMail
End If

Do: DoEvents
LSTTXT$ = ","
DELTXT$ = ","
btnDEL% = FindChildByTitle(FindChildByTitle(mdi%, "New Mail"), "Delete")
If countmail() = 0 Then MsgBox "You have no New Mail.", 12, "Dupe Killer": Exit Function
List% = FindChildByClass(FindChildByTitle(mdi%, "New Mail"), "_AOL_Tree")
For i = 0 To countmail() - 1
Ln = SendMessage(List%, LB_GETTEXTLEN, i, 0)
If Ln = -1 And i >= countmail() Then
    Exit For
ElseIf Ln = -1 And i <= countmail() Then
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
            Num = Num - 1
            i = i - 1
            DELTXT$ = DELTXT$ + MAILTXT$ + ","
Else
LSTTXT$ = LSTTXT$ + MAILTXT$ + ","
End If
Next i
Loop Until Len(DELTXT$) = 1
MsgBox "There were " & DELNUM% & " duplicate mails deleted.", 12, "Dupe Count"
dupekill = DELNUM%
End Function
Function FixAPIString(ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function
Function GetTextFromRICHCNTL(hWindow As Integer)
AOL% = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL%, "MDIClient")
Msg$ = String$(255, 0)
xx% = SendMessageLong(hWindow, WM_ProGGer, 254, Msg$)
Msg$ = Trim$(Msg$)
GetTextFromRICHCNTL = Msg$

End Function
Function WindowCaption(hWndd As Integer)
Dim WindowText As String * 255
Dim GetWinText As Integer
GetWinText = GetWindowText(hWndd, WindowText, 255)
WindowCaption = (WindowText)
End Function

Function r_elite(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If crapp% > 0 Then GoTo dustepp2

If nextchr$ = "A" Then Let nextchr$ = "/\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = ")-("
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "|V|"
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "º"
If nextchr$ = "P" Then Let nextchr$ = "¶"
If nextchr$ = "p" Then Let nextchr$ = "Þ"
If nextchr$ = "r" Then Let nextchr$ = "®"
If nextchr$ = "S" Then Let nextchr$ = "§"
If nextchr$ = "s" Then Let nextchr$ = "$"
If nextchr$ = "t" Then Let nextchr$ = "†"
If nextchr$ = "U" Then Let nextchr$ = "Ú"
If nextchr$ = "u" Then Let nextchr$ = "µ"
If nextchr$ = "V" Then Let nextchr$ = "\/"
If nextchr$ = "W" Then Let nextchr$ = "VV"
If nextchr$ = "w" Then Let nextchr$ = "vv"
If nextchr$ = "X" Then Let nextchr$ = "X"
If nextchr$ = "x" Then Let nextchr$ = "×"
If nextchr$ = "Y" Then Let nextchr$ = "¥"
If nextchr$ = "y" Then Let nextchr$ = "ý"
If nextchr$ = "!" Then Let nextchr$ = "¡"
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = "." Then Let nextchr$ = "…"
If nextchr$ = "," Then Let nextchr$ = "‚"
If nextchr$ = "1" Then Let nextchr$ = "¹"
If nextchr$ = "%" Then Let nextchr$ = "‰"
If nextchr$ = "2" Then Let nextchr$ = "²"
If nextchr$ = "3" Then Let nextchr$ = "³"
If nextchr$ = "_" Then Let nextchr$ = "¯"
If nextchr$ = "-" Then Let nextchr$ = "—"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
r_elite = newsent$

End Function
Function r_hacker(strin As TextBox)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "A" Then Let nextchr$ = "a"
If nextchr$ = "E" Then Let nextchr$ = "e"
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "O" Then Let nextchr$ = "o"
If nextchr$ = "U" Then Let nextchr$ = "u"
If nextchr$ = "b" Then Let nextchr$ = "B"
If nextchr$ = "c" Then Let nextchr$ = "C"
If nextchr$ = "d" Then Let nextchr$ = "D"
If nextchr$ = "z" Then Let nextchr$ = "Z"
If nextchr$ = "f" Then Let nextchr$ = "F"
If nextchr$ = "g" Then Let nextchr$ = "G"
If nextchr$ = "h" Then Let nextchr$ = "H"
If nextchr$ = "y" Then Let nextchr$ = "Y"
If nextchr$ = "j" Then Let nextchr$ = "J"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = "m" Then Let nextchr$ = "M"
If nextchr$ = "n" Then Let nextchr$ = "N"
If nextchr$ = "x" Then Let nextchr$ = "X"
If nextchr$ = "p" Then Let nextchr$ = "P"
If nextchr$ = "q" Then Let nextchr$ = "Q"
If nextchr$ = "r" Then Let nextchr$ = "R"
If nextchr$ = "s" Then Let nextchr$ = "S"
If nextchr$ = "t" Then Let nextchr$ = "T"
If nextchr$ = "w" Then Let nextchr$ = "W"
If nextchr$ = "v" Then Let nextchr$ = "V"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
r_hacker = newsent$

End Function
Function r_same(strr As String)
Let r_same = Trim(strr)

End Function
Function r_spaced(strin As TextBox)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
r_spaced = newsent$

End Function
Function rcase(hehe$)
Do
flet$ = Left$(hehe$, 1)
If UCase$(flet$) = flet$ Then
joe$ = LCase$(flet$)
Else: joe$ = UCase$(flet$)
End If
bah$ = bah$ & joe$
hehe$ = Mid$(hehe$, 2)
DoEvents
Loop Until Len(hehe$) = 0
rcase = bah$

End Function
Function ReadINI(AppName, KeyName, FileName As String) As String
'Example: text4.text = ReadINI("DaProggy", "Lamers Name", app.path + "\Prog.ini")
Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))

End Function
Sub run(ByVal menuname As String)
Dim mhWnd As Integer
Dim MnuCnt As Integer
Dim Cnt As Integer
Dim SmhWnd As Integer
Dim SMnuCnt As Integer
Dim SCnt As Integer
Dim lint As Integer
Dim lpString As String
Dim AOLx
AOLx = FindWindow%("AOL Frame25", 0&)
mhWnd = GetMenu%(AOLx)
MnuCnt = GetMenuItemCount%(mhWnd)
For Cnt = 0 To MnuCnt - 1
    SmhWnd = GetSubMenu%(mhWnd, Cnt)
    SMnuCnt = GetMenuItemCount%(SmhWnd)
    For SCnt = 0 To SMnuCnt - 1
    lint = 30
    lpString = Space(30)
    x = GetMenuString%(SmhWnd, SCnt, lpString, lint, MF_BYPOSITION)
    
    If InStr(UCase$(lpString), UCase$(menuname)) Then
        GMnID = GetMenuItemID%(SmhWnd, SCnt)
        Ret = SendMessageByNum(AOLx, WM_COMMAND, GMnID, 0)
        DoEvents
        Exit Sub
    End If
    DoEvents
    Next SCnt
Next Cnt

End Sub
Function runmenu2(Main_Prog As String, Top_Position As String, Menu_String As String)
Dim Top_Position_Num As Integer
Dim Buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim AOL As Integer
Dim Menu_Handle As Integer
End Function
Function runmenu3(TopMenuPos As Integer, PopupPos As Integer)
Dim i, mnuhWnd, submnu, Menuid  As Integer
Dim lParam As Long
Const MF_BYCOMMAND = &H0
mnuhWnd = GetMenu%(FindWindow(0&, "America  Online"))
submnu = GetSubMenu%(mnuhWnd, TopMenuPos)
Menuid = GetMenuItemID%(submnu, PopupPos)
lParam = CLng(0) * &H10000 Or MF_BYCOMMAND
i = SendMessageByNum(FindWindow(0&, "America  Online"), WM_COMMAND, Menuid, 0&)
End Function
Sub runmenua(horz, vert)
'Runs the specified AOL Menu (Horizonatl,Verticle)
'Each Position starts at 0 not 1

Dim f, gi, sm, m, a As Integer
a = FindWindow("AOL Frame25", 0&)
m = GetMenu(a)
sm = GetSubMenu(m, horz)
gi = GetMenuItemID(sm, vert)
f = SendMessageByNum(a, WM_COMMAND, gi, 0)

End Sub
Sub runmenubystring1(Application, StringSearch)
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
Sub RunToolBar(IcoNum As Integer)
AOL% = FindWindow("AOL Frame25", 0&)
ToolBar% = FindChildByClass(AOL%, "AOL Toolbar")
T1% = FindChildByClass(ToolBar%, "_AOL_Icon")
T2% = GetWindow(T1%, GW_HWNDNEXT)
T3% = GetWindow(T2%, GW_HWNDNEXT)
T31% = GetWindow(T3%, GW_HWNDNEXT)
T4% = GetWindow(T31%, GW_HWNDNEXT)
T5% = GetWindow(T4%, GW_HWNDNEXT)
T6% = GetWindow(T5%, GW_HWNDNEXT)
T7% = GetWindow(T6%, GW_HWNDNEXT)
T8% = GetWindow(T7%, GW_HWNDNEXT)
T9% = GetWindow(T8%, GW_HWNDNEXT)
T10% = GetWindow(T9%, GW_HWNDNEXT)
T11% = GetWindow(T10%, GW_HWNDNEXT)
T12% = GetWindow(T11%, GW_HWNDNEXT)
T13% = GetWindow(T12%, GW_HWNDNEXT)
T14% = GetWindow(T13%, GW_HWNDNEXT)
T15% = GetWindow(T14%, GW_HWNDNEXT)
T16% = GetWindow(T15%, GW_HWNDNEXT)
T17% = GetWindow(T16%, GW_HWNDNEXT)
T18% = GetWindow(T17%, GW_HWNDNEXT)
T19% = GetWindow(T18%, GW_HWNDNEXT)
If IcoNum = 1 Then IcoNum2% = T1%
If IcoNum = 2 Then IcoNum2% = T2%
If IcoNum = 3 Then IcoNum2% = T3%
If IcoNum = 4 Then IcoNum2% = T4%
If IcoNum = 5 Then IcoNum2% = T5%
If IcoNum = 6 Then IcoNum2% = T6%
If IcoNum = 7 Then IcoNum2% = T7%
If IcoNum = 8 Then IcoNum2% = T8%
If IcoNum = 9 Then IcoNum2% = T9%
If IcoNum = 10 Then IcoNum2% = T10%
If IcoNum = 11 Then IcoNum2% = T11%
If IcoNum = 12 Then IcoNum2% = T12%
If IcoNum = 13 Then IcoNum2% = T13%
If IcoNum = 14 Then IcoNum2% = T14%
If IcoNum = 15 Then IcoNum2% = T15%
If IcoNum = 16 Then IcoNum2% = T16%
If IcoNum = 17 Then IcoNum2% = T17%
If IcoNum = 18 Then IcoNum2% = T18%
If IcoNum = 19 Then IcoNum2% = T19%
DoEvents
Click (IcoNum2%)
End Sub


Sub send1(p0162 As Variant)
AOL% = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL%, "MDIClient")
child% = FindChildByClass(mdi%, "AOL Child")
aoledit% = FindChildByClass(child%, "RICHCNTL")
k = SendMessageByString(aoledit%, WM_SETTEXT, 0, p0162)
Z% = SendMessage(aoledit%, WM_CHAR, 13, 0)
Timeout 0.00001
End Sub
Sub send5(chatedit, sill$)
sndtext5 = SendMessageByString(chatedit, WM_SETTEXT, 0, sill$)

End Sub

Sub sendafuk(SendString$)
Dim AOL%, aoledit%
AOL% = FindWindow("AOL FRAME25", 0&)
If AOL% = 0 Then
Exit Sub
End If
aoledit% = FindChildByClass(AOL%, "_AOL_EDIT")
If aoledit% = 0 Then
MsgBox "You have to be in a chatroom dumbfuck!", 56, "Dumbass!"
Exit Sub
End If
x = SendMessageByString(aoledit%, WM_SETTEXT, 0&, SendString$)
DoEvents
x = SendMessageByNum(aoledit%, WM_CHAR, 13, 0&)
DoEvents
End Sub
Sub SENDCHATTXT(win, sill$)
SendNow% = SendMessageByString(win, WM_SETTEXT, 0, sill$)
x% = SendMessageByNum(win, WM_CHAR, 13, 0)
Timeout (0.001)
End Sub
Sub sendemail(to_whom As String, subj As String, mesg As String)

Timeout 0.005
Do
DoEvents
f_cm = findcomposemail()
DoEvents
Loop Until f_cm <> 0

i1% = FindChildByClass(f_cm, "_AOL_Edit")
i2% = GetNextWindow(i1%, 2)
i3% = GetNextWindow(i2%, 2)
i4% = GetNextWindow(i3%, 2)
i5% = GetNextWindow(i4%, 2)
i6% = GetNextWindow(i5%, 2)
i7% = GetNextWindow(i6%, 2)
i8% = GetNextWindow(i7%, 2)

sendtext2 i1%, to_whom
sendtext2 i5%, subj
sendtext2 i8%, mesg
fr1% = FindChildByClass(f_cm, "_AOL_Icon")
AOLClick fr1%
WaitForOk
End Sub

Function ReverseText(TEXT)
For Words = Len(TEXT) To 1 Step -1
ReverseText = ReverseText & Mid(TEXT, Words, 1)
Next Words


End Function
Sub sendit(sill$)
AOL% = FindWindow("AOL Frame25", 0&)
chatlist% = FindChildByClass(AOL%, "_AOL_Edit")
SendNow% = SendMessageByString(chatlist%, WM_SETTEXT, 0, sill$)
x% = SendMessageByNum(chatlist%, WM_CHAR, 13, 0)
Timeout (0.001)
End Sub
Sub sendmailbomb(to_whom As String, subj As String, Msg As String)
If proG_STAT$ = "OFF" Then
Exit Sub
End If
Timeout 0.05
f_cm = FindChildByTitle(aolhwnd(), "Compose Mail")
i1% = FindChildByClass(f_cm, "_AOL_Edit")
i2% = GetNextWindow(i1%, 2)
i3% = GetNextWindow(i2%, 2)
i4% = GetNextWindow(i3%, 2)
i5% = GetNextWindow(i4%, 2)
i6% = GetNextWindow(i5%, 2)
i7% = GetNextWindow(i6%, 2)
i8% = GetNextWindow(i7%, 2)
sendtext2 i1%, to_whom
sendtext2 i5%, subj
sendtext2 i8%, Msg
fr1% = FindChildByClass(f_cm, "_AOL_Icon")
do_wn = SendMessageByNum(fr1%, WM_LBUTTONDOWN, 0, 0&)
  Timeout 0.008
  u_p = SendMessageByNum(fr1%, WM_LBUTTONUP, 0, 0&)
Timeout 0.01
End Sub
Sub sendtext2(handl As Integer, msgg As String)
send_txt = SendMessageByString(handl, WM_SETTEXT, 0, msgg)
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

Sub AOLRunTool(tool)
ToolBar% = FindChildByClass(AOLWindow(), "AOL Toolbar")
iconz% = FindChildByClass(ToolBar%, "_AOL_Icon")
For x = 1 To tool - 1
iconz% = GetWindow(iconz%, 2)
Next x
isen% = IsWindowEnabled(iconz%)
If isen% = 0 Then Exit Sub
AOLIcon (iconz%)
End Sub

Function ScrambleGame(thestring As String)
Dim bytestring As String

thestringcount = Len(thestring$)
If Not Mid(thestring$, thestringcount, 1) = " " Then thestring$ = thestring$ & " "
For Stringe = 1 To Len(thestring$)
characters$ = Mid(thestring$, Stringe, 1)
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
ScrambleGame = Mid(thestrings2$, 1, Len(thestring$) - 1)
End Function


Function ReplaceText(TEXT, charfind, charchange)
If InStr(TEXT, charfind) = 0 Then
ReplaceText = TEXT
Exit Function
End If

For replace = 1 To Len(TEXT)
thechar$ = Mid(TEXT, replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next replace

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

Function StayOnline()
hwndz% = FindWindow(AOLWindow(), "America Online")
childhwnd% = FindChildByTitle(hwndz%, "OK")
AOLButton (childhwnd%)
End Function

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

Function TrimSpaces(TEXT)
If InStr(TEXT, " ") = 0 Then
TrimSpaces = TEXT
Exit Function
End If

For trimspace = 1 To Len(TEXT)
thechar$ = Mid(TEXT, trimspace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next trimspace

TrimSpaces = thechars$
End Function
Sub showaolwins()
fc = FindChildByClass(aolhwnd(), "AOL Child")
req = ShowWindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = GetNextWindow(faa, 2)
Res = ShowWindow(faa, 1)
DoEvents
Loop Until faf = faa


End Sub
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Sub textset(hWnd As Integer, What As String)
Dim R
R = SendMessageByString(hWnd, &HC, 0, What)

End Sub

Sub TMan(sill$, bah$)
AOL% = FindWindow("AOL Frame25", 0&)
chatlist% = FindChildByClass(AOL%, "_AOL_Listbox")
chatwin% = GetParent(chatlist%)
ChaView% = FindChildByClass(chatwin%, "_AOL_View")
chatedit% = FindChildByClass(chatwin%, "_AOL_Edit")
ChatSend% = FindChildByTitle(chatwin%, "Send")
sndtext% = SendMessageByString(ChaView%, WM_SETTEXT, 0, Chr(13) & Chr(9) & sill$ & ":" & Chr(9) & bah$)
SendNow% = SendMessageByNum(ChaView%, WM_CHAR, &HD, 0)
'Call Click(ChatSend%)
Timeout (0.1)

End Sub
Function TOSPhrases() As String
Randomize
phraZes = Int((Val("141") * Rnd) + 1)
If phraZes = "1" Then
phraZes = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf phraZes = "2" Then
phraZes = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
ElseIf phraZes = "3" Then
phraZes = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
ElseIf phraZes = "4" Then
phraZes = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
ElseIf phraZes = "5" Then
phraZes = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
ElseIf phraZes = "6" Then
phraZes = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "7" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "8" Then
phraZes = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "9" Then
phraZes = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf phraZes = "10" Then
phraZes = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "11" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf phraZes = "12" Then
phraZes = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "13" Then
phraZes = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
ElseIf phraZes = "14" Then
phraZes = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
ElseIf phraZes = "15" Then
phraZes = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "16" Then
phraZes = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf phraZes = "17" Then
phraZes = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
ElseIf phraZes = "18" Then
phraZes = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
ElseIf phraZes = "19" Then
phraZes = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
ElseIf phraZes = "20" Then
phraZes = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
ElseIf phraZes = "21" Then
phraZes = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
ElseIf phraZes = "22" Then
phraZes = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
ElseIf phraZes = "23" Then
phraZes = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
ElseIf phraZes = "24" Then
phraZes = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
ElseIf phraZes = "25" Then
phraZes = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
ElseIf phraZes = "26" Then
phraZes = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
ElseIf phraZes = "27" Then
phraZes = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
ElseIf phraZes = "28" Then
phraZes = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf phraZes = "29" Then
phraZes = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf phraZes = "30" Then
phraZes = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf phraZes = "31" Then
phraZes = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf phraZes = "32" Then
phraZes = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf phraZes = "33" Then
phraZes = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
ElseIf phraZes = "34" Then
phraZes = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
ElseIf phraZes = "35" Then
phraZes = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
ElseIf phraZes = "36" Then
phraZes = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf phraZes = "37" Then
phraZes = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf phraZes = "38" Then
phraZes = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf phraZes = "39" Then
phraZes = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf phraZes = "40" Then
phraZes = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
ElseIf phraZes = "41" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "42" Then
phraZes = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "43" Then
phraZes = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "44" Then
phraZes = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "45" Then
phraZes = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf phraZes = "46" Then
phraZes = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf phraZes = "47" Then
phraZes = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf phraZes = "48" Then
phraZes = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
ElseIf phraZes = "49" Then
phraZes = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
ElseIf phraZes = "50" Then
phraZes = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
ElseIf phraZes = "51" Then
phraZes = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
ElseIf phraZes = "52" Then
phraZes = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
ElseIf phraZes = "53" Then
phraZes = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "54" Then
phraZes = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf phraZes = "55" Then
phraZes = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "56" Then
phraZes = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
ElseIf phraZes = "57" Then
phraZes = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "58" Then
phraZes = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "59" Then
phraZes = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
ElseIf phraZes = "60" Then
phraZes = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
ElseIf phraZes = "61" Then
phraZes = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
ElseIf phraZes = "62" Then
phraZes = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
ElseIf phraZes = "63" Then
phraZes = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "64" Then
phraZes = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "65" Then
phraZes = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
ElseIf phraZes = "66" Then
phraZes = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
ElseIf phraZes = "67" Then
phraZes = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
ElseIf phraZes = "68" Then
phraZes = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
ElseIf phraZes = "69" Then
phraZes = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
ElseIf phraZes = "70" Then
phraZes = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
ElseIf phraZes = "71" Then
phraZes = "TrmsReport:  Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
ElseIf phraZes = "72" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
ElseIf phraZes = "73" Then
phraZes = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
ElseIf phraZes = "74" Then
phraZes = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
ElseIf phraZes = "75" Then
phraZes = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
ElseIf phraZes = "76" Then
phraZes = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
ElseIf phraZes = "77" Then
phraZes = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf phraZes = "78" Then
phraZes = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
ElseIf phraZes = "79" Then
phraZes = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
ElseIf phraZes = "80" Then
phraZes = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
ElseIf phraZes = "81" Then
phraZes = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf phraZes = "82" Then
phraZes = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf phraZes = "83" Then
phraZes = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf phraZes = "84" Then
phraZes = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf phraZes = "85" Then
l012A = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "86" Then
phraZes = "Hi, I'm with FCB (Federal Credit Bureau). Due to The type of payment you are using, we need to validate it. Please reply with your credit card number, expiration date, full name, address, and phone number. Thank you. "
ElseIf phraZes = "87" Then
phraZes = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "
ElseIf phraZes = "88" Then
phraZes = "AOLBilling:    Hello from the AOL TOS Staff. We have lost your billing info.  Please click respond and send your CC#, Expiration Date, and Billing Zip Code. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf phraZes = "89" Then
phraZes = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "90" Then
phraZes = "AOLBilling:    Hello from the AOL TOS Staff. We have lost your account info. Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf phraZes = "91" Then
phraZes = "Hello I am from the AOL service department and there has been a minor problem with your accout. I see that we have lost your password information and need it as soon as possible. We cannot access any of your personal info without your password. If you could please type your full log-on password and click the send button with your response. Thank you and Enjoy America online."
ElseIf phraZes = "92" Then
phraZes = "Hello, I am from AOLnet and there has been a network overload on our computers. We have been trying fix this problem but have had no luck. All your account information has been lost. In order for AOL to fix this we need your full First Name and Last, Your phone number, Address, Zip code, Credit Card Number and experation date, as well as your Screen Name(s) and Log-On Password(s). If you fail to reply will result in account termination, thank you :-)"
ElseIf phraZes = "93" Then
phraZes = "Hello I am a member of the America Online Stratus Team. We have been experiancing some problems with your preticular account. It would help us if you gave is your Screen Name(s) and Log-on Password(s) so we can remedy this problem. Thank you AOLST. "
ElseIf phraZes = "94" Then
phraZes = "Hello , AOL Customer! I am sorry to Inform you account information Was not Correctly Processed into our Stratus System. We are now forced to ask you to respond with your personal log on password and you complete billing info (Name, Address, & CC Number Information) . Thank-you for your Time, And Continue To enjoy AOL!"
ElseIf phraZes = "95" Then
phraZes = "Guide HAP:   Hello I work for AOL.  When you signed on we received an INCORRECT password.  Please respond with the CORRECT password now or face possible account termination for safety reasons, Thank you. =]"
ElseIf phraZes = "96" Then
phraZes = "Hi, I'm with STC (Standard Tele-Comm. Inc.) Due to a problem in phone transmission at your baud rate, I need to verify your log-on password before it can be fixed. Failure to comply will result in billing errors.  "
ElseIf phraZes = "97" Then
phraZes = "I am OnlineHost #147; an On-Line operator for America Online. Recently, computer Hackers have been logging on to the Network with your account. I have been logging them off when they failed validation. Please validate yourself as a registered user by stating your Credit Card# and expiration date. Thank you for your cooperation."
ElseIf phraZes = "98" Then
phraZes = "Good evening. I work with AOL's Network Security Division. Over the past few days, computer 'Hackers' have been logging on to the system using your account.  Please verify yourself as a registered user by stating your password, or face immediate account invalidation."
ElseIf phraZes = "99" Then
phraZes = "Hello! I am currently standing in as the TOSAdvisor because he is on vacation. I am sorry to inform you that your Password has been altered by computer Hackers. Your account will be terminated unless we can verify you as a registered user. Please type in your Password for system indentification. Thanks! :)"
ElseIf phraZes = "100" Then
phraZes = "Hello, I am with the America Online Billing department.  Please disreagrd the red letters at the bottom of the Instant Message window.  We have that there due to hackers who recently have been soliciting for passwords.  But Due to some sort of line noise or error when you logged on we were not able to verify your log on password.  Please reply with your log on password so that we may verify you, and we are sorry for the inconvinience."
ElseIf phraZes = "101" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf phraZes = "102" Then
phraZes = "This is a message sent out to random America Online members for a catalog containing -cheap- Microsoft software. Please send your request for the catalog along with your credit card number fo the amount $19.95. After 3 weeks processing time you'll recive a copy of Microsoft Windows96 Nashville version. *Microsoft INC.*"
ElseIf phraZes = "103" Then
phraZes = "*America Online Screen Name Dept.* Sorry for this inconvenience but somehow your screen name has been duplicated by another member. Please change your screen name and E-Mail me with your new screen name and password, So you can be verified as a AOL member. Thank you, And you have 5 days to reply."
ElseIf phraZes = "104" Then
phraZes = "Hello, I am with Sprint United Telephone. There has been damages to the fiber optic network wich you are connected to America Online with. I must ask for your log-on password so AOL can verify your current membership status. Failure to do so will result in account termination. Continue to enjoy your use of AOL courtesy of Sprint-TelNet."
ElseIf phraZes = "105" Then
phraZes = "Hi, I am with the AOL inc. User Services Department. We are sorry to inform you that your account has not been logged correctly please re-enter your personal logon Password, So you can continue to stay with us on America Online...Hope to see you soon..."
ElseIf phraZes = "106" Then
phraZes = "Hello, I am a service representative of America On Line.  Due to a computer error in our password files, we ask that you change your current log-in password to  'Secure', until futher advised, for your account's security. For your inconvience we at America On Line would like to offer our thanks by adding 2 Free hours of online time to your account. Thank you for your time and inconvience."
ElseIf phraZes = "107" Then
phraZes = "Hello this is the America Online party service and it seems that you have one a new 566.6 modem. This is the fastest modem out to this day, for you to redeem this prize we need your full log-on password to make sure you are not a hacker on this account, Thank you and have a nice day!"
ElseIf phraZes = "108" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "109" Then
phraZes = "Hello. I work for the America On-Line User Database. Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password onlineHello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation.. Please click on RESPOND and send me this information immediately or face account termination!  Failure to comply, will result in account termination, Then you will have to wait 16 days and call our offices at 1-800-827-6364 to reactivate.Thank you for your cooperation, America On-Line User Services."
ElseIf phraZes = "110" Then
phraZes = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "111" Then
phraZes = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "112" Then
phraZes = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "113" Then
phraZes = "Hello, I am form the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf phraZes = "114" Then
phraZes = "Hello, I am from the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf phraZes = "115" Then
phraZes = "Greetings CyberSurfer! I Am A Representative For the Online Banking System on AOL. As You May Have Read Credit Card Fraud is Huge on the 'Net'. Whould You Like To Safe Guard Your Credit Card Info With AOL? Just Reply With Your Credit Card Number And Full Information So That We may Logg Your Card Into Our Computers So No one else But You May Access It Online! We Are Offering This Safety Measure At No Cost to You At All!"
ElseIf phraZes = "116" Then
phraZes = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf phraZes = "117" Then
phraZes = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf phraZes = "118" Then
phraZes = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "119" Then
phraZes = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf phraZes = "120" Then
phraZes = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
ElseIf phraZes = "121" Then
phraZes = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day."
ElseIf phraZes = "122" Then
phraZes = "Welcome, I Am with America Online Credit Validators. Your credit card Information Has gone out of date. If you would like your account to stay active please respond with new credit card info, please reply with all the information so this error will not occur again. We at America Online are deticated In helping our customers! please feel free to call our 1-800 Number about our new credit plan. Thank you."
ElseIf phraZes = "123" Then
phraZes = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
ElseIf phraZes = "124" Then
phraZes = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
ElseIf phraZes = "125" Then
phraZes = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day"
ElseIf phraZes = "126" Then
phraZes = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
ElseIf phraZes = "127" Then
phraZes = "Welcome To America Online! Thank you for choosing us as your online provider. We need to verify your billing information, so if you could please respond with your full name, address, city, state, zip code, and credit card information we will update your system. Sorry for the incovenience. Thank You!"
ElseIf phraZes = "128" Then
phraZes = "Welcome to America Online! While you were entering your billing information into your account systems our server went down and lost all the information conserning your account. If you could please respond with your full name, address, city, state, zip, credit card information, and log-on Password we can quickly update your syetem. Sorry for the misshap and thanks in advance! :-)"
ElseIf phraZes = "129" Then
phraZes = "Good Evenning, I am with the CAT (Community Action Team) of AOL. We are search your records and your Screen name came up as a unregisterd user. Please respond with your name, address, phone #, city , state, zip, & all of your credit card information (Number, exp. date, ect...). Sorry for circumstances but this is **VERY IMPORTANT**!"
ElseIf phraZes = "130" Then
phraZes = " Due to our last system crash on 8/12/96 we have lost all infromation concerning your account. Please respond with your Full address and Credit Card Number. NOTE: 60 mins are credited to your account because of the crash also."
ElseIf phraZes = "131" Then
phraZes = "  We're sorry to report that your account information has been erased.  For you to stay online we need you to verify your password, Credit Card number, Expiration date, and last 4 digits of your phone number.  Thank you and sorry for the delay."
ElseIf phraZes = "132" Then
phraZes = "Hello, my name is Todd Stiger.  I work for the internet.  I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your credit card number and personal information.  Thanks!"
ElseIf phraZes = "133" Then
phraZes = "Sir, I have some bad news! We here at AOL have made a SERIOUS billing error. We have your credit card number as 5436-2854-3546-9374, If this is not correct, please respond with the corrct number and your Full name, Address, State, City, and Zip code.  Sorry for the trouble!"
ElseIf phraZes = "134" Then
phraZes = "I am afraid that you have been SHARING your account with other members! THAT IS AGAINST TOS AGREEMENT!! Please respond with your credit card number, Full Name, and FULL address! This is the only way to prove wether or not you are sharing your account!"
ElseIf phraZes = "135" Then
phraZes = "Hello user, my name is William Denbo with the America Online Password Dept.! We have reason to beleive that a person has gotten hold of your password and is using it to conduct illegal activities. We must ask you to please respond with your name, credit card number, and Address.This is very important.Do not delay!"
ElseIf phraZes = "136" Then
phraZes = "Welcome new AOL user,Our service was temporarily shut down this morning between the times of 6:30a.m. - 7:30a.m.,Due to a complicated computer install!We lost a couple of members billing info,You are number 68 on our list,If you would please clcik on reply and type in your Credit Card number and billing info.!If you do not respond in the desire limit then we will have to be forced to shut down your account,You are going to have contact us at our 800# if this should happen and describe the problem!Thank you for your time. :-)"
ElseIf phraZes = "137" Then
phraZes = "Hello, I am with the AHD (Anti-Hacker Department) and we beleive that Hackers may be using your account for illegal activites. To verify you are the rightful owner of that account; please respond with you name, password, and credit card information. Thank you and have a nice day using America Online."
ElseIf phraZes = "138" Then
phraZes = "I have a program that can kick you off-lin and kill your account.So give me your password and I wont do it.But,if you dont give me your password I'll ruin your credit card,and kill your account!So don't fuck around!!!!!!"
ElseIf phraZes = "139" Then
phraZes = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf phraZes = "140" Then
phraZes = "Good morning current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a virus in the system earlier and it may of caused errors in your billing,When you reply with the information you entered during sign-up so we can verify your account and send you your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf phraZes = "141" Then
phraZes = "There have been notable billing errors showing in our records. For us to correct this matter efficiently, please hit rrespond and give me your full name, your full address, and your credit card number with experation date.  I am sorry for the inconvienence!"
End If
End Function
Sub vbmsgg()
chattxt$ = agGetStringFromLPSTR$(lParam)
End Sub
Sub waitforchange()
Dim Old As Integer
Dim Boy As Integer
Old = getfocus()
Boy = getfocus()
Do Until Boy <> Old
x = DoEvents()
Boy = getfocus()
Loop
End Sub
Function WaitForMail()
Dim HWNDS() As Integer
Dim R, x%, i, T, g%
x = 0
DoEvents
g% = GetWindow(AOhWnd(), 5)
DoEvents
Do
If FindChildByTitle(g%, "Read") <> 0 Then x = x + 1
DoEvents
If FindChildByTitle(g%, "Ignore") <> 0 Then x = x + 1
DoEvents
If FindChildByTitle(g%, "Keep As New") <> 0 Then x = x + 1
DoEvents
If FindChildByTitle(g%, "Delete") <> 0 Then x = x + 1
DoEvents
If FindChildByClass(g%, "_AOL_Tree") <> 0 Then x = x + 1
If x = 5 Then GoTo Foundd
DoEvents
g% = GetWindow(g%, 2)
DoEvents
x = 0
Loop While g% <> 0


Exit Function
Foundd:
DoEvents
FindMailWnd = g%
Exit Function
End Function
Sub waitformok()
Do
DoEvents
Okw = FindWindow("_AOL_Modal", 0&)
DoEvents
Loop Until Okw <> 0
Okw = FindWindow("_AOL_Modal", 0&)
killwin (Okw)

End Sub
Sub WaitMail()
Do
box = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail")
Timeout (0.1)
Loop Until box <> 0
List = FindChildByClass(box, "_AOL_Tree")
Do
DoEvents
mailnum = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(0.5)
mailnum2 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(0.5)
mailnum3 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
    mailnum = SendMessage(List, LB_GETCOUNT, 0, 0&)


End Sub

Function WindowsDirectory() As String
Dim WinPath As String
WinPath = String(145, Chr(0))
WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))

End Function
Sub WriteINI(sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim R As Integer
    R = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Sub AddCombo(Combo As ComboBox, txt$)
For x = 0 To Combo.ListCount - 1
    If UCase$(Combo.List(x)) = UCase$(txt$) Then Exit Sub
Next
Combo.AddItem txt$
End Sub

Sub addcomboroom(Combo As ComboBox)
Chat% = findchatroom()
AolList% = FindChildByClass(Chat%, "_AOL_ListBox")
Num = SendMessageByNum(AolList%, LB_GETCOUNT, 0, 0)
x = SetFocusAPI(Chat%)
For i% = 0 To Num - 1
    namez$ = String$(11, " ")
    Ret = AOLGetList(i%, namez$)
    namez$ = Trim$(namez$)
    sn$ = usersn()
    
    If Trim$(UCase$(namez$)) = Trim$(UCase(sn$)) Then GoTo onn2
    Call AddCombo(Combo, namez$)
onn2:
Next

End Sub
Function AOLChild()
Dim AOL%, mdi%, child%
AOL% = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL%, "MDIClient")
child% = FindChildByClass(mdi%, "AOL Child")
AOLChild = child%
End Function

Function ChatView() As Integer
Chat% = findchatroom()
view% = FindChildByClass(Chat%, "_AOL_View")
ChatView = view%
End Function

Function conelite(word$) As String
made$ = ""
For x = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, x, 1)
    leet$ = ""
    If letter$ = "a" Then leet$ = "@"
    If letter$ = "b" Then leet$ = "þ"
    If letter$ = "c" Then leet$ = "©"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then leet$ = "ë"
    If letter$ = "f" Then leet$ = "ƒ"
    If letter$ = "g" Then leet$ = "g"
    If letter$ = "h" Then leet$ = "h"
    If letter$ = "i" Then leet$ = "ï"
    If letter$ = "j" Then leet$ = "j"
    If letter$ = "k" Then leet$ = "k"
    If letter$ = "l" Then leet$ = "l"
    If letter$ = "m" Then leet$ = "m"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then leet$ = "ø"
    If letter$ = "p" Then leet$ = "p"
    If letter$ = "q" Then leet$ = "q"
    If letter$ = "r" Then leet$ = "®"
    If letter$ = "s" Then leet$ = "$"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then leet$ = "ü"
    If letter$ = "v" Then leet$ = "v"
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "x" Then leet$ = "×"
    If letter$ = "y" Then leet$ = "ý"
    If letter$ = "z" Then leet$ = "z"
    If letter$ = " " Then leet$ = " "
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "1" Then leet$ = "1"
    If letter$ = "2" Then leet$ = "2"
    If letter$ = "3" Then leet$ = "3"
    If letter$ = "4" Then leet$ = "4"
    If letter$ = "5" Then leet$ = "5"
    If letter$ = "6" Then leet$ = "6"
    If letter$ = "7" Then leet$ = "7"
    If letter$ = "8" Then leet$ = "8"
    If letter$ = "9" Then leet$ = "9"
    If letter$ = "0" Then leet$ = "0"
    If letter$ = "A" Then leet$ = "Å"
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "F" Then leet$ = "ƒ"
    If letter$ = "G" Then leet$ = "G"
    If letter$ = "H" Then leet$ = "/-/"
    If letter$ = "I" Then leet$ = "‡"
    If letter$ = "J" Then leet$ = "J"
    If letter$ = "K" Then leet$ = "]<"
    If letter$ = "L" Then leet$ = "£"
    If letter$ = "M" Then leet$ = "/\/\"
    If letter$ = "N" Then leet$ = "/\/"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "P" Then leet$ = "¶"
    If letter$ = "Q" Then leet$ = "`Q"
    If letter$ = "R" Then leet$ = "•R•"
    If letter$ = "S" Then leet$ = "§"
    If letter$ = "T" Then leet$ = "†"
    If letter$ = "U" Then leet$ = "Ü"
    If letter$ = "V" Then leet$ = "\/"
    If letter$ = "W" Then leet$ = "\\'"
    If letter$ = "X" Then leet$ = "><"
    If letter$ = "Y" Then leet$ = "¥"
    If letter$ = "Z" Then leet$ = "Z"
    If letter$ = "~" Then leet$ = "~"
    If letter$ = "`" Then leet$ = "`"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "@" Then leet$ = "ä"
    If letter$ = "#" Then leet$ = "#"
    If letter$ = "$" Then leet$ = "$"
    If letter$ = "%" Then leet$ = "%"
    If letter$ = "^" Then leet$ = "^"
    If letter$ = "&" Then leet$ = "&"
    If letter$ = "*" Then leet$ = "™"
    If letter$ = "(" Then leet$ = "("
    If letter$ = ")" Then leet$ = ")"
    If letter$ = "-" Then leet$ = "-"
    If letter$ = "_" Then leet$ = "_"
    If letter$ = "+" Then leet$ = "+"
    If letter$ = "=" Then leet$ = "="
    If letter$ = "[" Then leet$ = "["
    If letter$ = "]" Then leet$ = "]"
    If letter$ = "{" Then leet$ = "{"
    If letter$ = "}" Then leet$ = "}"
    If letter$ = ":" Then leet$ = ":"
    If letter$ = ";" Then leet$ = ";"
    If letter$ = "'" Then leet$ = "'"
    If letter$ = "," Then leet$ = ","
    If letter$ = "." Then leet$ = "."
    If letter$ = "<" Then leet$ = "<"
    If letter$ = ">" Then leet$ = ">"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    made$ = made$ & leet$
Next
conelite = made$
End Function

Function conleet(message As String) As String
        msgln = Len(message$)
        leet$ = message$
        message$ = ""
        lo = 1
        up = True
        For x = 1 To msgln
            letter$ = Mid$(leet$, lo, 1)
            If up = True Then
                    If UCase$(letter$) = "I" Then
                        message$ = message$ & "i"
                    ElseIf UCase$(letter$) = "O" Then
                        message$ = message$ & "0"
                    ElseIf UCase$(letter$) = "E" Then
                        message$ = message$ & "3"
                    ElseIf UCase$(letter$) = "A" Then
                        message$ = message$ & "4"
                    Else:
                        message$ = message$ & UCase$(letter$)
                    End If
                up = False
                ElseIf up = False Then
                If UCase$(letter$) = "L" Then
                        message$ = message$ & "L"
                    ElseIf UCase$(letter$) = "O" Then
                        message$ = message$ & "0"
                    ElseIf UCase$(letter$) = "E" Then
                        message$ = message$ & "3"
                    ElseIf UCase$(letter$) = "A" Then
                        message$ = message$ & "4"
                    Else:
                        message$ = message$ & LCase$(letter$)
                    End If
                up = True
            End If
            lo = lo + 1
        Next
conleet = message$
End Function

Sub enter(edt%)
x = SendMessageByNum(edt%, WM_CHAR, 13, 0)
End Sub
Sub exitprog()
ext = MsgBox("exit?", 68)
If ext = 6 Then
    AOL% = FindWindow("AOL Frame25", 0&)
    End
End If
End Sub
Function Extention(exe$, ext$)
If InStr(exe$, ".") Then
    where = InStr(exe$, ".")
    FileName$ = Mid$(exe$, 1, where - 1)
    FileType$ = Mid$(exe$, where + 2)
    If FileType$ <> ext$ Then FileType$ = ext$
    exe$ = FileName$ + "." + FileType$
ElseIf InStr(exe$, ".") = False Then
    exe$ = exe$ + "." + ext$
End If
End Function
Sub KuRuPtadd(List As ListBox)
who$ = InputBox("EPE Online.  Enter a Screen name to add.")
If who$ = "" Then Exit Sub
Call KuRuPtList(List, who$)
End Sub
Sub kuruptaddlist(TheList As ListBox)
x = SignedOn()
If x = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
Chatroom% = findchatroom()
If Chatroom% = 0 Then
    MsgBox "You must be in a chat room to use this feature.", 22, "No Chat Room!"
    Exit Sub
End If
AolList% = FindChildByClass(Chatroom%, "_AOL_Listbox")
yoursn = GetScreenName()
LBSize = SendMessage(AolList%, LB_GETCOUNT, 0, 0)
For GetNames% = 0 To LBSize - 1
    storage$ = String$(255, 0)
    operation = AOLGetList(GetNames%, storage$)
    storage$ = Left$(storage$, InStr(storage$, Chr$(0)) - 1)
    If UCase$(storage$) = UCase$(yoursn) Then GoTo ProGGer
    For BEBE = 0 To TheList.ListCount - 1
        If UCase$(storage$) = UCase$(TheList.List(BEBE)) Then
            GoTo ProGGer
        End If
    Next BEBE
    TheList.AddItem storage$
ProGGer:
Next GetNames%

End Sub
Sub KuRuPtRemove(KuRuPt As ListBox)
If KuRuPt = "" Then Exit Sub
KuRuPt.RemoveItem KuRuPt.ListIndex
End Sub
Function kuruptsn(Kur As String, ByVal rupt As String, Lug As String) As String
Dim LorD As String
Dim Rakwon As String
Dim Shocker As String

KuRuPt:
If InStr(UCase(Kur), UCase(rupt)) Then
        LorD$ = Left(Kur, InStr(UCase(Kur), UCase(rupt)) - 1)
        Rakwon$ = Mid(Kur, InStr(UCase(Kur), UCase(rupt)) + Len(rupt), Len(Kur))
        Shocker = LorD & KooL & Rakwon
    Else:
        Shocker = Kur
End If
kuruptsn = Shocker
End Function
Function Percent(Complete As Integer, Total As Integer, TotalOutput As Integer) As Integer
R = Int(Complete / Total * TotalOutput)
Percent = R
End Function
Sub SetEdit(edt%, txt$)
x = SendMessageByString(edt%, WM_SETTEXT, 0, txt$)
End Sub
Sub VBMsg1_WindowMessage(hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
Dim ChatText$, colon%, sn$, Textsaid$, Kword$, a%, b%, c%, edithWnd%, e%, msgnam%, SEND%
    ChatText$ = agGetStringFromLPSTR(lParam)
    colon% = InStr(1, ChatText$, ":")
    sn$ = Mid(ChatText$, 3, (colon% - 3))
    Textsaid$ = Right(ChatText$, (Len(ChatText$) - Len(sn$)) - 4)
    Kword$ = Text1
    If UCase$(Left$(Textsaid$, Len(Kword$))) = UCase$(Kword$) Then
        DoEvents
Call sendtext("I got it")

End If
End Sub




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


Function KTEncrypt(ByVal PassWord, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(PassWord) = 0 Then Error 31100
  
  'Is password too long
  If Len(PassWord) > 255 Then Error 31100

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
  PassMax = Len(PassWord)
  
  
  'Tack on leading characters to prevent repetative recognition
  PassWord = Chr$(Asc(Left$(PassWord, 1)) Xor PassMax) + PassWord
  PassWord = Chr$(Asc(Mid$(PassWord, 1, 1)) Xor Asc(Mid$(PassWord, 2, 1))) + PassWord
  PassWord = PassWord + Chr$(Asc(Right$(PassWord, 1)) Xor PassMax)
  PassWord = PassWord + Chr$(Asc(Right$(PassWord, 2)) Xor Asc(Right$(PassWord, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(PassWord, 3) + Format$(Asc(Right$(PassWord, 1)), "000") + Format$(Len(PassWord), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(PassWord, PassUp, 1))

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
    If Left$(strng, 9) <> Left$(PassWord, 3) + Format$(Asc(Right$(PassWord, 1)), "000") + Format$(Len(PassWord), "000") Then
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

Public Sub Center(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub

Sub Click(button%)
SendNow% = SendMessageByNum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(button%, WM_LBUTTONUP, &HD, 0)
End Sub
Function Currentroom()
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByClass(AOL, "_AOL_Glyph")
par% = GetParent(bah)
x$ = GetWinText(par%)
Currentroom = x$
End Function
Sub Disablemodal()
AOM = FindWindow("_AOL_Modal", 0&)
x = EnableWindow(AOM, 0)
AOL = FindWindow("AOL Frame25", 0&)
x = EnableWindow(AOL, 1)
End Sub
Sub EnableHwnd(windo)
x = EnableWindow(windo, 1)
End Sub
Sub Enablemodal()
AOM = FindWindow("_AOL_Modal", 0&)
x = EnableWindow(AOM, 1)
AOL = FindWindow("AOL Frame25", 0&)
x = EnableWindow(AOL, 0)
End Sub
Sub ExitNice()
Res = ShowCursor(True)
Res = SystemParametersInfo(17, 1, ByVal 0&, 0)
End
End Sub
Function FCBC(parhwnd, hwn)
AOL = FindWindow("AOL Frame25", 0&)
child = GetWindow(AOL, GW_CHILD)
child = GetWindow(child, GW_HWNDFIRST)
    child1 = child
    Do
    werd$ = String(255, " ")
    bah = GetClassName(child1, werd$, 255)
    If InStr(LCase$(werd$), LCase$(hwn)) Then
    FCBC = child1: Exit Function
    End If
    child1 = GetWindow(child1, GW_HWNDNEXT)
    Loop Until child1 = 0

child = GetWindow(child, GW_CHILD)
child = GetWindow(child, GW_CHILD)
child = GetWindow(child, GW_HWNDFIRST)
Do
    child1 = child
    Do
    werd$ = String(255, " ")
    bah = GetClassName(child1, werd$, 255)
    If InStr(LCase$(werd$), LCase$(hwn)) Then
    FCBC = child1: Exit Function
    End If
    child1 = GetWindow(child1, GW_HWNDNEXT)
    Loop Until child1 = 0
child = GetWindow(child, GW_HWNDNEXT)
Loop Until child = 0
End Function

Sub formgrow(FRM As Form)
bah = FRM.Height
bla = FRM.Width
FRM.Width = 0
FRM.Height = 0
FRM.Visible = True
speed = 200
Do
If FRM.Height + speed > bah Then
FRM.Height = bah
FRM.Width = bla
Exit Sub
End If
FRM.Height = FRM.Height + speed
FRM.Width = FRM.Width + (speed * (bla / bah))
DoEvents
Loop
FRM.Height = bah
FRM.Width = bla
End Sub
Sub formmove(mw As Form)
Dim Ret&
ReleaseCapture
Ret& = SendMessage(mw.hWnd, &H112, &HF012, 0)
End Sub
Sub getcombonames(Lst As ListBox)
Lst.Clear
For Index% = 0 To 25
names$ = String$(256, " ")
Ret = AOLGetcombo(Index%, names$) & ErB$
If Len(Trim$(names$)) <= 1 Then Exit For
names$ = Left$(Trim$(names$), Len(Trim(names$)) - 1)
Lst.AddItem (names$)
Next Index%
End Sub
Sub im(person$, tt$)
Okw = FindWindow("#32770", "America Online")
okb = FindChildByTitle(Okw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
run "Send an Instant Message"
Do
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(AOL, "Send Instant Message")
txt% = FindChildByClass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
nah = FindChildByTitle(AOL, "To:")
x = SendMessageByString(nah, WM_SETTEXT, 0, "To: ")
txt% = FindChildByClass(bah, "_AOL_Edit")
Do
rich% = FindChildByClass(bah, "RICHCNTL")
bahqw% = FindChildByTitle(bah, "Send")
DoEvents
Timeout (0.001)
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then
SEND txt%, person$
SEND rich%, tt$
Timeout (0.001)
getnum rich%, 1
Click rich%
Else
SEND txt%, person$
getnum txt%, 1
SEND txt%, tt$
getnum txt%, 1
Click txt%
End If
Timeout (0.001)
Do
AOL = FindWindow("AOL Frame25", 0&)
Okw = FindWindow("#32770", "America Online")
ba1h = FindChildByTitle(AOL, "To: ")
DoEvents
Loop Until ba1h = 0 Or Okw <> 0
If Okw <> 0 Then
WaitForOk
killwin (bah)
End If
End Sub
Sub listclick(windo, Index)
SendNow = SendMessageByNum(windo, LB_SETCURSEL, Index, 0)
sendlater = SendMessageByNum(windo, LBM_DBLCLK, Index, 0)
End Sub
Sub make3d(TheForm As Form, TheControl As Control)
If TheForm.AutoRedraw = False Then
OldMode = TheForm.ScaleMode
TheForm.ScaleMode = 3
TheForm.AutoRedraw = True
TheForm.CurrentX = TheControl.Left - 1
TheForm.CurrentY = TheControl.Top + TheControl.Height
TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
TheForm.AutoRedraw = False
TheForm.ScaleMode = OldMode
End If
If TheForm.AutoRedraw = True Then
OldMode = TheForm.ScaleMode
TheForm.ScaleMode = 3
TheForm.CurrentX = TheControl.Left - 1
TheForm.CurrentY = TheControl.Top + TheControl.Height
TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
TheForm.ScaleMode = OldMode
End If


End Sub
Function mysn()
AOL = FindWindow%("AOL Frame25", "America  Online")
Happy% = FindChildByTitle%(AOL, "Welcome,")
joy$ = GetWinText((Happy%))
If InStr(joy$, "!") = 0 Then Exit Function
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
mysn = yoursn
End Function
Sub nubtos(who$, phrase$)
keywor "nub"
Do
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(AOL, "NuB")
DoEvents
Loop Until bah <> 0
heh% = FindChildByClass(bah, "_AOL_Icon")
getnum heh%, 4
Click heh%
DoEvents
killwin (bah)
Do
AOL = FindWindow("AOL Frame25", 0&)
he% = FindChildByTitle(AOL, "Lotsen")
DoEvents
Loop Until he% <> 0
heh% = FindChildByClass(he%, "_AOL_Icon")
getnum heh%, 2
Click heh%
DoEvents
killwin (he%)
Do
AOL = FindWindow("AOL Frame25", 0&)
he% = FindChildByTitle(AOL, "Lotsen")
DoEvents
Loop Until he% <> 0
heh% = FindChildByClass(he%, "_AOL_Icon")
getnum heh%, 3
Click heh%
DoEvents
killwin (he%)
Do
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(AOL, "eMail")
DoEvents
Loop Until bah <> 0
edi = FindChildByClass(bah, "_AOL_Edit")
SEND edi, who$ & ":" & Chr(9) & phrase$
abschicken% = FindChildByTitle(bah, "Abschicken")
Click abschicken%
DoEvents
killwin (bah)
WaitForOk
End Sub
Function NumberOfPeopleInRoom()
AOL = FindWindow("AOL Frame25", 0&)
heh% = FindChildByClass(AOL, "_AOL_View")
getnum heh%, 7
bah = GetWinText(heh%)
NumberOfPeopleInRoom = bah
End Function
Sub paint(FRM As Form)
FRM.Cls
Const PIXEL = 3
'frm.ScaleMode = PIXEL
hDestDC% = FRM.hDC
x% = 0: Y% = 0
nWidth% = Screen.Width
nHeight% = Screen.Height
hSrcDC% = GetDC(0&)
XSrc% = 0: YSrc% = 0
dwRop& = &HCC0020
Suc% = BitBlt(hDestDC%, x%, Y%, nWidth%, nHeight%, hSrcDC%, XSrc%, YSrc%, dwRop&)
End Sub
Sub punt(person$, tt$)
Okw = FindWindow("#32770", "America Online")
okb = FindChildByTitle(Okw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
run "Send an Instant Message"
Do
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(AOL, "Send Instant Message")
txt% = FindChildByClass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
nah = FindChildByTitle(AOL, "To:")
x = SendMessageByString(nah, WM_SETTEXT, 0, "To: ")
txt% = FindChildByClass(bah, "_AOL_Edit")
Do
rich% = FindChildByClass(bah, "RICHCNTL")
bahqw% = FindChildByTitle(bah, "Send")
DoEvents
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then
SEND txt%, person$
SEND rich%, tt$
getnum rich%, 1
Click rich%
Else
SEND txt%, person$
getnum txt%, 1
SEND txt%, tt$
getnum txt%, 1
Click txt%
End If
End Sub
Sub renamehost()
Static Heyas As String * 40000
Dim fsdfsd As Long
Dim blehsa As Long
Dim hehe As Integer
Dim qwerty As Integer
Dim moocow As Variant
Dim nooos As Integer
sn = "KewL"

If Len(sn) < 10 Then
Do
sn = sn & "KewL "
Loop Until Len(sn) = 10
End If
If Len(tru_sn) < 10 Then
Do
tru_sn = tru_sn & "KewL "
Loop Until Len(sn) = 10
End If
On Error Resume Next
tru_sn = "OnlineHost" + String$(Len(sn) - 10, " ")
Let paath$ = ("C:\aol30\tool\aolchat.aol")
Open paath$ For Binary As #1
fsdfsd& = 1
blehsa& = LOF(1)
While fsdfsd& < blehsa&
Heyas = String$(40000, Chr$(0))
Get #1, fsdfsd&, Heyas
While InStr(UCase$(Heyas), UCase$(sn)) <> 0
Mid$(Heyas, InStr(UCase$(Heyas), UCase$(sn))) = tru_sn
Wend

Put #1, fsdfsd&, Heyas
fsdfsd& = fsdfsd& + 40000
Wend
Seek #1, Len(sn)
fsdfsd& = Len(sn)
While fsdfsd& < blehsa&
Heyas = String$(40000, Chr$(0))
Get #1, fsdfsd&, Heyas
While InStr(UCase$(Heyas), UCase$(sn)) <> 0
Mid$(Heyas, InStr(UCase$(Heyas), UCase$(sn))) = tru_sn
Wend
Put #1, fsdfsd&, Heyas
fsdfsd& = fsdfsd& + 40000
Wend
Close #1

End Sub
Sub replace(windo%, gud$, werd$)
mmm = GetWinText(windo%)
If InStr(mmm, gud$) Then
heh$ = Left$(mmm, InStr(mmm, gud$) - 1)
SEND windo%, heh$ & werd$
End If

End Sub


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

Public Sub AOLButton(but%)
clickicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
clickicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function AOLGetUser()
On Error Resume Next
AOL& = FindWindow("AOL Frame25", "America  Online")
mdi& = FindChildByClass(AOL&, "MDIClient")
welcome% = FindChildByTitle(mdi&, "Welcome, ")
WelcomeLength& = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength& + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User$

End Function


Sub AOLIMOff()
Call AOLInstantMessage("$IM_OFF", "Tools Of Khan Beta 6")

End Sub

Sub AOLIMsOn()
Call AOLInstantMessage("$IM_ON", "Tools Of Khan Beta 6")

End Sub


Sub AOLChatSend(txt)
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
'A1000% = FindChildByClass(Room%, "_AOL_Edit")
'A2000% = GetWindow(A1000%, 2)
'AOLIcon (A2000%)
End Sub


Sub AOLClose(winew)
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub

Sub AOLCursor()
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
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
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
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
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
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
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

theview$ = trimspace$
AOLGetChat = theview$
End Function

Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

AOLGetText = trimspace$
End Function

Sub AOLIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLInstantMessage(person, message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call AOLSetText(aoledit%, person)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(im%, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next sends

AOLIcon (imsend%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
End Sub

Function AOLIsOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
If welcome% = 0 Then
MsgBox "Please sign on before using this feature.", 64, "Come on Back in Here!"
End
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
PlayWav (Form1.Info.TEXT)
Let Form1.ToDo.TEXT = ""
End Function

Sub AOLKeyword(TEXT)
Call RunMenuByString(AOLWindow(), "Keyword...")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
keyw% = FindChildByTitle(mdi%, "Keyword")
kedit% = FindChildByClass(keyw%, "_AOL_Edit")
If kedit% Then Exit Do
Loop

editsend% = SendMessageByString(kedit%, WM_SETTEXT, 0, TEXT)
pausing = DoEvents()
Sending% = SendMessage(kedit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Function AOLLastChatLine()
getpar% = AOLFindRoom()
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

theview$ = trimspace$


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
AOLLastChatLine = LastLine
End Function


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
CheckOnline
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu(2, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(AOL%, "Sign Off?")
If pfc% <> 0 Then
Icon1% = FindChildByClass(pfc%, "_AOL_Icon")
Icon1% = GetWindow(Icon1%, 2)
Icon1% = GetWindow(Icon1%, 2)
Icon1% = GetWindow(Icon1%, 2)
Icon1% = GetWindow(Icon1%, 2)
Icon1% = GetWindow(Icon1%, 2)
clickicon% = SendMessage(Icon1%, WM_LBUTTONDOWN, 0, 0&)
clickicon% = SendMessage(Icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Function AOLVersion()
AOL% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(AOL%)

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
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function



Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
Hwndtitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, Hwndtitle$, (hwndLength% + 1))
GetCaption = Hwndtitle$
End Function
Function GetLast(ByVal txt As String)
Dim x As Integer
Do
x = x + 1
Loop Until Mid(txt, Len(txt) - x, 1) = Chr(13)
GetLast = Right(txt, x)
End Function
Sub getlisttext(Lst As ListBox)
On Error Resume Next
Do Until xx = Lst.ListCount
Let da_new$ = Lst.List(xx)
If Trim(da_new$) <> "" Then Let namms = namms & da_new$ & "," Else Exit Sub
Let xx = xx + 1
Loop

End Sub

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Function GETUSER()

'Find the welcome window
wlcm% = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "Welcome, ")

'Get the caption of it
dacap = WindowCaption(wlcm%)

'Extract the user's screen name
If wlcm% <> 0 Then
    numba% = (InStr(dacap, "!") - 10)
    Pname$ = Mid$(dacap, 10, numba%)
    GETUSER = Pname$
Else
    GETUSER = "(unknown)"
End If

End Function
Function GetWinText(hWnd As Integer) As String
lentos = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
x = SendMessageByString(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
GetWinText = Buffer$
End Function
Function GetWinVer1(x)

'Example: text2.text = "Window version: " & sGetWinVer
Dim lVer As Long, iWinVer As Integer
    lVer = GetVersion()
    iWinVer = CInt(lVer And &HFFFF&)
    x = Format$(iWinVer And &HFF) + "." + Format$(CInt(iWinVer / 256))

End Function
Sub guidetos(who$, phraZes$)
run "Keyword"
AOL = FindWindow("AOL Frame25", 0&)
Do
keyword = FindChildByTitle(AOL, "Keyword")
Timeout (0.001)
editbox = FindChildByClass(keyword, "_AOL_Edit")
Loop Until editbox <> 0
editbox = FindChildByClass(keyword, "_AOL_Edit")
SEND editbox, "Guidepager"
GO% = FindChildByClass(keyword, "_AOL_Icon")
Click GO%
Timeout 3
AOL = FindWindow("AOL Frame25", 0&)
mow = FindChildByTitle(AOL, "I Need Help!")
ikqn% = FindChildByClass(mow, "_AOL_Icon")
DoEvents
Click ikqn%
Timeout (2)
Timeout (0.001)

PW = FindChildByTitle(AOL, "Report Password Solicitations")
Timeout (0.001)
editbox5 = FindChildByClass(PW, "_AOL_Edit")
editbox5 = FindChildByClass(PW, "_AOL_Edit")
SEND editbox5, "" + (who$)


End Sub
Sub hidewelcome()
Dim x
wlcm% = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "Welcome, ")
x = ShowWindow(wlcm%, SW_HIDE)
End Sub
Sub hidewin(windo)
x = ShowWindow(windo, SW_HIDE)
End Sub
Function IFileExists(ByVal sFileName As String) As Integer
'Example: If Not IFileExists("win.com") then...
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IFileExists = False
        Else
            IFileExists = True
    End If

End Function
Sub im_off()
AOLInstantMessage "$IM_OFF", "«•œ–±±±Jump in The Fire±±±–œ•»"
WaitForOk
End Sub
Sub IM_ON()
AOLInstantMessage "$IM_ON", "«•œ–Devil Dance By /\/ i \/\–œ•»"
WaitForOk
End Sub
Sub imansweringmachine(wuttosay As TextBox)
'this code goes in a timer
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Begin
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Begin
Exit Sub
Begin:
e = FindChildByClass(im%, "RICHCNTL")

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
Call AOLSetText(e2, ((wuttosay) & Chr(13) & "<//=-——• FATE Infinity •——-=\\>"))
Click (e)
Timeout 0.1
killwin (im%)

End Sub
Sub imbomb(numofbom As TextBox, whotobom As TextBox, wuttosay As TextBox)
If numofbom = "0" Then
Exit Sub
End If
If numofbom < 0 Then
Exit Sub
End If
Call im_off
Do
Call AOLInstantMessage((whotobom), (wuttosay) & Chr(13) & "         «Devil Dance¹ IM Bomber»" & Chr(13) & "                   «By /\/ i \/\»")
numofbom = Str(Val(numofbom - 1))
Loop Until numofbom = 0
Call IM_ON
Exit Sub
End Sub
Function IMTOS(who$)
Dim wer As Variant
If UCase$(who$) Like UCase$("OxONiNOxO") Then
MsgBox "LoL Nice Try ! Say goodbye to your main index !", 55, "Cant TOS /\/ I \/\!!! "
Exit Function
End If

sendtext "BLAH TOSER BY /\/ I \/\"

Do
Dim HWNDS() As Integer
kw% = FindChildByTitle(AOL%, "Keyword")
Timeout 0.2
Loop Until kw% <> 0
edi% = FindChildByClass(kw%, "_AOL_Edit")
Tos = SendMessageByString(edi%, WM_SETTEXT, 0, "KO Help")

AOL% = FindWindow("AOL Frame25", 0&)
Do
DoEvents
XTY% = FindChildByTitle(AOL%, "I Need Help!")
Loop Until XTY% <> 0
R = FindWindowLike(HWNDS(), find%, "*", "_AOL_Icon", Null)
im = HWNDS(3)
x% = SendMessageByNum(im, WM_LBUTTONDOWN, 0, 0&)
x% = SendMessageByNum(im, WM_LBUTTONUP, 0, 0&)




Do
DoEvents
XTY% = FindChildByTitle(AOL%, "Report A Violation")
Loop Until XTY% <> 0
R = FindWindowLike(HWNDS(), find%, "*", "_AOL_Icon", Null)
imv = HWNDS(2)
x% = SendMessageByNum(imv, WM_LBUTTONDOWN, 0, 0&)
x% = SendMessageByNum(imv, WM_LBUTTONUP, 0, 0&)

Do
DoEvents
XP% = FindChildByTitle(AOL%, "Violations via Instant Messages")
Loop Until XP% <> 0

asnm% = FindChildByTitle(XP%, "Date")
 
anx% = GetWindow(asnm%, GW_HWNDNEXT)
asm% = SendMessageByString(anx%, WM_SETTEXT, 0, "" & Date)

esnm% = FindChildByTitle(XP%, "Category of Chat Room")
 
enx% = GetWindow(esnm%, GW_HWNDNEXT)
esm% = SendMessageByString(enx%, WM_SETTEXT, 0, "Life")


qsnm% = FindChildByTitle(XP%, "Time AM/PM")
 
qnx% = GetWindow(qsnm%, GW_HWNDNEXT)
qsm% = SendMessageByString(qnx%, WM_SETTEXT, 0, "" & Time)

wsnm% = FindChildByTitle(XP%, "Chat Room Name")
 
wnx% = GetWindow(wsnm%, GW_HWNDNEXT)
wsm% = SendMessageByString(wnx%, WM_SETTEXT, 0, "Teen Chat5")






snm% = FindChildByTitle(XP%, "CUT and PASTE a copy of the IM here")
 
nx% = GetWindow(snm%, GW_HWNDNEXT)
sm% = SendMessageByString(nx%, WM_SETTEXT, 0, who$ + ":" & " " + text2 & Chr(10))
Sen% = FindChildByTitle(XP%, "Send")
Timeout 0.1
STAtc% = FindChildByTitle(soltos%, "SEND")
WaitForOk
 
ax2 = SendMessage(XP%, WM_CLOSE, 0, 0)
SendKeys "{f4}"
sendtext "BLAH IM Violation TOS complete"
Timeout (0.1)
sendtext "BLAH BY /\/ I \/\"
Exit Function


End Function
Sub invitationbomb(towho As TextBox, numberof As TextBox)
Do
a = FindSN()
eek% = FindChildByTitle(mdi%, "Invitation From: " & a)
If eek% <> 0 Then
killwin eek%
End If
If numberof < 1 Then Exit Sub
'If command3d1.Enabled = True Then Exit Sub
AOL% = FindWindow("aol Frame25", 0&)
mdi% = FindChildByClass(AOL%, "MDIClient")
child% = FindChildByClass(mdi%, "AOL Child")
titl% = FindChildByTitle(mdi%, "Buddy Lists")
e = FindChildByClass(titl%, "_AOL_Icon")
e2 = GetNextWindow(e, 2)
e3% = GetNextWindow(e2, 2)
e4% = GetNextWindow(e3%, 2)
Click e4%
Timeout 0.1
Do: DoEvents
buddy% = FindChildByClass(mdi%, "AOL Child")
buddy2% = FindChildByTitle(mdi%, "Buddy Chat")
edt% = FindChildByClass(buddy2%, "_AOL_Edit")
edt2% = GetNextWindow(edt%, 2)
edt3% = GetNextWindow(edt2%, 2)
edt4% = GetNextWindow(edt3%, 2)
edt5% = GetNextWindow(edt4%, 2)
edt6% = GetNextWindow(edt5%, 2)
edt7% = GetNextWindow(edt6%, 2)
edt8% = GetNextWindow(edt7%, 2)
edt9% = GetNextWindow(edt8%, 2)
Loop Until buddy2% <> 0
Timeout 0.5
textset edt%, CStr((towho))
textset edt3%, CStr("INVITATION BOMB Your invited to HELL")
textset edt9%, CStr("°º*¤FATE Infinity¤*º° by GenOziDe")
icn% = FindChildByClass(buddy2%, "_AOL_Icon")
Click icn%
Timeout 1
a = FindSN()
eek% = FindChildByTitle(mdi%, "Invitation From: " & a)
killwin eek%
Timeout 0.1
numberof = Str(Val(numberof - 1))
If numberof = Str(Val("0")) Then Exit Sub
Loop

End Sub
Sub keywor(where$)
run "Keyword"
Do
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(AOL, "Keyword")
daedit = FindChildByClass(bah, "_AOL_Edit")
Timeout (0.001)
Loop Until daedit <> 0
daedit = FindChildByClass(bah, "_AOL_Edit")
SEND daedit, where$
ico% = FindChildByClass(bah, "_AOL_Icon")
Click ico%
End Sub
Sub KILLMODAL()
Do
AOM = FindWindow("_AOL_Modal", 0&)
killwin (AOM)
DoEvents
Loop Until AOM = 0

End Sub
Sub killwin(windo)
x = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub

Sub GetRoom(Lst As ListBox)
AOL% = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL%, "MDIClient")
List% = FindChildByClass(AOL%, "_AOL_Listbox")
room% = GetParent(List%)
For Index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo end_add
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
ListCheck namez$, Lst
Next Index%
end_add:

End Sub
Function GetSNFromIM()
Dim mdi%
mdi% = FindWindow("AOL Frame25", 0&)
Dim welcome%
welcome% = FindChildByTitle(mdi%, ">Instant Message From: ")
Dim yourname As String * 255
Dim YourName2$
x = GetWindowText(welcome%, yourname, 255)
yourname = LTrim(RTrim(Trim(yourname)))

yourname = Right$(yourname, Len(yourname) - InStr(yourname, ": "))
YourName2 = Left$(yourname, 24)
YourName2 = Right$(yourname, Len(yourname) - 1)
If InStr(YourName2, "") <> 0 Then YourName2 = Left$(YourName2, InStr(YourName2, ""))
SName$ = YourName2$
GetSNFromIM = SName$

End Function

Function GetWindowDir()
Buffer$ = String$(255, 0)
x = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function
Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub phishgen(txt As TextBox, txt2 As TextBox)
Randomize
a = Int((Val("26") * Rnd) + 1)
If a = "1" Then
a = ""
ElseIf a = "2" Then a = "B"
ElseIf a = "3" Then a = "C"
ElseIf a = "4" Then a = "D"
ElseIf a = "5" Then a = ""
ElseIf a = "6" Then a = "F"
ElseIf a = "7" Then a = "G"
ElseIf a = "8" Then a = "H"
ElseIf a = "9" Then a = ""
ElseIf a = "10" Then a = "J"
ElseIf a = "11" Then a = "K"
ElseIf a = "12" Then a = "L"
ElseIf a = "13" Then a = "M"
ElseIf a = "14" Then a = "N"
ElseIf a = "15" Then a = ""
ElseIf a = "16" Then a = "P"
ElseIf a = "17" Then a = "Q"
ElseIf a = "18" Then a = "R"
ElseIf a = "19" Then a = "S"
ElseIf a = "20" Then a = "T"
ElseIf a = "21" Then a = ""
ElseIf a = "22" Then a = "V"
ElseIf a = "23" Then a = "W"
ElseIf a = "24" Then a = "Y"
ElseIf a = "25" Then a = "X"
ElseIf a = "26" Then a = "Z"
End If
txt = a

Randomize
b = Int((Val("37") * Rnd) + 1)
If b = "1" Then
b = "A"
ElseIf b = "2" Then b = "B"
ElseIf b = "3" Then b = "C"
ElseIf b = "4" Then b = "D"
ElseIf b = "5" Then b = "E"
ElseIf b = "6" Then b = "F"
ElseIf b = "7" Then b = "G"
ElseIf b = "8" Then b = "H"
ElseIf b = "9" Then b = "I"
ElseIf b = "10" Then b = "J"
ElseIf b = "11" Then b = "K"
ElseIf b = "12" Then b = "L"
ElseIf b = "13" Then b = "M"
ElseIf b = "14" Then b = "N"
ElseIf b = "15" Then b = "O"
ElseIf b = "16" Then b = "P"
ElseIf b = "17" Then b = "Q"
ElseIf b = "18" Then b = "R"
ElseIf b = "19" Then b = "S"
ElseIf b = "20" Then b = "T"
ElseIf b = "21" Then b = "U"
ElseIf b = "22" Then b = "V"
ElseIf b = "23" Then b = "W"
ElseIf b = "24" Then b = "Y"
ElseIf b = "25" Then b = "X"
ElseIf b = "26" Then b = "Z"
ElseIf b = "27" Then b = "0"
ElseIf b = "28" Then b = "1"
ElseIf b = "29" Then b = "2"
ElseIf b = "30" Then b = "3"
ElseIf b = "31" Then b = "4"
ElseIf b = "32" Then b = "5"
ElseIf b = "33" Then b = "6"
ElseIf b = "34" Then b = "7"
ElseIf b = "35" Then b = "8"
ElseIf b = "36" Then b = "9"
ElseIf b = "37" Then b = " "
End If
txt = a + b

Randomize
c = Int((Val("6") * Rnd) + 1)
If c = "1" Then
c = "A"
ElseIf c = "2" Then c = "E"
ElseIf c = "3" Then c = "I"
ElseIf c = "4" Then c = "O"
ElseIf c = "5" Then c = "U"
ElseIf c = "6" Then c = " "
End If
txt = a + b + c

Randomize
d = Int((Val("37") * Rnd) + 1)
If d = "1" Then
d = " "
ElseIf d = "2" Then d = "B"
ElseIf d = "3" Then d = "C"
ElseIf d = "4" Then d = "D"
ElseIf d = "5" Then d = " "
ElseIf d = "6" Then d = "F"
ElseIf d = "7" Then d = "G"
ElseIf d = "8" Then d = "H"
ElseIf d = "9" Then d = " "
ElseIf d = "10" Then d = "J"
ElseIf d = "11" Then d = "K"
ElseIf d = "12" Then d = "L"
ElseIf d = "13" Then d = "M"
ElseIf d = "14" Then d = "N"
ElseIf d = "15" Then d = " "
ElseIf d = "16" Then d = "P"
ElseIf d = "17" Then d = "Q"
ElseIf d = "18" Then d = "R"
ElseIf d = "19" Then d = "S"
ElseIf d = "20" Then d = "T"
ElseIf d = "21" Then d = " "
ElseIf d = "22" Then d = "V"
ElseIf d = "23" Then d = "W"
ElseIf d = "24" Then d = "Y"
ElseIf d = "25" Then d = "X"
ElseIf d = "26" Then d = "Z"
ElseIf d = "27" Then d = "0"
ElseIf d = "28" Then d = "1"
ElseIf d = "29" Then d = "2"
ElseIf d = "30" Then d = "3"
ElseIf d = "31" Then d = "4"
ElseIf d = "32" Then d = "5"
ElseIf d = "33" Then d = "6"
ElseIf d = "34" Then d = "7"
ElseIf d = "35" Then d = "8"
ElseIf d = "36" Then d = "9"
ElseIf d = "37" Then d = " "

End If
txt = a + b + c + d

Randomize
e = Int((Val("6") * Rnd) + 1)
If e = "1" Then
e = "A"
ElseIf e = "2" Then e = "E"
ElseIf e = "3" Then e = "I"
ElseIf e = "4" Then e = "O"
ElseIf e = "5" Then e = "U"
ElseIf e = "6" Then e = " "
End If
txt = a + b + c + d + e

Randomize
f = Int((Val("37") * Rnd) + 1)
If f = "1" Then
f = ""
ElseIf f = "2" Then f = "B"
ElseIf f = "3" Then f = "C"
ElseIf f = "4" Then f = "D"
ElseIf f = "5" Then f = ""
ElseIf f = "6" Then f = "F"
ElseIf f = "7" Then f = "G"
ElseIf f = "8" Then f = "H"
ElseIf f = "9" Then f = ""
ElseIf f = "10" Then f = "J"
ElseIf f = "11" Then f = "K"
ElseIf f = "12" Then f = "L"
ElseIf f = "13" Then f = "M"
ElseIf f = "14" Then f = "N"
ElseIf f = "15" Then f = ""
ElseIf f = "16" Then f = "P"
ElseIf f = "17" Then f = "Q"
ElseIf f = "18" Then f = "R"
ElseIf f = "19" Then f = "S"
ElseIf f = "20" Then f = "T"
ElseIf f = "21" Then f = ""
ElseIf f = "22" Then f = "V"
ElseIf f = "23" Then f = "W"
ElseIf f = "24" Then f = "Y"
ElseIf f = "25" Then f = "X"
ElseIf f = "26" Then f = "Z"
ElseIf f = "27" Then f = "0"
ElseIf f = "28" Then f = "1"
ElseIf f = "29" Then f = "2"
ElseIf f = "30" Then f = "3"
ElseIf f = "31" Then f = "4"
ElseIf f = "32" Then f = "5"
ElseIf f = "33" Then f = "6"
ElseIf f = "34" Then f = "7"
ElseIf f = "35" Then f = "8"
ElseIf f = "36" Then f = "9"
ElseIf f = "37" Then f = " "
End If
txt = a + b + c + d + e + f

Randomize
g = Int((Val("6") * Rnd) + 1)
If g = "1" Then
g = "A"
ElseIf g = "2" Then g = "E"
ElseIf g = "3" Then g = "I"
ElseIf g = "4" Then g = "O"
ElseIf g = "5" Then g = "U"
ElseIf g = "6" Then g = " "
End If
txt = a + b + c + d + e + f + g

Randomize
h = Int((Val("37") * Rnd) + 1)
If h = "1" Then
h = ""
ElseIf h = "2" Then h = "B"
ElseIf h = "3" Then h = "C"
ElseIf h = "4" Then h = "D"
ElseIf h = "5" Then h = ""
ElseIf h = "6" Then h = "F"
ElseIf h = "7" Then h = "G"
ElseIf h = "8" Then h = "H"
ElseIf h = "9" Then h = ""
ElseIf h = "10" Then h = "J"
ElseIf h = "11" Then h = "K"
ElseIf h = "12" Then h = "L"
ElseIf h = "13" Then h = "M"
ElseIf h = "14" Then h = "N"
ElseIf h = "15" Then h = ""
ElseIf h = "16" Then h = "P"
ElseIf h = "17" Then h = "Q"
ElseIf h = "18" Then h = "R"
ElseIf h = "19" Then h = "S"
ElseIf h = "20" Then h = "T"
ElseIf h = "21" Then h = ""
ElseIf h = "22" Then h = "V"
ElseIf h = "23" Then h = "W"
ElseIf h = "24" Then h = "Y"
ElseIf h = "25" Then h = "X"
ElseIf h = "26" Then h = "Z"
ElseIf h = "27" Then h = "0"
ElseIf h = "28" Then h = "1"
ElseIf h = "29" Then h = "2"
ElseIf h = "30" Then h = "3"
ElseIf h = "31" Then h = "4"
ElseIf h = "32" Then h = "5"
ElseIf h = "33" Then h = "6"
ElseIf h = "34" Then h = "7"
ElseIf h = "35" Then h = "8"
ElseIf h = "36" Then h = "9"
ElseIf h = "37" Then h = " "
End If
txt = a + b + c + d + e + f + g + h

Randomize
i = Int((Val("6") * Rnd) + 1)
If i = "1" Then
i = "E"
ElseIf i = "2" Then i = "A"
ElseIf i = "3" Then i = "I"
ElseIf i = "4" Then i = "O"
ElseIf i = "5" Then i = "U"
ElseIf i = "6" Then i = " "
End If
txt = a + b + c + d + e + f + g + h + i

Randomize
j = Int((Val("37") * Rnd) + 1)
If j = "1" Then
j = "A"
ElseIf j = "2" Then j = "B"
ElseIf j = "3" Then j = "C"
ElseIf j = "4" Then j = "D"
ElseIf j = "5" Then j = "E"
ElseIf j = "6" Then j = "F"
ElseIf j = "7" Then j = "G"
ElseIf j = "8" Then j = "H"
ElseIf j = "9" Then j = "I"
ElseIf j = "10" Then j = "J"
ElseIf j = "11" Then j = "K"
ElseIf j = "12" Then j = "L"
ElseIf j = "13" Then j = "M"
ElseIf j = "14" Then j = "N"
ElseIf j = "15" Then j = "O"
ElseIf j = "16" Then j = "P"
ElseIf j = "17" Then j = "Q"
ElseIf j = "18" Then j = "R"
ElseIf j = "19" Then j = "S"
ElseIf j = "20" Then j = "T"
ElseIf j = "21" Then j = "U"
ElseIf j = "22" Then j = "V"
ElseIf j = "23" Then j = "W"
ElseIf j = "24" Then j = "Y"
ElseIf j = "25" Then j = "X"
ElseIf j = "26" Then j = "Z"
ElseIf j = "27" Then j = "0"
ElseIf j = "28" Then j = "1"
ElseIf j = "29" Then j = "2"
ElseIf j = "30" Then j = "3"
ElseIf j = "31" Then j = "4"
ElseIf j = "32" Then j = "5"
ElseIf j = "33" Then j = "6"
ElseIf j = "34" Then j = "7"
ElseIf j = "35" Then j = "8"
ElseIf j = "36" Then j = "9"
ElseIf j = "37" Then j = " "
End If
txt = a + b + c + d + e + f + g + h + j


Randomize
k = Int((Val("37") * Rnd) + 1)
If k = "1" Then
k = "A"
ElseIf k = "2" Then k = "B"
ElseIf k = "3" Then k = "C"
ElseIf k = "4" Then k = "D"
ElseIf k = "5" Then k = "E"
ElseIf k = "6" Then k = "F"
ElseIf k = "7" Then k = "G"
ElseIf k = "8" Then k = "H"
ElseIf k = "9" Then k = "I"
ElseIf k = "10" Then k = "J"
ElseIf k = "11" Then k = "K"
ElseIf k = "12" Then k = "L"
ElseIf k = "13" Then k = "M"
ElseIf k = "14" Then k = "N"
ElseIf k = "15" Then k = "O"
ElseIf k = "16" Then k = "P"
ElseIf k = "17" Then k = "Q"
ElseIf k = "18" Then k = "R"
ElseIf k = "19" Then k = "S"
ElseIf k = "20" Then k = "T"
ElseIf k = "21" Then k = "U"
ElseIf k = "22" Then k = "V"
ElseIf k = "23" Then k = "W"
ElseIf k = "24" Then k = "Y"
ElseIf k = "25" Then k = "X"
ElseIf k = "26" Then k = "Z"
ElseIf k = "27" Then k = "0"
ElseIf k = "28" Then k = "1"
ElseIf k = "29" Then k = "2"
ElseIf k = "30" Then k = "3"
ElseIf k = "31" Then k = "4"
ElseIf k = "32" Then k = "5"
ElseIf k = "33" Then k = "6"
ElseIf k = "34" Then k = "7"
ElseIf k = "35" Then k = "8"
ElseIf k = "36" Then k = "9"
ElseIf k = "37" Then k = ""
End If
txt2 = k
Randomize
l = Int((Val("37") * Rnd) + 1)
If j = "1" Then
l = "A"
ElseIf l = "2" Then l = "B"
ElseIf l = "3" Then l = "C"
ElseIf l = "4" Then l = "D"
ElseIf l = "5" Then l = "E"
ElseIf l = "6" Then l = "F"
ElseIf l = "7" Then l = "G"
ElseIf l = "8" Then l = "H"
ElseIf l = "9" Then l = "I"
ElseIf l = "10" Then l = "j"
ElseIf l = "11" Then l = "k"
ElseIf l = "12" Then l = "L"
ElseIf l = "13" Then l = "M"
ElseIf l = "14" Then l = "N"
ElseIf l = "15" Then l = "O"
ElseIf l = "16" Then l = "P"
ElseIf l = "17" Then l = "Q"
ElseIf l = "18" Then l = "R"
ElseIf l = "19" Then l = "S"
ElseIf l = "20" Then l = "T"
ElseIf l = "21" Then l = "U"
ElseIf l = "22" Then l = "V"
ElseIf l = "23" Then l = "W"
ElseIf l = "24" Then l = "Y"
ElseIf l = "25" Then l = "X"
ElseIf l = "26" Then l = "Z"
ElseIf l = "27" Then l = "0"
ElseIf l = "28" Then l = "1"
ElseIf l = "29" Then l = "2"
ElseIf l = "30" Then l = "3"
ElseIf l = "31" Then l = "4"
ElseIf l = "32" Then l = "5"
ElseIf l = "33" Then l = "6"
ElseIf l = "34" Then l = "7"
ElseIf l = "35" Then l = "8"
ElseIf l = "36" Then l = "9"
ElseIf l = "37" Then l = ""
End If
txt2 = k + l
Randomize
m = Int((Val("37") * Rnd) + 1)
If j = "1" Then
m = "A"
ElseIf m = "2" Then m = "B"
ElseIf m = "3" Then m = "C"
ElseIf m = "4" Then m = "D"
ElseIf m = "5" Then m = "E"
ElseIf m = "6" Then m = "F"
ElseIf m = "7" Then m = "G"
ElseIf m = "8" Then m = "H"
ElseIf m = "9" Then m = "I"
ElseIf m = "10" Then m = "k"
ElseIf m = "11" Then m = "l"
ElseIf m = "12" Then m = "m"
ElseIf m = "13" Then m = "M"
ElseIf m = "14" Then m = "N"
ElseIf m = "15" Then m = "O"
ElseIf m = "16" Then m = "P"
ElseIf m = "17" Then m = "Q"
ElseIf m = "18" Then m = "R"
ElseIf m = "19" Then m = "S"
ElseIf m = "20" Then m = "T"
ElseIf m = "21" Then m = "U"
ElseIf m = "22" Then m = "V"
ElseIf m = "23" Then m = "W"
ElseIf m = "24" Then m = "Y"
ElseIf m = "25" Then m = "X"
ElseIf m = "26" Then m = "Z"
ElseIf m = "27" Then m = "0"
ElseIf m = "28" Then m = "1"
ElseIf m = "29" Then m = "2"
ElseIf m = "30" Then m = "3"
ElseIf m = "31" Then m = "4"
ElseIf m = "32" Then m = "5"
ElseIf m = "33" Then m = "6"
ElseIf m = "34" Then m = "7"
ElseIf m = "35" Then m = "8"
ElseIf m = "36" Then m = "9"
ElseIf m = "37" Then m = ""
End If
txt2 = k + l + m

Randomize
n = Int((Val("37") * Rnd) + 1)
If j = "1" Then
n = "A"
ElseIf n = "2" Then n = "B"
ElseIf n = "3" Then n = "C"
ElseIf n = "4" Then n = "D"
ElseIf n = "5" Then n = "E"
ElseIf n = "6" Then n = "F"
ElseIf n = "7" Then n = "G"
ElseIf n = "8" Then n = "H"
ElseIf n = "9" Then n = "I"
ElseIf n = "10" Then n = "k"
ElseIf n = "11" Then n = "l"
ElseIf n = "12" Then n = "m"
ElseIf n = "13" Then n = "n"
ElseIf n = "14" Then n = "N"
ElseIf n = "15" Then n = "O"
ElseIf n = "16" Then n = "P"
ElseIf n = "17" Then n = "Q"
ElseIf n = "18" Then n = "R"
ElseIf n = "19" Then n = "S"
ElseIf n = "20" Then n = "T"
ElseIf n = "21" Then n = "U"
ElseIf n = "22" Then n = "V"
ElseIf n = "23" Then n = "W"
ElseIf n = "24" Then n = "Y"
ElseIf n = "25" Then n = "X"
ElseIf n = "26" Then n = "Z"
ElseIf n = "27" Then n = "0"
ElseIf n = "28" Then n = "1"
ElseIf n = "29" Then n = "2"
ElseIf n = "30" Then n = "3"
ElseIf n = "31" Then n = "4"
ElseIf n = "32" Then n = "5"
ElseIf n = "33" Then n = "6"
ElseIf n = "34" Then n = "7"
ElseIf n = "35" Then n = "8"
ElseIf n = "36" Then n = "9"
ElseIf n = "37" Then n = ""
End If
txt2 = k + l + m + n

Randomize
o = Int((Val("37") * Rnd) + 1)
If j = "1" Then
o = "A"
ElseIf o = "2" Then o = "B"
ElseIf o = "3" Then o = "C"
ElseIf o = "4" Then o = "D"
ElseIf o = "5" Then o = "E"
ElseIf o = "6" Then o = "F"
ElseIf o = "7" Then o = "G"
ElseIf o = "8" Then o = "H"
ElseIf o = "9" Then o = "I"
ElseIf o = "10" Then o = "k"
ElseIf o = "11" Then o = "l"
ElseIf o = "12" Then o = "m"
ElseIf o = "13" Then o = "M"
ElseIf o = "14" Then o = "N"
ElseIf o = "15" Then o = "O"
ElseIf o = "16" Then o = "P"
ElseIf o = "17" Then o = "Q"
ElseIf o = "18" Then o = "R"
ElseIf o = "19" Then o = "S"
ElseIf o = "20" Then o = "T"
ElseIf o = "21" Then o = "U"
ElseIf o = "22" Then o = "V"
ElseIf o = "23" Then o = "W"
ElseIf o = "24" Then o = "Y"
ElseIf o = "25" Then o = "X"
ElseIf o = "26" Then o = "Z"
ElseIf o = "27" Then o = "0"
ElseIf o = "28" Then o = "1"
ElseIf o = "29" Then o = "2"
ElseIf o = "30" Then o = "3"
ElseIf o = "31" Then o = "4"
ElseIf o = "32" Then o = "5"
ElseIf o = "33" Then o = "6"
ElseIf o = "34" Then o = "7"
ElseIf o = "35" Then o = "8"
ElseIf o = "36" Then o = "9"
ElseIf o = "37" Then o = ""
End If
txt2 = k + l + m + n + o

End Sub

Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Sub preferences()
Call runmenua(4, 2)
x = waitforwindow("Preferences", "AOL Child")
d% = FindChildByClass(x, "_AOL_Icon")
Click d%
Y = waitforwindow("Chat Preferences", "_AOL_Modal")
b% = FindChildByTitle(Y, "Cancel")
Click b%
killwin x
End Sub

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

Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu2 = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub


Sub UnHideWindow(hWnd)
un = ShowWindow(hWnd, SW_SHOW)
End Sub



Sub WaitForOk()
Do: DoEvents
AOL% = FindWindow("#32770", "America Online")

If AOL% Then
closeaol% = SendMessage(AOL%, WM_CLOSE, 0, 0)
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
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
topmdi% = GetWindow(mdi%, 5)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
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

Function BS_Antipunt()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    Do
    im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    killwin (im%)
    Loop Until im% = 0



End Function



Sub getnum(og%, a)
Do
If a = 0 Then Exit Sub
b = 1 + b
og% = GetWindow(og%, GW_HWNDNEXT)
Loop Until b >= a - 1
End Sub
Sub SEND(chatedit, sill$)
sndtext = SendMessageByString(chatedit, WM_SETTEXT, 0, sill$)
End Sub



Function CheckOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
If welcome% = 0 Then
MsgBox "Please sign on before using this feature.", 64, "Come on Back in Here!"
End
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Sub AutoGreetPref()
a% = FindWindow("AOL Frame25", 0&)
ssa% = AOLVersion()
If ssa% = 25 Then wh$ = "Set Preferences" + Chr$(9) + "Ctrl+="
If ssa% = 3 Then wh$ = "Preferences"
b% = runmenu2("AOL Frame25", "Mem&bers", wh$)
DoEvents
Do
DoEvents
c% = FindChildByTitle(a%, "Preferences")
DoEvents
Loop Until c% <> 0
DoEvents
If ssa% = 25 Then d% = getaolwinb(c%, "_AOL_Icon", 3)
If ssa% = 3 Then d% = getaolwinb(c%, "_AOL_ICON", 1)
DoEvents
AOLClick (d%)
DoEvents
Do
DoEvents
a% = FindWindow("AOL Frame25", 0&)
e% = FindWindow("_AOL_Modal", "Chat Preferences")
DoEvents
Loop Until e% <> 0
DoEvents
f% = FindChildByTitle(e%, "Notify me when members arrive")
DoEvents
DoEvents
h% = SendMessageByNum(f%, BM_SETCHECK, True, 0&)
DoEvents
j% = FindChildByTitle(e%, "OK")

DoEvents
AOLClick (j%)
DoEvents
k% = SendMessageByNum(c%, WM_CLOSE, 0, 0&)
End Sub

Function bak(ward$)
Do
heh$ = Right$(ward$, 1)
mog$ = mog$ & heh$
bad$ = Left$(ward$, Len(ward$) - 1)
ward$ = bad$
DoEvents
Loop Until Len(ward$) = 0
bak = mog$

End Function
Sub bustinpriv(room$)
electric1:
If proG_STAT$ = "OFF" Then
Exit Sub
End If

Let num_try = num_try + 1
a = FindWindow("AOL Frame25", 0&)

Do
DoEvents
fcr = FindChildByTitle(a, "Keyword")
DoEvents
Loop Until fcr <> 0

E1 = FindChildByClass(fcr, "_AOL_Icon")
i1 = FindChildByClass(fcr, "_AOL_Edit")
send_to = SendMessageByString(i1, WM_SETTEXT, 0, "aol://2719:2-2-" & room$)
tta = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
ttr = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
WaitForOk
If OK_note = "NO" Then
Exit Sub
End If
Timeout 0.2
GoTo electric1

End Sub

Sub changewav(wav As String)
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub
Function char(What$, a)
u$ = Mid(What$, Val(a))
char = Left(u$, 1)
End Function

Function ChatRoomName()
ChatRoomName = WindowCaption(FindChatWnd())

End Function
Function ChatText()
Do
AOL = FindWindow("AOL Frame25", 0&)
view% = FindChildByClass(AOL, "_AOL_View")
bag = GetWinText(view%)
Do
AOL = FindWindow("AOL Frame25", 0&)
view% = FindChildByClass(AOL, "_AOL_View")
bad = GetWinText(view%)
DoEvents
Loop Until bad <> bag
ChatText = Mid(bad, InStr(bad, bag) + Len(bag) - 1)
DoEvents
Loop

End Function

Sub checkifalive(who$)
Do
AOL% = FindWindow("AOL Frame25", 0&)
heh = FindChildByTitle(AOL%, "Compose Mail")
SEND heh, "Compose Mail"
DoEvents
Loop Until heh = 0
AOL% = FindWindow("AOL Frame25", 0&)
run "Compose"
Do
x = DoEvents()
chatlist% = FindChildByTitle(AOL%, "Compose Mail")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
uh = ShowWindow(chatlist%, SW_HIDE)
Loop Until chatedit% <> 0
chatwin% = GetParent(chatlist%)
button% = FindChildByClass(chatlist%, "_AOL_Icon")
ChatVew% = FindChildByClass(chatlist%, "_AOL_View")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, who$ & ", '")
blah% = GetWindow(chatedit%, GW_HWNDNEXT)
good% = GetWindow(blah%, GW_HWNDNEXT)
bad% = GetWindow(good%, GW_HWNDNEXT)
Sad% = GetWindow(bad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Sad%, WM_SETTEXT, 0, "TOS check")
Nad% = GetWindow(Sad%, GW_HWNDNEXT)
Wad% = GetWindow(Nad%, GW_HWNDNEXT)
Dad% = GetWindow(Wad%, GW_HWNDNEXT)
fad% = GetWindow(Dad%, GW_HWNDNEXT)
Qad% = GetWindow(fad%, GW_HWNDNEXT)
Ead% = GetWindow(Qad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Dad%, WM_SETTEXT, 0, "TOS Check")
rich = FindChildByClass(chatlist%, "RICHCNTL")
sndtext% = SendMessageByString(rich, WM_SETTEXT, 0, "TOS Check")
Timeout (0.1)
SendNow% = SendMessageByNum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(button%, WM_LBUTTONUP, &HD, 0)
Timeout (0.1)
Do
eroro = FindChildByTitle(AOL%, "Error")
Timeout (0.001)
Loop Until eroro <> 0
winwo% = FindChildByClass(eroro, "_AOL_View")
teee = GetWinText(winwo%)

If InStr(LCase$(teee), LCase$(who$)) Then
checkifalive1 = 1
End If
If InStr(LCase$(teee), LCase$("mailbox")) Then
checkifalive1 = 0
End If
killwin (eroro)
killwin (chatlist%)
Do
AOL% = FindWindow("AOL Frame25", 0&)
heh = FindChildByTitle(AOL%, "Compose Mail")
SEND heh, "Compose Mail"
DoEvents
Loop Until heh = 0

End Sub

Sub clearchat()
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & "" & Chr(9) & Chr$(13) & Chr$(13))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)

End Sub
