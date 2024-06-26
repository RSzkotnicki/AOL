Attribute VB_Name = "CPU"
' ‹‹››-C.P.U.
' ‹‹››-This bas contains subs and functions
' ‹‹››-from different members of the grewp
' ‹‹››-and also from various other bas files.
' ‹‹››-None of us at C.P.U. take full credit for
' ‹‹››-this bas.  Many of us took part in the
' ‹‹››-creation of this bas and therefore can
' ‹‹››-not take full credit for its contents.

Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const SPI_SCREENSAVERRUNNING = 97
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1

Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function Sendmessege Lib "user32" Alias "SendMessegeA" (ByValwMsg As Long, ByVal wParam As Long, Param As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
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
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
'****************************************************************
'Windows API/Global Declarations for :Rip a cd
'****************************************************************


Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2


Private Declare Function PutFocus Lib "user32" Alias "SetFocus" _
       (ByVal hWnd As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
       (ByVal hWnd As Long, _
       ByVal wMsg As Long, _
       ByVal wParam As Integer, _
       ByVal lParam As Long) As Long
       Private Const EM_LINESCROLL = &HB6
'****************************************************************
'Windows API/Global Declarations for :PrintText
'****************************************************************
Dim NumLinesOnPageToPrint As Integer
Dim Length_ChrsInlineOfText As Integer
Dim FirstPageNum As Integer
Dim NextPageNum As Integer
Dim LineNum As Integer
Dim MarginSize As Integer
Dim CheckThisLineNum As Integer
Dim NumLines As Integer
Dim TotalPageCount As Integer

Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230


Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
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



Const EM_UNDO = &HC7
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

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

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

Private Type MailInfo
    MailDate As Date
    sn As String * 10
    Title As String * 53
End Type
Private StopNow, LoadMail As Boolean
Dim RetStr As String * 71
Dim Mails() As MailInfo

Sub AOL_CloseInvalid()
' closes "invalid" window
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then
AOLClose (aolcl%)
Pause (1)
End If
End Sub


Sub Scorekeeper(txtright As TextBox, txtnameswitch As TextBox, txtorig As TextBox, txtlast As TextBox, tmr As Timer, lstscore As ListBox, lblpoint As Label)
If UCase(txtright.Text) = UCase(txtorig.Text) Then
tmr.Enabled = False
txtorig.Locked = False
txtlast.Text = SNFromLastChatLine
         If lstscore.ListCount <> 0 Then
         For x9x = 0 To lstscore.ListCount - 1
         If InStr(UCase(lstscore.List(x9x)), UCase(txtlast.Text)) Then
         SendChat ("(—·•°|[ " & txtlast.Text & " Got It Correct [ " & lblpoint.Caption & " ]")
         txtnameswitch.Text = lstscore.List(x9x)
         l08 = InStr((txtnameswitch.Text), " -")
         l09 = Mid$((txtnameswitch.Text), l08, 2)
         l10 = Left((txtnameswitch.Text), Len(l09) + 1)
         x1x = Trim(l10)
         x2x = Val(x1x) + lblpoint.Caption
         txtnameswitch.Text = x2x & " - " & txtlast.Text
         lx9 = Val(x9x)
         lstscore.List(lx9) = txtnameswitch.Text
         Exit Sub
         End If
         Next x9x
         SendChat ("(—·•°|[ " & txtlast.Text & " Got It Correct [ " & lblpoint.Caption & " ]")
         lstscore.AddItem lblpoint.Caption & " - " & txtlast.Text: Exit Sub
         Else
         SendChat ("(—·•°|[ " & txtlast.Text & " Got It Correct [ " & lblpoint.Caption & " ]")
         lstscore.AddItem lblpoint.Caption & " - " & txtlast.Text: Exit Sub
         End If
         Else
         Exit Sub
         End If
End Sub


Sub SentenceLink(First As String, Link As String, Addy As String, CCL As String, CCL2 As String, Second As String, Underlined As Boolean)
'First = first part of sentence (before the link)
'Link = what link should say in chat room
'Addy = web address (URL)
'CCL = color of link
'CCL2 = color of link after it is clicked
'Second = second part of sentence (after the link)
'Underlined = link is underlined (True) or is not underlined (False)
'Example:
'Call SentenceLink("Test ", "Test", "http://www.test.com", "ff0000", "fffffe", " Test", True)

If Underlined = False Then SendChat "<body vlink=#" + CCL2 + "><font color=#000000>" + First + "</a><a href=""""><a href=""" + Addy + """><font color=#" + CCL + "></u>" + Link + "</html>" + Second + "<font color=#fffffe></a>"
If Underlined = True Then SendChat "<body vlink=#" + CCL2 + "><font color=#000000>" + First + "</a><a href=""""><a href=""" + Addy + """><font color=#" + CCL + ">" + Link + "</html>" + Second + "<font color=#fffffe></a>"
End Sub

Sub sendok()
View% = FindChildByClass(AOLFindRoom, "RICHCNTL")
Edito% = GetWindow(View%, 2)
Edit1% = GetWindow(Edito%, 2)
Edit2% = GetWindow(Edit1%, 2)
Edit3% = GetWindow(Edit2%, 2)
Edit4% = GetWindow(Edit3%, 2)
EditBox% = GetWindow(Edit4%, 2)
textsend1% = SendMessageByString(EditBox%, WM_CHAR, 13, 0&)
textsend1% = SendMessageByString(EditBox%, WM_CHAR, 13, 0&)
End Sub
Function AddListToString(thelist As ListBox)
'Makes a list into a string a "comma" after each word
For DoList = 0 To thelist.ListCount - 1
AddListToString = AddListToString & thelist.List(DoList) & "</Html>"
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function

Sub AddStringToList(theitems, thelist As ListBox)
'Adds a string to a list box
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

Public Sub AOL4_Button(but%)
'Clicks an _AOL_Button
Dim ClickIcon%
ClickIcon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
ClickIcon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub
Sub AOL4_ChangeCaption(newcaption As String)
'This changes the "America  Online" Caption
Call AOL4_SetText(AOL4_Window(), newcaption)
End Sub
Sub AOL4_ChatSend(txt)
'Sendz txt to a chat room
    Room% = AOL4_FindRoom()
    If Room% Then
        hChatEdit% = Find2ndChildByClass(Room%, "RICHCNTL")
        ret = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, txt)
        ret = SendMessageByNum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub
Function AOL4_ClearChat()
'This gets all the txt from chat room
childs% = AOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
AOL4_ClearChat = theview$
End Function

Function AOL4_ClickList(hWnd)
ClickList% = SendMessageByNum(hWnd, &H203, 0, 0&)
End Function


Function AOL4_GetChat()
'This gets all the txt from chat room
childs% = AOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
AOL4_GetChat = theview$
End Function

Function AOL4_GetChatSN()
'This getz the last chat line without a SN in front of it
heh$ = (AOL4_LastChatLine)
heh$ = LCase(heh$)
nwe$ = Mid(heh$, InStr(heh$, ":") + 2)
AOL4_GetChatSN = nwe$
End Function

Public Function AOL4_GetList(Index As Long, Buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = AOL4_FindRoom()
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

Function AOL4_GetUser()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOL4_GetUser = user
End Function

Sub AOL4_Hide()
'This hidez AOL
a = ShowWindow(AOL4_Window(), 0)
End Sub

Sub AOL4_Icon(icon%)
'Clickz on an AOL icon
click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOL4_IM_AutoAnswer(message)
'Res'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(IM%, "RICHCNTL")

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
e2 = GetWindow(E, GW_HWNDNEXT) 'Send Text
E = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
AOL4_Icon (E)
Pause 2#

End Sub
Sub AOL4_InstantMessage(Person, message)
'This openz an IM and fillz it out with Person and Message
AOL4_KEYWORD ("aol://9293:" & Person)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
Loop Until (IM% <> 0 And aolrich% <> 0 And imsend% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, message)
For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends
AOL4_Icon imsend%
If IM% Then Call AOL4_KillWin(IM%)
End Sub

Sub AOL4_Invite(Person)
'This will send an Invite to a person
FreeProcess
On Error GoTo ErrHandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
bud% = FindChildByTitle(MDI%, "Buddy List Window")
E = FindChildByClass(bud%, "_AOL_Icon")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
AOL4_Icon (E)
Pause (1#)
Chat% = FindChildByTitle(MDI%, "Buddy Chat")
AOLEdit% = FindChildByClass(Chat%, "_AOL_Edit")
If Chat% Then GoTo FILL
FILL:
Call AOL4_SetText(AOLEdit%, Person)
de = FindChildByClass(Chat%, "_AOL_Icon")
AOL4_Icon (de)
Killit% = FindChildByTitle(MDI%, "Invitation From:")
AOL4_KillWin (Killit%)
FreeProcess
ErrHandler:
Exit Sub
End Sub




Sub AOL4_KEYWORD(txt)
'This goes to an AOL Keyword
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub

Sub AOL4_KillWin(windo)
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_locateMember(name As String)
'This will Locate a member....if online
Call AOL4_KEYWORD("aol://3548:" + name)
End Sub

Sub AOL4_Mail(Person, subject, message)
'This openz a mail and fills it out
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(Tool2%, "_AOL_Icon")
Icon2% = GetWindow(ico3n%, 2)
click% = SendMessageByNum(Icon2%, WM_LBUTTONDOWN, 0&, 0&)
click% = SendMessageByNum(Icon2%, WM_LBUTTONUP, 0&, 0&)
Pause (4)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    mail% = FindChildByTitle(MDI%, "Write Mail")
    AOLEdit% = FindChildByClass(mail%, "_AOL_Edit")
    aolrich% = FindChildByClass(mail%, "RICHCNTL")
    subjt% = FindChildByTitle(mail%, "Subject:")
    subjec% = GetWindow(subjt%, 2)
        Call AOL4_SetText(AOLEdit%, Person)
        Call AOL4_SetText(subjec%, subject)
        Call AOL4_SetText(aolrich%, message)
E = FindChildByClass(mail%, "_AOL_Icon")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
AOL4_Icon (E)
End Sub

Sub AOL4_MassIMer(Person, message)
'This openz an IM and fillz it out with Person and Message
AOL4_KEYWORD ("aol://9293:" & Person)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
Loop Until (IM% <> 0 And aolrich% <> 0 And imsend% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, message)
For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends
AOL4_Icon imsend%

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop

End Sub


Function AOL4_MDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL4_MDI = FindChildByClass(AOL%, "MDIClient")
End Function
Sub AOL4_OpenMail()
'This openz up the Users Mail box
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(Tool2%, "_AOL_Icon")
Icon2% = GetWindow(ico3n%, 0)
click% = SendMessageByNum(Icon2%, WM_LBUTTONDOWN, 0&, 0&)
click% = SendMessageByNum(Icon2%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub AOL4_Read1Mail()
'This will read the very first mail in the User's box
MailBox% = FindChildByTitle(AOL4_MDI(), AOL4_GetUser + "'s Online Mailbox")
E = FindChildByClass(MailBox%, "_AOL_Icon")
AOL4_Icon (E)
End Sub

Function AOL4_RoomCount()
'Countz people in a chat room and returnz it
thechild% = AOL4_FindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")
getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOL4_RoomCount = getcount
End Function

Public Sub AOL4_RoomSNs(Listboxes As ListBox)
'This adds AOL's room list to a VB listbox
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = AOL4_FindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
lst.RemoveItem Listboxes.ListCount - 1
i = GetListIndex(Listboxes, AOL4_GetUser())
If i <> -2 Then Listboxes.RemoveItem i
End Sub

Sub AOL4_RunMenu(menu1 As Integer, menu2 As Integer)
'This will run one of the drop down menu's like  Edit/Paste
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub
Public Sub AOL4_Scroll(Text)
    For i = 1 To 4
Call AOL4_ChatSend(Text + String(116 - Len(Text), Chr(9)) + Text)
Pause (0.01)
Next i
End Sub


Sub AOL4_SetText(win, txt)
'This is usually used for an _AOL_Edit or RICHCNTL
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Function AOL4_SignedON()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
If welcome% = 0 Then
AOL4_SignedON = 0
Exit Function
End If
AOL4_SignedON = 1
End Function

Sub AOL4_signoff()
'This will sign the User off of AOL
AppActivate "America  Online"
SendKeys "%SS"
End Sub

Function AOL4_SNfromIM()
'This will return the Screen Name from an IM
IM% = FindChildByTitle(AOL4_MDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOL4_MDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IM%)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
AOL4_SNfromIM = Naw$
End Function

Sub AOL4_Tool(wch%)
'You can only use  AOL4_Tool(0) or AOL4_Tool(2)
'0 = Open mail box    2 = Write Mail
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(Tool2%, "_AOL_Icon")
Icon2% = GetWindow(ico3n%, wch%)
click% = SendMessageByNum(Icon2%, WM_LBUTTONDOWN, 0&, 0&)
click% = SendMessageByNum(Icon2%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub AOL4_UNHide()
'This will Un Hide the AOL window after hidden
a = ShowWindow(AOL4_Window(), 5)
End Sub

Function AOL4_Window()
'This findz the AOL window
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL4_Window = AOL%
End Function

Sub AOLCursor()
Call RunMenuByString(AOL4_Window(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
MDI% = FindChildByClass(FindWindow("_AOL_Modal", vbNullString), "_AOL_Icon")
Call AOL4_Icon(MDI%)
'Call AOLClose(FindWindow("_AOL_Modal", vbNullString))
End Sub

Function AOLGetListString(Parent, Index, Buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
aolhandle = Parent
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

Function AOLGetTopWindow()
'This getz the window ontop of all others
AOLGetTopWindow = GetTopWindow(AOL4_MDI())
End Function


Function Find2ndChildByClass(parentw, childhand)
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


Function AOL4_FindRoom()
'Finds the chat room and sets focus on it
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
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
        AOL4_FindRoom = firs%
        Exit Function
    End If
    AOL4_FindRoom = 0
End Function
Function GetWinText(GetThis As Integer) As String
'This can get a window's caption or get text from just
'about anything that has text including _AOL_EDIT.

'Example:
'WinCaption$ = GetWinText(Pref%)

BufLen% = SendMessageByNum(GetThis%, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String(BufLen%, 0)
Q% = SendMessageByString(GetThis%, WM_GETTEXT, BufLen% + 1, Buffer$)


DoEvents
GetWinText$ = Buffer$
End Function
Sub AOLClose(winew)
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub
Sub FadePreview(PreTxtMain As Control, FadedText As String, PreTxt As TextBox)
'by monk-e-god
'-FADE PREVIEW-
'To use the fadepreview you need a
'rich textbox (which requires an ocx)
'and a regular text box in which the
'HTML will be interpreted.

'example:
'HTMLbox.Text = FadeByColor4(FADE_RED, FADE_BLACK, FADE_GREY, FADE_GREEN, "Red/Black/Grey/Green Fade Preview", False)
'Call FadePreview(PreviewBox, HTMLbox.Text, InvisBox)

'now in the rich textbox, PreviewBox, you
'will see a Red to Black to Grey to Green
'fade saying "Red/Black/Grey/Green Fade Preview"

'NOTE: You cannot preview wavy fades.
'NOTE: PreTxtMain MUST be a rich textbox!

PreTxtMain.Text = ""
Dim Starts()
Dim Lengths()
Dim Colors()
Dim LastHtml%
Dim CurStart%
Dim CurLen%
Dim CurColor$
Dim NumFades%
PreTxt.Text = FadedText
NumFades% = 0
LastHtml% = 2
findhtml% = 1
While findhtml%
If NumFades% = 0 Then findhtml% = 0

NumFades% = NumFades% + 1
findhtml% = InStr(findhtml% + 1, PreTxt.Text, "<Font Color=#") 'InStr(LastHtml - 1, PreTxt.Text, "<Font Color=#")
If findhtml% = 0 Then GoTo Blah
LastHtml% = InStr(findhtml% + 1, PreTxt.Text, ">")
thecolor = Mid(PreTxt.Text, findhtml% + 13, 6)
htmlblue$ = Right(thecolor, 2)
htmlgreen$ = Mid(thecolor, 3, 2)
htmlred$ = Left(thecolor, 2)
vbcolor = "&H00" + htmlblue$ + htmlgreen$ + htmlred$ + "&"

nexthtml% = InStr(findhtml% + 1, PreTxt.Text, "<Font Color=#")
CurLen% = 1
Firstpart$ = Left(PreTxt.Text, findhtml% - 1)
Secondpart$ = Mid(PreTxt.Text, LastHtml% + 1)
PreTxt.Text = Firstpart$ + Secondpart$
CurStart% = findhtml%
CurColor = vbcolor
ReDim Preserve Starts(NumFades%)
ReDim Preserve Lengths(NumFades%)
ReDim Preserve Colors(NumFades%)
Starts(NumFades%) = CurStart% - 1
Lengths(NumFades%) = CurLen%
Colors(NumFades%) = CurColor

Blah:
H = H
Wend
PreTxtMain.Text = PreTxt.Text
For cc% = 1 To NumFades% - 1
PreTxtMain.SelStart = Starts(cc%)
PreTxtMain.SelLength = Lengths(cc%)
PreTxtMain.SelColor = Val(Colors(cc%))
H = H
Next cc%
PreTxtMain.SelLength = 0

End Sub
Sub AOLSetFocus()
'SetFocusAPI doesn't work AOL because AOL has added
'a safeguard against other programs calling certain
'API functions (like owner-drawn things and like.)
'This is the only way known for setting the focus
'to AOL.  This is a normal VB command!
If AOLWindow() = 0 Then: Exit Sub
AppActivate GetCaption(AOLWindow())
End Sub

Function WordWrap(sText As String, ByVal lMaxWidth As Long) As String
    Dim lStart As Long
    Dim lEnd As Long
    Dim lTextLen As Long
    Dim sSep As String
    
    ' setup length and starting positions
    lTextLen = Len(sText)
    lStart = 1
    lEnd = lMaxWidth
    ' look for the following separator
    sSep = " "
    Do While lEnd < lTextLen
        ' Parse back to white space
        Do While InStr(sSep, Mid$(sText, lEnd, 1)) = 0
            lEnd = lEnd - 1
            ' Don't send us text with words longer than the lines!
            If lEnd <= lStart Then
                WordWrap = sText
                Exit Function
            End If
        Loop
        ' build wrapped string
        WordWrap = WordWrap & Mid$(sText, lStart, lEnd - lStart + 1) & Chr(13) + Chr(10)
        WordWrap = WordWrap & Chr(9)
        ' adjust pointers into original string
        lStart = lEnd + 1
        lEnd = lStart + lMaxWidth
    Loop
    ' get last bit of string if necessary
    WordWrap = WordWrap & Mid$(sText, lStart)
End Function
Function FileExist(ByVal sFileName As String) As Integer
'Example: If Not FileExist(app.Path & "\test.ini") then...
Dim i As Integer
On Error Resume Next
i = Len(Dir$(sFileName))
    If Err Or i = 0 Then
        AC_FileExist = False
        Else
        AC_FileExist = True
        End If
Resume Next
End Function
Sub AddMailToList(lst As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
AOLOpenMail
MailBox% = FindChildByTitle(MDI%, AOLUserSN & "'s Online MailBox")
TAD1% = FindChildByClass(MailBox%, "_AOL_TabControl")
TaD2% = FindChildByClass(TAD1%, "_AOL_TabPage")
TheTree% = FindChildByClass(TaD2%, "_AOL_Tree")
l0034 = SendMessageByNum(TheTree%, LB_GETCOUNT, 0, 0)
OldLoad% = l0034
DoEvents
DoEvents
lst.Clear
For l0038 = 0 To l0034 - 1
l003C$ = String(255, 0)
l0040 = SendMessageByString(TheTree%, LB_GETTEXT, l0038, l003C$)
l003C$ = Right$(l003C$, Len(l003C$) - 12)
l003C$ = Right$(l003C$, Len(l003C$) - InStr(l003C$, Chr(9)))
l003C$ = "<" & l0038 & ">" & Trim(l003C$)
lst.AddItem l003C$
Next l0038
End Sub

Function AOL_ForwMail(sn As String, mail_index As Integer)

themail% = FindChildByTitle(AOLMDI(), GetSN & "'s Online Mailbox")
one% = FindChildByClass(themail%, "_AOL_TabControl")
Two% = FindChildByClass(one%, "_AOL_TabPage")
MailList% = FindChildByClass(Two%, "_AOL_Tree")
SendNow = SendMessageByNum(MailList%, LB_SETCURSEL, mail_index - 1, 0)
Pause 0.5
ico% = FindChildByClass(themail%, "_AOL_Icon")
AOLClick ico%
Ico1% = GetWindow(ico%, 2)
ico2% = GetWindow(Ico1%, 2)

SendNow = SendMessageByNum(MailList%, LB_SETCURSEL, mail_index - 1, 0)
AOLClick ico2%
Pause 0.5
AOLClick ico2%
AOLClick ico2%
AOLClick ico2%
AOLClick ico2%
AOLClick ico2%

Do
Jess = AOLFindmail()
Loop Until Jess <> 0
tm% = FindChildByClass(Jess, "_AOL_Icon")
but% = GetWindow(tm%, 2)
but1% = GetWindow(but%, 2)
but2% = GetWindow(but1%, 2)
but3% = GetWindow(but2%, 2)
but4% = GetWindow(but3%, 2)
but5% = GetWindow(but4%, 2)
but6% = GetWindow(but5%, 2)
butf% = GetWindow(but6%, 2)
AOLSetFocus
Pause 0.4

List% = FindChildByClass(AOLFindmail(), "RICHCNTL")
waitForchange1 (List%)
Pause 0.5

dad% = FindChildByClass(List%, "RICHTRACKERBAR")
DADA% = FindChildByClass(dad%, "BUTTON")

If dad% <> 0 Then
D = SendMessageByNum(DADA%, WM_LBUTTONDOWN, &HD, 0)
Pause 0.07
R = SendMessageByNum(DADA%, WM_LBUTTONUP, &HD, 0)
Pause 0.5
Do: DoEvents
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): Exit Do
If IM2% = 0 Then Exit Do
Loop
End If

Do
D = SendMessageByNum(butf%, WM_LBUTTONDOWN, &HD, 0)
Pause 0.07
R = SendMessageByNum(butf%, WM_LBUTTONUP, &HD, 0)
Pause 0.9
Loop Until AOLFindmail1 <> 0
Pause 0.5


writ% = AOLFindmail1()
AOLSetFocus
Pause 0.1
Ed% = FindChildByClass(writ%, "_AOL_Edit")
X = SendMessageByString(Ed%, WM_SETTEXT, 0, sn)
ed2% = GetWindow(Ed%, 2)
ed3% = GetWindow(ed2%, 2)
ed4% = GetWindow(ed3%, 2)
ed5% = GetWindow(ed4%, 2)

getSW$ = GetWinText(ed5%)
txt = Mid(getSW$, 6)

w = SendMessageByString(ed5%, WM_SETTEXT, 0, txt)
rich% = FindChildByClass(writ%, "RICHCNTL")

snd% = FindChildByClass(writ%, "_AOL_Icon")
snd1% = GetWindow(snd%, 2)
snd2% = GetWindow(snd1%, 2)
snd3% = GetWindow(snd2%, 2)
snd4% = GetWindow(snd3%, 2)
snd5% = GetWindow(snd4%, 2)
snd6% = GetWindow(snd5%, 2)
snd7% = GetWindow(snd6%, 2)
snd8% = GetWindow(snd7%, 2)
snd9% = GetWindow(snd8%, 2)
snd10% = GetWindow(snd9%, 2)
snd11% = GetWindow(snd10%, 2)
snd12% = GetWindow(snd11%, 2)
snd13% = GetWindow(snd12%, 2)
snd14% = GetWindow(snd13%, 2)
snd15% = GetWindow(snd14%, 2)
snd16% = GetWindow(snd15%, 2)
snd17% = GetWindow(snd16%, 2)
snd18% = GetWindow(snd17%, 2)
D = SendMessageByNum(snd14%, WM_LBUTTONDOWN, &HD, 0)
Pause 0.07
R = SendMessageByNum(snd14%, WM_LBUTTONUP, &HD, 0)
Pause 0.1

writ% = AOLFindmail1()
If writ% <> 0 Then
D = SendMessageByNum(snd14%, WM_LBUTTONDOWN, &HD, 0)
Pause 0.07
R = SendMessageByNum(snd14%, WM_LBUTTONUP, &HD, 0)
Pause 0.1
End If
Pause 5
AOLClose (Jess)
Jess = AOLFindmail()
AOLClose (Jess)

End Function
Function AOLFindmail()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Icon")
listere% = FindChildByClass(childfocus%, "RICHCNTL")
listerf% = FindChildByTitle(AOLMDI, "Welcome, ")

If listers% <> 0 And listere% <> 0 And GetCaption(childfocus%) <> GetCaption(listerf%) And childfocus% <> AOLFindRoom And GetCaption(childfocus%) <> "AOL Today -- Welcome" Then AOLFindmail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function
Function AOLFindmail1()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Icon")
listere% = FindChildByClass(childfocus%, "RICHCNTL")
listerf% = FindChildByClass(childfocus%, "_AOL_Edit")
listerg% = FindChildByTitle(AOLMDI, "Welcome, ")

If listers% <> 0 And listere% <> 0 And listerf% <> 0 And GetCaption(childfocus%) <> GetCaption(listerf%) And childfocus% <> AOLFindRoom And GetCaption(childfocus%) <> "AOL Today -- Welcome" Then AOLFindmail1 = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function
Sub waitForchange1(Box%)
Dim old
Dim boy
Dim sid
sid = 0
ttop:
Do
old = GetWinText(Box%)
Pause 0.3
boy = GetWinText(Box%)
Loop Until boy = old
sid = sid + 1
If sid <> 2 Then GoTo ttop



End Sub
Function ChatSendBox()
 Dim Room%
 Dim ar1%
 Dim ar2%
 Dim ar3%
 Dim ar4%
 Dim ar5%
 Dim ar6%
 Dim ar7%
 
 
 
Room% = AOLFindRoom()
ar1% = FindChildByClass(Room%, "RICHCNTL")
ar2% = GetWindow(ar1%, 2)
ar3% = GetWindow(ar2%, 2)
ar4% = GetWindow(ar3%, 2)
ar5% = GetWindow(ar4%, 2)
ar6% = GetWindow(ar5%, 2)
ar7% = GetWindow(ar6%, 2)
ChatSendBox = ar7%
End Function



Private Sub Timer1_Timer()

       PauseTime = 5 ' Set duration.
       ProgressBar1.Max = ((PauseTime + 1) * Timer1.interval)
       Start = Timer ' Set start time.

              Do While Timer < Start + PauseTime
                     ProgressBar1.Value = ProgressBar1.Value + 10

                            DoEvents ' Yield to other processes.
                            Loop

                     Finish = Timer ' Set end time.
                     TotalTime = Finish - Start ' Calculate total time.
                     MsgBox "Paused for " & TotalTime & " seconds"
                     Timer1.Enabled = False
              

End Sub



Sub SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Sub
Sub SetText(Window, Text)

eltext% = SendMessageByString(Window, WM_SETTEXT, 0, "")
eltext% = SendMessageByString(Window, WM_SETTEXT, 0, Text)
End Sub

Public Sub ShowWelcome()
Dim X
wlcm% = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "Welcome, ")
X = ShowWindow(wlcm%, SW_SHOW)
End Sub
Sub SizeFormToWindow(frm As Form, win%)
'This will make a frm the size of a win
'ex: Call SizeFormToWindow(form1, IM%)
Dim wndRect As RECT, lRet As Long
lRet = GetWindowRect(win%, wndRect)
With frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub


Function StayOnline()
hwndz% = FindWindow("_AOL_Palette", "America Online")
childhwnd% = FindChildByTitle(hwndz%, "OK")
AOL4_Button (childhwnd%)
End Function


Function Scanfile(Filename As String, SearchString As String) As Long

       Free = FreeFile
       Dim Where As Long
       Open Filename$ For Binary Access Read As #Free

              For X = 1 To LOF(Free) Step 32000
                     Text$ = Space(32000)
                     Get #Free, X, Text$
                     Debug.Print X

                            If InStr(1, Text$, SearchString$, 1) Then
                                   Where = InStr(1, Text$, SearchString$, 1)
                                   Scanfile = (Where + X) - 1
                                   Close #Free
                                   Exit For
                            End If

              Next X

       Close #Free
End Function



Sub StopCDRecord()
 Dim i As Long, RS As String, cb As Long
       RS = Space$(128)
       i = mciSendString("stop cdaudio", RS, 128, cb)
       i = mciSendString("close cdaudio", RS, 128, cb)
End Sub

Sub ToChat(Chat)
Room% = AOL4_FindRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub addroom(lst As ListBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
namez$ = String$(256, " ")
ret = AOLGetList(Index, namez$)
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)))
ADD_AOL_LB namez$, lst
Next Index
end_addr:
lst.RemoveItem lst.ListCount - 1
i = GetListIndex(lst, AOLGetUser())
If i <> -2 Then lst.RemoveItem i
End Sub
Sub addroomtotext(thelist As ListBox, Text As TextBox)
' addroomtotext list1, text1
Dim Y
Call addroom(thelist)
For Y = 0 To thelist.ListCount - 1
tt$ = tt$ + thelist.List(Y) + ","
Next Y
TimeOut (0.01)
Text.Text = tt$

End Sub
Sub aol4_macroScroll(Text As String)
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    SendChat Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If Counter = 4 Then
        TimeOut (2.9)
        Counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub
Sub aol4_SpiralScroll(txt As TextBox)
X = txt.Text
thastar:
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
mystr = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = mystr
SendChat "•[" + X + "]•"
If txt.Text = X Then
Exit Sub
End If
GoTo thastar

End Sub

Sub Answerbot()
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
Dim X As Integer
DoEvents
a = LastChatLine
last = Len(a)
For X = 1 To last
name = Mid(a, X, 1)
Final = Final & name
If name = ":" Then Exit For
Next X
Final = Left(Final, Len(Final) - 1)
If Final = AOLGetUser Then
Exit Sub
Else
If InStr(a, "/Vv KoBe vV") Then
 SendChat (" Don't Waste Time on a Server")
Call TimeOut(0.6)
End If
End If
End Sub
Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Function AOLGetUser()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = user
End Function
Sub AOLSetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub AOLSNReset(sn$, aoldir$, Replace$)
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
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
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
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
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
Sub AOLversion()

AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome, " + UserSN())
aol3% = FindChildByClass(Wel%, "RICHCNTL")
If aol3% = 0 Then AC_AOLVersion = 25: Exit Sub
If aol3% <> 0 Then
    If GetCaption(AOL%) <> "America Online" Then AC_AOLVersion = 3 Else AC_AOLVersion = 4
    End If
    End Sub
Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
Function Chat_RoomName()
Call GetCaption(AOLFindChatRoom)
End Function
Sub AOLIcon(icon%)
click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Public Function AOLSupRoom()
IsUserOnline
If AOLIsOnline = 0 Then GoTo last
AOL4_FindRoom
If AOLFindRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call SendChat("SuP 2  " & Person$)
TimeOut (1)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function
Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Public Function AOLGetNewMail(Index) As String
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
mail% = FindChildByTitle(MDI%, AOLGetUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

'de = sendmessage(aoltree%, LB_GETCOUNT, 0, 0)
txtlen% = SendMessageByNum(AOLTree%, LB_GETTEXTLEN, Index, 0&)
txt$ = String(txtlen% + 1, 0&)
X = SendMessageByString(AOLTree%, LB_GETTEXT, Index, txt$)
AOLGetNewMail = txt$
End Function
Sub ADD_AOL_LB(itm As String, lst As ListBox)
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
Public Sub AOLClick(Iconhwnd%)
'Simulates clicking the mouse button.
Dim click%
click% = SendMessage(Iconhwnd%, WM_LBUTTONDOWN, 0, 0&)
click% = SendMessage(Iconhwnd%, WM_LBUTTONUP, 0, 0&)
End Sub

Public Function AOLOpenMail()
'Opens the mail box.
'Returns the handle of _AOL_TREE
Dim AOL%
Dim TabControl%
Dim TabPage%
Dim MDI%
Dim MailBox%
Dim TabPageNew%
Dim TabPageOld%
Dim TabPageSent%
Dim TreeNew%
Dim TreeOld%
Dim TreeSent%
Dim Edit%
Dim Button%
Dim TheTree%
Dim Num1%
Dim Num2%
Dim TAD1%
Dim TaD2%
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Edit% = FindChildByClass(AOL%, "AOL Toolbar")
Edit% = FindChildByClass(Edit%, "_AOL_Toolbar")
Button% = FindChildByClass(Edit%, "_AOL_Icon")
AOLClick Button%
Do
DoEvents
MailBox% = FindChildByTitle(MDI%, AOLUserSN & "'s Online MailBox")
Loop Until MailBox% <> 0
TimeOut (3)
TAD1% = FindChildByClass(MailBox%, "_AOL_TabControl")
TaD2% = FindChildByClass(TAD1%, "_AOL_TabPage")
TheTree% = FindChildByClass(TaD2%, "_AOL_Tree")
Do
DoEvents
Num2% = Num1%
Num1% = SendMessageByNum(TheTree%, LB_GETCOUNT, 0, 0)
TimeOut (1)
Loop Until Num1% = Num2% And Num1% <> 0

End Function

Public Function AOLGetList(Index As Long, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
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
Public Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone
AOLFindRoom = 0
GoTo 50
firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone

Wend

bone:
Room% = firs%
AOLFindRoom = Room%
50
End Function
Public Sub AOLClearChat()
getpar% = AOL4_FindRoom()
child = FindChildByClass(getpar%, "RICHCNTL")
End Sub
Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear

Room = AOL4_FindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = UserSN Then GoTo Na
ListBox.AddItem Person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub


Sub AddSysFonts(lst)
'lst can be a ListBox or ComboBox
For i = 0 To Screen.FontCount - 1  ' Determine number of fonts.
    lst.AddItem Screen.Fonts(i)  ' Put each font into list box.
Next i
End Sub


Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function GetWindowDir()
'finds the window's directory
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function

Public Sub HideWelcome()
Dim X
wlcm% = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "Welcome, ")
X = ShowWindow(wlcm%, SW_HIDE)
End Sub
Sub HideWindow(hWnd)
'This hides the hwnd window
hi = ShowWindow(hWnd, SW_HIDE)
End Sub


Private Sub InitializeTextBoxFast()

        
       'This routine assigns the string to temporary string variabl
       '     e
       '     'as the string is being built.
       Dim tmp As String
       Dim i As Integer
       Dim J As Integer
       Text1.Text = ""
       lblStatus = "Performing fast load..."
        
       '     'just a pause to let the textbox and label update

              DoEvents

                            For i% = 1 To 100
                                   tmp$ = tmp$ + "This is line " + Str$(i%)
                                    
                                   '     'Add 10 words to a line of text

                                          For J% = 1 To 10
                                                 tmp$ = tmp$ + " ...Word " + Str$(J%)
                                          Next J%

                                    
                                   '     'Force a carriage return and linefeed
                                   '     'VB3 users need to use chr$(13) & chr$(10)
                                   tmp$ = tmp$ + vbCrLf
                            Next i%

                      
                     '     'Now it's time to assign it to the text property.
                     Text1.Text = tmp$
                      
              End Sub
Private Sub InitializeTextBoxSlow()

        
       'This routine assigns the string to the textbox text propert
       '     y
       '     'as the string is being built. This is the method that
       '     'the MS VBKB detailed. I named it InitializeTextBoxSlow.
       Dim i As Integer
       Dim J As Integer
       Text1.Text = ""
       lblStatus = "Performing slow load..."
        
       '     'just a pause to let the textbox and label update

              DoEvents

                            For i% = 1 To 100
                                   Text1.Text = Text1.Text + "This is line " + Str$(i%)
                                    
                                   '     'Add 10 words to a line of text.

                                          For J% = 1 To 10
                                                 Text1.Text = Text1.Text + " ...Word " + Str$(J%)
                                          Next J%

                                    
                                   '     'Force a carriage return and linefeed
                                   '     'VB3 users need to use chr$(13) & chr$(10)
                                   Text1.Text = Text1.Text + vbCrLf
                            Next i%

                     Text1.Text = Text1.Text
              End Sub
Function IsUserOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Sub MailMe(Recipiants, subject, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, messege)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub
Sub PhishPhrases()
Randomize X
phraZes = Int((Val("140") * Rnd) + 1)
If phraZes = "1" Then
Text2 = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf phraZes = "2" Then
Text2 = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
ElseIf phraZes = "3" Then
Text2 = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
ElseIf phraZes = "4" Then
Text2 = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
ElseIf phraZes = "5" Then
Text2 = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
ElseIf phraZes = "6" Then
Text2 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "7" Then
Text2 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "8" Then
Text2 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "9" Then
Text2 = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf phraZes = "10" Then
Text2 = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "11" Then
Text2 = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf phraZes = "12" Then
Text2 = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "13" Then
Text2 = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
ElseIf phraZes = "14" Then
Text2 = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
ElseIf phraZes = "15" Then
Text2 = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "16" Then
Text2 = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf phraZes = "17" Then
Text2 = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
ElseIf phraZes = "18" Then
Text2 = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
ElseIf phraZes = "19" Then
Text2 = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
ElseIf phraZes = "20" Then
Text2 = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
ElseIf phraZes = "21" Then
Text2 = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
ElseIf phraZes = "22" Then
Text2 = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
ElseIf phraZes = "23" Then
Text2 = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
ElseIf phraZes = "24" Then
Text2 = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
ElseIf phraZes = "25" Then
Text2 = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
ElseIf phraZes = "26" Then
Text2 = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
ElseIf phraZes = "27" Then
Text2 = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
ElseIf phraZes = "28" Then
Text2 = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf phraZes = "29" Then
Text2 = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf phraZes = "30" Then
Text2 = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf phraZes = "31" Then
Text2 = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf phraZes = "32" Then
Text2 = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf phraZes = "33" Then
Text2 = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
ElseIf phraZes = "34" Then
Text2 = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
ElseIf phraZes = "35" Then
Text2 = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
ElseIf phraZes = "36" Then
Text2 = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf phraZes = "37" Then
Text2 = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf phraZes = "38" Then
Text2 = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf phraZes = "39" Then
Text2 = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf phraZes = "40" Then
Text2 = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
ElseIf phraZes = "41" Then
Text2 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "42" Then
Text2 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "43" Then
Text2 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "44" Then
Text2 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "45" Then
Text2 = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf phraZes = "46" Then
Text2 = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf phraZes = "47" Then
Text2 = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf phraZes = "48" Then
Text2 = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
ElseIf phraZes = "49" Then
Text2 = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
ElseIf phraZes = "50" Then
Text2 = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
ElseIf phraZes = "51" Then
Text2 = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
ElseIf phraZes = "52" Then
Text2 = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
ElseIf phraZes = "53" Then
Text2 = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "54" Then
Text2 = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf phraZes = "55" Then
Text2 = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "56" Then
Text2 = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
ElseIf phraZes = "57" Then
Text2 = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "58" Then
Text2 = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "59" Then
Text2 = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
ElseIf phraZes = "60" Then
Text2 = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
ElseIf phraZes = "61" Then
Text2 = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
ElseIf phraZes = "62" Then
Text2 = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
ElseIf phraZes = "63" Then
Text2 = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "64" Then
Text2 = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "65" Then
Text2 = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
ElseIf phraZes = "66" Then
Text2 = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
ElseIf phraZes = "67" Then
Text2 = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
ElseIf phraZes = "68" Then
Text2 = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
ElseIf phraZes = "69" Then
Text2 = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
ElseIf phraZes = "70" Then
Text2 = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
ElseIf phraZes = "71" Then
Text2 = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
ElseIf phraZes = "72" Then
Text2 = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
ElseIf phraZes = "73" Then
Text2 = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
ElseIf phraZes = "74" Then
Text2 = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
ElseIf phraZes = "75" Then
Text2 = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
ElseIf phraZes = "76" Then
Text2 = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
ElseIf phraZes = "77" Then
Text2 = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf phraZes = "78" Then
Text2 = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
ElseIf phraZes = "79" Then
Text2 = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
ElseIf phraZes = "80" Then
Text2 = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
ElseIf phraZes = "81" Then
Text2 = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf phraZes = "82" Then
Text2 = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf phraZes = "83" Then
Text2 = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf phraZes = "84" Then
Text2 = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf phraZes = "85" Then
Text2 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."

End If
Text2 = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "

End Sub
Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub

Sub KillModal()
MODAL% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(MODAL%, WM_CLOSE, 0, 0)
End Sub
Sub killwait()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Sub killwin(hWnd%)
'Closes a chosen window
Dim KillNow%
KillNow% = SendMessageByNum(hWnd%, WM_CLOSE, 0, 0)
End Sub
Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub
Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
Function RemoveSpace(thetext$)
Dim Text$
Dim theloop%
Text$ = thetext$
For theloop% = 1 To Len(thetext$)
If Mid(Text$, theloop%, 1) = " " Then
Text$ = Left$(Text$, theloop% - 1) + Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
RemoveSpace = Text$
End Function
Sub ResetNew(sn As String, pth As String)
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(sn)) = "NEWUSER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo no_reset
If Len(sn) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
tru_sn = "NewUser" + String$(Len(sn) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(sn)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(sn))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(sn)
l9E68& = Len(sn)
While l9E68& < l9E6A&
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(sn)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(sn))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next

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
Function ScrambleText(thetext)
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
'Full bas by eLeSsDee == eLeSsDee@mindless.com
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
Scrambled$ = Scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
Scrambled$ = Scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = Scrambled$

Exit Function
End Function
Function ScrollText&(TextBox As Control, vLines As Integer)

       Dim Success As Long
       Dim SavedWnd As Long
       Dim R As Long
       Dim lines As Long
       'save the window handle of the control that currently has fo
       '     cus
       SavedWnd = Screen.ActiveControl.hWnd
       lines& = vLines
        
       '     'Set the focus to the passed control (text control)
       TextBox.SetFocus
        
       '     'Scroll the lines.
       Success = SendMessageLong(TextBox.hWnd, EM_LINESCROLL, 0, lines&)
        
       '     'Restore the focus to the original control
       R = PutFocus(SavedWnd)
        
       '     'Return the number of lines actually scrolled
       ScrollText& = Success
End Function
Public Sub TB4(Number As Integer)
AOL% = FindWindow("AOL Frame25", vbNullString)
TB% = FindChildByClass(AOL%, "AOL Toolbar")
tc% = FindChildByClass(TB%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")

If Number = 1 Then
    Call AOLIcon(td%)
    Exit Sub
End If

For T = 0 To Number - 2
td% = GetWindow(td%, 2)
Next T

Call AOLIcon(td%)

End Sub
Function UnEliteText(Text As String)
Dim X$
Dim Char%
Dim T$
For Char% = 1 To Len(Text)
    X$ = Mid$(Text, Char%, 1)
    If X$ = "â" Or X$ = "å" Or X$ = "ã" Then X$ = "a"
    If X$ = "ç" Then X$ = "c"
    If X$ = "ê" Or X$ = "ë" Or X$ = "é" Then X$ = "e"
    If X$ = "ƒ" Then X$ = "f"
    If X$ = "ï" Or X$ = "î" Or X$ = "í" Then X$ = "i"
    If X$ = "¡" Then X$ = "j"
    If X$ = "|" Then X$ = "l"
    If X$ = "ñ" Then X$ = "n"
    If X$ = "ô" Or X$ = "ð" Or X$ = "õ" Then X$ = "o"
    If X$ = "þ" Then X$ = "p"
    If X$ = "š" Then X$ = "s"
    If X$ = "†" Then X$ = "t"
    If X$ = "û" Or X$ = "ü" Or X$ = "ú" Then X$ = "u"
    If X$ = "×" Then X$ = "x"
    If X$ = "ÿ" Or X$ = "ý" Then X$ = "y"
    If X$ = "Ä" Or X$ = "Å" Or X$ = "Ã" Then X$ = "A"
    If X$ = "ß" Then X$ = "B"
    If X$ = "Ç" Or X$ = "©" Then X$ = "C"
    If X$ = "Ð" Then X$ = "D"
    If X$ = "Ê" Or X$ = "Ë" Or X$ = "É" Then X$ = "E"
    If X$ = "Î" Or X$ = "Ï" Or X$ = "Í" Then X$ = "I"
    If X$ = "£" Then X$ = "L"
    If X$ = "Ñ" Then X$ = "N"
    If X$ = "Ö" Or X$ = "Ô" Or X$ = "Õ" Then X$ = "O"
    If X$ = "Þ" Then X$ = "P"
    If X$ = "®" Then X$ = "R"
    If X$ = "§" Or X$ = "Š" Then X$ = "S"
    If X$ = "Ü" Or X$ = "Û" Then X$ = "U"
    If X$ = "¥" Then X$ = "Y"
    T$ = T$ + X$
Next Char%
UnEliteText = T$
End Function

Sub UnUpChat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub
Sub UpChat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Function WinCaption(win)
Dim GetWinText%
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function
Function GetChatText()
Room% = AOL4_FindRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(AORich%)
GetChatText = chattext
End Function
Function SNFromLastChatLine()
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        sn = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = sn
End Function
Function LastChatLine()
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Function LineFromText(Text, theline)
'This returnz a line from text
theview$ = Text
For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
C = C + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If theline = C Then GoTo ex
thechars$ = ""
End If

Next FindChar
Exit Function
ex:
thechatext$ = ReplaceText(thechatext$, Chr(13), "")
thechatext$ = ReplaceText(thechatext$, Chr(10), "")
LineFromText = thechatext$
End Function

Sub ListToList(Source, destination)
'Copies 1 list to another
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(destination, LB_ADDSTRING, 0, Buffer$)
Next Adding
End Sub

Sub MaxWindow(hWnd)
'Maximizes the hwnd window
ma = ShowWindow(hWnd, SW_MAXIMIZE)
End Sub


Sub MiniWindow(hWnd)
'This minimizes the hwnd window
mi = ShowWindow(hWnd, SW_MINIMIZE)
End Sub

Sub NotOnTop(the As Form)
'This makes the form not stayontop
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub ParentChange(Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub




Function LastChatLineWithSN()
chattext$ = GetChatText
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
Function Decrypt(Text As String)
Dim X%
Dim Char$
Dim TextX$
Text = LCase(Text)
For X% = 1 To Len(Text)
Char$ = Mid$(Text, X%, 1)
If Char$ = "q" Then Char$ = "a": GoTo Hell
If Char$ = "w" Then Char$ = "s": GoTo Hell
If Char$ = "e" Then Char$ = "d": GoTo Hell
If Char$ = "r" Then Char$ = "f": GoTo Hell
If Char$ = "t" Then Char$ = "g": GoTo Hell
If Char$ = "y" Then Char$ = "h": GoTo Hell
If Char$ = "u" Then Char$ = "j": GoTo Hell
If Char$ = "i" Then Char$ = "k": GoTo Hell
If Char$ = "o" Then Char$ = "l": GoTo Hell
If Char$ = "p" Then Char$ = ":": GoTo Hell
If Char$ = "a" Then Char$ = "z": GoTo Hell
If Char$ = "s" Then Char$ = "x": GoTo Hell
If Char$ = "d" Then Char$ = "c": GoTo Hell
If Char$ = "f" Then Char$ = "v": GoTo Hell
If Char$ = "g" Then Char$ = "b": GoTo Hell
If Char$ = "h" Then Char$ = "n": GoTo Hell
If Char$ = "j" Then Char$ = "m": GoTo Hell
If Char$ = "k" Then Char$ = "<": GoTo Hell
If Char$ = "l" Then Char$ = ">": GoTo Hell
If Char$ = ":" Then Char$ = "/": GoTo Hell
If Char$ = "z" Then Char$ = "1": GoTo Hell
If Char$ = "x" Then Char$ = "2": GoTo Hell
If Char$ = "c" Then Char$ = "3": GoTo Hell
If Char$ = "v" Then Char$ = "4": GoTo Hell
If Char$ = "b" Then Char$ = "5": GoTo Hell
If Char$ = "n" Then Char$ = "6": GoTo Hell
If Char$ = "m" Then Char$ = "7": GoTo Hell
If Char$ = "<" Then Char$ = "8": GoTo Hell
If Char$ = ">" Then Char$ = "9": GoTo Hell
If Char$ = "/" Then Char$ = "0": GoTo Hell
If Char$ = "1" Then Char$ = "q": GoTo Hell
If Char$ = "2" Then Char$ = "w": GoTo Hell
If Char$ = "3" Then Char$ = "e": GoTo Hell
If Char$ = "4" Then Char$ = "r": GoTo Hell
If Char$ = "5" Then Char$ = "t": GoTo Hell
If Char$ = "6" Then Char$ = "y": GoTo Hell
If Char$ = "7" Then Char$ = "u": GoTo Hell
If Char$ = "8" Then Char$ = "i": GoTo Hell
If Char$ = "9" Then Char$ = "o": GoTo Hell
If Char$ = "0" Then Char$ = "p": GoTo Hell

Hell:
TextX$ = TextX$ + Char$
Next X%
Decrypt = TextX$
End Function

Function Encrypt(Text As String)
Dim X%
Dim Char$
Dim TextX$
Text = LCase(Text)
For X% = 1 To Len(Text)
Char$ = Mid$(Text, X%, 1)
If Char$ = "q" Then Char$ = "1": GoTo Hell
If Char$ = "w" Then Char$ = "2": GoTo Hell
If Char$ = "e" Then Char$ = "3": GoTo Hell
If Char$ = "r" Then Char$ = "4": GoTo Hell
If Char$ = "t" Then Char$ = "5": GoTo Hell
If Char$ = "y" Then Char$ = "6": GoTo Hell
If Char$ = "u" Then Char$ = "7": GoTo Hell
If Char$ = "i" Then Char$ = "8": GoTo Hell
If Char$ = "o" Then Char$ = "9": GoTo Hell
If Char$ = "p" Then Char$ = "0": GoTo Hell
If Char$ = "a" Then Char$ = "q": GoTo Hell
If Char$ = "s" Then Char$ = "w": GoTo Hell
If Char$ = "d" Then Char$ = "e": GoTo Hell
If Char$ = "f" Then Char$ = "r": GoTo Hell
If Char$ = "g" Then Char$ = "t": GoTo Hell
If Char$ = "h" Then Char$ = "y": GoTo Hell
If Char$ = "j" Then Char$ = "u": GoTo Hell
If Char$ = "k" Then Char$ = "i": GoTo Hell
If Char$ = "l" Then Char$ = "o": GoTo Hell
If Char$ = ":" Then Char$ = "p": GoTo Hell
If Char$ = "z" Then Char$ = "a": GoTo Hell
If Char$ = "x" Then Char$ = "s": GoTo Hell
If Char$ = "c" Then Char$ = "d": GoTo Hell
If Char$ = "v" Then Char$ = "f": GoTo Hell
If Char$ = "b" Then Char$ = "g": GoTo Hell
If Char$ = "n" Then Char$ = "h": GoTo Hell
If Char$ = "m" Then Char$ = "j": GoTo Hell
If Char$ = "<" Then Char$ = "k": GoTo Hell
If Char$ = ">" Then Char$ = "l": GoTo Hell
If Char$ = "/" Then Char$ = ":": GoTo Hell
If Char$ = "1" Then Char$ = "z": GoTo Hell
If Char$ = "2" Then Char$ = "x": GoTo Hell
If Char$ = "3" Then Char$ = "c": GoTo Hell
If Char$ = "4" Then Char$ = "v": GoTo Hell
If Char$ = "5" Then Char$ = "b": GoTo Hell
If Char$ = "6" Then Char$ = "n": GoTo Hell
If Char$ = "7" Then Char$ = "m": GoTo Hell
If Char$ = "8" Then Char$ = "<": GoTo Hell
If Char$ = "9" Then Char$ = ">": GoTo Hell
If Char$ = "0" Then Char$ = "/": GoTo Hell

Hell:
TextX$ = TextX$ + Char$
Next X%
Encrypt = TextX$
End Function
Sub CopyToClipBoard(Text As String)
Clipboard.SetText (Text)
'Example: CopyToClipBoard(Text1)
End Sub

Sub DoubleClick(Button%)
'This double clicks a button of your choice
Dim DoubleClickNow%
DoubleClickNow% = SendMessageByNum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub
Function fader(thetext$)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 8
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    v$ = Mid$(G$, w + 4, 1)
    Q$ = Mid$(G$, w + 5, 1)
    X$ = Mid$(G$, w + 6, 1)
    Y$ = Mid$(G$, w + 7, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & T$ & "<FONT COLOR=" & Chr$(34) & "#DCDCDC" & Chr$(34) & ">" & v$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & Q$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & X$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & Y$
Next w
SendChat P$
End Function

Sub UnloadAll()
'makes sure that all Forms are Unloaded
'better than the End Statement
Dim frm As Form
For Each frm In Forms
Unload frm
Next frm
End Sub
Sub SaveText(lst As TextBox, File As String)

On Error GoTo error
Dim mystr As String
Open File For Output As #1
Print #1, lst
Close 1
Exit Sub
error:
X = MsgBox("Error!!", vbOKOnly, "Error!!")
End Sub
Sub LoadText(lst As TextBox, File As String)
'You need a common dialog box for this
On Error GoTo error
Dim mystr As String
Open File For Input As #1
Do While Not EOF(1)
            Line Input #1, a$
            texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
        Loop
        lst = texto$
Close #1
Exit Sub
error:
X = MsgBox("File Not Found", vbOKOnly, "Error!!")
End Sub
Sub SaveList(lst As ListBox, File As String)
'you need a common dialog box for this
On Error GoTo error
Open File For Output As #1
For i = 0 To lst.ListCount - 1
a$ = lst.List(i)
Print #1, a$
Next
Close 1
Exit Sub
error:
X = MsgBox("Error!!", vbOKOnly, "Error!!")

End Sub
Sub LoadList(lst As ListBox, File As String)
'you need a common dialog box for this
On Error GoTo error
Open File For Input As #1
Do Until EOF(1)
Input #1, a$
lst.AddItem a$
Loop
Close 1
Exit Sub
error:
X = MsgBox("File Not Found", vbOKOnly, "Error!!")

End Sub
Sub ProgressBar(pb As Control, ByVal percent)
'This allows you to make a Progress Bar with no
'OCX file.
' example:  call ProgressBar(form1.picture1,33)
Dim num$
If Not pb.AutoRedraw Then
    pb.AutoRedraw = -1
    End If
    pb.Cls
    pb.ScaleWidth = 100
    pb.DrawMode = 10
    num$ = Format$(percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$
    pb.Line (0, 0)-(percent, pb.ScaleHeight), , BF
    pb.Refresh
End Sub
Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)
End Sub
Function ReplaceText(Text As String, charfind As String, charchange As String)
Dim X%
Dim thechar$, thechars$
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If
 For X% = 1 To Len(Text)
thechar$ = Mid$(Text, X%, Len(charfind))
If thechar$ = charfind Then
Text = Left$(Text, X% - 1) + charchange + Right$(Text, Len(Text) - X% - Len(charfind) + 1)
End If
Next X%
ReplaceText = Text
End Function

Function ReverseText(Text)
For Words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, Words, 1)
Next Words
End Function


Sub SendMail(Recipiants, subject, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub
Function MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
sn = SNfromIM()
snlen = Len(SNfromIM()) + 3
Blah = Mid(IMmessage, InStr(IMmessagge, sn) + snlen)
MessageFromIM = Left(Blah, Len(Blah) - 1)
End Function
Function SNfromIM()

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = theSN$

End Function
Sub RespondIM(message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(IM%, "RICHCNTL")

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
List1.AddItem SNfromIM
List1.AddItem MessageFromIM
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
e2 = GetWindow(E, GW_HWNDNEXT) 'Send Text
E = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, Text1)
ClickIcon (E)
Call TimeOut(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
E = FindChildByClass(IM%, "RICHCNTL")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (E)
End Sub
Sub ClickIcon(icon%)
click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Function AOLWindow()
'This sets focus on the AOL window
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function
Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = user
End Function
Function KoolLink(Address As String, WhatToSay As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)
KoolLink = "<font color=#fffffe><pre=<a href=""" + Address + """></u>" + ReverseBoldFadedThree(WhatToSay, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy) + "</a></a>"
End Function


Function ReverseBold(TextBox As String, Wavy As Boolean)
Dim Text$
Dim X%
Dim Tot%
Tot% = Len(TextBox)
If Wavy = False Then
For X% = 1 To Tot% Step 2
Text$ = Text$ + "<b>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</b>" + Mid(TextBox, X% + 1, 1)
Next X%
End If
If Wavy = True Then
For X% = 1 To Tot% Step 4
Text$ = Text$ + "<sub><b>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</sub></b>" + Mid(TextBox, X% + 1, 1)
Text$ = Text$ + "<sup><b>" + Mid(TextBox, X% + 2, 1)
Text$ = Text$ + "</sup></b>" + Mid(TextBox, X% + 3, 1)
Next X%
End If
ReverseBold = Text$
End Function

Function ReverseItalic(TextBox As String, Wavy As Boolean)
Dim Text$
Dim X%
Dim Tot%
Tot% = Len(TextBox)
If Wavy = False Then
For X% = 1 To Tot% Step 2
Text$ = Text$ + "<i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</i>" + Mid(TextBox, X% + 1, 1)
Next X%
End If
If Wavy = True Then
For X% = 1 To Tot% Step 4
Text$ = Text$ + "<sub><i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</sub></i>" + Mid(TextBox, X% + 1, 1)
Text$ = Text$ + "<sup><i>" + Mid(TextBox, X% + 2, 1)
Text$ = Text$ + "</sup></i>" + Mid(TextBox, X% + 3, 1)
Next X%
End If
ReverseItalic = Text$
End Function

Function ReverseBoldItalic(TextBox As String, Wavy As Boolean)
Dim Text$
Dim X%
Dim Tot%
Tot% = Len(TextBox)
If Wavy = False Then
For X% = 1 To Tot% Step 2
Text$ = Text$ + "<b></i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</b><i>" + Mid(TextBox, X% + 1, 1)
Next X%
End If
If Wavy = True Then
For X% = 1 To Tot% Step 4
Text$ = Text$ + "<sub><b></i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</sub></b><i>" + Mid(TextBox, X% + 1, 1)
Text$ = Text$ + "<sup></i><b>" + Mid(TextBox, X% + 2, 1)
Text$ = Text$ + "</sup></b><i>" + Mid(TextBox, X% + 3, 1)
Next X%
End If
ReverseBoldItalic = Text$
End Function

Function ReverseBoldFadedThree(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = ThreeColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedThree = Text$
End Function

Function ReverseBoldFadedTwo(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TwoColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedTwo = Text$
End Function

Function ReverseBoldFadedFour(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FourColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedFour = Text$
End Function

Function ReverseBoldFadedFive(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FiveColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedFive = Text$
End Function

Function ReverseBoldFadedSix(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SixColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedSix = Text$
End Function

Function ReverseBoldFadedSeven(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SevenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedSeven = Text$
End Function

Function ReverseBoldFadedEight(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = EightColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedEight = Text$
End Function

Function ReverseBoldFadedNine(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = NineColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedNine = Text$
End Function

Function ReverseBoldFadedTen(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedTen = Text$
End Function
Function ReverseBoldItalicFadedThree(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = ThreeColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedThree = Text$
End Function

Function ReverseBoldItalicFadedTwo(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TwoColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedTwo = Text$
End Function

Function ReverseBoldItalicFadedFour(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FourColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedFour = Text$
End Function

Function ReverseBoldItalicFadedFive(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FiveColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedFive = Text$
End Function

Function ReverseBoldItalicFadedSeven(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SevenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedSeven = Text$
End Function

Function ReverseBoldItalicFadedEight(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = EightColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedEight = Text$
End Function

Function ReverseBoldItalicFadedNine(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = NineColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedNine = Text$
End Function

Function ReverseBoldItalicFadedTen(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedTen = Text$
End Function
Function ReverseBoldItalicFadedSix(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SixColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedSix = Text$
End Function
Function ReverseItalicFadedThree(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = ThreeColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedThree = Text$
End Function

Function ReverseItalicFadedTwo(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TwoColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedTwo = Text$
End Function

Function ReverseItalicFadedFour(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FourColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedFour = Text$
End Function

Function ReverseItalicFadedFive(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FiveColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedFive = Text$
End Function

Function ReverseItalicFadedSix(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SixColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedSix = Text$
End Function

Function ReverseItalicFadedSeven(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SevenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedSeven = Text$
End Function

Function ReverseItalicFadedEight(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = EightColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedEight = Text$
End Function

Function ReverseItalicFadedNine(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = NineColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedNine = Text$
End Function

Function ReverseItalicFadedTen(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedTen = Text$
End Function
Function TrimTime()
B$ = Left$(Time$, 5)
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(B$, 3) & " " & Ap$
End Function

Function TrimTime2()
B$ = Time$
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(B$, 5) & " " & Ap$
End Function

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub




Sub AOL40_Keyword(Keyword)
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Tool2%, "_AOL_Icon")
For GetIcon = 1 To 20
icon% = GetWindow(icon%, 2)
Next GetIcon
Call Pause(0.05)
Call ClickIcon(icon%)
Do: DoEvents
MDI% = FindChildByClass(AOLWindow(), "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
Edit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
Icon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And Edit% <> 0 And Icon2% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, Keyword)
Call TimeOut(0.05)
Call ClickIcon(Icon2%)
Call ClickIcon(Icon2%)
End Sub
Public Sub AOLScrollList(lst As ListBox)
For X% = 0 To List1.ListCount - 1
SendChat ("Scrolling Name [" & X% & "]" & List1.List(X%))
TimeOut (0.75)
Next X%
End Sub
Sub SendIM(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call AOL4_KEYWORD("aol://9293:")

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

Function LongScroll(Text)
Dim T$
Dim a%
a% = 1900 / Len(Text)
T$ = " <pre="
For X% = 1 To a%
T$ = T$ + Text
DoEvents
Next X%
LongScroll = T$
SendChat (T$)
End Function


Sub Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
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







Sub CtrlAltDeleteOff(TrueOrFalse As Boolean)
Dim OnOff As Integer
Dim Test As Boolean
OnOff = SystemParametersInfo(SPI_SCREENSAVERRUNNING, TrueOrFalse, Test, 0)
End Sub
Function BackwardsText(Text)
Dim T$
Dim Char%
T$ = ""
Char% = Len(Text)
Do While Char% <> 0
T$ = T$ + Mid$(Text, Char%, 1)
Char% = Char% - 1
Loop
BackwardsText = T$
End Function
Function EliteText(Text)
Dim Char%
Dim T$
Dim X$
Dim C%
T$ = ""
For Char% = 1 To Len(Text)
    X$ = ""
    If Mid$(Text, Char%, 1) = "a" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "â"
    If C% = 2 Then X$ = "å"
    If C% = 3 Then X$ = "ã"
    End If
    If Mid$(Text, Char%, 1) = "c" Then
    X$ = "ç"
    End If
    If Mid$(Text, Char%, 1) = "e" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "ê"
    If C% = 2 Then X$ = "ë"
    If C% = 3 Then X$ = "é"
    End If
    If Mid$(Text, Char%, 1) = "f" Then
    X$ = "ƒ"
    End If
    If Mid$(Text, Char%, 1) = "i" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "ï"
    If C% = 2 Then X$ = "î"
    If C% = 3 Then X$ = "í"
    End If
    If Mid$(Text, Char%, 1) = "j" Then
    X$ = "¡"
    End If
    If Mid$(Text, Char%, 1) = "l" Then
    X$ = "|"
    End If
    If Mid$(Text, Char%, 1) = "n" Then
    X$ = "ñ"
    End If
    If Mid$(Text, Char%, 1) = "o" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "ô"
    If C% = 2 Then X$ = "ð"
    If C% = 3 Then X$ = "õ"
    End If
    If Mid$(Text, Char%, 1) = "p" Then
    X$ = "þ"
    End If
    If Mid$(Text, Char%, 1) = "s" Then
    X$ = "š"
    End If
    If Mid$(Text, Char%, 1) = "t" Then
    X$ = "†"
    End If
    If Mid$(Text, Char%, 1) = "u" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "û"
    If C% = 2 Then X$ = "ü"
    If C% = 3 Then X$ = "ú"
    End If
    If Mid$(Text, Char%, 1) = "x" Then
    X$ = "×"
    End If
    If Mid$(Text, Char%, 1) = "y" Then
    C% = Int(Rnd * 2 + 1)
    If C% = 1 Then X$ = "ÿ"
    If C% = 2 Then X$ = "ý"
    End If
    If Mid$(Text, Char%, 1) = "A" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "Ä"
    If C% = 2 Then X$ = "Å"
    If C% = 3 Then X$ = "Ã"
    End If
    If Mid$(Text, Char%, 1) = "B" Then
    X$ = "ß"
    End If
    If Mid$(Text, Char%, 1) = "C" Then
    C% = Int(Rnd * 2 + 1)
    If C% = 1 Then X$ = "Ç"
    If C% = 2 Then X$ = "©"
    End If
    If Mid$(Text, Char%, 1) = "D" Then
    X$ = "Ð"
    End If
    If Mid$(Text, Char%, 1) = "E" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "Ê"
    If C% = 2 Then X$ = "Ë"
    If C% = 3 Then X$ = "É"
    End If
    If Mid$(Text, Char%, 1) = "I" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "Î"
    If C% = 2 Then X$ = "Ï"
    If C% = 3 Then X$ = "Í"
    End If
    If Mid$(Text, Char%, 1) = "L" Then
    X$ = "£"
    End If
    If Mid$(Text, Char%, 1) = "N" Then
    X$ = "Ñ"
    End If
    If Mid$(Text, Char%, 1) = "O" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "Ö"
    If C% = 2 Then X$ = "Ô"
    If C% = 3 Then X$ = "Õ"
    End If
    If Mid$(Text, Char%, 1) = "P" Then
    X$ = "Þ"
    End If
    If Mid$(Text, Char%, 1) = "R" Then
    X$ = "®"
    End If
    If Mid$(Text, Char%, 1) = "S" Then
    C% = Int(Rnd * 3 + 1)
    If C% = 1 Then X$ = "§"
    If C% = 2 Then X$ = "Š"
    End If
    If Mid$(Text, Char%, 1) = "U" Then
    C% = Int(Rnd * 2 + 1)
    If C% = 1 Then X$ = "Ü"
    If C% = 2 Then X$ = "Û"
    End If
    If Mid$(Text, Char%, 1) = "Y" Then
    X$ = "¥"
    End If
    If X$ = "" Then
    X$ = Mid$(Text, Char%, 1)
    End If
    T$ = T$ + X$
Next Char%
EliteText = T$
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
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


Sub IMIgnore(thelist As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Sub IMBuddy(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If Buddy% = 0 Then
    AOL40_Keyword ("BuddyView")
    Do: DoEvents
    Loop Until Buddy% <> 0
End If

AOIcon% = FindChildByClass(Buddy%, "_AOL_Icon")

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
Sub IMKeyword(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call AOL40_Keyword("aol://9293:")

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
Function FindChildByTitle(Parent%, TitleToFind$) As Integer
'Finds a child by the window title.
Dim ChildHandle%
Dim TitleOfChild$
Dim LengthOfTitleOfChild%
ChildHandle% = GetWindow(Parent%, GW_CHILD)
While ChildHandle%
TitleOfChild$ = String(200, 0)
LengthOfTitleOfChild% = GetWindowText(ChildHandle%, TitleOfChild$, 199)
TitleOfChild$ = Left$(TitleOfChild$, LengthOfTitleOfChild%)
If InStr(UCase$(TitleOfChild$), UCase$(TitleToFind$)) Then GoTo ExitWhile
ChildHandle% = GetWindow(ChildHandle%, GW_HWNDNEXT)
Wend
ChildHandle% = 0
ExitWhile:
FindChildByTitle = ChildHandle%
End Function

Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function

Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
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
GetListIndex = -2
End Function
Sub SendChat(Chat)
Room% = AOL4_FindRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub SendCharNum(win, chars)
E = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub

Sub Sendtext1(txt)
Room% = AOL4_FindRoom()
If Room% Then
   hChatEdit% = Find2ndChildByClass(Room%, "RICHCNTL")
   ret = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, txt)
   buto% = FindChildByClass(Room%, "_AOL_Icon")
    buto% = GetWindow(buto%, GW_HWNDNEXT)
    buto% = GetWindow(buto%, GW_HWNDNEXT)
    buto% = GetWindow(buto%, GW_HWNDNEXT)
    buto% = GetWindow(buto%, GW_HWNDNEXT)
    buto% = GetWindow(buto%, GW_HWNDNEXT)
    AOL4_Icon (buto%)
End If
End Sub

Sub CenterForm(Form As Form)
Form.Top = (Screen.Height * 0.85) / 2 - Form.Height / 2
Form.Left = Screen.Width / 2 - Form.Width / 2
End Sub
Sub DragForm(Form As Form, Button As Integer, Shift As Integer, X As Single, Y As Single)
'to use this do the following...
'put this in the MouseDown Proc of the form or object
'Call DragForm(Me, Button, Shift, x, Y)
If Button <> 1 Then Exit Sub
ReleaseCapture
G% = SendMessage(Form.hWnd, WM_NCLBUTTONDOWN, 2, 0)
End Sub


Sub ShutFuckingWindowsDown()
Dim MsgRes As Long
MsgRes = MsgBox("Do you really want to Shut Down Windows 95?", vbYesNo Or vbQuestion)
If MsgRes = vbNo Then Exit Sub
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End
End Sub

Function SpacedText(Text)
Dim Char%
Dim T$
T$ = ""
For Char% = 1 To Len(Text)
    T$ = T$ + Mid$(Text, Char%, 1) + " "
Next Char%
SpacedText = T$
End Function

Function ThirteenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Red12, Green12, Blue12, Red13, Green13, Blue13, Wavy As Boolean)

If Len(Text) < 7 Then
    Do Until Len(Text) = 7
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 6 <> 0 Then
    Do Until Len(Text) Mod 6 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 6
Thirteen1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Thirteen2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Thirteen3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Thirteen4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Thirteen5 = ThreeColors(Mid(Text, P + P + P + P + 1, P), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, False)
Thirteen6 = ThreeColors(Right(Text, P), Red11, Green11, Blue11, Red12, Green12, Blue12, Red13, Green13, Blue13, False)
ThirteenColors = Thirteen1 + Thirteen2 + Thirteen3 + Thirteen4 + Thirteen5 + Thirteen6
If Wavy = True Then
For X% = 1 To Len(ThirteenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(ThirteenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(ThirteenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(ThirteenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(ThirteenColors, X% + 63, 21)
Next X%
ThirteenColors = TextX$
End If
End Function
Function TwelveColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Red12, Green12, Blue12, Wavy As Boolean)

If Len(Text) < 7 Then
    Do Until Len(Text) = 7
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 6 <> 0 Then
    Do Until Len(Text) Mod 6 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 6
Twelve1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Twelve2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Twelve3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Twelve4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Twelve5 = ThreeColors(Mid(Text, P + P + P + P + 1, P), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, False)
Twelve6 = TwoColors(Right(Text, P), Red11, Green11, Blue11, Red12, Green12, Blue12, False)
TwelveColors = Twelve1 + Twelve2 + Twelve3 + Twelve4 + Twelve5 + Twelve6
If Wavy = True Then
For X% = 1 To Len(TwelveColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(TwelveColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(TwelveColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(TwelveColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(TwelveColors, X% + 63, 21)
Next X%
TwelveColors = TextX$
End If
End Function

Function ElevenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Wavy As Boolean)

If Len(Text) < 6 Then
    Do Until Len(Text) = 6
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 5 <> 0 Then
    Do Until Len(Text) Mod 5 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 5
Eleven1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Eleven2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Eleven3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Eleven4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Eleven5 = ThreeColors(Right(Text, P), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, False)
ElevenColors = Eleven1 + Eleven2 + Eleven3 + Eleven4 + Eleven5
If Wavy = True Then
For X% = 1 To Len(ElevenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(ElevenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(ElevenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(ElevenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(ElevenColors, X% + 63, 21)
Next X%
ElevenColors = TextX$
End If
End Function

Function TenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)

If Len(Text) < 6 Then
    Do Until Len(Text) = 6
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 5 <> 0 Then
    Do Until Len(Text) Mod 5 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 5
Ten1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Ten2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Ten3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Ten4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Ten5 = TwoColors(Right(Text, P), Red9, Green9, Blue9, Red10, Green10, Blue10, False)
TenColors = Ten1 + Ten2 + Ten3 + Ten4 + Ten5
If Wavy = True Then
For X% = 1 To Len(TenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(TenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(TenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(TenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(TenColors, X% + 63, 21)
Next X%
TenColors = TextX$
End If
End Function
Function NineColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 8 <> 0 Then
    Do Until Len(Text) Mod 8 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 8
Nine1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Nine2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Nine3 = TwoColors(Mid(Text, P + P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Nine4 = TwoColors(Mid(Text, P + P + P + 1, P), Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Nine5 = TwoColors(Mid(Text, P + P + P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Nine6 = TwoColors(Mid(Text, P + P + P + P + P + 1, P), Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Nine7 = TwoColors(Mid(Text, P + P + P + P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, False)
Nine8 = TwoColors(Right(Text, P), Red8, Green8, Blue8, Red9, Green9, Blue9, False)
NineColors = Nine1 + Nine2 + Nine3 + Nine4 + Nine5 + Nine6 + Nine7 + Nine8
If Wavy = True Then
For X% = 1 To Len(NineColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(NineColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(NineColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(NineColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(NineColors, X% + 63, 21)
Next X%
NineColors = TextX$
End If
End Function
Function SeeFade(R1, G1, B1, R2, B2, G2, pctre)
'i have often found that this will only work once,
'so for this reason i recomend u copy and paste
'the code into the Paint Proc of a picture box.
'This only shows 2 colors faded at a time.

On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double

Static SplitNum(3) As Double
Static DivideNum(3) As Double

Dim FadeW As Integer
Dim Loo As Integer
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100
FadeW = pctre.Width / 100
For Loo = 0 To 100

pctre.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)

Next Loo

End Function
Function EightColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 7 <> 0 Then
    Do Until Len(Text) Mod 7 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 7
Eight1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Eight2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Eight3 = TwoColors(Mid(Text, P + P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Eight4 = TwoColors(Mid(Text, P + P + P + 1, P), Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Eight5 = TwoColors(Mid(Text, P + P + P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Eight6 = TwoColors(Mid(Text, P + P + P + P + P + 1, P), Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Eight7 = TwoColors(Right(Text, P), Red7, Green7, Blue7, Red8, Green8, Blue8, False)
EightColors = Eight1 + Eight2 + Eight3 + Eight4 + Eight5 + Eight6 + Eight7
If Wavy = True Then
For X% = 1 To Len(EightColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(EightColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(EightColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(EightColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(EightColors, X% + 63, 21)
Next X%
EightColors = TextX$
End If
End Function

Function SevenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 6 <> 0 Then
    Do Until Len(Text) Mod 6 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 6
Seven1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Seven2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Seven3 = TwoColors(Mid(Text, P + P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Seven4 = TwoColors(Mid(Text, P + P + P + 1, P), Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Seven5 = TwoColors(Mid(Text, P + P + P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Seven6 = TwoColors(Right(Text, P), Red6, Green6, Blue6, Red7, Green7, Blue7, False)
SevenColors = Seven1 + Seven2 + Seven3 + Seven4 + Seven5 + Seven6
If Wavy = True Then
For X% = 1 To Len(SevenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(SevenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(SevenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(SevenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(SevenColors, X% + 63, 21)
Next X%
SevenColors = TextX$
End If
End Function
Function SixColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 5 <> 0 Then
    Do Until Len(Text) Mod 5 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 5
Six1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Six2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Six3 = TwoColors(Mid(Text, P + P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Six4 = TwoColors(Mid(Text, P + P + P + 1, P), Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Six5 = TwoColors(Right(Text, P), Red5, Green5, Blue5, Red6, Green6, Blue6, False)
SixColors = Six1 + Six2 + Six3 + Six4 + Six5
If Wavy = True Then
For X% = 1 To Len(SixColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(SixColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(SixColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(SixColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(SixColors, X% + 63, 21)
Next X%
SixColors = TextX$
End If
End Function

Function FourColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 3 <> 0 Then
    Do Until Len(Text) Mod 3 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 3
Four1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Four2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Four3 = TwoColors(Right(Text, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
FourColors = Four1 + Four2 + Four3
If Wavy = True Then
For X% = 1 To Len(FourColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(FourColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(FourColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(FourColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(FourColors, X% + 63, 21)
Next X%
FourColors = TextX$
End If
End Function


Function FiveColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)

If Len(Text) < 3 Then
    Do Until Len(Text) = 3
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 2 <> 0 Then
    Do Until Len(Text) Mod 2 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 2
Five1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Five2 = ThreeColors(Right(Text, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
FiveColors = Five1 + Five2
If Wavy = True Then
For X% = 1 To Len(FiveColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(FiveColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(FiveColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(FiveColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(FiveColors, X% + 63, 21)
Next X%
FiveColors = TextX$
End If
End Function



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
















'Variable color fade functions begin here


Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
    C1BAK = C1
    C2BAK = C2
    C3BAK = C3
    C4BAK = C4
    C = 0
    o = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For X = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(Text) * X) + Red1
        VAL2 = (BVAL2 / Len(Text) * X) + Green1
        VAL3 = (BVAL3 / Len(Text) * X) + Blue1
        
        C1 = RGB2HEX(VAL1, VAL2, VAL3)
        C2 = RGB2HEX(VAL1, VAL2, VAL3)
        C3 = RGB2HEX(VAL1, VAL2, VAL3)
        C4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then C = 1: msg = msg & "<FONT COLOR=#" + C1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If C <> 1 Then
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        
        If Wavy = True Then
            If o2 = 1 Then msg = msg + "<SUB>"
            If o2 = 3 Then msg = msg + "<SUP>"
            msg = msg + Mid$(Text, X, 1)
            If o2 = 1 Then msg = msg + "</SUB>"
            If o2 = 3 Then msg = msg + "</SUP>"
            If Q2 = 2 Then
                Q = 1
                Q2 = 1
                If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
                If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
                If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
                If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
            End If
        ElseIf Wavy = False Then
            msg = msg + Mid$(Text, X, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        End If
nc:     Next X
    C1 = C1BAK
    C2 = C2BAK
    C3 = C3BAK
    C4 = C4BAK
    TwoColors = msg
End Function

Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)

    D = Len(Text)
        If D = 0 Then GoTo TheEnd
        If D = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If D = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If D = X Then GoTo Odds
    Next X
Evens:
    C = D \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C)
    GoTo TheEnd
Odds:
    C = D \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If Wavy = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If Wavy = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If Wavy = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If Wavy = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    msg = FadeA + FadeB
    ThreeColors = msg
End Function

Function RGB2HEX(R, G, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
        For xx& = 1 To 2
            Divide = Color& / 16
            Answer& = Int(Divide)
            Remainder& = (10000 * (Divide - Answer&)) / 625
            If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
            If Remainder& = 10 Then Configuring$ = "A" + Configuring$
            If Remainder& = 11 Then Configuring$ = "B" + Configuring$
            If Remainder& = 12 Then Configuring$ = "C" + Configuring$
            If Remainder& = 13 Then Configuring$ = "D" + Configuring$
            If Remainder& = 14 Then Configuring$ = "E" + Configuring$
            If Remainder& = 15 Then Configuring$ = "F" + Configuring$
            Color& = Answer&
        Next xx&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
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
Sub UnHideWindow(hWnd)
'This will Unhide the hwnd window
un = ShowWindow(hWnd, SW_SHOW)
End Sub



Function UntilWindowClass(parentw, childhand)
GoBack:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
Wend
GoTo GoBack
FindClassLike = 0
Greed:
Room% = firs%
UntilWindowClass = Room%
End Function

Function UntilWindowTitle(parentw, childhand)
GoBac:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(parentw, GW_CHILD)
While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo Greed
Wend
GoTo GoBac
FindWindowLike = 0
Greed:
Room% = firs%
UntilWindowTitle = Room%
End Function


Sub WaitWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
topmdi% = GetWindow(MDI%, 5)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
topmdi2% = GetWindow(MDI%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop
End Sub


Sub WriteSend(Text)
Dim a%
Dim X%
Dim T$
a% = Len(Text)
SendChat (" ")
Pause (0.7)
For X% = 1 To a%
T$ = T$ + Mid(Text, X%, 1)
SendChat (T$)
Pause (0.7)
Next X%
X% = a%
Do While X% <> 0
X% = X% - 1
T$ = Left(T$, X%)
SendChat (T$)
Pause (0.7)
Loop
SendChat (" ")
End Sub

Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
    BlackRedBlack = msg
    SendChat (msg)
End Function
