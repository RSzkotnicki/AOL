Attribute VB_Name = "Laby"
Global string2, string1, im, room1, timers, timers2, tuierl, stoper, ti, color, olor, names$(200), Scores(200), TotalScore(200), nscores, sn(200), h
Global can As Boolean, blNcancelsave As Boolean, new1 As Boolean
Global charcount, num(150), linecount, speed
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function FillRect Lib "user" (ByVal hDC As Integer, lpRect As Rect, ByVal hBrush As Integer) As Integer
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function getnextwindow Lib "user32" Alias "GetNextWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long

Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nByteS As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function Movewindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (Object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function Sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function Drawmenubar Lib "user32" Alias "DrawMenuBar" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function Getparent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Declare Function MenuItemFromPoint Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Declare Function Getmenu Lib "user32" Alias "GetMenu" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function Gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function SndPlaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Global Const SPI_SCREENSAVERRUNNING = 97
Global Const WM_CHAR = &H102
Global Const WM_SETTEXT = &HC
Global Const WM_USER = &H400
Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101
Global Const WM_Close = &H10
Global Const WM_COMMAND = &H111
Global Const WM_CLEAR = &H303
Global Const WM_DESTROY = &H2
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
Global Const WM_LBUTTONDBLCLK = &H203

Global Const BM_GETCHECK = &HF0
Global Const BM_GETSTATE = &HF2
Global Const BM_SETCHECK = &HF1
Global Const BM_SETSTATE = &HF3

Global Const LB_GETITEMDATA = &H199
Global Const LB_ADDSTRING = &H180
Global Const LB_DELETESTRING = &H182
Global Const LB_FINDSTRING = &H18F
Global Const LB_FINDSTRINGEXACT = &H1A2
Global Const LB_GETCURSEL = &H188
Global Const LB_GETTEXT = &H189
Global Const LB_GETTEXTLEN = &H18A
Global Const LB_SELECTSTRING = &H18C
Global Const LB_SETCOUNT = &H1A7
Global Const LB_SETCURSEL = &H186
Global Const LB_SETSEL = &H185
Global Const LB_INSERTSTRING = &H181

Global Const VK_HOME = &H24
Global Const VK_RIGHT = &H27
Global Const VK_CONTROL = &H11
Global Const VK_DELETE = &H2E
Global Const VK_DOWN = &H28
Global Const VK_LEFT = &H25
Global Const VK_RETURN = &HD
Global Const VK_SPACE = &H20
Global Const VK_TAB = &H9

Global Const HWND_TOP = 0
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1

Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Global Const GW_CHILD = 5
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_MAX = 5
Global Const GW_OWNER = 4

Global Const SW_MAXIMIZE = 3
Global Const SW_MINIMIZE = 6
Global Const SW_Hide = 0
Global Const SW_RESTORE = 9
Global Const SW_Show = 5
Global Const SW_SHOWDEFAULT = 10
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_SHOWNORMAL = 1

Global Const MF_APPEND = &H100&
Global Const MF_DELETE = &H200&
Global Const MF_CHANGE = &H80&
Global Const MF_ENABLED = &H0&
Global Const MF_DISABLED = &H2&
Global Const MF_REMOVE = &H1000&
Global Const MF_POPUP = &H10&
Global Const MF_String = &H0&
Global Const MF_UNCHECKED = &H0&
Global Const MF_CHECKED = &H8&
Global Const MF_GRAYED = &H1&
Global Const MF_BYPOSITION = &H400&
Global Const MF_BYCOMMAND = &H0&

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Global Const GWW_HINSTANCE = (-6)
Global Const GWW_ID = (-12)

Global Const GWL_STYLE = (-16)

Global Const PROCESS_VM_READ = &H10

Global Const STANDARD_RIGHTS_REQUIRED = &HF0000

Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   y As Long
End Type


Private Type MailInfo
    MailDate As Date
    sn As String * 10
    Title As String * 53
End Type
Private StopNow, LoadMail As Boolean
Dim RetStr As String * 71
Dim Mails() As MailInfo
Sub AOL4_Read1Mail()
'This will read the very first mail in the User's box
MailBox% = findchildbytitle(AOL4_MDI(), AOL4_GetUser + "'s Online Mailbox")
E = findchildbyclass(MailBox%, "_AOL_Icon")
AOL4_Icon (E)
End Sub

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
Menuitem% = SubCount%
GoTo MatchString
End If
Next GetString
Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, Menuitem%, 0)
End Sub
Function AOL4_GetUser()
'This will tell what SN is usin da Prog
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
mdi% = findchildbyclass(AOL%, "MDIClient")
welcome% = findchildbytitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOL4_GetUser = user
End Function

Sub AOL4_Invite(Person)
'This will send an Invite to a person
freeprocess
On Error GoTo ErrHandler
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
bud% = findchildbytitle(mdi%, "Buddy List Window")
E = findchildbyclass(bud%, "_AOL_Icon")
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
AOL4_Icon (E)
pause (1#)
chat% = findchildbytitle(mdi%, "Buddy Chat")
aoledit% = findchildbyclass(chat%, "_AOL_Edit")
If chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = findchildbyclass(chat%, "_AOL_Icon")
AOL4_Icon (de)
Killit% = findchildbytitle(mdi%, "Invitation From:")
AOL4_KillWin (Killit%)
freeprocess
ErrHandler:
Exit Sub
End Sub
Function AOL4_GetChatSN()
'This getz the last chat line without a SN in front of it
heh$ = (AOL4_LastChatLine)
heh$ = LCase(heh$)
nwe$ = Mid(heh$, InStr(heh$, ":") + 2)
AOL4_GetChatSN = nwe$
End Function

Sub AOL4_Tool(wch%)
'You can only use  AOL4_Tool(0) or AOL4_Tool(2)
'0 = Open mail box    2 = Write Mail
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = findchildbyclass(AOL%, "AOL Toolbar")
Tool2% = findchildbyclass(tool%, "_AOL_Toolbar")
ico3n% = findchildbyclass(Tool2%, "_AOL_Icon")
Icon2% = getwindow(ico3n%, wch%)
'Click% = SendMessageByNum(icon2%, WM_LBUTTONDOWN, 0&, 0&)
'Click% = SendMessageByNum(icon2%, WM_LBUTTONUP, 0&, 0&)
End Sub

Function AOL4_Saying()
'This will generate a random saying
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: aol4_chatsend (Text1)
Case 2: aol4_chatsend (Text1)
Case 3: aol4_chatsend (Text1)
Case 4: aol4_chatsend (Text1)
Case 5: aol4_chatsend (Text1)
Case 6: aol4_chatsend (Text1)
Case 7: aol4_chatsend (Text1)
Case Else: aol4_chatsend (Text1)
End Select
End Function
Sub AOL4_Elite_Talker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If letter$ = "`" Then Leet$ = "´"
    If letter$ = "!" Then Leet$ = "¡"
    If letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
 aol4_chatsend (Made$)
End Sub

Sub AOL4_MemberProfile(name As String)
'This getz a members profile
End Sub

Sub AOL4_locateMember(name As String)
'This will Locate a member....if online
Call AOL4_Keyword("aol://3548:" + name)
End Sub


Sub AOL4_MassIMer(Person, Message)
Dim im
'This openz an IM and fillz it out with Person and Message
AOL4_Keyword ("aol://9293:" & Person)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
im = findchildbytitle(mdi%, "Send Instant Message")
aolrich% = findchildbyclass(im, "RICHCNTL")
imsend% = findchildbyclass(im, "_AOL_Icon")
Loop Until (im <> 0 And aolrich% <> 0 And imsend% <> 0)
Call sendmessagebystring(aolrich%, WM_SETTEXT, 0, Message)
For sends = 1 To 9
imsend% = getwindow(imsend%, GW_HWNDNEXT)
Next sends
Ao4Click imsend%

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
im = findchildbytitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_Close, 0, 0): closer2 = SendMessage(im, WM_Close, 0, 0): Exit Do
If im = 0 Then Exit Do
Loop

End Sub


Sub SizeFormToWindow(frm As Form, win%)
'This will make a frm the size of a win
'ex: Call SizeFormToWindow(form1, IM%)
Dim wndRect As Rect, lRet As Long
lRet = GetWindowRect(win%, wndRect)
With frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub
Function AOL4_SNfromIM()
Dim im
'This will return the Screen Name from an IM
im = findchildbytitle(AOL4_MDI(), ">Instant Message From:")
If im Then GoTo greed
im = findchildbytitle(AOL4_MDI(), "  Instant Message From:")
If im Then GoTo greed
Exit Function
greed:
heh$ = GetCaption(im)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
AOL4_SNfromIM = Naw$
End Function

Public Sub CenterFormTop(frm As Form)
'This will center the form in the top center
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub

Sub AOL4_Ctrl_Alt_Del(Index)
        Dim ret As Integer
        Dim pOld As Boolean
Select Case Index
Case 1:
'Disables Ctrl+Alt+Del
        ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
Case 2:
'Enables Ctrl+Alt+Del
        ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Select
End Sub
Sub waitforok()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
okb = findchildbytitle(okw, "OK")
DoEvents
Loop Until okb <> 0
Do
okw = FindWindow("#32770", "America Online")
    okb = findchildbytitle(okw, "OK")
    okd = Sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = Sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)
DoEvents
Loop Until okw = 0

End Sub


Sub HideWindow(hWnd)
'This hides the hwnd window
hi = showwindow(hWnd, SW_Hide)
End Sub


Sub AOL4_UNHide()
'This will Un Hide the AOL window after hidden
a = showwindow(AOL4_Window(), 5)
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
Sub AOL4_Hide()
'This hidez AOL
a = showwindow(AOL4_Window(), 0)
End Sub
Sub UnHideWindow(hWnd)
'This will Unhide the hwnd window
un = showwindow(hWnd, SW_Show)
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



Sub playwav(File)
'This plaz a wav file
'example:  Playwav("filename.wav")
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = SndPlaysound(SoundName$, wFlags%)

End Sub

Public Sub AOL4_RoomSNs(ListBOXES As ListBox)
'This adds AOL's room list to a VB listbox
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = AOL4_FindRoom()
aolhandle = findchildbyclass(room, "_AOL_Listbox")
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
ListBOXES.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
lst.RemoveItem ListBOXES.ListCount - 1
'i = GetListIndex(Listboxes, AOL4_GetUser())
'If i <> -2 Then Listboxes.RemoveItem i
End Sub
Public Function AOL4_GetList(Index As Long, Buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = AOL4_FindRoom()
aolhandle = findchildbyclass(room, "_AOL_Listbox")
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
Function AddListToString(thelist As ListBox)
'Makes a list into a string a "comma" after each word
For DoList = 0 To thelist.ListCount - 1
AddListToString = AddListToString & thelist.List(DoList) & ","
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 1)
End Function


Sub Timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub
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


Function AOL4_ClickList(hWnd)
clicklist% = Sendmessagebynum(hWnd, &H203, 0, 0&)
End Function


Function AOL4_CountMail()
'This countz the mails in your box
'ex: Msgbox "There are a total of " & AOL4_CountMail & " mailz in your box..."
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = findchildbyclass(hWndAOL, "MDIClient")
hWndMail = findchildbytitle(hWndAOLClient, "*OnLine Mailbox")
hWndTabControl = findchildbyclass(hWndMail, "_AOL_TabControl")
hWndTabPage = findchildbyclass(hWndTabControl, "_AOL_TabPage")
hWndMailLB = findchildbyclass(hWndTabPage, "_AOL_Tree")
AOL4_CountMail = Sendmessagebynum(hWndMailLB, LB_GETCOUNT, 0&, 0&)
End Function

Sub AOL4_OpenMail()
'This openz up the Users Mail box
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = findchildbyclass(AOL%, "AOL Toolbar")
Tool2% = findchildbyclass(tool%, "_AOL_Toolbar")
ico3n% = findchildbyclass(Tool2%, "_AOL_Icon")
Icon2% = getwindow(ico3n%, 0)
Icon2% = Sendmessagebynum(Icon2%, WM_LBUTTONDOWN, 0&, 0&)
Icon2% = Sendmessagebynum(Icon2%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub AOLRunMenuByString(stringer As String)

End Sub



Function findchildbytitle(parentw, childhand)
'This finds a child window by it'z title
firs% = getwindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo greed
firs% = getwindow(parentw, GW_CHILD)
While firs%
firs% = getwindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo greed
firs% = getwindow(firs%, GW_HWNDNEXT)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo greed
Wend
findchildbytitle = 0
greed:
room% = firs%
findchildbytitle = room%
End Function
Function findchildbyclass(parentw, childhand)
'This findz child by it'z class
firs% = getwindow(parentw, GW_MAX)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
firs% = getwindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
While firs%
firss% = getwindow(parentw, GW_MAX)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
firs% = getwindow(firs%, GW_HWNDNEXT)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
Wend
findchildbyclass = 0
greed:
room% = firs%
findchildbyclass = room%
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

Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Function LineFromText(Text, theline)
'This returnz a line from text
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

Sub ListToList(Source, destination)
'Copies 1 list to another
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = sendmessagebystring(Source, LB_GETTEXT, Adding, Buffer$)
addstrings% = sendmessagebystring(destination, LB_ADDSTRING, 0, Buffer$)
Next Adding
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
End Function

Sub ParentChange(Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub

Function RandomNumber(finished)
'Returnz a random number
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
Function ReverseText(Text)
For Words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, Words, 1)
Next Words
End Function

Sub AOLRunTool(tool)
'This will click on one of the toolbar Iconz
End Sub

Function AOLGetTopWindow()
'This getz the window ontop of all others
AOLGetTopWindow = Gettopwindow(AOL4_MDI())
End Function

Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If
For Replace = 1 To Len(Text)
thechar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & thechar$
If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace
ReplaceText = thechars$
End Function
Sub AOL4_Reset_Pref()

End Sub

Function StayOnline()
'Clickz that damn 45 min window when it popz up
hwndz% = FindWindow("_AOL_Palette", "America Online Timer")
childhwnd% = findchildbytitle(hwndz%, "OK")
AOL4_Button (childhwnd%)

End Function



Sub MiniWindow(hWnd)
'This minimizes the hwnd window
mi = showwindow(hWnd, SW_MINIMIZE)
End Sub

Sub MaxWindow(hWnd)
'Maximizes the hwnd window
ma = showwindow(hWnd, SW_MAXIMIZE)
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
Function AOL4_MDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL4_MDI = findchildbyclass(AOL%, "MDIClient")
End Function

Function UntilWindowClass(parentw, childhand)
GoBack:
DoEvents
firs% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
firs% = getwindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
While firs%
firss% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
firs% = getwindow(firs%, GW_HWNDNEXT)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo greed
Wend
GoTo GoBack
FindClassLike = 0
greed:
room% = firs%
UntilWindowClass = room%
End Function
Function UntilWindowTitle(parentw, childhand)
GoBac:
DoEvents
firs% = getwindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo greed
firs% = getwindow(parentw, GW_CHILD)
While firs%
firss% = getwindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) Then GoTo greed
firs% = getwindow(firs%, GW_HWNDNEXT)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo greed
Wend
GoTo GoBac
FindWindowLike = 0
greed:
room% = firs%
UntilWindowTitle = room%
End Function
Public Sub Centerform(frmForm As Form)
'This will center a form in very center of screen
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub

Public Sub CenterCorner(frmForm As Form)
'This centerz a form in the top right of the screen
   With frmForm
      .Left = (Screen.Width - .Width) / 1
      .Top = (Screen.Height - .Height) / 2000
   End With
End Sub
Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = getwindow(hWnd, GW_CHILD)
   

While hChild
hChild = getwindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Public Sub AOL4_Button(but%)
'Clicks an _AOL_Button
ClickIcon% = Sendmessagebynum(but%, WM_KEYDOWN, VK_SPACE, 0)
ClickIcon% = Sendmessagebynum(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Public Sub AOL4_Scroll(Text)
For i = 1 To 4
    Call aol4_chatsend(Text + String(116 - Len(Text), Chr(9)) + Text)
    pause (0.01)
Next i
End Sub
Sub AOL4_IMOff()
'Turnz IMz Off
Call AOL4_instantmessage("$IM_OFF", "sdfd")
End Sub

'Function SNFromLastChatLine()
'Chattext$ = LastChatLineWithSN
'ChatTrim$ = Left$(Chattext$, 11)
'For z = 1 To 11
'    If Mid$(ChatTrim$, z, 1) = ":" Then
'        sn = Left$(ChatTrim$, z - 1)
'    End If
'Next z
'SNFromLastChatLine = sn
'End Function

'Function AOL4_WavColorO(Text1 As String)
'G$ = Text1
'a = Len(G$)
'For w = 1 To a Step 7
'    r$ = Mid$(G$, w, 1)
'    U$ = Mid$(G$, w + 1, 1)
'    s$ = Mid$(G$, w + 2, 1)
'    t$ = Mid$(G$, w + 3, 1)
'    z$ = Mid$(G$, w + 4, 1)
'    Y$ = Mid$(G$, w + 5, 1)
'    c$ = Mid$(G$, w + 6, 1)
'    V$ = Mid$(G$, w + 7, 1)
'    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FFC800" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#FFAF00" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF9600" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "FF7D00" & Chr$(34) & ">" & t$ & "<FONT COLOR=" & Chr$(34) & "#FF6400" & Chr$(34) & "><sup>" & z$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#FF6400" & Chr$(34) & ">" & Y$ & "<FONT COLOR=" & Chr$(34) & "#FF4B00" & Chr$(34) & "><sub>" & c$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "FF3200>" & Chr$(34) & V
'Next w
'aol4_chatsend (p$)
'End Function
Sub AOL4_IM_AutoAnswer(Message)
Dim im
'Res'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")

im = findchildbytitle(mdi%, ">Instant Message From:")
If im Then GoTo greed
im = findchildbytitle(mdi%, "  Instant Message From:")
If im Then GoTo greed
Exit Sub
greed:
E = findchildbyclass(im, "RICHCNTL")

E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
e2 = getwindow(E, GW_HWNDNEXT) 'Send Text
E = getwindow(e2, GW_HWNDNEXT) 'Send Button
Call sendmessagebystring(e2, WM_SETTEXT, 0, Message)
AOL4_Icon (E)
pause 2#

End Sub

Sub AOL4_IMsOn()
'Turnz IMz On
Call AOL4_instantmessage("$IM_ON", "WiZoRd's")
End Sub

Sub aol4_chatsend(txt)
'Sendz txt to a chat room
    room% = AOL4_FindRoom()
    If room% Then
        hChatEdit% = Find2ndChildByClass(room%, "RICHCNTL")
        ret = sendmessagebystring(hChatEdit%, WM_SETTEXT, 0, txt)
        ret = Sendmessagebynum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub

Sub AOL4_ChangeCaption(newcaption As String)
'This changes the "America  Online" Caption
Call AOL4_SetText(AOL4_Window(), newcaption)
End Sub
Function AOL4_FindRoom()
'Findz the chat room/setz focus on it
    AOL% = FindWindow("AOL Frame25", vbNullString)
    mdi% = findchildbyclass(AOL%, "MDIClient")
    firs% = getwindow(mdi%, 5)
    listers% = findchildbyclass(firs%, "RICHCNTL")
    listere% = findchildbyclass(firs%, "RICHCNTL")
    listerb% = findchildbyclass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (l <> 100)
            DoEvents
            firs% = getwindow(firs%, 2)
            listers% = findchildbyclass(firs%, "RICHCNTL")
            listere% = findchildbyclass(firs%, "RICHCNTL")
            listerb% = findchildbyclass(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
        AOL4_FindRoom = firs%
        Exit Function
    End If
    AOL4_FindRoom = 0
End Function

Function AOL4_GetChat()
'This gets all the txt from chat room
childs% = AOL4_FindRoom()
child = findchildbyclass(childs%, "RICHCNTL")
GetTrim = Sendmessagebynum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = sendmessagebystring(child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
AOL4_GetChat = theview$
End Function
Function AOL4_ClearChat()
'This gets all the txt from chat room
childs% = AOL4_FindRoom()
child = findchildbyclass(childs%, "RICHCNTL")
GetTrim = Sendmessagebynum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = sendmessagebystring(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
AOL4_ClearChat = theview$
End Function

Sub AOL4_Icon(icon%)
'Clickz on an AOL icon
icon% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
icon% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub


Function AOL4_IsOnline()
'This tellz if the User is Online
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
welcome% = findchildbytitle(mdi%, Text1)
If welcome% = 0 Then
MsgBox "This feature is meant for use while signed on.", 64, "Attention:"
AOL4_IsOnline = 0
Exit Function
End If
AOL4_IsOnline = 1
End Function

Function AOL4_SignedON()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
welcome% = findchildbytitle(mdi%, "Welcome, ")
If welcome% = 0 Then
AOL4_SignedON = 0
Exit Function
End If
AOL4_SignedON = 1
End Function

Sub AOL4_Keyword(txt)
'This goes to an AOL Keyword
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = findchildbyclass(AOL%, "AOL Toolbar")
    temp% = findchildbyclass(temp%, "_AOL_Toolbar")
    temp% = findchildbyclass(temp%, "_AOL_Combobox")
    KWBox% = findchildbyclass(temp%, "Edit")
    Call sendmessagebystring(KWBox%, WM_SETTEXT, 0, txt)
    Call Sendmessagebynum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call Sendmessagebynum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub





Sub AOL4_Mail(Person, subject, Message)
'This openz a mail and fills it out
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = findchildbyclass(AOL%, "AOL Toolbar")
Tool2% = findchildbyclass(tool%, "_AOL_Toolbar")
ico3n% = findchildbyclass(Tool2%, "_AOL_Icon")
Icon2% = getwindow(ico3n%, 2)
Ao4Click Icon2%
pause (4)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    mdi% = findchildbyclass(AOL%, "MDIClient")
    mail% = findchildbytitle(mdi%, "Write Mail")
    aoledit% = findchildbyclass(mail%, "_AOL_Edit")
    aolrich% = findchildbyclass(mail%, "RICHCNTL")
    subjt% = findchildbytitle(mail%, "Subject:")
    subjec% = getwindow(subjt%, 2)
        Call AOL4_SetText(aoledit%, Person)
        Call AOL4_SetText(subjec%, subject)
        Call AOL4_SetText(aolrich%, Message)
E = findchildbyclass(mail%, "_AOL_Icon")
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
E = getwindow(E, GW_HWNDNEXT)
Ao4Click (E)
'FIND A WAY TO CLICK mail has been sent window
End Sub

Function AOL4_RoomCount()
'Countz people in a chat room and returnz it
thechild% = AOL4_FindRoom()
lister% = findchildbyclass(thechild%, "_AOL_Listbox")
getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOL4_RoomCount = getcount
End Function

Sub AOL4_SetText(win, txt)
'This is usually used for an _AOL_Edit or RICHCNTL
thetext% = sendmessagebystring(win, WM_SETTEXT, 0, txt)
End Sub

Sub AOL4_signoff()
'This will sign the User off of AOL
AppActivate "America  Online"
SendKeys "%SS"
End Sub


Function AOL4_Version()
'This getz the version of AOL the User has
End Function

Function AOL4_Window()
'This findz the AOL window
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL4_Window = AOL%
End Function



Function GetCaption(hWnd)
'returns the caption of "hWnd" window
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

Function GetWindowDir()
'finds the window's directory
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function
Sub NotOnTop(the As Form)
'This makes the form not stayontop
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub pause(interval)
'Pause/waits for "interval" seconds
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub SendCharNum(win, chars)
E = Sendmessagebynum(win, WM_CHAR, chars, 0)
End Sub

Sub SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Sub

Sub AOL4_Set_Pref()

End Sub

Sub Stayontop(the As Form)
'This keepz a form on top of all the other windows
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub AOL4_RunMenu(menu1 As Integer, menu2 As Integer)
'This will run one of the drop down menu's like  Edit/Paste
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = Getmenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = Sendmessagebynum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub


Sub WaitWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
topmdi% = getwindow(mdi%, 5)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
topmdi2% = getwindow(mdi%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop
End Sub

Sub AOL4_KillWin(windo)
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = Sendmessagebynum(windo, WM_Close, 0, 0)
End Sub

Function freeprocess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Function Find2ndChildByClass(parentw, childhand)
    firs% = getwindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    firs% = getwindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    While firs%
        firs% = getwindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
        firs% = getwindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Wend
    Find2ndChildByClass = 0
Found:
    firs% = getwindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = getwindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = getwindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = getwindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = firs%
End Function
Function AOL4_MessageFromIM()
Dim im
'This gets the Message from an IM
im = findchildbytitle(AOL4_MDI(), ">Instant Message From:")
If im Then GoTo DAMN
im = findchildbytitle(AOL4_MDI(), "  Instant Message From:")
If im Then GoTo DAMN
Exit Function
DAMN:
imtext% = findchildbyclass(im, "RICHCNTL")
'IMmessage = AOL4_GetText(imtext%)
FUK$ = IMmessage
Naw$ = Mid(FUK$, InStr(FUK$, ":") + 2)
AOL4_MessageFromIM = Naw$
End Function

'Function AOL4_WavColorb(Text1 As String)
'G$ = Text1
'a = Len(G$)
'For w = 1 To a Step 7
'    r$ = Mid$(G$, w, 1)
'    U$ = Mid$(G$, w + 1, 1)
'    s$ = Mid$(G$, w + 2, 1)
'    t$ = Mid$(G$, w + 3, 1)
'    z$ = Mid$(G$, w + 4, 1)
'    Y$ = Mid$(G$, w + 5, 1)
'    c$ = Mid$(G$, w + 6, 1)
'    V$ = Mid$(G$, w + 7, 1)
'    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#000080" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#8470ff" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000c0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "0000ff" & Chr$(34) & ">" & t$ & "<FONT COLOR=" & Chr$(34) & "#1e90ff" & Chr$(34) & "><sup>" & z$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#00bfff" & Chr$(34) & ">" & Y$ & "<FONT COLOR=" & Chr$(34) & "#87cefa" & Chr$(34) & "><sub>" & c$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "000000>" & Chr$(34) & V
'Next w
'aol4_chatsend (p$)
'End Function

'Function AOL4_WavColorp(Text1 As String)
'G$ = Text1
'a = Len(G$)
'For w = 1 To a Step 7
'    r$ = Mid$(G$, w, 1)
'    U$ = Mid$(G$, w + 1, 1)
'    s$ = Mid$(G$, w + 2, 1)
'    t$ = Mid$(G$, w + 3, 1)
'    z$ = Mid$(G$, w + 4, 1)
'    Y$ = Mid$(G$, w + 5, 1)
'    c$ = Mid$(G$, w + 6, 1)
'    V$ = Mid$(G$, w + 7, 1)
'    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#feb6fe" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#fe90fe" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#fe34fe" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#fe34fe" & Chr$(34) & ">" & t$ & "<FONT COLOR=" & Chr$(34) & "#c200c2" & Chr$(34) & "><sup>" & z$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#c200c2" & Chr$(34) & ">" & Y$ & "<FONT COLOR=" & Chr$(34) & "#700070" & Chr$(34) & "><sub>" & c$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "500050>" & Chr$(34) & V
'Next w
'aol4_chatsend (p$)
'End Function
'Function aol4_wavcolorw(Text1 As String)
'G$ = Text1
'a = Len(G$)
'For w = 1 To a Step 7
'    r$ = Mid$(G$, w, 1)
'    U$ = Mid$(G$, w + 1, 1)
'    s$ = Mid$(G$, w + 2, 1)
'    t$ = Mid$(G$, w + 3, 1)
'    z$ = Mid$(G$, w + 4, 1)
'    Y$ = Mid$(G$, w + 5, 1)
'    c$ = Mid$(G$, w + 6, 1)
'    V$ = Mid$(G$, w + 7, 1)
'    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#980000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#fe7c00" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "000066" & Chr$(34) & ">" & t$ & "<FONT COLOR=" & Chr$(34) & "#980000" & Chr$(34) & "><sup>" & z$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#fe7c00" & Chr$(34) & ">" & Y$ & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & "><sub>" & c$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#000066" & Chr$(34) & V
'Next w
'aol4_chatsend (p$)
'End Function

Function AOL4_LastChatLine()
a = FindChat
If a = 0 Then Timer1.Enabled = False
AOL = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL, "MDIClient")
chat = findchildbyclass(mdi%, "AOL Child")
child = findchildbyclass(chat, "RICHCNTL")
GetTrim = Sendmessagebynum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = sendmessagebystring(child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$

If chattemp$ = theview$ Then Exit Function
chattemp$ = theview$
For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
   thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
   thechars$ = ""
End If
Next FindChar
lastlen = Val(FindChar) - Len(thechars$)
If lastlen = 1 Then Exit Function
lastline = Mid(theview$, lastlen, Len(thechars$))
AOLLastChatLine = lastline
If AOLLastChatLine = "" Then Exit Function
spcpos1 = InStr(1, AOLLastChatLine, ":")
sn1 = Left(AOLLastChatLine, spcpos1 - 1)
helps = Len(AOLLastChatLine) - (spcpos1 + 1)
chat = Right(AOLLastChatLine, helps - 1)
Text1.Text = sn1
Text2.Text = chat
End Function

Sub AOL4_instantmessage(Person, Message)
Dim im
'This openz an IM and fillz it out with Person and Message
AOL4_Keyword ("aol://9293:" & Person)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
im = findchildbytitle(mdi%, "Send Instant Message")
aolrich% = findchildbyclass(im, "RICHCNTL")
imsend% = findchildbyclass(im, "_AOL_Icon")
Loop Until (im <> 0 And aolrich% <> 0 And imsend% <> 0)
Call sendmessagebystring(aolrich%, WM_SETTEXT, 0, Message)
For sends = 1 To 9
    imsend% = getwindow(imsend%, GW_HWNDNEXT)
Next sends
Ao4Click (imsend%)
pause 0.5
a% = FindWindow("#32770", vbNullString)
'MsgBox a%
If a% = 0 Then Exit Sub
AOL4_KillWin a%
If im Then Call AOL4_KillWin(im)
End Sub
Sub Ao4Click(Button%)
SendNow% = Sendmessagebynum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = Sendmessagebynum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub

Function wm(who As String, subject As String, body As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
Toolbar% = findchildbyclass(AOL%, "AOL Toolbar")
toolbar2% = findchildbyclass(Toolbar%, "_AOL_Toolbar")
icn% = findchildbyclass(toolbar2%, "_AOL_Icon")
icn2% = getwindow(icn%, GW_HWNDNEXT)
Ao4Click icn2%
Do:
chld% = findchildbyclass(mdi%, "AOL Child")
gds$ = String$(200, 150)
blabla = GetWindowText(chld%, gds$, 100)
mstr = InStr(1, gds$, "Write", 1)
Loop Until mstr <> 0
pause 1
B% = findchildbyclass(chld%, "_AOL_Edit")
sndtext% = sendmessagebystring(B%, WM_SETTEXT, 0, who)
B% = getwindow(B%, GW_HWNDNEXT)
B% = getwindow(B%, GW_HWNDNEXT)
B% = getwindow(B%, GW_HWNDNEXT)
B% = getwindow(B%, GW_HWNDNEXT)
sndtext% = sendmessagebystring(B%, WM_SETTEXT, 0, subject)
B% = findchildbyclass(chld%, "RICHCNTL")
sndtext% = sendmessagebystring(B%, WM_SETTEXT, 0, body)
cutbut% = findchildbyclass(chld%, "_AOL_Icon")
For X = 1 To 18
    cutbut% = getwindow(cutbut%, GW_HWNDNEXT)
Next
Button% = findchildbyclass(chld%, "_AOL_Icon")
SendNow% = Sendmessagebynum(cutbut%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = Sendmessagebynum(cutbut%, WM_LBUTTONUP, &HD, 0)
pause 1
B% = FindWindow("_AOL_Modal", vbNullString)
AOL4_KillWin B%

End Function

'Sub moveform(hwnd As Form)


'This will move a form that has no title just by clicking
''and dragging.  Place in "Form×_MouseDown"''

'Dim mpos As Long
'Dim p As ConvertPointAPI
'Dim ret As Integer
'Call GetCursorPos(mpos)
'ret = SendMessage(hwnd.hwnd, WM_LBUTTONUP, 0, mpos)
'ret = SendMessage(hwnd.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, mpos) ''

'End Sub
Sub Ao4Kill45min()
Do
a% = FindWindow("_Aol_Palette", 0&)
B% = findchildbyclass(a%, "_Aol_Icon")
Call pause(0.001)
If a% = 0 Then Exit Sub
Loop Until B% <> 0
Ao4Click (B%)
End Sub
Sub killwait2()
Do: DoEvents
aom% = FindWindow("_AOL_Modal", 0&)
Loop Until aom% <> 0
OK% = findchildbyclass(aom%, "_AOL_Icon")
Do: DoEvents
Call Ao4Click(OK%)
Loop Until OK% <> 0
End Sub




Sub APP3D(myForm As Form, MyCtl As Control)
'Place in Form_Paint, works best with a grey background

'Example: Call APP3D(Me,Text1)

    myForm.ScaleMode = 3
    myForm.CurrentX = MyCtl.Left - 1
    myForm.CurrentY = MyCtl.Top + MyCtl.Height
    myForm.Line -Step(0, -(MyCtl.Height + 1)), RGB(92, 92, 92)
    myForm.Line -Step(MyCtl.Width + 1, 0), RGB(92, 92, 92)
    myForm.Line -Step(0, MyCtl.Height + 1), RGB(255, 255, 255)
    myForm.Line -Step(-(MyCtl.Width + 1), 0), RGB(255, 255, 255)
End Sub



Function Scramble1(phrase As String)
If phrase = "" Then Exit Function
Randomize Timer
txt$ = phrase
One = 1
Txt3$ = ""
txt$ = txt$ + " "
For X = 1 To Len(txt$)
  If Mid$(txt$, X, 1) = " " Then
    Two = X
    Txt2$ = Mid$(txt$, One, Two - One)
    For y = 1 To (Two - One)
      Rand = Int(Len(Txt2$) * Rnd) + 1
      Txt2$ = Mid$(Txt2$, Rand, 1) + Mid$(Txt2$, 1, Rand - 1) + Mid$(Txt2$, Rand + 1, Two)
    Next y
    Txt3$ = Txt3$ + Txt2$ + " "
    One = Two + 1
  End If
Next X
Txt3$ = Left$(Txt3$, Len(Txt3$) - 1)
txt$ = Txt3$
Scramble1 = txt$
End Function

Function generatecolor()
Dim red, blue, green, redf, bluef, greenf
Randomize
red = Hex((255 * Rnd) + 1)
blue = Hex((255 * Rnd) + 1)
green = Hex((255 * Rnd) + 1)
If Len(red) = 1 Then redf = "0" + red Else redf = red
If Len(blue) = 1 Then bluef = "0" + blue Else bluef = blue
If Len(green) = 1 Then greenf = "0" + green Else greenf = green
generatecolor = "#" + redf + bluef + greenf
'aol4_chatsend "<font color=" + Text1.Text + ">HI"
End Function
Function randomcolor(word As Control)
Dim hey
randomcolor = "<font color=" + generatecolor + ">"
For X = 0 To Len(word)
    word.SelStart = X
    word.SelLength = 1
    hey = word.SelText
    If hey = " " Then randomcolor = randomcolor + " "
    If hey <> " " Then randomcolor = randomcolor + hey + "<font color=" + generatecolor + ">"
Next

End Function
Function wavy(Text2 As Control)
Dim d, f
timers = 0
f = ""
For X = 0 To Len(Text2.Text)
    If timers = 0 Then d = "<Sup>"
    If timers = 1 Then d = "</sup>"
    If timers = 2 Then d = ""
    If timers = 3 Then d = "<Sub>"
    If timers = 4 Then d = "</sub>"
    timers = timers + 1
    If timers = 5 Then timers = 0
    Text2.SelStart = X
    Text2.SelLength = 1
    f = f + Text2.SelText + d
Next
aol4_chatsend f
End Function
Function wavycolor(Text2 As Control)
Dim d, f
timers = 0
f = ""
For X = 0 To Len(Text2.Text)
    If timers = 0 Then d = "<Sup><Font Color=" + generatecolor + ">"
    If timers = 1 Then d = "</sup><Font Color=" + generatecolor + ">"
    If timers = 2 Then d = "<Font Color=" + generatecolor + ">"
    If timers = 3 Then d = "<Sub><Font Color=" + generatecolor + ">"
    If timers = 4 Then d = "</sub><Font Color=" + generatecolor + ">"
    timers = timers + 1
    If timers = 5 Then timers = 0
    Text2.SelStart = X
    Text2.SelLength = 1
    f = f + Text2.SelText + d
Next
aol4_chatsend f
End Function
Function WavYChat(thetext As String)

G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "" & r$ & "" & u$ & "" & S$ & "" & T$
Next W
WavYChat = P$

End Function
Sub AOL4_punt(Person, Message)
Dim im
'This openz an IM and fillz it out with Person and Message
AOL4_Keyword ("aol://9293:" & Person)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = findchildbyclass(AOL%, "MDIClient")
im = findchildbytitle(mdi%, "Send Instant Message")
aolrich% = findchildbyclass(im, "RICHCNTL")
imsend% = findchildbyclass(im, "_AOL_Icon")
Loop Until (im <> 0 And aolrich% <> 0 And imsend% <> 0)
Call sendmessagebystring(aolrich%, WM_SETTEXT, 0, Message)
For sends = 1 To 9
imsend% = getwindow(imsend%, GW_HWNDNEXT)
Next sends
Ao4Click (imsend%)
a% = FindWindow("#32770", vbNullString)
If a% = 0 Then Exit Sub
AOL4_KillWin a%
End Sub

Sub Attention(thetext As String)
'G$ = WavYChaT("Surge ")
'L$ = WavYChaT(" by JoLT")
'aa$ = WavYChaT("Attention")
aol4_chatsend ("       ·¤±»»»»»»» ATTENTION «««««««¶¤·")
Call Timeout(0.15)
aol4_chatsend (thetext)
Call Timeout(0.15)
aol4_chatsend ("       ·¤±»»»»»»» ATTENTION «««««««¶¤·")
Call Timeout(0.15)
'aol4_chatsend ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub
Sub WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
aol4_chatsend (P$)
End Sub

Sub LineScroll()
aol4_chatsend ("     ·¤±»»»»»»»Greetz Out To «««««««¶¤·")
 pause (0.5)
aol4_chatsend ("        ·¤±»»»»»»»ACIDSK8ES«««««««¶¤·")
 pause (0.5)
aol4_chatsend ("         ·¤±»»»»»»»COOTER«««««««¶¤·")
 pause (0.5)
aol4_chatsend ("          ·¤±»»»»»»»LIVID«««««««¶¤· ")
 pause (0.5)
aol4_chatsend ("       ·¤±»»»»»»»Cheshire9X«««««««¶¤·")
 pause (0.5)
aol4_chatsend ("           ·¤±»»»»»»»    «««««««¶¤·")
 pause (0.5)
aol4_chatsend ("           ·¤±»»»»»»»    «««««««¶¤·")
 pause (0.5)
aol4_chatsend ("          ·¤±»»»»»»»FrOzEn«««««««¶¤·")
 pause (0.5)
End Sub

