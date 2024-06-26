Attribute VB_Name = "my"
'Sups Peeps Using my phat bas most of
'all this code i wrote totally mah self
'but like pws scanner and delete tree scanner
'was ripped but i aint got time
'to write one mah self, but just dont rip
'mah shit
'if u want some help or find ne probs
'email me at il1cli@aol.com
'lata peeps
'This is a 32-bit bas for AOL4
'By 1©
Declare Function getnextwindow1 Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
'-gdi32 Dlcaritons-'
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'-Kernal/Kernal32 Declaritons-'
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''----For Wav's-----''''''''''''''''''''''''''''''''''''''''''''
Const SND_SYNC = 0
Const SND_ASYNC = 1
Const SND_NODEFAULT = 2
Const SND_MEMORY = 4
Const SND_NOSTOP = 16
Declare Function sndPlaySound2 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function sndPlaySound Lib "MMSYSTEM.DLL" (ByVal WAVFILE As Any, ByVal WFLF As Integer) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''---Other----'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Global R%       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Option Explicit
Dim MailSettings(1 To 8)
Dim Normal%
Global StopButton As Variant
Global StopVariable As Variant
Type Color
 Red As Integer
 Green As Integer
 Blue As Integer
End Type
Declare Function sndPlaySoundA Lib "c:\WINDOWS\SYSTEM\WINMM.DLL" (ByVal lpszSoundName$, ByVal valueflags As Long) As Long




Global Const SND_LOOP = &H8


Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)

Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Declare Function setwindowtext Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const PROCESS_VM_READ = &H10


Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWMINIMIZED = 2


Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20

     'Result Code from WritePrivateProfileString
  'Passed to WritePrivateProfileString
 'Path to .ini file

Public Function AOLKillRoomFull(Delay%)
'Used to kill the "Room Full" message.
'The delay paramater is used by the room buster
'included in this bas. All it does it put a short
'delay when it closes a roomfull message.
Dim AoL%
Dim Button%
AoL% = FindWindow("#32770", "America Online")
Button% = FindChildByClass(AoL%, "Button")
If Button% = 0 Then
AOLKillRoomFull = 0
Exit Function
Else
aolclick Button%
AoL% = SendMessage(AoL%, WM_CLOSE, 0, 0)
AOLKillRoomFull = 1
Timeout Delay%
End If
End Function
Public Sub PlayWAV(FileName As String)
'Plays a wav file.
'Filename is the name of the wav that u want to
'play.
Dim Constants%
Dim PlayIt%
PlayIt% = sndPlaySoundA(FileName, SND_ASYNC Or SND_NODEFAULT)
End Sub

Function FindChildByClass(Parent%, ClassToFind$) As Integer
'Finds a child by the class name.
Dim ChildHandle%
Dim ClassOfChild$
Dim LengthOfClassOfChild%
ChildHandle% = GetWindow(Parent%, GW_CHILD)
While ChildHandle%
ClassOfChild$ = String(200, " ")
LengthOfClassOfChild% = GetClassName(ChildHandle%, ClassOfChild$, 199)
ClassOfChild$ = Left$(ClassOfChild$, LengthOfClassOfChild%)
If UCase$(ClassOfChild$) = UCase$(ClassToFind$) Then GoTo ExitWhile
ChildHandle% = GetWindow(ChildHandle%, GW_HWNDNEXT)
Wend
ExitWhile:
FindChildByClass = ChildHandle%
End Function

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

Sub randompassword(Text1 As TextBox)
Dim num
Dim first
Dim second
Dim third
Dim fourth
Dim fifth
num = RandomNumber("6")
If num = "1" Then
first = "1"
End If
If num = "2" Then
first = "2"
End If
If num = "3" Then
first = "3"
End If
If num = "4" Then
first = "4"
End If
If num = "5" Then
first = "5"
End If
If num = "6" Then
first = "6"
End If
num = ""
num = RandomNumber("6")
If num = "1" Then
second = "a"
End If
If num = "2" Then
second = "b"
End If
If num = "3" Then
second = "c"
End If
If num = "4" Then
second = "d"
End If
If num = "5" Then
second = "e"
End If
If num = "6" Then
second = "f"
End If
num = ""
num = RandomNumber("6")
If num = "1" Then
third = "1"
End If
If num = "2" Then
third = "2"
End If
If num = "3" Then
third = "3"
End If
If num = "4" Then
third = "4"
End If
If num = "5" Then
third = "5"
End If
If num = "6" Then
third = "6"
End If
num = ""
num = RandomNumber("6")
If num = "1" Then
fourth = "a"
End If
If num = "2" Then
fourth = "b"
End If
If num = "3" Then
fourth = "c"
End If
If num = "4" Then
fourth = "d"
End If
If num = "5" Then
fourth = "e"
End If
If num = "6" Then
fourth = "f"
End If
num = ""
num = RandomNumber("6")
If num = "1" Then
fifth = "1"
End If
If num = "2" Then
fifth = "2"
End If
If num = "3" Then
fifth = "3"
End If
If num = "4" Then
fifth = "4"
End If
If num = "5" Then
fifth = "5"
End If
If num = "6" Then
fifth = "6"
End If
Text1.Text = first + second + third + fourth + fifth
End Sub
Sub Timeout(HowLong)
'Delays your program by a certime specified in
'howlong.
Dim Sample
Sample = Timer
Do: DoEvents
Loop While Timer - Sample < HowLong
End Sub

Public Function AOLFindChatRoom()
'Finds the chatroom, if found returns the
'handle of the chatroom.
Dim AoL%
Dim mdi%
Dim ChildHandle%
Dim AolGlyph%
Dim AolStatic%
Dim Richcntl%
Dim AOLComboBox%
Dim AOLListBox%
Dim AOLIcon%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
ChildHandle% = GetWindow(mdi%, GW_CHILD)
While ChildHandle%
AolGlyph% = FindChildByClass(ChildHandle%, "_AOL_Glyph")
AolStatic% = FindChildByClass(ChildHandle%, "_AOL_Static")
Richcntl% = FindChildByClass(ChildHandle%, "RICHCNTL")
AOLComboBox% = FindChildByClass(ChildHandle%, "_AOL_Combobox")
AOLListBox% = FindChildByClass(ChildHandle%, "_AOL_Listbox")
AOLIcon% = FindChildByClass(ChildHandle%, "_AOL_ICON")
If AolGlyph% > 0 And AolStatic% > 0 And Richcntl% > 0 And AOLComboBox% > 0 And AOLListBox% > 0 And AOLIcon% > 0 Then GoTo ExitWhile
ChildHandle% = GetWindow(ChildHandle%, GW_HWNDNEXT)
Wend
ChildHandle% = 0
ExitWhile:
AOLFindChatRoom = ChildHandle%
End Function

Public Sub SendText(TextToSend)
'Sends a line of text to the chatroom.
Dim ChatRoom%
Dim Richcntl%
Dim X%
ChatRoom% = AOLFindChatRoom
If ChatRoom% = 0 Then Exit Sub
Richcntl% = FindChildByClass(ChatRoom%, "RichCntl")
For X% = 1 To 6
Richcntl% = GetWindow(Richcntl%, GW_HWNDNEXT)
Next X%
X% = SendMessageByString(Richcntl%, WM_SETTEXT, 0, TextToSend)
X% = SendMessageByNum(Richcntl%, WM_CHAR, 13, 0)
End Sub

Public Sub AolSendtext(TextToSend)
'Sends a line of text to the chatroom.
Dim ChatRoom%
Dim Richcntl%
Dim X%
ChatRoom% = AOLFindChatRoom
If ChatRoom% = 0 Then Exit Sub
Richcntl% = FindChildByClass(ChatRoom%, "RichCntl")
For X% = 1 To 6
Richcntl% = GetWindow(Richcntl%, GW_HWNDNEXT)
Next X%
X% = SendMessageByString(Richcntl%, WM_SETTEXT, 0, TextToSend)
X% = SendMessageByNum(Richcntl%, WM_CHAR, 13, 0)
End Sub

Public Sub AOLKeyword(WhatToSend$)
'Goes to a certain keyword that u specifiy.
Dim AoL%
Dim Edit%
Dim X%
AoL% = FindWindow("AOL Frame25", vbNullString)
Edit% = FindChildByClass(AoL%, "AOL Toolbar")
Edit% = FindChildByClass(Edit%, "_AOL_Toolbar")
Edit% = FindChildByClass(Edit%, "_AOL_Combobox")
Edit% = FindChildByClass(Edit%, "Edit")
X% = SendMessageByString(Edit%, WM_SETTEXT, 0, WhatToSend$)
X% = SendMessageByNum(Edit%, WM_CHAR, VK_SPACE, 0)
X% = SendMessageByNum(Edit%, WM_CHAR, VK_RETURN, 0)
End Sub

Public Sub aolclick(Iconhwnd%)
'Simulates clicking the mouse button.
Dim Click%
Click% = SendMessage(Iconhwnd%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Iconhwnd%, WM_LBUTTONUP, 0, 0&)
End Sub

Public Function AOLColor(Text$, bold As Variant, italic As Variant, underline As Variant, Arrangement$, Font$, color1 As Color, color2 As Color, color3 As Color, color4 As Color, Wavy As Variant, Freeze As Variant)
Dim blender As String
Dim Index%
Dim newred
Dim NewGreen
Dim NewBlue
Dim RedOne
Dim GreenOne
Dim BlueOne
Dim RedMinus
Dim GreenMinus
Dim BlueMinus
Dim i
Dim Dummy$
Dim Build$
Dim Stage%
Dim DummyTwo$
Dim counter%
Dim CounterTwo%
Dim ROne%
Dim GOne%
Dim bone%
Dim RTwo%
Dim GTwo%
Dim BTwo%
Select Case UCase$(Arrangement$)
Case "SINGLE"
   Text$ = "<font color=" + RGB2HEX(color1.Red, color1.Green, color1.Blue) + ">" + Text$
Case "MULTI"
   For i = 1 To Len(Text$)
      Dummy$ = Mid(Text$, i, 1)
      Stage% = Stage% + 1
      If Stage% = 5 Then Stage% = 1
      If Stage% = 1 Then Dummy$ = "<font color=" + RGB2HEX(color1.Red, color1.Green, color1.Blue) + ">" + Dummy$
      If Stage% = 2 Then Dummy$ = "<font color=" + RGB2HEX(color2.Red, color2.Green, color2.Blue) + ">" + Dummy$
      If Stage% = 3 Then Dummy$ = "<font color=" + RGB2HEX(color3.Red, color3.Green, color3.Blue) + ">" + Dummy$
      If Stage% = 4 Then Dummy$ = "<font color=" + RGB2HEX(color4.Red, color4.Green, color4.Blue) + ">" + Dummy$
      Build$ = Build$ + Dummy$
   Next i
   Text$ = Build$
Case "BLEND"
ROne% = color1.Red
GOne% = color1.Green
bone% = color1.Blue
RTwo% = color4.Red
GTwo% = color4.Green
BTwo% = color4.Blue
If ROne% > RTwo% Then
RedOne = True
RedMinus = (ROne% - RTwo%) / Len(Text$)
Else
RedOne = False
RedMinus = (RTwo% - ROne%) / Len(Text$)
End If
If GOne% > GTwo% Then
GreenOne = True
GreenMinus = (GOne% - GTwo%) / Len(Text$)
Else
GreenOne = False
GreenMinus = (GTwo% - GOne%) / Len(Text$)
End If
If bone% > BTwo% Then
BlueOne = True
BlueMinus = (bone% - BTwo%) / Len(Text$)
Else
BlueOne = False
BlueMinus = (BTwo% - bone%) / Len(Text$)
End If
newred = ROne%
NewGreen = GOne%
NewBlue = bone%
blender = blender + "<font color=" + RGB2HEX(Int(newred), Int(NewGreen), Int(NewBlue)) + ">" + Mid(Text$, Index% + 1, 1)
For Index% = 2 To Len(Text$) - 1
If RedOne = True Then newred = newred - RedMinus Else newred = newred + RedMinus
If GreenOne = True Then NewGreen = NewGreen - GreenMinus Else NewGreen = NewGreen + GreenMinus
If BlueOne = True Then NewBlue = NewBlue - BlueMinus Else NewBlue = NewBlue + BlueMinus
blender = blender + "<font color=" + RGB2HEX(Int(newred), Int(NewGreen), Int(NewBlue)) + ">" + Mid(Text$, Index%, 1)
Next Index%
If Len(Text$) > 1 Then blender = blender + "<font color=" + RGB2HEX(RTwo%, GTwo%, BTwo%) + ">" + Right(Text$, 1)
Text$ = blender
End Select
If Wavy = True Then
   If UCase$(Arrangement$) = "MULTI" Or UCase$(Arrangement$) = "BLEND" Then
      Build$ = ""
      Stage% = 0
      For i = 23 To Len(Text$) Step 23
         Dummy$ = Mid(Text$, i - 22, 23)
         Stage% = Stage% + 1
         If Stage% = 5 Then Stage% = 1
         If Stage% = 1 Then Dummy$ = "<sub>" + Dummy$
         If Stage% = 2 Then Dummy$ = "</sub>" + Dummy$
         If Stage% = 3 Then Dummy$ = "<sup>" + Dummy$
         If Stage% = 4 Then Dummy$ = "</sup>" + Dummy$
         Build$ = Build$ + Dummy$
      Next i
   Text$ = Build$
   Else
   Build$ = Left(Text$, 22)
   For i = 23 To Len(Text$)
      Dummy$ = Mid(Text$, i, 1)
      Stage% = Stage% + 1
      If Stage% = 5 Then Stage% = 1
      If Stage% = 1 Then Dummy$ = "<sub>" + Dummy$
      If Stage% = 2 Then Dummy$ = "</sub>" + Dummy$
      If Stage% = 3 Then Dummy$ = "<sup>" + Dummy$
      If Stage% = 4 Then Dummy$ = "</sup>" + Dummy$
      Build$ = Build$ + Dummy$
   Next i
   Text$ = Build$
   End If
End If
If Font$ = "" Then GoTo skipthis
Text$ = "<font face=" + Chr$(34) + Font$ + Chr$(34) + ">" + Text$
skipthis:
If underline = True Then Text$ = "<U>" + Text$
If bold = True Then Text$ = "<B>" + Text$
If italic = True Then Text$ = "<I>" + Text$
If Freeze = True Then Text$ = Text$ + "<binary>"
AOLColor = Text$
End Function



Public Sub AOLPRbuster(WhatRoom$)
'This is a Room Runner
Dim ChatRoom%
Dim StartRoom$
Dim X%
Dim i%
Dim Room$
Dim finial
finial = False
ChatRoom% = AOLFindChatRoom
StartRoom$ = String$(100, " ")
X% = GetWindowText(ChatRoom%, StartRoom$, 100)
For i% = 0 To 100
If i% = 0 Then i% = 0
If i% = 0 Then
Call AOLKeyword("aol://2719:2-2-" + WhatRoom$)
GoTo OK
End If
Call AOLKeyword("aol://2719:2-2-" + WhatRoom$)
OK:
Do
X% = AOLKillRoomFull(0)
If X% = 1 Then GoTo IsFull
ChatRoom% = AOLFindChatRoom
If ChatRoom% > 1 Then
Room$ = String(100, " ")
X% = GetWindowText(ChatRoom%, Room$, 100)
If Trim(UCase(Room$)) = Trim(UCase(StartRoom$)) Then
Else
Room$ = Trim(Room$)
Room$ = Left$(Room$, Len(Room$) - 1)
Exit Sub
End If
End If
Loop
IsFull:
If finial = True Then
GoTo finial:
End If
Next i%
finial:
End Sub

Public Sub AOLPRRunner(WhatRoom$)
'This is a Room Runner
Dim ChatRoom%
Dim StartRoom$
Dim X%
Dim i%
Dim Room$
ChatRoom% = AOLFindChatRoom
StartRoom$ = String$(100, " ")
X% = GetWindowText(ChatRoom%, StartRoom$, 100)
For i% = 0 To 100
If i% = 1 Then i% = 2
If i% = 0 Then
Call AOLKeyword("aol://2719:2-2-" + WhatRoom$)
GoTo OK
End If
Call AOLKeyword("aol://2719:2-2-" + WhatRoom$ + Str(i%))
OK:
Do
X% = AOLKillRoomFull(0)
If X% = 1 Then GoTo IsFull
ChatRoom% = AOLFindChatRoom
If ChatRoom% > 1 Then
Room$ = String(100, " ")
X% = GetWindowText(ChatRoom%, Room$, 100)
If Trim(UCase(Room$)) = Trim(UCase(StartRoom$)) Then
Else
Room$ = Trim(Room$)
Room$ = Left$(Room$, Len(Room$) - 1)
Exit Sub
End If
End If
Loop
IsFull:
Next i%
End Sub


Public Function RemoveSpace(TheText$)
'This function removes all spaces from a string.
'Returns the spaceless string.
Dim Text$
Dim theloop%
Text$ = TheText$
For theloop% = 1 To Len(TheText$)
If Mid(Text$, theloop%, 1) = " " Then
Text$ = Left$(Text$, theloop% - 1) + Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
RemoveSpace = Text$
End Function

Public Function AOLStayOnline()
'This function is good to put inside of a
'timer to stay online. It kills both the
'45 minute window and the idle window.
Dim Timer45%
Dim TimerIdle%
Dim Timer45Button%
Dim TimerIdleButton%
Dim AolStatic%
Dim StaticText$
Dim Dummy%
Dim SupposedToBe$
Timer45% = FindWindow("_AOL_PALETTE", "America Online Timer")
TimerIdle% = FindWindow("_AOL_MODAL", vbNullString)
If Timer45% > 0 Then
   Timer45Button% = FindChildByClass(Timer45%, "_aol_icon")
   aolclick Timer45Button%
   AOLStayOnline = 1
End If
If TimerIdle% > 0 Then
TimerIdleButton% = FindChildByClass(TimerIdle%, "_AOL_ICON")
AolStatic% = FindChildByClass(TimerIdle%, "_AOL_STATIC")
StaticText$ = String(100, " ")
Dummy% = SendMessageByString(AolStatic%, WM_GETTEXT, 100, StaticText$)
StaticText$ = UCase(Trim$(Left$(StaticText$, 21)))
SupposedToBe$ = UCase("You have been idle")
If StaticText$ = SupposedToBe$ Then
aolclick TimerIdleButton%
AOLStayOnline = 2
End If
End If
End Function

Public Sub AOLIM(Person$, WhatToSend$)
'Im's somebody
Dim AoL%
Dim mdi%
Dim im%
Dim ImText$
Dim ImTextLength%
Dim X%
Dim Colon%
Dim Dummy$
Dim Richcntl%
Dim Button%
Dim Buddy%
Dim Edit%
Dim Modal%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
im% = FindChildByTitle(mdi%, Person$)
If im% = FindChildByTitle(mdi%, "Instant Message") And im% > 0 Then
Richcntl% = FindChildByClass(im%, "RICHCNTL")
For X% = 1 To 13
Richcntl% = GetWindow(Richcntl%, GW_HWNDNEXT)
Next X%
Button% = FindChildByClass(im%, "_AOL_ICON")
For X% = 1 To 13
Button% = GetWindow(Button%, GW_HWNDNEXT)
Next X%
X% = SendMessageByString(Richcntl%, WM_SETTEXT, 0, WhatToSend$)
aolclick Button%
Exit Sub
Else
AOLKeyword "aol://9293:" + Person$
Do: DoEvents
im% = FindChildByTitle(mdi%, "Send Instant Message")
Richcntl% = FindChildByClass(im%, "RichCntl")
Button% = FindChildByClass(im%, "_AOL_ICON")
For X% = 1 To 9
Button% = GetWindow(Button%, GW_HWNDNEXT)
Next X%
Loop Until Button% > 0 And Richcntl% > 0
X% = SendMessageByString(Richcntl%, WM_SETTEXT, 0, WhatToSend$)
aolclick Button%
Do
im% = FindChildByTitle(mdi%, "Send Instant Message")
Modal% = FindWindow("#32770", "America Online")
If Modal% > 0 Then
X% = SendMessage(Modal%, WM_CLOSE, 0, 0)
X% = SendMessage(im%, WM_CLOSE, 0, 0)
End If
Loop Until im% = 0
End If
End Sub

Function EliteString(strng$)
'Changes to elite text
Dim InputString$
Dim length%
Dim Spaces%
Dim NextLetter$
Dim DoubleLetter$
Dim newstring$
Dim Skip%
InputString$ = strng$
length% = Len(InputString$)
Do While Spaces% <= length%
DoEvents
Spaces% = Spaces% + 1
NextLetter$ = Mid$(InputString$, Spaces%, 1)
If NextLetter$ = "A" Then NextLetter$ = "À"
If NextLetter$ = "a" Then NextLetter$ = "á"
If NextLetter$ = "B" Then NextLetter$ = "ß"
If NextLetter$ = "C" Then NextLetter$ = "Ç"
If NextLetter$ = "c" Then NextLetter$ = "ç"
If NextLetter$ = "D" Then NextLetter$ = "Ð"
If NextLetter$ = "E" Then NextLetter$ = "È"
If NextLetter$ = "e" Then NextLetter$ = "é"
If NextLetter$ = "f" Then NextLetter$ = "ƒ"
If NextLetter$ = "I" Then NextLetter$ = "Ì"
If NextLetter$ = "i" Then NextLetter$ = "í"
If NextLetter$ = "l" Then NextLetter$ = "Ï"
If NextLetter$ = "N" Then NextLetter$ = "Ñ"
If NextLetter$ = "n" Then NextLetter$ = "ñ"
If NextLetter$ = "O" Then NextLetter$ = "Ò"
If NextLetter$ = "o" Then NextLetter$ = "ó"
If NextLetter$ = "r" Then NextLetter$ = "®"
If NextLetter$ = "S" Then NextLetter$ = "Š"
If NextLetter$ = "s" Then NextLetter$ = "š"
If NextLetter$ = "t" Then NextLetter$ = "†"
If NextLetter$ = "U" Then NextLetter$ = "Ù"
If NextLetter$ = "u" Then NextLetter$ = "ú"
If NextLetter$ = "x" Then NextLetter$ = "×"
If NextLetter$ = "Y" Then NextLetter$ = "Ÿ"
If NextLetter$ = "y" Then NextLetter$ = "ý"
If NextLetter$ = "<" Then NextLetter$ = "«"
If NextLetter$ = ">" Then NextLetter$ = "»"
If NextLetter$ = "`" Then NextLetter$ = "“"
If NextLetter$ = "'" Then NextLetter$ = "”"
newstring$ = newstring$ + NextLetter$
Loop
EliteString = newstring$
End Function
Function HackerString(strng$)
'Changes to Hacker text
Dim newstring$
Dim InputString$
Dim length%
Dim Spaces%
Dim NextLetter$
InputString$ = strng$
length% = Len(InputString$)
Do While Spaces% <= length%
Spaces% = Spaces% + 1
Let NextLetter$ = Mid$(InputString$, Spaces%, 1)
If NextLetter$ = "A" Then NextLetter$ = "a"
If NextLetter$ = "E" Then NextLetter$ = "e"
If NextLetter$ = "I" Then NextLetter$ = "i"
If NextLetter$ = "O" Then NextLetter$ = "o"
If NextLetter$ = "U" Then NextLetter$ = "u"
If NextLetter$ = "b" Then NextLetter$ = "B"
If NextLetter$ = "c" Then NextLetter$ = "C"
If NextLetter$ = "d" Then NextLetter$ = "D"
If NextLetter$ = "z" Then NextLetter$ = "Z"
If NextLetter$ = "f" Then NextLetter$ = "F"
If NextLetter$ = "g" Then NextLetter$ = "G"
If NextLetter$ = "h" Then NextLetter$ = "H"
If NextLetter$ = "y" Then NextLetter$ = "Y"
If NextLetter$ = "j" Then NextLetter$ = "J"
If NextLetter$ = "k" Then NextLetter$ = "K"
If NextLetter$ = "l" Then NextLetter$ = "L"
If NextLetter$ = "m" Then NextLetter$ = "M"
If NextLetter$ = "n" Then NextLetter$ = "N"
If NextLetter$ = "x" Then NextLetter$ = "X"
If NextLetter$ = "p" Then NextLetter$ = "P"
If NextLetter$ = "q" Then NextLetter$ = "Q"
If NextLetter$ = "r" Then NextLetter$ = "R"
If NextLetter$ = "s" Then NextLetter$ = "S"
If NextLetter$ = "t" Then NextLetter$ = "T"
If NextLetter$ = "w" Then NextLetter$ = "W"
If NextLetter$ = "v" Then NextLetter$ = "V"
If NextLetter$ = " " Then NextLetter$ = " "
newstring$ = newstring$ + NextLetter$
Loop
HackerString = newstring$
End Function
Function SpacedString(strng$)
'Puts one space after every character
Dim InputString$
Dim length%
Dim Spaces%
Dim NextLetter$
Dim newstring$
InputString$ = strng$
length% = Len(InputString$)
Do While Spaces% <= length%
Spaces% = Spaces% + 1
NextLetter$ = Mid$(InputString$, Spaces%, 1)
NextLetter$ = NextLetter$ + " "
newstring$ = newstring$ + NextLetter$
Loop
SpacedString = Left$(newstring$, Len(newstring$) - 2)
End Function

Public Sub AOLIMS(OnorOff$)
'Turns Im's on or off, depending on OnorOff$
'Values are "ON" or "OFF"
Dim Modal%
Dim Button%
If UCase$(OnorOff$) = "ON" Then
AOLIM "$IM_ON", "1c Ownz U!"
Exit Sub
End If
If UCase$(OnorOff$) = "OFF" Then
AOLIM "$IM_OFF", "1c Ownz U!"
Exit Sub
End If
Exit Sub
End Sub

Public Sub AOLComposeMail(Person$, Subject$, Body$)
'Sends an email to someone.
Dim AoL%
Dim mdi%
Dim Toolbar%
Dim ComposeIcon%
Dim ComposeMail%
Dim sendto%
Dim X%
Dim SubjectBox%
Dim BodyBox%
Dim SendIcon%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
Toolbar% = FindChildByClass(AoL%, "AOL TOOLBAR")
Toolbar% = FindChildByClass(Toolbar%, "_AOL_TOOLBAR")
ComposeIcon% = FindChildByClass(Toolbar%, "_AOL_ICON")
ComposeIcon% = GetWindow(ComposeIcon%, GW_HWNDNEXT)
aolclick ComposeIcon%
Do
ComposeMail% = FindChildByTitle(mdi%, "Write Mail")
sendto% = FindChildByClass(ComposeMail%, "_AOL_EDIT")
SubjectBox% = GetWindow(sendto%, GW_HWNDNEXT)
For X% = 1 To 3
SubjectBox% = GetWindow(SubjectBox%, GW_HWNDNEXT)
Next X%
BodyBox% = GetWindow(SubjectBox%, GW_HWNDNEXT)
For X% = 1 To 14
BodyBox% = GetWindow(BodyBox%, GW_HWNDNEXT)
Next X%
SendIcon% = GetWindow(BodyBox%, GW_HWNDNEXT)
For X% = 1 To 6
SendIcon% = GetWindow(SendIcon%, GW_HWNDNEXT)
Next X%
Loop Until ComposeMail% > 0 And sendto% > 0 And SubjectBox% > 0 And BodyBox% > 0 And SendIcon% > 0
X% = SendMessageByString(sendto%, WM_SETTEXT, 0, Person$)
X% = SendMessageByString(SubjectBox%, WM_SETTEXT, 0, Subject$)
X% = SendMessageByString(BodyBox%, WM_SETTEXT, 0, Body$)
aolclick SendIcon%
Do
ComposeMail% = FindChildByTitle(mdi%, "Write Mail")
Loop Until ComposeMail% = 0
End Sub

Sub AOLHide(welcome As Variant, Toolbar As Variant)
'Hides aol unless welcome or toolbar is set to true
Dim AoL%
Dim mdi%
Dim X%
Dim Wel%
Dim Toollbar%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
Toollbar% = FindChildByClass(AoL%, "AOL Toolbar")
Wel% = FindChildByTitle(mdi%, "Welcome,")
If welcome = True Then X% = ShowWindow(Wel%, SW_HIDE): Exit Sub
If Toolbar = True Then X% = ShowWindow(Toollbar%, SW_HIDE): Exit Sub
X% = ShowWindow(AoL%, SW_HIDE)
End Sub

Sub AOLShow(welcome As Variant, Toolbar As Variant)
'Show's aol, unless welcome is set to true
Dim AoL%
Dim mdi%
Dim X%
Dim Wel%
Dim Toollbar%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
Toollbar% = FindChildByClass(AoL%, "AOL Toolbar")
Wel% = FindChildByTitle(mdi%, "Welcome,")
If welcome = True Then X% = ShowWindow(Wel%, SW_SHOW): Exit Sub
If Toolbar = True Then X% = ShowWindow(Toollbar%, SW_SHOW): Exit Sub
X% = ShowWindow(AoL%, SW_SHOW)
End Sub

Public Function AOLAddRoom()
'Adds the room.
'Returns the list of people in string format.
On Error Resume Next

Dim AOLProcess&
Dim ListItemHold&
Dim Person$
Dim ListPersonHold&
Dim ReadBytes&
Dim Room%
Dim aolhandle%
Dim aolthread%
Dim AOLProcessThread%
Dim PersonIndex%
Dim RoomCount%
Dim Buffer$
Room% = AOLFindChatRoom()
aolhandle% = FindChildByClass(Room%, "_AOL_Listbox")
RoomCount% = SendMessage(aolhandle%, LB_GETCOUNT, 0, 0)

For PersonIndex% = 0 To RoomCount% - 1

aolthread% = GetWindowThreadProcessId(aolhandle%, AOLProcess&)
AOLProcessThread% = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess&)

If AOLProcessThread% Then
Person$ = String$(4, vbNullChar)
ListItemHold& = SendMessage(aolhandle%, LB_GETITEMDATA, ByVal CLng(PersonIndex%), ByVal 0&)
ListItemHold& = ListItemHold& + 24
Call ReadProcessMemory(AOLProcessThread%, ListItemHold&, Person$, 4, ReadBytes&)
                        
Call RtlMoveMemory(ListPersonHold&, ByVal Person$, 4)
ListPersonHold& = ListPersonHold& + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread%, ListPersonHold&, Person$, Len(Person$), ReadBytes&)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread%)
End If
If Not Person$ = UserSN Then AOLAddRoom = AOLAddRoom + ", " + Person$
Next PersonIndex%
AOLAddRoom = Right$(AOLAddRoom, Len(AOLAddRoom) - 1)
End Function



Public Function Combo2String(TheCombo As ComboBox)
Dim X%
For X% = 0 To TheCombo.ListCount - 1
Combo2String = Combo2String + "," + TheCombo.List(X%)
Next X%
Combo2String = Right$(Combo2String, Len(Combo2String) - 1)
End Function

Public Sub Combo2List(TheCombo As ComboBox, TheList As ListBox)
Dim X%
For X% = 0 To TheCombo.ListCount - 1
TheList.AddItem TheCombo.List(X%)
Next X%
End Sub





Sub AOLUpchat(OnorOff$)
Dim AoL%
Dim Modal%
Dim X%
If UCase(OnorOff$) = "ON" Then
AoL% = FindWindow("AOL Frame25", vbNullString)
Modal% = FindWindow("_AOL_Modal", vbNullString)
X% = ShowWindow(Modal%, SW_SHOWMINIMIZED)
X% = ShowWindow(Modal%, SW_HIDE)
X% = SetFocusAPI(AoL%)
End If
If UCase(OnorOff$) = "OFF" Then
Modal% = FindWindow("_AOL_Modal", vbNullString)
X% = ShowWindow(Modal%, SW_SHOW)
X% = ShowWindow(Modal%, SW_RESTORE)
X% = SetFocusAPI(Modal%)
End If
End Sub

Public Sub AOLScroller(Lines10or14%)
If Lines10or14% = 10 Then AolSendtext "<I><B><font color=" + Chr$(34) + "#800080" + Chr$(34) + ">@@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@"
If Lines10or14% = 14 Then AolSendtext "<font face=" + Chr$(34) + "Stencil" + Chr$(34) + "><I><B><font color=" + Chr$(34) + "#04040" + Chr$(34) + ">|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
End Sub

Sub AOLKillModal()
Dim Modal%
Dim X%
Modal% = FindWindow("_AOL_Modal", vbNullString)
X% = SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub AOLKillWait()
Dim Modal%
Dim Button%
Dim AoL%
Dim Toolbar%
AoL% = FindWindow("AOL Frame25", vbNullString)
Call RunMenuByString(AoL%, "&About America Online")
Do
Modal% = FindWindow("_AOL_Modal", vbNullString)
Button% = FindChildByClass(Modal%, "_AOL_Icon")
Loop Until Button% > 0
aolclick Button%
End Sub

Sub RunMenuByString(WhichProgram%, thestring$)
Dim MainMenu%
Dim MainCount%
Dim MainLoop%
Dim SubMenu%
Dim SubCount%
Dim SubLoop%
Dim TheGottenString$
Dim SubID%
Dim ExecuteMenuCommand%
Dim GetTheString%
MainMenu% = GetMenu(WhichProgram%)
MainCount% = GetMenuItemCount(MainMenu%)
For MainLoop% = 0 To MainCount% - 1
   SubMenu% = GetSubMenu(MainMenu%, MainLoop%)
   SubCount% = GetMenuItemCount(SubMenu%)
   For SubLoop% = 0 To SubCount% - 1
      SubID% = GetMenuItemID(SubMenu%, SubLoop%)
      TheGottenString$ = String$(100, " ")
      GetTheString% = GetMenuString(SubMenu%, SubID%, TheGottenString$, 100, 1)
      If InStr(UCase(TheGottenString$), UCase(thestring$)) Then GoTo RunTheMenu
   Next SubLoop%
Next MainLoop%
Exit Sub
RunTheMenu:
ExecuteMenuCommand% = SendMessage(WhichProgram%, WM_COMMAND, SubID%, 0)
End Sub


Public Sub AOLMassIm(whattosay$, names$)
StopButton = False
'Mass ims a room or a list of people
'To mass im the room use: Call AOLMassIm("Blah","")
'To mass im a list it must be in string format.
'To stop the mass im set stopbutton=true
Dim RoomSns$
Dim Person$
If names$ = "" Then
RoomSns$ = AOLAddRoom
Else
RoomSns$ = names$
End If
Do: DoEvents
If InStr(RoomSns$, ",") = 0 Then Exit Do
Person$ = Left$(RoomSns$, InStr(RoomSns$, ",") - 1)
If StopButton = True Then Exit Sub
If InStr(whattosay$, "$$") Then
whattosay$ = ReplaceText(whattosay$, "$$", Person$)
End If
AOLIM Person$, whattosay$
Timeout 0.2
RoomSns$ = Right(RoomSns$, Len(RoomSns$) - InStr(RoomSns$, ","))
Loop
If StopButton = True Then Exit Sub
AOLIM RoomSns$, whattosay$
End Sub



Sub Fader4(texttofade$)
Dim color1 As Color
Dim color2 As Color
Dim color3 As Color
Dim color4 As Color
color1.Blue = 255
color2.Red = 0
color3.Green = 125
color4.Red = 500
AolSendtext (AOLColor(texttofade$, False, False, False, "BLEND", "Arial", color1, color2, color3, color4, False, False))
End Sub




Public Sub AOLEmailBomb(Person$, Subject$, Body$, HowMany%)
'Emails somebody a bunch of times
'To stop set stopbutton=true
StopButton = False
Dim X%
For X% = 1 To HowMany%
DoEvents
If StopButton = True Then Exit Sub
AOLComposeMail Person$, Subject$, Body$
Next X%
End Sub

Public Function AOLGetChat()
'Gets the whole Chatroom Text.
Dim ChatRoom%
Dim Richcntl%
Dim GetTrim%
Dim TrimSpace$
Dim getstring%
Dim trimspace2$
ChatRoom% = AOLFindChatRoom
Richcntl% = FindChildByClass(ChatRoom%, "richcntl")
GetTrim% = SendMessageByNum(Richcntl%, WM_GETTEXTLENGTH, 0&, 0&)
TrimSpace$ = GetTrim%
getstring% = SendMessageByString(Richcntl%, WM_GETTEXT, GetTrim% + 1, TrimSpace$)
trimspace2$ = Right$(TrimSpace$, Len(TrimSpace$) - 1)
AOLGetChat = trimspace2$
End Function

Public Function AOLGetLastLine(TheText$, PersonOrText$)
'Gets the last line of chat in the chatroom.
'The whole chatroom text must be supplied.
'"Person" or "Text" must be answered.
Dim Character$
Dim CurrentLine$
Dim X%
For X% = 1 To Len(TheText$)
Character$ = Mid(TheText$, X%, 1)
CurrentLine$ = CurrentLine$ + Character$
If Character$ = Chr(13) Then CurrentLine$ = ""
Next X%
If CurrentLine$ = "" Then Exit Function
If UCase$(PersonOrText$) = "PERSON" Then AOLGetLastLine = Left$(CurrentLine$, InStr(CurrentLine$, ":") - 1)
If UCase$(PersonOrText$) = "TEXT" Then AOLGetLastLine = Right$(CurrentLine$, (Len(CurrentLine$) - InStr(CurrentLine$, ":") - 2))
End Function

Public Function AOLIMAnswer(YourAnswer$)
'Ims a person, gets there message, and
'closes the im, whenever someone ims you.
'This works good in a timer.
Dim AoL%
Dim mdi%
Dim im%
Dim Richcntl%
Dim TheText$
Dim X%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
im% = FindChildByTitle(mdi%, ">Instant Message From:")
If im% = 0 Then Exit Function
Richcntl% = FindChildByClass(im%, "RICHCNTL")
TheText$ = String(520, " ")
X% = SendMessageByString(Richcntl%, WM_GETTEXT, 520, TheText$)
AOLIMAnswer = Trim(TheText$)
TheText$ = String(40, " ")
X% = GetWindowText(im%, TheText$, 40)
TheText$ = Trim(TheText$)
TheText$ = Right(TheText$, Len(TheText$) - 23)
AOLIM Left(TheText$, Len(TheText$) - 1), YourAnswer$
X% = SendMessage(im%, WM_CLOSE, 0, 0)
End Function

Public Sub FormEffect(Form As Form, How$, HowFar%, Delay%)
'Form = which form to effect
'Correct values of How$ are:
'RollUp, RollDown, SlideLeft, SlideRight,
'SlideDown, SlideUp, and Center
'HowFar% vary's according to your choice:
'RollUp - how far to roll up
'RollDown - how far to roll down
'SlideLeft - how far to slide left
'SlideRight - how far to slide right
'SlideDown - how far to slide down
'SlideUp - how far to slide up
'Center - value is ignored
'Move - drags the form when you put in
'       the form_mousedown
'Delay% = how slow, bigger the slower
On Error Resume Next
Dim X%
Dim xx%
Select Case UCase$(How$)
   Case "ROLLUP"
      Form.Show
      For X% = 0 To HowFar%
      DoEvents
      For xx% = 1 To Delay% * 100
      Next xx%
      Form.Height = Form.Height - 1
      Next X%
   Case "ROLLDOWN"
      Form.Show
      For X% = 0 To HowFar%
      DoEvents
      For xx% = 1 To Delay% * 100
      Next xx%
      Form.Height = Form.Height + 1
      Next X%
   Case "CENTER"
      With Form
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
      Form.Top = (Screen.Height - Form.Height) / 2
      Form.Left = (Screen.Width - Form.Width) / 2
   Case "SLIDEUP"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Top = Form.Top - 1
      Next X%
   Case "SLIDEDOWN"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Top = Form.Top + 1
      Next X%
   Case "SLIDERIGHT"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Left = Form.Left + 1
      Next X%
   Case "SLIDELEFT"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Left = Form.Left - 1
      Next X%
End Select
End Sub

Public Sub AOLIMIgnore(Person$)
'Im Ignores a person u specifiy
AOLIM "$IM_OFF " + Person$, "1c"
End Sub

Public Sub AOLIMunIgnore(Person$)
'Im Ignores a person u specifiy
AOLIM "$IM_ON " + Person$, "1c"
End Sub


Public Sub AOLSignOff(Quick)
'Signs u off
'Quick can be true or false
Dim AoL%
AoL% = FindWindow("AOL Frame25", vbNullString)
Select Case Quick
   Case True
      AOLHide False, False
      AoL% = SendMessage(AoL%, WM_CLOSE, 0, 0)
   Case False
      Call RunMenuByString(AoL%, "&Sign Off")
End Select
End Sub

Public Sub AoLTitle(NewTitle$)
'Changes the title of aol
Dim AoL%
AoL% = FindWindow("AOL Frame25", vbNullString)
AoL% = setwindowtext(AoL%, NewTitle$)
End Sub

Public Function AOLCountMail(WhichBox$, closeit)
'Counts your mail
'WhichBox$ can be NEW, OLD, or SENT
'Closeit can be true or false
Dim Tree%
Dim AoL%
Dim mdi%
Dim MailBox%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
MailBox% = FindChildByTitle(mdi%, "Fairite's Online MailBox")
Tree% = AOLOpenMail(WhichBox$)
AOLCountMail = AOLWaitMail(Tree%)
If closeit = True Then MailBox% = SendMessage(MailBox%, WM_CLOSE, 0, 0)
End Function

Public Function AOLWaitMail(Tree%)
'Tree% is the handle of the _AOL_TREE
'Returns the number of mails
Dim FirstCount%
Do: DoEvents
FirstCount% = SendMessage(Tree%, LB_GETCOUNT, 0, 0)
Timeout 2
AOLWaitMail = SendMessage(Tree%, LB_GETCOUNT, 0, 0)
Loop Until FirstCount% = AOLWaitMail
End Function

Private Sub OpenMenu(which%)
'This is used by RunMenu and is not meant to
'be used by any other program.
Dim AoL%
Dim X%
AoL% = FindWindow("ThunderRT5Form", "AOL Popup Ctrl")
Select Case which%
   Case 0
      X% = WinExec(App.Path + "\apc.exe", SW_HIDE)
   Case 1
      X% = setwindowtext(AoL%, "1")
   Case 2
      X% = setwindowtext(AoL%, "2")
   Case 3
      X% = setwindowtext(AoL%, "3")
   Case 4
      X% = setwindowtext(AoL%, "4")
   Case 5
      X% = setwindowtext(AoL%, "5")
   Case 6
      X% = setwindowtext(AoL%, "6")
   Case 7
      X% = setwindowtext(AoL%, "7")
   Case 8
      X% = setwindowtext(AoL%, "X")
End Select
End Sub

Public Sub RunMenu(Main%, popup%)
Dim AoL%
Dim Toolbar%
Dim X%
Dim Icon%
Normal% = FindWindow("#32768", vbNullString)
AoL% = FindWindow("AOL Frame25", vbNullString)
Toolbar% = FindChildByClass(AoL%, "AOL Toolbar")
Toolbar% = FindChildByClass(Toolbar%, "_AOL_TOOLBAR")
Select Case Main%
   Case 0
      OpenMenu 0
   Case 1
      Icon% = FindChildByClass(Toolbar%, "_AOL_ICON")
      aolclick Icon%
   Case 2
      Icon% = FindChildByClass(Toolbar%, "_AOL_ICON")
      Icon% = GetWindow(Icon%, GW_HWNDNEXT)
      aolclick Icon%
   Case 3
      Select Case popup%
         Case 1
            AOLKeyword "aol://4344:1204.MAILcent.9385917.565821970"
         Case 2
            Icon% = FindChildByClass(Toolbar%, "_AOL_ICON")
            aolclick Icon%
         Case 3
            Icon% = FindChildByClass(Toolbar%, "_AOL_ICON")
            Icon% = GetWindow(Icon%, GW_HWNDNEXT)
            aolclick Icon%
         Case 4
            OpenMenu 1
            X% = SendMessage(Wait, WM_CHAR, 111, 0)
         Case 5
            OpenMenu 1
            X% = SendMessage(Wait, WM_CHAR, 115, 0)
         Case 6
            OpenMenu 1
            X% = SendMessage(Wait, WM_CHAR, 97, 0)
         Case 7
            OpenMenu 1
            X% = SendMessage(Wait, WM_CHAR, 112, 0)
         Case 8
            AOLKeyword "aol://4344:1204.mailcnt.9376642.525891810"
         Case 9
            AOLKeyword "aol://4344:1204.mxmain.9384765.562532827"
         Case 10
            OpenMenu 1
            X% = SendMessage(Wait, WM_CHAR, 116, 0)
         Case 11
            OpenMenu 1
            X% = SendMessage(Wait, WM_CHAR, 117, 0)
      End Select
   Case 4
      Icon% = FindChildByClass(Toolbar%, "_AOL_ICON")
      For X% = 1 To 3
      Icon% = GetWindow(Icon%, GW_HWNDNEXT)
      Next X%
      aolclick Icon%
   Case 5
      Select Case popup%
         Case 1
            OpenMenu 2
            X% = SendMessage(Wait, WM_CHAR, 112, 0)
         Case 3
            OpenMenu 2
            X% = SendMessage(Wait, WM_CHAR, 111, 0)
         Case 4
            OpenMenu 2
            X% = SendMessage(Wait, WM_CHAR, 100, 0)
         Case 5
            OpenMenu 2
            X% = SendMessage(Wait, WM_CHAR, 109, 0)
         Case 6
            OpenMenu 2
            X% = SendMessage(Wait, WM_CHAR, 110, 0)
         Case 7
            OpenMenu 2
            X% = SendMessage(Wait, WM_CHAR, 108, 0)
      End Select
   Case 6
      Select Case popup%
         Case 1
            AOLKeyword "aol://4344:1225.mymain.6935573.560703915"
         Case 2
            AOLKeyword "aol://4344:1225.w4scr.6898937.566019331"
         Case 3
            OpenMenu 3
            X% = SendMessage(Wait, WM_CHAR, 112, 0)
         Case 4
            OpenMenu 3
            X% = SendMessage(Wait, WM_CHAR, 121, 0)
         Case 5
            AOLKeyword "Names"
         Case 6
            OpenMenu 3
            X% = SendMessage(Wait, WM_CHAR, 97, 0)
         Case 7
            AOLKeyword "aol://4344:1580.prntcon.12263709.564517913"
         Case 8
            OpenMenu 3
            X% = SendMessage(Wait, WM_CHAR, 111, 0)
         Case 9
            AOLKeyword "BuddyView"
         Case 10
            OpenMenu 3
            X% = SendMessage(Wait, WM_CHAR, 101, 0)
         Case 11
            AOLKeyword "aol://1391:40-19546"
         Case 12
            AOLKeyword "aol://4344:775.gremind.6896459.520364252"
         Case 13
            AOLKeyword "aol://5862:146"
      End Select
   Case 7
      Select Case popup%
         Case 1
            OpenMenu 4
            X% = SendMessage(Wait, WM_CHAR, 102, 0)
         Case 2
            OpenMenu 4
            X% = SendMessage(Wait, WM_CHAR, 97, 0)
         Case 3
            OpenMenu 4
            X% = SendMessage(Wait, WM_CHAR, 103, 0)
      End Select
   Case 8
      Select Case popup%
         Case 1
            AOLKeyword "aol://4344:2021.n_spoke3.17630876.558238150"
         Case 2
            OpenMenu 5
            X% = SendMessage(Wait, WM_CHAR, 103, 0)
         Case 3
            AOLKeyword "http://www.aol.com/netfind/"
         Case 4
            AOLKeyword "aol://5863:126"
         Case 5
            AOLKeyword "http://www.aol.com/netfind/newsgroups.html"
         Case 6
            AOLKeyword "aol://4344:206.ftpmain.8358103.531330427"
      End Select
   Case 9
      Select Case popup%
         Case 1
            AOLKeyword "aol://1722:aoltoday"
         Case 2
            AOLKeyword "aol://1722:news"
         Case 3
            AOLKeyword "aol://1722:sports"
         Case 4
            AOLKeyword "aol://1722:influence"
         Case 5
            AOLKeyword "aol://1722:travel"
         Case 6
            AOLKeyword "aol://1722:international"
         Case 7
            AOLKeyword "aol://1722:personalfinance"
         Case 8
            AOLKeyword "aol://1722:workplace"
         Case 9
            AOLKeyword "aol://1722:computing"
         Case 10
            AOLKeyword "aol://1722:researchandlearn"
         Case 11
            AOLKeyword "aol://1722:entertainment"
         Case 12
            AOLKeyword "aol://1722:games"
         Case 13
            AOLKeyword "aol://1722:interests"
         Case 14
            AOLKeyword "aol://1722:lifestyles"
         Case 15
            AOLKeyword "aol://1722:shopping"
         Case 16
            AOLKeyword "aol://1722:health"
         Case 17
            AOLKeyword "aol://1722:families"
         Case 18
            AOLKeyword "aol://1722:kidsonly"
         Case 19
            AOLKeyword "aol://1722:digitalcity"
      End Select
   Case 10
      Select Case popup%
         Case 1
            AOLKeyword "aol://4344:619.pcp1ttop.3747944.550692449"
         Case 2
            OpenMenu 7
            X% = SendMessage(Wait, WM_CHAR, 99, 0)
         Case 3
            OpenMenu 7
            X% = SendMessage(Wait, WM_CHAR, 102, 0)
         Case 4
            OpenMenu 7
            X% = SendMessage(Wait, WM_CHAR, 115, 0)
         Case 5
            AOLKeyword "aol://4344:227.mooze.3035633.520439342"
         Case 6
            AOLKeyword "aol://9293:"
         Case 7
            AOLKeyword "BuddyView"
         Case 8
            AOLKeyword "SendPage"
         Case 9
            OpenMenu 7
            X% = SendMessage(Wait, WM_CHAR, 100, 0)
         Case 10
            OpenMenu 7
            X% = SendMessage(Wait, WM_CHAR, 108, 0)
         Case 11
            OpenMenu 7
            X% = SendMessage(Wait, WM_CHAR, 103, 0)
         Case 12
            AOLKeyword "http://aol.switchboard.com/aol/default.htm"
      End Select
   Case Else
      OpenMenu 8
End Select
End Sub

Private Function Wait()
'This is used by RunMenu and is not meant to
'be used by any other program.
Dim Modal%
Do
Modal% = FindWindow("#32768", vbNullString)
Loop While Modal% = Normal%
Wait = Modal%
End Function


Public Function AOLOpenMail(which$)
'Opens the mail box.
'Returns the handle of _AOL_TREE
If UCase(which$) = "NEW" Then RunMenu 1, 0
If UCase(which$) = "OLD" Then RunMenu 3, 4
If UCase(which$) = "SENT" Then RunMenu 3, 5
Dim AoL%
Dim mdi%
Dim MailBox%
Dim TabControl%
Dim TabPageNew%
Dim TabPageOld%
Dim TabPageSent%
Dim TreeNew%
Dim TreeOld%
Dim TreeSent%
Do
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
MailBox% = FindChildByTitle(mdi%, "Fairite's Online MailBox")
TabControl% = FindChildByClass(MailBox%, "_AOL_TabControl")
TabPageNew% = FindChildByClass(TabControl%, "_AOL_TabPage")
TabPageOld% = GetWindow(TabPageNew%, GW_HWNDNEXT)
TabPageSent% = GetWindow(TabPageOld%, GW_HWNDNEXT)
TreeNew% = FindChildByClass(TabPageNew%, "_AOL_TREE")
TreeOld% = FindChildByClass(TabPageOld%, "_AOL_TREE")
TreeSent% = FindChildByClass(TabPageSent%, "_AOL_TREE")
Loop Until TreeSent% > 0
If UCase(which$) = "NEW" Then AOLOpenMail = TreeNew%
If UCase(which$) = "OLD" Then AOLOpenMail = TreeOld%
If UCase(which$) = "SENT" Then AOLOpenMail = TreeSent%
End Function

Public Function AOLGetMail(PersonSubjectorDate$, Index%, Tree%)
'PersonSubjectOrDate$ can be "PERSON", "SUBJECT", or "DATE
'index% is the mail number to get i couldn't
'do this like addroom cuz there can be commas
'in the subject text
'Tree%=the handle of _AOL_TREE
Dim length%
Dim MailString$
Dim MailDate$
Dim MailPerson$
Dim MailSubject$
If Tree% = 0 Then Exit Function
MailString$ = String(500, " ")
length% = SendMessageByString(Tree%, LB_GETTEXT, Index%, MailString$)
MailString$ = Left$(MailString$, length%)
MailDate$ = Left$(MailString$, InStr(MailString$, Chr$(9)) - 1)
MailString$ = Right(MailString$, Len(MailString) - Len(MailDate$) - 1)
MailPerson$ = Left$(MailString$, InStr(MailString$, Chr(9)) - 1)
MailSubject$ = Right(MailString$, Len(MailString$) - Len(MailPerson$) - 1)
If UCase(PersonSubjectorDate$) = "DATE" Then AOLGetMail = MailDate$
If UCase(PersonSubjectorDate$) = "PERSON" Then AOLGetMail = MailPerson$
If UCase(PersonSubjectorDate$) = "SUBJECT" Then AOLGetMail = MailSubject$
End Function


Public Sub AOLMailPref(Action$, chek1%, chek2%, chek3%, chek4%, chek5%, chek6%, chek7%, chek8%)
'Action$ can be "RESTORE", or "SET"
'Values of the chek's can be 1 for checked or
'0 for unchecked
Dim Preferences%
Dim Check1%
Dim Check2%
Dim Check3%
Dim Check4%
Dim Check5%
Dim Check6%
Dim Check7%
Dim Check8%
Dim X%
Dim OkButton%
RunMenu 3, 7
Do
Preferences% = FindWindow("_AOL_MODAL", "Mail Preferences")
Loop Until Preferences% > 0
Do
Check1% = GetWindow(Preferences%, GW_CHILD)
Check2% = GetWindow(Check1%, GW_HWNDNEXT)
Check3% = GetWindow(Check2%, GW_HWNDNEXT)
Check4% = GetWindow(Check3%, GW_HWNDNEXT)
Check5% = GetWindow(Check4%, GW_HWNDNEXT)
Check6% = GetWindow(Check5%, GW_HWNDNEXT)
Check7% = GetWindow(Check6%, GW_HWNDNEXT)
Check8% = GetWindow(Check7%, GW_HWNDNEXT)
Loop Until Check8% > 0
If UCase(Action$) = "SET" Then
MailSettings(1) = SendMessage(Check1%, BM_GETCHECK, 0, 0)
MailSettings(2) = SendMessage(Check2%, BM_GETCHECK, 0, 0)
MailSettings(3) = SendMessage(Check3%, BM_GETCHECK, 0, 0)
MailSettings(4) = SendMessage(Check4%, BM_GETCHECK, 0, 0)
MailSettings(5) = SendMessage(Check5%, BM_GETCHECK, 0, 0)
MailSettings(6) = SendMessage(Check6%, BM_GETCHECK, 0, 0)
MailSettings(7) = SendMessage(Check7%, BM_GETCHECK, 0, 0)
MailSettings(8) = SendMessage(Check8%, BM_GETCHECK, 0, 0)
X% = SendMessage(Check1%, BM_SETCHECK, chek1%, 0)
X% = SendMessage(Check2%, BM_SETCHECK, chek2%, 0)
X% = SendMessage(Check3%, BM_SETCHECK, chek3%, 0)
X% = SendMessage(Check4%, BM_SETCHECK, chek4%, 0)
X% = SendMessage(Check5%, BM_SETCHECK, chek5%, 0)
X% = SendMessage(Check6%, BM_SETCHECK, chek6%, 0)
X% = SendMessage(Check7%, BM_SETCHECK, chek7%, 0)
X% = SendMessage(Check8%, BM_SETCHECK, chek8%, 0)
End If
If UCase(Action$) = "RESTORE" Then
X% = SendMessage(Check1%, BM_SETCHECK, MailSettings(1), 0)
X% = SendMessage(Check2%, BM_SETCHECK, MailSettings(2), 0)
X% = SendMessage(Check3%, BM_SETCHECK, MailSettings(3), 0)
X% = SendMessage(Check4%, BM_SETCHECK, MailSettings(4), 0)
X% = SendMessage(Check5%, BM_SETCHECK, MailSettings(5), 0)
X% = SendMessage(Check6%, BM_SETCHECK, MailSettings(6), 0)
X% = SendMessage(Check7%, BM_SETCHECK, MailSettings(7), 0)
X% = SendMessage(Check8%, BM_SETCHECK, MailSettings(8), 0)
End If
OkButton% = FindChildByClass(Preferences%, "_AOL_ICON")
aolclick OkButton%
End Sub

Public Function AOLLocate(Person$, Room$, Section$)
'Locates some one online.
'Returns Status, Room, and, Section
'The status can be: NOTONLINE, ONLINEPRIVATE,
'ONLINEMEMBER,ONLINEPUBLIC, or ONLINECONFERENCE
'Room will be the room if member,public or conference
'Section will be the catagory if member, or public
Dim AoL%
Dim mdi%
Dim locate%
Dim LocateEdit%
Dim LocateButton%
Dim LocateStatic%
Dim StaticTextLength%
Dim Modal%
Dim StaticText$
Dim X%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
RunMenu 10, 10
Do: DoEvents
locate% = FindChildByTitle(mdi%, "Locate Member Online")
LocateEdit% = FindChildByClass(locate%, "_AOL_EDIT")
LocateButton% = FindChildByClass(locate%, "_AOL_ICON")
Loop Until LocateButton% > 0
X% = SendMessageByString(LocateEdit%, WM_SETTEXT, 0, Person$)
aolclick LocateButton%
X% = SendMessage(locate%, WM_CLOSE, 0, 0)
StaticText$ = String(100, " ")
Do
Modal% = FindWindow("#32770", "America Online")
locate% = FindChildByTitle(mdi%, "Locate " + Person$)
LocateStatic% = FindChildByClass(locate%, "_AOL_STATIC")
StaticTextLength% = SendMessageByString(LocateStatic%, WM_GETTEXT, 100, StaticText$)
If Modal% > 0 Then
X% = SendMessage(Modal%, WM_CLOSE, 0, 0)
AOLLocate = "NOTONLINE"
Exit Function
End If
Loop Until StaticTextLength% > 0
X% = SendMessage(locate%, WM_CLOSE, 0, 0)
StaticText$ = Left(StaticText$, StaticTextLength%)
If InStr(StaticText$, "but not in a chat area") Then
AOLLocate = "ONLINENOCHAT"
Exit Function
End If
If InStr(StaticText$, "private") Then
AOLLocate = "ONLINEPRIVATE"
Exit Function
End If
If InStr(StaticText$, "member") Then
AOLLocate = "ONLINEMEMBER"
StaticText$ = Right$(StaticText$, Len(StaticText$) - InStr(StaticText$, Chr$(34)))
Room$ = Left$(StaticText$, InStr(StaticText$, Chr$(34)) - 1)
StaticText$ = Right$(StaticText$, Len(StaticText$) - Len(Room$) - 22)
Section$ = Left(StaticText$, Len(StaticText$) - 3)
Exit Function
End If
If InStr(StaticText$, "public") Then
AOLLocate = "ONLINEPUBLIC"
StaticText$ = Right$(StaticText$, Len(StaticText$) - InStr(StaticText$, Chr$(34)))
Room$ = Left$(StaticText$, InStr(StaticText$, Chr$(34)) - 1)
StaticText$ = Right$(StaticText$, Len(StaticText$) - Len(Room$) - 22)
Section$ = Left(StaticText$, Len(StaticText$) - 3)
Exit Function
End If
If InStr(StaticText$, "conference") Then
AOLLocate = "ONLINECONFERENCE"
StaticText$ = Right$(StaticText$, Len(StaticText$) - InStr(StaticText$, Chr$(34)))
Room$ = Left$(StaticText$, Len(StaticText$) - 3)
Exit Function
End If
AOLLocate = StaticText$
End Function

Public Function AOLProfile(Person$)
'Returns a members profile
Dim AoL%
Dim mdi%
Dim locate%
Dim LocateEdit%
Dim LocateIcon%
Dim Profile%
Dim ProfileView%
Dim ProfileInfo$
Dim ProfileInfoLength%
Dim Modal%
Dim X%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
RunMenu 10, 11
Do
locate% = FindChildByTitle(mdi%, "Get a Member's Profile")
LocateEdit% = FindChildByClass(locate%, "_AOL_EDIT")
LocateIcon% = FindChildByClass(locate%, "_AOL_ICON")
Loop Until locate% > 0 And LocateEdit% > 0 And LocateIcon% > 0
X% = SendMessageByString(LocateEdit%, WM_SETTEXT, 0, Person$)
aolclick LocateIcon%
Do
Modal% = FindWindow("#32770", "America Online")
If Modal% > 0 Then
X% = SendMessage(Modal%, WM_CLOSE, 0, 0)
Do
X% = SendMessage(locate%, WM_CLOSE, 0, 0)
locate% = FindChildByTitle(mdi%, "Get a Member's Profile")
Loop Until locate% = 0
Exit Function
End If
Profile% = FindChildByTitle(mdi%, "Member Profile")
ProfileView% = FindChildByClass(Profile%, "_AOL_VIEW")
Loop Until Profile% > 0 And ProfileView% > 0
Do
ProfileInfo$ = String(10000, " ")
ProfileInfoLength% = SendMessageByString(ProfileView%, WM_GETTEXT, 1000, ProfileInfo$)
Timeout 0.1
ProfileInfo$ = String(10000, " ")
X% = SendMessageByString(ProfileView%, WM_GETTEXT, 1000, ProfileInfo$)
Loop Until ProfileInfoLength% = X%
AOLProfile = Left$(ProfileInfo$, ProfileInfoLength%)
X% = SendMessage(Profile%, WM_CLOSE, 0, 0)
Do
X% = SendMessage(locate%, WM_CLOSE, 0, 0)
locate% = FindChildByTitle(mdi%, "Get a Member's Profile")
Loop Until locate% = 0
End Function



Public Function RandomNumber(MaxNum%)
'Generates a random number not bigger than MaxNum%
Randomize
RandomNumber = Int(Rnd * MaxNum%) + 1
End Function
Function ReplaceText(thestring$, StringToFind$, ReplaceString$)
'Replaces one text in a string with another text
Dim StartPos%
Dim Beggining$
Dim ending$
StartPos% = InStr(UCase(thestring$), UCase(StringToFind$))
Beggining$ = Left(thestring$, StartPos% - 1)
If Len(thestring$) - (StartPos% + Len(StringToFind$)) + 1 = 0 Then ReplaceText = Beggining$ + ReplaceString$
ending$ = Right(thestring$, Len(thestring$) - (StartPos% + Len(StringToFind$)) + 1)
ReplaceText = Beggining$ + ReplaceString$ + ending$
End Function

Public Function SlideText(TheText$)
'Takes the last character of thetext$ and puts it at the beggining
SlideText = Right(TheText$, 1) + Left(TheText$, Len(TheText$) - 1)
End Function

Function ScrambleText(ByVal TheText$)
Dim StringToScramble$
Dim NumberRecord$
Dim Place%
Dim X%
Dim ScrambledString$
Do
NumberRecord$ = ""
ScrambledString$ = ""
If InStr(TheText$, " ") = 0 Then
StringToScramble$ = TheText$
TheText$ = ""
Else
StringToScramble$ = Left$(TheText$, InStr(TheText$, " "))
TheText$ = Right(TheText$, Len(TheText$) - Len(StringToScramble$))
StringToScramble$ = Left$(StringToScramble$, Len(StringToScramble$) - 1)
End If
For X% = 1 To Len(StringToScramble$)
Do
Place% = RandomNumber(Len(StringToScramble$))
Loop Until FindTextInString(NumberRecord$, Str(Place%)) = False
NumberRecord$ = NumberRecord$ + Str(Place%) + ","
ScrambledString$ = ScrambledString$ + Mid(StringToScramble$, Place%, 1)
Next X%
ScrambleText = ScrambleText + ScrambledString$ + " "
Loop Until Len(StringToScramble$) = 0
ScrambleText = Left(ScrambleText, Len(ScrambleText) - 1)
End Function


Public Function FindTextInString(ByVal TheText$, ByVal WhatToFind$)
'Finds a certain string in a string-list format
'If finds a match the function returns true
'This is different from using instr, because it matches a whole
'string. For example if your list was Jack,Jill,Bill and you searched
'for Jac this function will return false
TheText$ = UCase(TheText$)
WhatToFind$ = UCase(WhatToFind$)
If TheText$ = WhatToFind$ Then
FindTextInString = True
Exit Function
End If
If InStr(TheText$, "," + WhatToFind$ + ",") Then
FindTextInString = True
Exit Function
End If
If Left(TheText$, Len(WhatToFind$) + 1) = WhatToFind$ + "," Then
FindTextInString = True
Exit Function
End If
If Right$(TheText$, Len(WhatToFind$) + 1) = "," + WhatToFind$ Then
FindTextInString = True
Exit Function
End If
End Function



Public Function RGB2HEX(R%, G%, B%)
'Converts the rgb values to the aol hexadecimal
'format.
Dim X%
Dim xx%
Dim Color%
Dim Divide
Dim Answer%
Dim Remainder%
Dim Configuring$
For X% = 1 To 3
If X% = 1 Then Color% = B%
If X% = 2 Then Color% = G%
If X% = 3 Then Color% = R%
For xx% = 1 To 2
Divide = Color% / 16
Answer% = Int(Divide)
Remainder% = (10000 * (Divide - Answer%)) / 625
If Remainder% < 10 Then Configuring$ = Str(Remainder%) + Configuring$
If Remainder% = 10 Then Configuring$ = "A" + Configuring$
If Remainder% = 11 Then Configuring$ = "B" + Configuring$
If Remainder% = 12 Then Configuring$ = "C" + Configuring$
If Remainder% = 13 Then Configuring$ = "D" + Configuring$
If Remainder% = 14 Then Configuring$ = "E" + Configuring$
If Remainder% = 15 Then Configuring$ = "F" + Configuring$
Color% = Answer%
Next xx%
Next X%
Configuring$ = Chr$(34) + "#" + RemoveSpace(Configuring$) + Chr$(34)
RGB2HEX = Configuring$
End Function

Public Function blender(thestring$, ByVal ROne%, ByVal GOne%, ByVal bone%, ByVal RTwo%, ByVal GTwo%, ByVal BTwo%)
Dim Index%
Dim newred
Dim NewGreen
Dim NewBlue
Dim RedOne
Dim GreenOne
Dim BlueOne
Dim RedMinus
Dim GreenMinus
Dim BlueMinus
If ROne% > RTwo% Then
RedOne = True
RedMinus = (ROne% - RTwo%) / Len(thestring$)
Else
RedOne = False
RedMinus = (RTwo% - ROne%) / Len(thestring$)
End If
If GOne% > GTwo% Then
GreenOne = True
GreenMinus = (GOne% - GTwo%) / Len(thestring$)
Else
GreenOne = False
GreenMinus = (GTwo% - GOne%) / Len(thestring$)
End If
If bone% > BTwo% Then
BlueOne = True
BlueMinus = (bone% - BTwo%) / Len(thestring$)
Else
BlueOne = False
BlueMinus = (BTwo% - bone%) / Len(thestring$)
End If
newred = ROne%
NewGreen = GOne%
NewBlue = bone%
blender = blender + "<font color=" + RGB2HEX(Int(newred), Int(NewGreen), Int(NewBlue)) + ">" + Mid(thestring$, Index% + 1, 1)
For Index% = 2 To Len(thestring$) - 1
If RedOne = True Then newred = newred - RedMinus Else newred = newred + RedMinus
If GreenOne = True Then NewGreen = NewGreen - GreenMinus Else NewGreen = NewGreen + GreenMinus
If BlueOne = True Then NewBlue = NewBlue - BlueMinus Else NewBlue = NewBlue + BlueMinus
blender = blender + "<font color=" + RGB2HEX(Int(newred), Int(NewGreen), Int(NewBlue)) + ">" + Mid(thestring$, Index%, 1)
Next Index%
If Len(thestring$) > 1 Then blender = blender + "<font color=" + RGB2HEX(RTwo%, GTwo%, BTwo%) + ">" + Right(thestring$, 1)
End Function

Sub fillchat()
Dim kll1 As String
Dim kll2 As String
Dim kll3 As String
Dim kll4 As String
kll3 = " <Font Color=#FFFFFE><pre="
kll1 = "——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————"
kll4 = "——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————1© Iz Your God"
kll2 = " <Font Color=#FFFFFE><pre="

AolSendtext (kll3 + kll1 + kll4)
Timeout 0.5
AolSendtext (kll2 + kll1 + kll4)
Timeout 0.5

End Sub






Function HyperLink(txt As String, URL As String)
HyperLink = ("<A HREF=" & Chr$(34) & URL & Chr$(34) & ">" & txt & "</A>")
End Function
Sub attentioner(Text$)
Dim att As String
att = ("<HTML><FONT  COLOR=#e80000 SIZE=3 PTSIZE=10><B>Â</FONT><FONT  COLOR=#d10000 SIZE=3 PTSIZE=10>†</FONT><FONT  COLOR=#ba0000 SIZE=3 PTSIZE=10>†</FONT><FONT  COLOR=#a30000 SIZE=3 PTSIZE=10>ê</FONT><FONT  COLOR=#8c0000 SIZE=3 PTSIZE=10>ñ</FONT><FONT  COLOR=#750000 SIZE=3 PTSIZE=10>†</FONT><FONT  COLOR=#5e0000 SIZE=3 PTSIZE=10>î</FONT><FONT  COLOR=#470000 SIZE=3 PTSIZE=10>ø</FONT><FONT  COLOR=#300000 SIZE=3 PTSIZE=10>ñ</B></HTML>")
Call AolSendtext("<HTML><s><FONT  COLOR=#280000 SIZE=3 PTSIZE=10 FAMILY=DECORATIVE FACE=Stop LANG=0><B>| || ||| |||| ||||| |||||| </s>" + att + " <FONT  COLOR=#280000 SIZE=3 PTSIZE=10 FAMILY=DECORATIVE FACE=Stop LANG=0><s> ||||||| |||||| ||||| |||| ||| || |</B></s></HTML> {S Im")
Call Timeout(0.5)
Call Fader4(Text$)
Call Timeout(0.5)
Call AolSendtext("<HTML><s><FONT  COLOR=#280000 SIZE=3 PTSIZE=10 FAMILY=DECORATIVE FACE=Stop LANG=0><B>| || ||| |||| ||||| |||||| </s>" + att + " <FONT  COLOR=#280000 SIZE=3 PTSIZE=10 FAMILY=DECORATIVE FACE=Stop LANG=0><s> ||||||| |||||| ||||| |||| ||| || |</B></s></HTML> {S Im")
End Sub



Sub Fader(texttofade$)
Dim color1 As Color
Dim color2 As Color
Dim color3 As Color
Dim color4 As Color
color1.Green = 255
color2.Blue = 0
color3.Green = 125
color4.Blue = 500
AolSendtext (AOLColor(texttofade$, False, False, False, "BLEND", "", color1, color2, color3, color4, False, False))
End Sub




Function sys_time() As String
'  text2.text = "It is Currently " + sys_timeanddate()
sys_time$ = Format$(Now, "h:mm:ss AM/PM ")
End Function










Function EncryptText(WhereToType As TextBox, Hidden As TextBox)

Dim a%, X%, Y%, z%, i%, Temp, Pharse$
Dim T As String
If WhereToType.Text = "" Then
Exit Function
Else
WhereToType.Enabled = False
Hidden.Text = ""
Hidden.Text = WhereToType.Text
WhereToType.Text = ""
Pharse$ = Hidden.Text
For i = 1 To Len(Pharse$)
      Temp = Asc(Mid$(Pharse$, i, 1))
      Mid$(Pharse$, i, 1) = Chr$(Abs(Temp - 255))
T$ = T$ + Mid$(Pharse$, i, 1)
Next i
T$ = ""
Pharse$ = Hidden.Text
Hidden.Text = ""
Hidden.Text = WhereToType.Text
WhereToType.Text = ""
For i = 1 To Len(Pharse$)
      Temp = Asc(Mid$(Pharse$, i, 1))
      Mid$(Pharse$, i, 1) = Chr$(Abs(Temp - 255))
T$ = T$ + Mid$(Pharse$, i, 1)
Next i
WhereToType.Enabled = True
WhereToType.Text = T$
End If
End Function





Sub ascii()
Call Fader4(" ×´¯)")
Timeout (0.5)
Call Fader(" ¸.·´ ¸.·´¯)")
Timeout (0.5)
Call Fader4("(_.·´     (_.·×")
Call Fader("`·._.·º·._.·×")

Call Fader(".·´¯·.¸.·´¯`·×")
End Sub

Function FindChildByClass2(parentw, childhand)
Dim firs%
Dim firss%
Dim Room%
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
FindChildByClass2 = 0

bone:
Room% = firs%
FindChildByClass2 = Room%

End Function



Sub AddRoom(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Dim Room
Dim aolhandle
Dim aolthread
Dim AOLProcessThread
Dim Index


Room = AOLFindChatRoom()
aolhandle = FindChildByClass2(Room, "_AOL_Listbox")

aolthread = GetWindowThreadProcessId(aolhandle, AOLProcess)
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

Function BackwardsString(strng$)
'Reverses a string
Dim InputString$
Dim length%
Dim Spaces%
Dim NextLetter$
Dim newstring$
InputString$ = strng$
length% = Len(InputString$)
Do While Spaces% <= length%
Spaces% = Spaces% + 1
NextLetter$ = Mid$(InputString$, Spaces%, 1)
newstring$ = NextLetter$ & newstring$
Loop
BackwardsString = newstring$
End Function



Public Sub String2List(ByVal thestring$, List1)
'String must be in this form:
'Item1,Item2,Item3,Item4,etc...
Dim ThisItem$
Do
If InStr(thestring$, ",") = 0 Then GoTo Done
ThisItem$ = Left$(thestring$, InStr(thestring$, ",") - 1)
thestring$ = Right$(thestring$, Len(thestring$) - InStr(thestring$, ","))
List1.AddItem (ThisItem$)
Loop Until ThisItem$ = ""
Done:
List1.AddItem (thestring$)
End Sub

Function phishphrases()
Dim phish As String
Dim num As String
num = RandomNumber(8)
If num = 1 Then
phish = ("Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You.")
End If
If num = 1 Then
phish = ("Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you.")
End If
If num = 1 Then
phish = ("Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience.")
End If
If num = 1 Then
phish = ("Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience.")
End If
If num = 1 Then
phish = ("Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff")
End If
If num = 1 Then
phish = ("This is a message sent out to random America Online members for a catalog containing -cheap- Microsoft software. Please send your request for the catalog along with your credit card number fo the amount $19.95. After 3 weeks processing time you'll recive a copy of Microsoft Windows96 Nashville version. *Microsoft INC.*")
End If
If num = 1 Then
phish = ("Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL.")
End If
If num = 1 Then
phish = ("Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you.")
End If
End Function

Function slowtalk(strng$)
Dim InputString$
Dim length%
Dim Spaces%
Dim NextLetter$
Dim newstring$
InputString$ = strng$
length% = Len(InputString$)
Do While Spaces% <= length%
Spaces% = Spaces% + 1
NextLetter$ = Mid$(InputString$, Spaces%, 1)
NextLetter$ = NextLetter$ + "</html><html>"
newstring$ = newstring$ + NextLetter$
Loop
slowtalk = Left$(newstring$, Len(newstring$))
End Function

Function LastChatLineWithSN()
Dim chattext$
Dim thechar$
Dim thechars$
Dim thechattext$
Dim FindChar
Dim lastlen
Dim LastLine
chattext$ = GetChatText

For FindChar = 1 To Len(chattext$)

thechar$ = Mid(chattext$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechattext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(chattext$, lastlen, Len(thechars$))

LastChatLineWithSN = LastLine
End Function
Function LastChatLine()
Dim chattext
Dim ChatTrimNum
Dim ChatTrim$

chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Function SNFromLastChatLine()
Dim chattext$
Dim ChatTrim$
Dim z As ListBoxConstants
Dim SN
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For z = 1 To 11
    If Mid$(ChatTrim$, z, 1) = ":" Then
        SN = Left$(ChatTrim$, z - 1)
    End If
Next z
SNFromLastChatLine = SN
End Function

Function GetChatText()
Dim Room%
Dim aorich%
Dim chattext
Room% = AOLFindChatRoom
aorich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(aorich%)
GetChatText = chattext
End Function

Function GetText(child)
Dim GetTrim
Dim TrimSpace$
Dim getstring
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function EliteText(word$)
Dim made$
Dim q
Dim X
Dim letter$
Dim leet$
made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = letter$
    made$ = made$ & leet$
Next q

EliteText = made$

End Function


Function UserSN()
On Error Resume Next
Dim AoL%
Dim mdi%
Dim welcome%
Dim welcomelength%
Dim welcometitle$
Dim a%
Dim user
AoL% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(AoL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
welcomelength% = GetWindowTextLength(welcome%)
welcometitle$ = String$(200, 0)
a% = GetWindowText(welcome%, welcometitle$, (welcomelength% + 1))
user = Mid$(welcometitle$, 10, (InStr(welcometitle$, "!") - 10))
UserSN = user
End Function

Sub playwave(File)
Dim SoundName$
Dim valueflags%
Dim X%
SoundName$ = File
valueflags% = SND_ASYNC Or SND_NODEFAULT
X% = sndPlaySoundA(SoundName$, valueflags%)
End Sub



Function GetClass(child)
Dim Buffer$
Dim getclas%
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Sub AOL4_KillWin(Windo)
Dim closethemofo
'Closes a window....ex: AOL4_Killwin (IM%)
closethemofo = SendMessage(Windo, WM_CLOSE, 0, 0)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
'turns the "number" so vb recognizes it for
'addition, subtraction, ect.

End Function
Sub BlueFade(vForm As Object)
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

Sub GreenFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub RedFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub FireFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub
Sub PlatinumFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub IceFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub


Function INI_GetFrom(AppName$, Keyname$, FileName$) As String
Dim GetFromINI
'ok, in form load put this     iniPath$ = App.Path + "\ProgName.ini"
'ok, in ur options or whatever put this....
'entry$ = text2.text
'R% = WritePrivateProfileString("Text1=", "<—1©", entry$, iniPath$)
'If R% <> 1 Then MsgBox "An error occurred while writing Name."




   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal Keyname$, "", RetStr, Len(RetStr), FileName$))
End Function




Sub Form_ScrollGreets(vLabel As Label)
Dim vlabel1
vlabel1.Top = Val(vlabel1.Top) - 8
End Sub

Sub password(txt As TextBox)
Dim pw
pw = "1c"
If Not txt = pw Then
MsgBox "Invalid Password try again", vbOK, "UnKnown"
End If
If txt = pw Then
MsgBox "Right Password"
'put form.show here
End If
End Sub

Public Sub Form_Effect(Form As Form, How$, HowFar%, Delay%)
'Form = which form to effect
'Correct values of How$ are:
'RollUp, RollDown, SlideLeft, SlideRight,
'SlideDown, SlideUp, and Center
'HowFar% vary's according to your choice:
'RollUp - how far to roll up
'RollDown - how far to roll down
'SlideLeft - how far to slide left
'SlideRight - how far to slide right
'SlideDown - how far to slide down
'SlideUp - how far to slide up
'Center - value is ignored
'Move - drags the form when you put in
'       the form_mousedown
'Delay% = how slow, bigger the slower
On Error Resume Next
Dim X%
Dim xx%
Select Case UCase$(How$)
   Case "ROLLUP"
      Form.Show
      For X% = 0 To HowFar%
      DoEvents
      For xx% = 1 To Delay% * 100
      Next xx%
      Form.Height = Form.Height - 1
      Next X%
   Case "ROLLDOWN"
      Form.Show
      For X% = 0 To HowFar%
      DoEvents
      For xx% = 1 To Delay% * 100
      Next xx%
      Form.Height = Form.Height + 1
      Next X%
   Case "CENTER"
      With Form
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
      Form.Top = (Screen.Height - Form.Height) / 2
      Form.Left = (Screen.Width - Form.Width) / 2
   Case "SLIDEUP"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Top = Form.Top - 1
      Next X%
   Case "SLIDEDOWN"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Top = Form.Top + 1
      Next X%
   Case "SLIDERIGHT"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Left = Form.Left + 1
      Next X%
   Case "SLIDELEFT"
      Form.Show
      For X% = 0 To HowFar%
      For xx% = 1 To Delay% * 100
      Next xx%
      DoEvents
      Form.Left = Form.Left - 1
      Next X%
End Select
End Sub

Sub PopUpMenu(vForm As Form)
'ok ... do not put call popupmenu(form1)!!!!!!!!!!!
'put the code in a butten
'so if it would be form1 put Form.PopUpMenu Form.mnufile, 0 in a butten!
vForm.PopUpMenu vForm.mnufile, 0

End Sub


Sub FadeFormBlue1(vForm As Form)

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
Sub FadeFormGreen1(vForm As Form)
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
Sub FadeFormGrey1(vForm As Form)
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
Sub FadeFormPurple1(vForm As Form)
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
Sub FadeFormRed1(vForm As Form)
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
Sub FadeFormYellow1(vForm As Form)
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

Function Percentage(Complete As Integer, total As Integer, TotalOutput As Integer) As Integer
R = Int(Complete / total * TotalOutput)
Percentage = R
End Function
Sub remove(List As ListBox)
If List.ListIndex = -1 Then Exit Sub
List.RemoveItem List.ListIndex
End Sub

Sub RemoveList(List As ListBox, Who$)
Dim X
On Error Resume Next
For X = 0 To List.ListCount - 1
If UCase(Who$) = UCase(List.List(X)) Then List.RemoveItem X: Exit Sub
Next X
End Sub
Sub ScrollMacro(Text As String)
Dim counter
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    counter = counter + 1
    SendText Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If counter = 4 Then
        Timeout (2.9)
        counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub

Function FadeToTop(Frm As Form, frm2 As Form)

Do
Frm.Top = Frm.Top - 25
Frm.Width = Frm.Width - 25
Frm.Height = Frm.Height - 25
Frm.Left = Screen.Width / 2 - Frm.Width / 2
Loop Until Frm.Top < 200
frm2.Show
Frm.Hide
StayOnTop frm2
End Function
Function Form_Up(Frm As Form)

Do
Frm.Top = Frm.Top - 25
Loop Until Frm.Top < 200
End Function
Function Form_Down(Frm As Form, frm2 As Form, hi, wi)
Do
Frm.Top = Frm.Top + 25
Frm.Width = Frm.Width + 20
Frm.Height = Frm.Height + 20
Frm.Left = Screen.Width / 2 - Frm.Width / 2
Loop Until Frm.Top > Screen.Height / 2 - Frm.Height / 2
Unload Frm
frm2.Height = hi
frm2.Width = wi

End Function
Sub Form_Maximize(Frm As Form)

Frm.WindowState = 2
End Sub

Public Sub String2Combo(thestring$, TheCombo As ComboBox)
Dim ThisItem$
Do
If InStr(thestring$, ",") = 0 Then GoTo Done
ThisItem$ = Left$(thestring$, InStr(thestring$, ",") - 1)
thestring$ = Right$(thestring$, Len(thestring$) - InStr(thestring$, ",") - 1)
TheCombo.AddItem ThisItem$
Loop Until ThisItem$ = ""
Done:
TheCombo.AddItem thestring$
End Sub
Function Trm(ByVal txt As String)
Dim X As Integer
Dim Y As String
Dim z As String
For X = 1 To Len(txt)
Y = Mid(txt, X, 1)
If Y = Chr(0) Then Y = ""
z = z & Y
Next X
Trm = z
End Function
Sub Form_Minimize(Frm As Form)

Frm.WindowState = 1
End Sub

Sub Form_Move2(Frm As Form)
Dim ReturnVal%
ReleaseCapture
ReturnVal% = SendMessage(Frm.hwnd, &HA1, 2, 0)

End Sub

Public Sub Form_CenterTop(Frm As Form)
'This will center the form in the top center
   With Frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub

Public Sub Form_TopRight(frmForm As Form)
'This centerz a form in the top right of the screen
   With frmForm
      .Left = (Screen.Width - .Width) / 1
      .Top = (Screen.Height - .Height) / 2000
   End With
End Sub
Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function
Sub NotOnTop(The As Form)
Dim setwinontop
setwinontop = SetWindowPos(The.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Wav_Loop(File$)
'Plays a wav and loop's it
Dim i%
i% = sndPlaySound(File$, SND_ASYNC Or SND_LOOP)

End Sub

Sub Wav_Play(File$)
Dim i%
'this plays a wav/mdi

i% = sndPlaySound(File$, 0)

End Sub

Public Sub center(frmForm As Form)
69   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub
Sub Form_Color(Frm As Form)
' Place this in a timer and watch the colors =)
Dim cx, cy, Radius, Limit
    Frm.ScaleMode = 3
    cx = Frm.ScaleWidth / 2
    cy = Frm.ScaleHeight / 2
    If cx > cy Then Limit = cy Else Limit = cx
    For Radius = 0 To Limit
Frm.Circle (cx, cy), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next Radius
End Sub
Sub Pause(interval)
Dim current
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Sub DeleteFile(FileName$)

If Not IFileExists(FileName$) Then MsgBox FileName$ & Chr(13) & "Bad File Name!", 16, "Error": Exit Sub
On Error GoTo ErrorInDeletion
Kill FileName$
Exit Sub
ErrorInDeletion:
MsgBox Error$
Resume Exitinga
Exitinga:
Exit Sub
End Sub

Sub CopyFile(FileName$, CopyTo$)

If FileName$ = "" Then Exit Sub
If CopyTo$ = "" Then Exit Sub
If Not IFileExists(FileName$) Then Exit Sub
On Error GoTo AnErrOccured
If InStr(Right$(FileName$, 4), ".") = 0 Then Exit Sub
If InStr(Right$(CopyTo$, 4), ".") = 0 Then Exit Sub
FileCopy FileName$, CopyTo$
Exit Sub
AnErrOccured:
MsgBox "An Unexpected Error Occured!", 16, "Error"
End Sub

Sub Explode(Frm As Form, CFlag As Integer, STEPS As Integer)

Dim FRect As RECT
Dim FWidth, FHeight As Integer
Dim i, X, Y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush
    
    GetWindowRect Frm.hwnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    FHeight = FRect.Bottom - FRect.Top
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(Frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    For i = 1 To STEPS
        cx = FWidth * (i / STEPS)
        cy = FHeight * (i / STEPS)
        If CFlag Then
            X = FRect.Left + (FWidth - cx) / 2
            Y = FRect.Top + (FHeight - cy) / 2
        Else
            X = FRect.Left
            Y = FRect.Top
        End If
        Rectangle hScreen, X, Y, X + cx, Y + cy
    Next i
    If ReleaseDC(0, hScreen) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    DeleteObject (Brush)
    Frm.Show
End Sub
Sub DitherForm(vForm As Form)
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
Sub File_OpenEXE(File$)
Dim openit!
Dim NoFreeze%
openit! = Shell(File$, 1): NoFreeze% = DoEvents()
End Sub
Sub File_ParseExe(File As String, txt As TextBox)
Dim filedata As String
Dim newstring As String
Open File For Binary As #1
Dim i
filedata = Space$(LOF(1))
Get #1, , filedata
Close #1
For i = 1 To Len(filedata)
filedata = newstring + Hex$(Mid(filedata, i)) & " "
' Add chr$(13) and 10 around here somewhere as needed
Next i
txt.Text = txt.Text + filedata
' Show the data
End Sub
Sub File_ReName(File$, NewName$)
Dim NoFreeze%
Name File$ As NewName$
NoFreeze% = DoEvents()
End Sub

Sub RenameFile(OldFile$, NewFile$)

If Not IFileExists(OldFile$) Then MsgBox "Bad File Name!", 16, "Error": Exit Sub
On Error GoTo TakeCareOfError
Name OldFile$ As NewFile$
Exit Sub
TakeCareOfError:
MsgBox "An unexpected error occured!", 16, "Error"
End Sub

Function IFileExists(ByVal sFilename As String) As Integer
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |Checks if a file you chose exists                   | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim TheFileLength As Integer
On Error Resume Next
TheFileLength = Len(dir$(sFilename))
If err Or TheFileLength = 0 Then
IFileExists = False
Else
IFileExists = True
End If
End Function
Function IfDirExists(TheDirectory)
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |Check's if Directory exsists on user's computer.    | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(dir$(TheDirectory))
If err Or Check = 0 Then
    IfDirExists = False
Else
    IfDirExists = True
End If
End Function
Sub File_RenameDirectory(old$, NewName$)
Dim NoFreeze%
If Not IfDirExists(old$) Then Exit Sub
Name old$ As NewName$
NoFreeze% = DoEvents()
End Sub

Sub File_Size(File$)
Dim exists
Dim err
Dim thelength
exists = Len(dir$(File$))
If err Or exists = 0 Then Exit Sub
thelength = FileLen(File$)
End Sub
Sub File_MakeDirectory(DirName$)
MkDir DirName$
End Sub
Sub File_Delete(File$)
Dim NoFreeze%
If Not IFileExists(File$) Then Exit Sub
Kill File$
NoFreeze% = DoEvents()
End Sub

Sub File_DeleteDirectory(DirName$)
If Not IfDirExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub

Sub File_Copy(File$, DestFile$)
If Not IFileExists(File$) Then Exit Sub
FileCopy File$, DestFile$
End Sub


Sub PWSScanner(FilePath$, FileName$, status As Label)
Dim PentiumRest%
Dim StopPWScanner
Dim TheFileLen, NumOne, GenOiZBack, GenOziDe, TheFileInfo$, PWS, PWS2, PWS3, VirusedFile, LengthOfFile, TotalRead, TheTab, TheMSg, TheMsg2, TheMsg3, TheMsg4, TheMsg5, TheDots
StopPWScanner = 0
If FileName$ = "" Then GoTo Errorr
FileName$ = FilePath$ & "\" & FileName$
If Right$(FilePath$, 1) = "\" Then FileName$ = FilePath$ & FileName$
If Not IFileExists(FileName$) Then MsgBox "File Not Found!", 16, "Error": GoTo Errorr
TheFileLen = FileLen(FileName$)
status.Caption = TheFileLen
NumOne = 1
GenOiZBack = 2
GenOziDe = 3
Do While GenOziDe > GenOiZBack
PentiumRest% = DoEvents()
If StopPWScanner = 1 Then GoTo Errorr
Open FileName$ For Binary As #1
If err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": GoTo Errorr
TheFileInfo$ = String(32000, 0)
Get #1, NumOne, TheFileInfo$
Close #1
Open FileName$ For Binary As #2
If err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": GoTo Errorr
PWS = InStr(1, LCase$(TheFileInfo$), "main.idx" + Chr(0), 1)
If PWS Then
Geno:
Mid(TheFileInfo$, PWS) = "GenOziDe  "
PWS2 = Mid(LCase$(TheFileInfo$), PWS + 8 + 1, 8)
PWS2 = Trm(PWS2)
PWS3 = Mid(LCase$(TheFileInfo$), PWS + 8 + 1 + Len(PWS), 1)
If PWS3 <> Chr(0) Then GoTo DeliriuM
If Len(PWS2) < 4 Then GoTo DeliriuM
If Len(PWS2) = "" Then GoTo DeliriuM
DeliriuM:
PWS = InStr(1, LCase$(TheFileInfo$), "main.idx" + Chr(0), 1)
If PWS <> 0 Then VirusedFile = FileName$: MsgBox VirusedFile & " is a Password Stealer!", 16, "Password Stealer": Close #2: Exit Sub
End If
TotalRead = TotalRead + 32000
status.Caption = Val(TotalRead)
LengthOfFile = LOF(2)
Close #2
If TotalRead > LengthOfFile Then: status.Caption = LengthOfFile: GoTo GOD
DoEvents
Loop
GOD:
TheTab = Chr$(9) & Chr$(9)
TheDots = "---------------------------------------------------------"
TheMSg = TheDots & Chr(13) & "File Information:" & Chr(13) & Chr(13)
TheMsg2 = TheMSg & FileName$ & " is clean from trojans." & Chr(13) & Chr(13)
TheMsg3 = TheMsg2 & FileName$ & " was scanned by GenOziDe." & Chr(13) & Chr(13)
TheMsg4 = TheMsg3 & "Scanned - 100% of - " & FileName$ & Chr(13) & Chr(13)
TheMsg5 = TheMsg3 & FileName$ & " is safe to use!" & Chr(13) & TheDots
MsgBox TheMsg5, 55, "File Is Clean!"
Errorr:
PentiumRest% = DoEvents()
status.Caption = ""
Close #1
PentiumRest% = DoEvents()
Close #2
PentiumRest% = DoEvents()
Exit Sub
End Sub

Sub Wav_Stop(File$)
Dim i%
i% = sndPlaySound(0&, 0)

End Sub


Sub getNum(og%, a)
Dim B
Do: DoEvents
If a = 0 Then Exit Sub
B = 1 + B
og% = GetWindow(og%, GW_HWNDNEXT)
Loop Until B >= a - 1

End Sub




Sub ClipBoard_CopyTo(What$)
Clipboard.SetText What$


End Sub





Function Readini(AppName, Keyname, FileName As String) As String

'This will read an INI file, good for making options on a prog

'Example: text55.text = ReadINI("Progs Name", "Intro Wave", app.path + "\Prog.ini")
Dim sRet As String
    sRet = String(255, Chr(0))
    Readini = Left(sRet, GetPrivateProfileString(AppName, ByVal Keyname, "", sRet, Len(sRet), FileName))



End Function
Function Text_Search(SearchFor, SearchThis)
Dim X
X = InStr(1, SearchThis, SearchFor)
Text_Search = X
End Function



Function INI_Read(AppName, Keyname As String, FileName As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    INI_Read = Left(sRet, GetPrivateProfileString(AppName, ByVal Keyname, "", sRet, Len(sRet), FileName))
End Function
Sub INI_Write(sAppname As String, sKeyName As String, sNewString As String, sFilename As String)
Dim ret&
ret& = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFilename)

End Sub
Public Sub ListDeleteItem(lst As ListBox, Item$)
Dim NoFreeze%
Dim a
Do:
NoFreeze% = DoEvents()
If LCase$(lst.List(a)) = LCase$(Item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Sub WriteToLog(What As String, FilePath As String)
Dim f%
Dim p$
If FilePath = "" Then Exit Sub
f% = FreeFile
Open FilePath For Binary Access Write As f%
p$ = What & Chr(10)
Put #1, LOF(1) + 1, p$
Close f%
End Sub
Sub Directory_Create(dir)

MkDir dir
End Sub

Sub Directory_Delete(dir)

RmDir (dir)
End Sub

Sub Form_Move(Frm As Form)
Dim ReturnVal%
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(Frm.hwnd, &HA1, 2, 0)
End Sub



Function deltreescan(Ffiillee As String)
Dim FileLenn As Variant
Dim FileLennn As Variant
Dim l003A As Variant
Dim l003E As Variant
Dim l0042 As String
Dim l0044 As Single
Dim l0046 As Single
Dim l0048 As Single
Dim l004a As Single
Dim l004c As Single
Dim l004e As Single
Dim l0050 As Single
Dim l0052 As Single
Dim l0054 As Single
Dim l0056 As Single
Dim l0058 As Single
Dim l005A As Variant
Dim l0045!
Open Ffiillee For Binary As #2
DoEvents
FileLenn = LOF(2)
FileLennn = FileLenn
l003A = 1
While FileLennn >= 0
    If FileLennn > 32000 Then
        l003E = 32000
    ElseIf FileLennn = 0 Then
        l003E = 1
    Else
        l003E = FileLennn
    End If
    l0042$ = String$(l003E, " ")
    Get #2, l003A, l0042$
    l0044! = InStr(1, l0042$, "deltree \y", 1)
    l0045! = InStr(1, l0042$, "MZÿ C:\*.*", 1)
If l0044! Then deltreescan = True
Close: Exit Function

If Not l0044! Then deltreescan = False
Close: Exit Function



Wend
End Function


Sub AOL4_IM_AutoAnswer(message)
Dim AoL%
Dim mdi%
Dim im%
Dim e
Dim e2
'Res'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")

im% = FindChildByTitle(mdi%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(mdi%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(im%, "RICHCNTL")

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
e2 = GetWindow(e, GW_HWNDNEXT) 'Send Text
e = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
AOL4_Icon (e)
Pause 2#
End Sub


Sub AOL4_Icon(Icon%)
Dim Click%
'Clickz on an AOL icon
Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function chat(Text1 As TextBox)
'This gets tha text
Dim rich%
Dim S$
rich% = FindChildByClass(AOLFindChatRoom, "RICHCNTL")
S$ = GetText(rich%)
S$ = ReplaceCharacter(S$, Chr(13), Chr(13) + Chr(10))
Text1 = S$
'This keeps the textbox scrolling down, take it out and u will see what i mean
Text1.SelStart = Len(Text1.Text)
'interval at 500
End Function


Function ReplaceCharacter(Text, charfind, charchange)
'Replaces the text what your own charectors
Dim replace
Dim thechar$
Dim thechars$
If InStr(Text, charfind) = 0 Then
ReplaceCharacter = Text
Exit Function
End If
For replace = 1 To Len(Text)
thechar$ = Mid(Text, replace, 1)
thechars$ = thechars$ & thechar$
If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next replace
ReplaceCharacter = thechars$
End Function

Function getsendboxtxt()
Dim AoL%
Dim mdi%

Dim listers%
Dim listere%
Dim nxt%
Dim nxt1%
Dim nxt2%
Dim nxt3%
Dim nxt4%
Dim firs%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = getnextwindow1(listers%, 2)
nxt% = getnextwindow1(listere%, 2)
nxt1% = getnextwindow1(nxt%, 2)
nxt2% = getnextwindow1(nxt1%, 2)
nxt3% = getnextwindow1(nxt2%, 2)
nxt4% = getnextwindow1(nxt3%, 2)
getsendboxtxt = GetText(nxt4%)
End Function
Sub clearchatbox()
Dim AoL%
Dim mdi%
Dim txt$
Dim listers%
Dim listere%
Dim nxt%
Dim nxt1%
Dim nxt2%
Dim nxt3%
Dim nxt4%
Dim firs%
AoL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AoL%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = getnextwindow1(listers%, 2)
nxt% = getnextwindow1(listere%, 2)
nxt1% = getnextwindow1(nxt%, 2)
nxt2% = getnextwindow1(nxt1%, 2)
nxt3% = getnextwindow1(nxt2%, 2)
nxt4% = getnextwindow1(nxt3%, 2)
txt$ = GetText(nxt4%)

Call SendMessageByString(nxt4%, WM_SETTEXT, 0, "")

End Sub

Function Find2ndChildByClass(parentw, childhand)
    Dim firs%
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


