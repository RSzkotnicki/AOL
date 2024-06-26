Attribute VB_Name = "FeaR32"
'FeaR32.Bas (A.K.A. - Zero G)
'Works with AOL4.0 and AOL3.0 for Win95
'This bas was made by FeaR.
'This will be updated every week
'Do not change the code.

'If you have anything to say, add,
'comment etc... Email me at -
'FeaR_ZeroG@Hotmail.com

'This .bas file and the program Zero G
'can be found on the site The Source-
'Address - http://thesource.winlabs.com/thesource/
'or at Zero G web site
'Address - Not up yet ... working on it

'I'd like to thank VvVCryoVvV for
'the text fades. THANX! Saved me some
'time.
'Also Thanx to all the P-ple that took
'the time to actually help me out,
'you all know who you are!

'Shouts to Linus! Neato!
'And anyone else i forgot!

'Subs & Functions starting with AOL_
'are for AOL3.0 although some may
'work for AOL4.0 but the ones starting
'with AOL4_ are for AOL4.0.

'-You can discover what your enemy
'fears most by observing the means
'he uses to frighten you.-

'-FeaR                  *Peace*

'I'm a very organized person.
'So all the Declarations and Const's
'are all in Alphabetical order ... Well
'almost =P

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String)

Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByValnIndex As Long, ByVal nNewWord As Long) As Long

Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage2 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function SendMessage3 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function SetSysModalWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Boolean, ByVal fuWinIni As Long) As Long

Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
        (ByVal hWnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
Global Const GFM_BACKSHADOW = 1
Global Const GFM_DROPSHADOW = 2

Public Const BM_SETCHECK = &HF1

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const LB_GETCOUNT = &H18B

Public Const SPI_SCREENSAVERRUNNING = 97

Global Const GWW_HWNDPARENT = (-8)

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2

Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SETTEXT = &HC

Const c00B0 = 2 '&H2%
Const c0100 = 258 ' &H102%
Const c00F6 = 513 ' &H201%
Const c00F8 = 514 ' &H202%
Const c014E = 12 ' &HC%
Const c008E = 14 ' &HE%
Const c0090 = 13 ' &HD%

Global AFKBack
Global FortyFiveMinKillStop
Global StopPWScanner

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

Dim vTextDummy$

Global Const WM_USER = &H400
Global Const EM_SETREADONLY = (WM_USER + 31)

Type RECT
       Left As Integer
       Top As Integer
       Right As Integer
       Bottom As Integer

End Type
Type POINTAPI
   x As Long
   y As Long
End Type

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

Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs%
FindChildByTitle = Room%
End Function

Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub AOL_SendChat(chat)
Room% = AOL_FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Function AOL_FindChatRoom()
MDI% = AOL_MDI()
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   AOL_FindChatRoom = Room%
Else:
   AOL_FindChatRoom = 0
End If
End Function

Function AOL_FindChatRoom()
MDI% = AOL_MDI()
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   AOL_FindChatRoom = Room%
Else:
   AOL_FindChatRoom = 0
End If
End Function

Function AOL_MDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL_MDI = FindChildByClass(AOL%, "MDIClient")
End Function

Function AOL_Win()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL_Win = AOL%
End Function

Sub StayOnTop(frm As Form)
SetWinOnTop = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub AOL_InstantMessage(sn, msg)
Call RunMenuByString(AOL_Win(), "Send an Instant Message")
Do: DoEvents
MDI% = AOL_MDI()
IM% = FindChildByTitle(MDI%, "Send Instant Message")
AolEdit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
If AolEdit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOL_SetText(AolEdit%, sn)
Call AOL_SetText(aolrich%, msg)
imsend% = FindChildByClass(IM%, "_AOL_Icon")
For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends
AOL_Icon (imsend%)
Do: DoEvents
MDI% = AOL_MDI()
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
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

Sub AOL_SetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub AOL_Icon(Icon%)
Click2% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click2% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOL_IMOff()
Call AOL_InstantMessage("$IM_OFF", "-Zero G...turning Instant Messages Off.")
End Sub

Sub AOL_IMOn()
Call AOL_InstantMessage("$IM_ON", "-Zero G...turning Instant Messages On.")
End Sub

Function AOL_Version()
AOL% = AOL_Win()
hMenu% = GetMenu(AOL%)
SubMenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")
FindString% = GetMenuString(SubMenu%, subitem%, MenuString$, 100, 1)
If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOL_Version = 3 'This will get AOL3.0
Else
AOL_Version = 4 'If not 3.0, 4.0 assumed
End If
End Function

Function AOL_Online()
MDI% = AOL_MDI()
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
If Welcome% = 0 Then
AOL_Online = 0
Exit Function
End If
AOL_Online = 1
End Function

Sub AOL4_KillWait()
Call RunMenuByString(AOL_Win(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Sub AOL_KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub AOL_45MinKill()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
AOL_Icon (AOIcon%)
End Sub

Sub AOL_10MinKill()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
AOL_Icon (AOIcon%)
End Sub

Sub AOL_AntiPunt()
Do
ANT% = FindChildByTitle(AOL_MDI(), "Untitled")
IMRICH% = FindChildByClass(ANT%, "RICHCNTL")
sts% = FindChildByClass(ANT%, "_AOL_Static")
ST% = GetWindow(sts%, GW_HWNDNEXT)
ST% = GetWindow(ST%, GW_HWNDNEXT)
Call AOL_SetText(ST%, "FeaR - This IM should be left open.")
mi = ShowWindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRICH% <> 0 Then
lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
End If
Loop
End Sub

Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Function AOL_UpChat()
Do
x% = DoEvents()
aolmod = FindWindow("_AOL_Modal", 0&)
killwin (aolmod)
Loop Until aolmod = 0
End Function

Sub killwin(win)
x = SendMessageByNum(win, WM_CLOSE, 0, 0)
End Sub

Sub Form_Center(frm As Form)
frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Sub Form_Grad_Blue(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(0, 0, i), BF
y = y + 2
Next i
End Sub

Sub Form_Grad_Blue_Light(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(0, i, 256), BF
y = y + 2
Next i
End Sub

Sub Form_Grad_Blue_Purple(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(i, 0, 256), BF
y = y + 2
Next i
End Sub

Sub Form_Grad_Green(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(0, i, 0), BF
y = y + 2
Next i
End Sub

Sub Form_Grad_Green_Yellow(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(i, 256, 0), BF
y = y + 2
Next i
End Sub

Sub Form_Grad_Random(frm As Form)
Randomize
x% = Int(Rnd * 8) + 1
 If x% = 1 Then Call Form_Grad_Blue(frm)
 If x% = 2 Then Call Form_Grad_Blue_Light(frm)
 If x% = 3 Then Call Form_Grad_Blue_Purple(frm)
 If x% = 4 Then Call Form_Grad_Green(frm)
 If x% = 5 Then Call Form_Grad_Green_Yellow(frm)
 If x% = 6 Then Call Form_Grad_Red(frm)
 If x% = 7 Then Call Form_Grad_Red_yellow(frm)
 If x% = 8 Then Call Form_Grad_Silver(frm)
End Sub

Function Randomize_Text_Color(Text1)
Randomize
x% = Int(Rnd * 7) + 1
 If x% = 1 Then Call Text_BlackBlue(Text1)
 If x% = 2 Then Call Text_BlackGreen(Text1)
 If x% = 3 Then Call Text_BlackGrey(Text1)
 If x% = 4 Then Call Text_BlackGreyBlack(Text1)
 If x% = 5 Then Call Text_BlackPurple(Text1)
 If x% = 6 Then Call Text_BlackPurpleBlack(Text1)
 If x% = 7 Then Call Text_BlackRed(Text1)

End Function


Sub Form_Grad_Red(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(i, 0, 0), BF
y = y + 2
Next i
End Sub

Sub Form_Grad_Red_yellow(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(256, i, 0), BF
y = y + 2
Next i
End Sub

Sub Form_Grad_Silver(frm As Form)
Dim i, y
frm.AutoRedraw = True
frm.DrawStyle = 6
frm.DrawMode = 13
frm.DrawWidth = 2
frm.ScaleMode = 3
frm.ScaleHeight = (256 * 2)
For i = 0 To 255
frm.Line (0, y)-(frm.Width, y + 2), RGB(i, i, i), BF
y = y + 2
Next i
End Sub

Sub AOL_Hide()
AOL% = AOL_Win()
Call ShowWindow(AOL%, 0)
End Sub

Sub AOL_Show()
AOL% = AOL_Win()
Call ShowWindow(AOL%, 5)
End Sub

Sub AOL4_KillLargeIcon()
AOL% = AOL_Win()
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub

Sub AOL_KillWait()
AOL% = AOL_Win()
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")
For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
Pause (0.05)
AOL_Icon (AOIcon%)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0
Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub

Sub NotOnTop(frm As Form)
SetWinOnTop = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub PlayWav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndPlaySound(SoundName$, wFlags%)
End Sub

Sub Timeout(duration)
'If you like typing the word TimeOut.
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop
End Sub

Function Pause(duration)
'If you like typing the word Pause.
current = Timer
Do While Timer - current < Val(duration)
DoEvents
Loop
End Function

Function AOL_UserSN()
On Error Resume Next
MDI% = AOL_MDI()
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOL_UserSN = User
End Function

Sub waitforok()
Do
    DoEvents
okw = FindWindow("#32770", "America Online")
    DoEvents
Loop Until okw <> 0
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub

Function AOL4_Text_Wavy(Text As String)
G$ = Text
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<sup>" & r$ & "</sup>" & u$ & "<sub>" & S$ & "</sub>" & T$
Next w
WavY = p$
End Function

Sub AOL4_Hide()
a = ShowWindow(AOL_Win(), SW_HIDE)
End Sub

Sub AOL4_Show()
a = ShowWindow(AOL_Win(), SW_SHOW)
End Sub

Sub AOL4_InstantMessage(sn, msg)
Call AOL4_Keyword("aol://9293:" & sn)
Timeout (2) 'Or Pause
Do
DoEvents
MDI% = AOL_MDI()
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
Loop Until (IM% <> 0 And aolrich% <> 0 And imsend% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, msg)
For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends
AOL_Icon imsend%
If IM% Then Call killwin(IM%)
Call waitforok
Call waitforok
End Sub

Sub AOL4_Keyword(kw)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    Temp% = FindChildByClass(AOL%, "AOL Toolbar")
    Temp% = FindChildByClass(Temp%, "_AOL_Toolbar")
    Temp% = FindChildByClass(Temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(Temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, kw)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub

Sub AOL4_SetFocus()
x = GetCaption(AOL_Win())
AppActivate x
End Sub

Sub AOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(aolmod%, SW_RESTORE)
Call AOL4_SetFocus
End Sub

Function AOL4_UpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(die%, SW_HIDE)
x = ShowWindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function

Sub Form_AddToAOL(frm As Form, xpos, ypos)
frm.Top = ypos
frm.Left = xpos
AOL% = FindWindow("AOL FRAME25", vbNullString)
TL% = FindChildByClass(AOL%, "AOL TOOLBAR")
sett = SetParent(frm.hWnd, AOL%)
ack = ShowWindow(AOL%, 2)
ack = ShowWindow(AOL%, 3)
End Sub

Sub AOL_Caption(YourCaption As String)
AOL% = AOL_Win()
Call AOL_SetText(AOL%, YourCaption)
End Sub

Sub AOL_HideWelcome()
MDI% = AOL_MDI()
wlcm% = FindChildByTitle(MDI%, "Welcome, ")
x = ShowWindow(wlcm%, SW_MINIMIZE)
End Sub

Sub Form_MoveWOTitle(frm As Form)
ReleaseCapture
G% = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, 2, 0)
End Sub

Sub AOL_8Liner(txt As String)
lonh = String(116, Chr(32))
D = 116 - Len(txt)
c$ = Left(lonh, D)
AOL_SendChat ("" & txt & c$ & txt)
AOL_SendChat ("" & txt & c$ & txt)
lonh = String(116, Chr(32))
D = 116 - Len(txt)
c$ = Left(lonh, D)
AOL_SendChat ("" & txt & c$ & txt)
AOL_SendChat ("" & txt & c$ & txt)
End Sub

Function Text_Search(SearchFor, SearchThis)
x = InStr(1, SearchThis, SearchFor)
Text_Search = x
End Function

Function AOL_CountMail()
TheMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(TheMail%, "_AOL_Tree")
AOL_CountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
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

Function AOL_GetPW(AolPath As String, mysn As String) As String
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As Variant
Dim l0058 As String
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0066 As Variant
l004C = Len(mysn)
Select Case l004C
Case 3
l0050 = mysn + "       "
Case 4
l0050 = mysn + "      "
Case 5
l0050 = mysn + "     "
Case 6
l0050 = mysn + "    "
Case 7
l0050 = mysn + "   "
Case 8
l0050 = mysn + "  "
Case 9
l0050 = mysn + " "
Case 10
l0050 = mysn
End Select
l0054 = 1
Do Until 2 > 3
l0058$ = ""
DoEvents
On Error Resume Next
Open AolPath$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Function
l0058$ = String(32000, 0)
Get #1, l0054, l0058$
Close #1
Open AolPath$ + "\idb\main.idx" For Binary As #2
l005A = InStr(1, l0058$, l0050 + Chr(0), 1)
If l005A Then
40:
DoEvents
Mid(l0058$, l005A) = "Pass Word "
l005E = Mid(l0058$, l005A + Len(l0050) + 1, 8)
l005E = CutStringDown(l005E)
l0062 = Mid(l0058$, l005A + Len(l0050) + 1 + Len(l005E), 1)
If l0062 <> Chr(0) Then GoTo 45
If Len(l005E) < 4 Then GoTo 45
If Len(l005E) = "" Then GoTo 45
AOL_GetPW = l005E
45:
l005A = InStr(1, l0058$, l0050 + Chr(0), 1)
If l005A Then DoEvents: GoTo 40
End If
l0054 = l0054 + 32000
l0066 = LOF(2)
Close #2
If l0054 > l0066 Then GoTo 30
Loop
30:
End Function

Function CutStringDown(ByVal PWS As String) As String
On Error Resume Next
CutStringDown = Trim(Left$(PWS, InStr(PWS, Chr$(0)) - 1))
End Function

Function AOL_FindAgain() As Integer
Dim AOLHandle As Integer
Dim AOLBox As Integer
Dim l00A6 As Integer
Dim l00A8 As String
Dim l00AA As Variant
Dim l00AE As Integer
AOLHandle% = FindWindow("_AOL_MODAL", 0&)
AOLBox% = FindChildByClass(AOLHandle%, "_AOL_Edit")
l00A6% = GetWindow(GetWindow(AOLBox%, c00B0), c00B0)
l00A8$ = Space(256)
l00AA = GetClassName(l00A6%, l00A8$, 256)
l00A8$ = CutStringDown(l00A8$)
If l00A8$ = "_AOL_Edit" Then l00AE% = l00A6%
If l00A8$ <> "_AOL_Edit" Then l00AE% = AOLBox%
AOL_FindAgain = l00AE%
End Function

'Function AOL_FindEdit() As Integer
'Dim AOLHandle As Integer
'Dim AolModal As Integer
'Dim Wel As Integer
'Dim GB As Integer
'Dim Welcome As Integer
'Dim AOLBox As Integer
'AOLHandle% = FindWindow("AOL Frame25", 0&)
'AolModal% = FindWindow("_AOL_Modal", 0&)
'AolModal% = GetParent(FindChildByTitle(AolModal%, "Cancel"))
'Wel% = FindChildByTitle(AOLHandle%, "Welcome")
'GB% = FindChildByTitle(AOLHandle%, "Goodbye from America Online!")
'If GB% <> 0 Then Welcome% = GB%
'If Wel% <> 0 Then Welcome% = Wel%
'If Welcome% <> 0 Then AOLBox% = FindChildByClass(Welcome%, "_AOL_Edit")
'If AolModal% <> 0 Then AOLBox% = AOL_FindAgain()
'AOL_FindEdit = AOLBox%
'End Function

Sub Form_Hide(frm As Form)
frm.Left = (Screen.Width - frm.Width) / 2
frm.Top = (Screen.Height - frm.Height) / 2
End Sub

Sub Run1(p00FA As Integer)
Dim l00FC As Variant
l00FC = SendMessage2(p00FA%, c0100, 13, 0)
End Sub

Sub RunButton(p00F0 As Integer)
Dim l00F2 As Variant
DoEvents
l00F2 = SendMessage2(p00F0%, c00F6, 0, 0&)
l00F2 = SendMessage2(p00F0%, c00F8, 0, 0&)
DoEvents
End Sub

Sub RunButton2(p0146 As Integer, ByVal p0148 As String)
Dim l014A As Variant
l014A = SendMessage3(p0146%, c014E, 0, p0148$)
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

Function SendIt(whut As Integer) As String
Dim l0088 As Variant
Dim l008C As String
    l0088 = SendMessage2(whut%, c008E, 0, 0)
    l008C$ = Space(l0088 + 1)
    l0088 = SendMessage3(whut%, c0090, l0088 + 1, l008C$)
    SendIt = CutStringDown(Trim$(l008C$))
End Function

Sub AOL_SendMail(SNs, subject, msg)

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

AOL_Icon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, SNs)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, msg)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

AOL_Icon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
AOL_Icon (AOIcon%)
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

Sub AOL_SetPreference()
Call RunMenuByString(AOL_Win(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOL_MDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
AOL_Icon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

AOL_Button (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Public Sub AOL_Button(but%)
ClickIco% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
ClickIco% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Sub AOL4_45MinKill()
Do
a% = FindWindow("_Aol_Palette", 0&)
B% = FindChildByClass(a%, "_Aol_Icon")
Call Pause(0.001)
Loop Until B% <> 0
AOL4_Click (B%)
End Sub

Sub AOL4_Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub

Sub SEND(chatedit, sill$)
sndtext = SendMessageByString(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub AOL4_mailcenter()
AOL% = AOL_Win()
Toolbar% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
MCenter% = GetWindow(TooLBaRB%, 2)
AOL4_Click MCenter%
End Sub

Sub AOL4_openmail()
AOL% = AOL_Win()
Toolbar% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
AOL4_Click TooLBaRB%
End Sub

Sub AOL4_Title(NewTitle$)
AOL% = AOL_Win()
AOL_SetText AOL%, NewTitle$
End Sub

Sub AOL4_writemail()
AOL% = AOL_Win()
Toolbar% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
AOL4_Click TooLBaRB%
End Sub

Sub AOL_GetMemberProfile(name As String)
AOL_RunMenuByString ("Get a Member's Profile")
Pause 0.3
MDI% = AOL_MDI()
prof% = FindChildByTitle(MDI%, "Get a Member's Profile")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOL_SetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOL_Button okbutton%
End Sub

Sub AOL_RunMenuByString(stringer As String)
Call RunMenuByString(AOL_Win(), stringer)
End Sub

Sub AOL_locateMember(name As String)
AOL_RunMenuByString ("Locate a Member Online")
Pause 0.3
MDI% = AOL_MDI()
prof% = FindChildByTitle(MDI%, "Locate Member Online")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOL_SetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOL_Button okbutton%
closes = SendMessage(prof%, WM_CLOSE, 0, 0)
End Sub

Sub AOL_Mail(sn, subject, msg)
Call RunMenuByString(AOL_Win(), "Compose Mail")

Do: DoEvents
MDI% = AOL_MDI()
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, sn)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(Mess%, WM_SETTEXT, 0, msg)

AOL_Icon (icone%)

Do: DoEvents
MDI% = AOL_MDI()
MailWin% = FindChildByTitle(MDI%, "Compose Mail")
erro% = FindChildByTitle(MDI%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
AOL_Icon (FindChildByTitle(aolw%, "OK"))
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

Sub AOL_OpenMail(which)
If which = 1 Then
Call AOL_RunMenuByString("Read &New Mail")
End If

If which = 2 Then
Call AOL_RunMenuByString("Check Mail You've &Read")
End If

If Not which = 1 Or Not which = 2 Then
Call AOL_RunMenuByString("Check Mail You've &Sent")
End If

End Sub

Sub AOL_RespondIM(msg)
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
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
e2 = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(e2, GW_HWNDNEXT)
Call AOL_SetText(e2, msg)
AOL_Icon (E)
Pause 0.8
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
E = GetWindow(E, GW_HWNDNEXT)
AOL_Icon (E)
End Sub

Sub AOL_SignOff()
AOL_RunMenuByString ("&Sign Off")
End Sub

Function Text_BlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackBlue = msg
End Function

Function Text_BlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackGreen = msg
End Function

Function Text_BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 220 / a
        f = E * B
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackGrey = msg
End Function

Function Text_BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackPurple = msg
End Function

Function Text_BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackRed = msg
End Function

Function Text_BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackYellow = msg
End Function

Function Text_BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueBlack = msg
End Function

Function Text_BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueGreen = msg
End Function

Function Text_BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BluePurple = msg
End Function

Function Text_BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueRed = msg
End Function

Function Text_BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueYellow = msg
End Function

Function Text_GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenBlack = msg
End Function

Function Text_GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenBlue = msg
End Function

Function Text_GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenPurple = msg
End Function

Function Text_GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenRed = msg
End Function

Function Text_GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenYellow = msg
End Function

Function Text_GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 220 / a
        f = E * B
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyBlack = msg
End Function

Function Text_GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyBlue = msg
End Function

Function Text_GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyGreen = msg
End Function

Function Text_GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyPurple = msg
End Function

Function Text_GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyRed = msg
End Function

Function Text_GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyYellow = msg
End Function

Function Text_PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleBlack = msg
End Function

Function Text_PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleBlue = msg
End Function

Function Text_PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleGreen = msg
End Function

Function Text_PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleRed = msg
End Function

Function Text_PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleYellow = msg
End Function

Function Text_RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedBlack = msg
End Function

Function Text_RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedBlue = msg
End Function

Function Text_RedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedGreen = msg
End Function

Function Text_RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedPurple = msg
End Function

Function Text_RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedYellow = msg
End Function

Function Text_YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowBlack = msg
End Function

Function Text_YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowBlue = msg
End Function

Function Text_YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowGreen = msg
End Function

Function Text_YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowPurple = msg
End Function

Function Text_YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowRed = msg
End Function


'Pre-set 3 Color fade combinations begin here


Function Text_BlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackBlueBlack = msg
End Function

Function Text_BlackGreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackGreenBlack = msg
End Function

Function Text_BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackGreyBlack = msg
End Function

Function Text_BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
   Text_BlackPurpleBlack = msg
End Function

Function BlankChatLineString() As String
    BlankChatLineString = Chr$(32) & Chr$(160) 'Makes up a blank chat line string
End Function
Function BlankChatLinesString() As String
    For B = 1 To 46
        sChatString = sChatString & Chr$(32) & Chr$(160) 'Puts these two characters in a string 46 times
    Next B
    BlankChatLinesString = sChatString
End Function
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

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
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
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function

Function AOLFindIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
AOLFindIM = IM%
End Function

Function Text_BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackRedBlack = msg
End Function

Function Text_BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlackYellowBlack = msg
End Function

Function Text_BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueBlackBlue = msg
End Function

Function Text_BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueGreenBlue = msg
End Function

Function Text_BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BluePurpleBlue = msg
End Function

Function Text_BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueRedBlue = msg
End Function

Function Text_BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_BlueYellowBlue = msg
End Function

Function Text_GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenBlackGreen = msg
End Function

Function Text_GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenBlueGreen = msg
End Function

Function Text_GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenPurpleGreen = msg
End Function

Function Text_GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenRedGreen = msg
End Function

Function Text_GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreenYellowGreen = msg
End Function

Function Text_GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyBlackGrey = msg
End Function

Function Text_GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyBlueGrey = msg
End Function

Function Text_GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyGreenGrey = msg
End Function

Function Text_GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyPurpleGrey = msg
End Function

Function Text_GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyRedGrey = msg
End Function

Function Text_GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_GreyYellowGrey = msg
End Function

Function Text_PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleBlackPurple = msg
End Function

Function Text_PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleBluePurple = msg
End Function

Function Text_PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleGreenPurple = msg
End Function

Function Text_PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleRedPurple = msg
End Function

Function Text_PurpleYellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_PurpleYellowPurple = msg
End Function

Function Text_RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedBlackRed = msg
End Function

Function Text_RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedBlueRed = msg
End Function

Function Text_RedGreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedGreenRed = msg
End Function

Function Text_RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedPurpleRed = msg
End Function

Function Text_RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_RedYellowRed = msg
End Function

Function Text_YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowBlackYellow = msg
End Function

Function Text_YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowBlueYellow = msg
End Function

Function Text_YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowGreenYellow = msg
End Function

Function Text_YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowPurpleYellow = msg
End Function

Function Text_YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Text_YellowRedYellow = msg
End Function


Function Text_RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    Text_RGBtoHEX = a
End Function

Sub Form_Grad_Grey(vForm As Form)
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

Sub Form_Grad_Yellow(vForm As Form)
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

Function Text_TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = C3
    C4BAK = C4
    c = 0
    o = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For x = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        val1 = (BVAL1 / Len(Text) * x) + Red1
        val2 = (BVAL2 / Len(Text) * x) + Green1
        VAL3 = (BVAL3 / Len(Text) * x) + Blue1
        
        c1 = RGB2HEX(val1, val2, VAL3)
        c2 = RGB2HEX(val1, val2, VAL3)
        C3 = RGB2HEX(val1, val2, VAL3)
        C4 = RGB2HEX(val1, val2, VAL3)
        
        If c1 = c2 And c2 = C3 And C3 = C4 And C4 = c1 Then c = 1: msg = msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        
        If WavY = True Then
            If o2 = 1 Then msg = msg + "<SUB>"
            If o2 = 3 Then msg = msg + "<SUP>"
            msg = msg + Mid$(Text, x, 1)
            If o2 = 1 Then msg = msg + "</SUB>"
            If o2 = 3 Then msg = msg + "</SUP>"
            If Q2 = 2 Then
                Q = 1
                Q2 = 1
                If o2 = 1 Then msg = msg + "<FONT COLOR=#" + c1 + ">"
                If o2 = 2 Then msg = msg + "<FONT COLOR=#" + c2 + ">"
                If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
                If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
            End If
        ElseIf WavY = False Then
            msg = msg + Mid$(Text, x, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        End If
nc:     Next x
    c1 = C1BAK
    c2 = C2BAK
    C3 = C3BAK
    C4 = C4BAK
    Text_TwoColors = msg
End Function

Function Text_RGB2HEX(r, G, B)
    Dim x&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For x& = 1 To 3
        If x& = 1 Then Color& = B
        If x& = 2 Then Color& = G
        If x& = 3 Then Color& = r
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
    Next x&
    Configuring$ = TrimSpaces(Configuring$)
    Text_RGB2HEX = Configuring$
End Function

Public Sub Form_CenterTop(frm As Form)
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
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

Sub Form_exit_Down(frm As Form)
Do
frm.Top = Trim(Str(Int(frm.Top) + 300))
DoEvents
Loop Until frm.Top > 7200
If frm.Top > 7200 Then End
End Sub


Sub Form_exit_Left(frm As Form)
Do
frm.Left = Trim(Str(Int(frm.Left) - 300))
DoEvents
Loop Until frm.Left < -6300
If frm.Left < -6300 Then End
End Sub


Sub Form_exit_right(frm As Form)
Do
frm.Left = Trim(Str(Int(frm.Left) + 300))
DoEvents
Loop Until frm.Left > 9600
If frm.Left > 9600 Then End
End Sub


Sub Form_exit_up(frm As Form)
Do
frm.Top = Trim(Str(Int(frm.Top) - 300))
DoEvents
Loop Until frm.Top < -4500
If frm.Top < -4500 Then End
End Sub



Function AOL_GetLastChatLine()
getpar = AOL_FindChatRoom()
child = FindChildByClass(getpar, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

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
If thechars = "" Then GoTo bad
LastLine = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
If LastLine <> "" Then
AOL_GetLastChatLine = LastLine
Else
bad:
AOL_GetLastChatLine = " "
End If
End Function

Sub All_AOL_Load()
Dim x
If IFileExists("C:\aol40\waol.exe") Then x = Shell("C:\America Online 4.0\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
If IFileExists("C:\America Online 4.0\waol.exe") Then x = Shell("C:\America Online 4.0\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
If IFileExists("C:\aol30\waol.exe") Then x = Shell("C:\aol30\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
If IFileExists("C:\aol30a\waol.exe") Then x = Shell("C:\aol30a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
If IFileExists("C:\aol30b\waol.exe") Then x = Shell("C:\aol30b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
If IFileExists("C:\aol25\waol.exe") Then x = Shell("C:\aol25\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
If IFileExists("C:\aol25a\waol.exe") Then x = Shell("C:\aol25a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
If IFileExists("C:\aol25b\waol.exe") Then x = Shell("C:\aol25b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Function IsUserOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome,")
If Welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function

Function IFileExists(ByVal sFileName As String) As Integer
Dim TheFileLength As Integer
On Error Resume Next
TheFileLength = Len(Dir$(sFileName))
If Err Or TheFileLength = 0 Then
IFileExists = False
Else
IFileExists = True
End If
End Function

Sub AOL_PWSScanner(FilePath$, FileName$, status As Label)
Dim TheFileLen, NumOne, sirkillaBacK, sirkilla, TheFileInfo$, PWS, PWS2, PWS3, VirusedFile, LengthOfFile, TotalRead, TheTab, TheMSg, TheMsg2, TheMsg3, TheMsg4, TheMsg5, TheDots
StopPWScanner = 0
If FileName$ = "" Then GoTo Errorr
FileName$ = FilePath$ & "\" & FileName$
If Right$(FilePath$, 1) = "\" Then FileName$ = FilePath$ & FileName$
If Not IFileExists(FileName$) Then MsgBox "File Not Found!", 16, "Error": GoTo Errorr
TheFileLen = FileLen(FileName$)
status.Caption = TheFileLen
NumOne = 1
sirkillaBacK = 2
sirkilla = 3
Do While sirkilla > sirkillaBacK
PentiumRest% = DoEvents()
If StopPWScanner = 1 Then GoTo Errorr
Open FileName$ For Binary As #1
If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": GoTo Errorr
TheFileInfo$ = String(32000, 0)
Get #1, NumOne, TheFileInfo$
Close #1
Open FileName$ For Binary As #2
If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": GoTo Errorr
PWS = InStr(1, LCase$(TheFileInfo$), "main.idx" + Chr(0), 1)
If PWS Then
sir:
Mid(TheFileInfo$, PWS) = "sirkilla  "
PWS2 = Mid(LCase$(TheFileInfo$), PWS + 8 + 1, 8)
'PWS2 = TRM(PWS2)
PWS3 = Mid(LCase$(TheFileInfo$), PWS + 8 + 1 + Len(PWS), 1)
If PWS3 <> Chr(0) Then GoTo genocide
If Len(PWS2) < 4 Then GoTo genocide
If Len(PWS2) = "" Then GoTo genocide
genocide:
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

'Sub AOL_Tos_Guide(who$, phraZes$)
'RUN "Keyword"
'AOL = AOL_Win()
'Do
'AOL4_Keyword = FindChildByTitle(AOL, "Keyword")
'Timeout (0.001)
'editbox = FindChildByClass(KeyWord, "_AOL_Edit")
'Loop Until editbox <> 0
'editbox = FindChildByClass(KeyWord, "_AOL_Edit")
'SEND editbox, "Guidepager"
'GO% = FindChildByClass(KeyWord, "_AOL_Icon")
'AOL_Button GO%
'Timeout 3
'AOL = AOL_Win()
'mow = FindChildByTitle(AOL, "I Need Help!")
'ikqn% = FindChildByClass(mow, "_AOL_Icon")
'DoEvents
'AOL_Button ikqn%
'Timeout (2)
'Timeout (0.001)

'PW = FindChildByTitle(AOL, "Report Password Solicitations")
'Timeout (0.001)
'editbox5 = FindChildByClass(PW, "_AOL_Edit")
'editbox5 = FindChildByClass(PW, "_AOL_Edit")
'SEND editbox5, "" + (who$)
'End Sub

Function AOL_Tos_Phrazes() As String
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

Sub DoubleClick(Button%)
Dim DoubleClickNow%
DoubleClickNow% = SendMessageByNum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub

Function GetWinText(hWnd As Integer) As String
Dim LengthOfText, Buffer$, GetTheText
LengthOfText = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(LengthOfText)
GetTheText = SendMessageByString(hWnd, WM_GETTEXT, LengthOfText + 1, Buffer$)
GetWinText = Buffer$
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function

Sub Form_LeftToRight(M As Form)
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 5000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000
End Sub

Sub Control_Shadow(f As Form, c As Control, shadow_effect As Integer, shadow_width As Integer, shadow_color As Long)
       Dim shColor As Long
       Dim shWidth As Integer
       Dim oldWidth As Integer
       Dim oldScale As Integer
       shWidth = shadow_width
       shColor = shadow_color
       oldWidth = f.DrawWidth
       oldScale = f.ScaleMode
       f.ScaleMode = 3
       f.DrawWidth = 1
        
        Select Case shadow_effect
        Case GFM_DROPSHADOW
       f.Line (c.Left + shWidth, c.Top + shWidth)-Step(c.Width - 1, c.Height - 1), shColor, BF
        Case GFM_BACKSHADOW
       f.Line (c.Left - shWidth, c.Top - shWidth)-Step(c.Width - 1, c.Height - 1), shColor, BF
End Select

f.DrawWidth = oldWidth
f.ScaleMode = oldScale
End Sub

Sub Control_BackgroundFlash(contrl As Control)
'Best in timer with interval of 50
DoEvents
For x = 0 To 15
contrl.BackColor = QBColor(x)
Pause 0.00001
Next x
End Sub

Sub Control_ForgroundFlash(contrl As Control)
'Best in timer with interval of 50
DoEvents
For x = 0 To 15
contrl.ForeColor = QBColor(x)
Pause 0.00001
Next x
End Sub

Sub AOL4_IMOff()
Call AOL4_InstantMessage("$IM_OFF", "-Zero G...turning Instant Messages Off.")
End Sub

Sub AOL4_IMOn()
Call AOL4_InstantMessage("$IM_On", "-Zero G...turning Instant Messages On.")
End Sub

Sub RunMenuByString2(StringSearch)
Dim SubCount%, FindString, MenuCount%, ToSearchSub%, getstring%, RunTheMenu%, MenuItemCount% 'runthemenu returns a 0 and locks up
hWnd% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(hWnd)
MenuCount% = GetMenuItemCount(hMenu%) 'tosearch
For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(hMenu%, FindString)
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
RunTheMenu% = SendMessage(hWnd, WM_COMMAND, MenuItem%, 0)
End Sub

Function AOL_SNfromIM()
MDI% = AOL_MDI()
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$
End Function



Sub AOL4_IMBomb(Text2 As TextBox, Text1 As TextBox, Text3 As TextBox)
If Text2 = "0" Then
Exit Sub
End If
If Text2 < 0 Then
Exit Sub
End If
Call AOL4_IMOff
Do
Pause (0.05)
Call AOL4_InstantMessage((Text1), (Text3) & Chr(13) & "Prog Name Here")
Text2 = Str(Val(Text2 - 1))
Loop Until Text2 = 0
Call AOL4_IMOn
Exit Sub
End Sub

Function AOL_IMScan()
aolcl% = FindWindow("#32770", "America Online")
If aolcl% > 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs OFF and can't be punted."
AOLIMScan = 1
End If
If aolcl% = 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs ON and can be punted."
AOLIMScan = 0
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

Function RGB2HEX(r, G, B)
    Dim x&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For x& = 1 To 3
        If x& = 1 Then Color& = B
        If x& = 2 Then Color& = G
        If x& = 3 Then Color& = r
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
    Next x&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Sub LoadText(lst As TextBox, File As String)

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
x = MsgBox("File Not Found", vbOKOnly, "Error!!")
End Sub


Sub SaveText(lst As TextBox, File As String)

On Error GoTo error
Dim mystr As String
Open File For Output As #1
Print #1, lst
Close 1
Exit Sub
error:
x = MsgBox("Error!!", vbOKOnly, "Error!!")
End Sub

Sub LoadList(lst As ListBox, File As String)

On Error GoTo error
Open File For Input As #1
Do Until EOF(1)
Input #1, a$
lst.AddItem a$
Loop
Close 1
Exit Sub
error:
x = MsgBox("File Not Found", vbOKOnly, "Error!!")

End Sub


Sub SaveList(lst As ListBox, File As String)

On Error GoTo error
Open File For Output As #1
For i = 0 To lst.ListCount - 1
a$ = lst.List(i)
Print #1, a$
Next
Close 1
Exit Sub
error:
x = MsgBox("Error!!", vbOKOnly, "Error!!")

End Sub

Sub AOL4_AddRoomCombo(ListBox As ListBox, ComboBox As ComboBox)
Call AOL4_AddRoomList(ListBox)
For Q = 0 To ListBox.ListCount
ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub

Sub AOL4_AddRoomList(lst As ListBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
names$ = String$(256, " ")
'ret = AOLGetList(Index, names$)
names$ = Left$(Trim$(names$), Len(Trim(names$)))
ADD_AOL_LB names$, lst
Next Index
endaddroom:
lst.RemoveItem lst.ListCount - 1
'i = GetListIndex(lst, AOL_UserSN())
If i <> -2 Then lst.RemoveItem i
End Sub

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

Sub Text_NumericsOnly()
'Put this in the Keypress of a text box
Const Numbers$ = "0123456789"
If KeyAscii = 8 Then
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
MsgBox "Error"
KeyAscii = 0
Exit Sub
End If
End If
End Sub

Sub Form_ShapeChange()
'put this in form load
'Show the form
SetWindowRgn hWnd, CreateEllipticRgn(0, 0, 300, 200), True
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



Sub Form_Float()
'In the General Declarations section
'of the form you will float put this
'Dim OriginalParenthWnd As Long
'In the Load event of the form put
'OriginalParenthWnd = SetWindowLong(Me.hwnd, GWW_HWNDPARENT, Parent.hwnd)End Sub
'In the unload form event put
'Dim r As Long
'r = SetWindowLong(Me.hwnd, GWW_HWNDPARENT, OriginalParenthWnd)End Sub
'In the button that Ends you prog put
'       Unload Me
End Sub

Sub Modals()
'oldSysModal = SetSysModalWindow([Form].hWnd)
End Sub

Sub Text_ReadOnly(Text1 As TextBox)
Call SendMessage(Text1.hWnd, EM_SETREADONLY, 1, 0)
End Sub

Sub Text_3D()
ScaleMode = 3 'pixel
FontSize = 24 'set font to 24 pts
ForeColor = &H808080 'dark grey
thetop% = 15 'top of lettering
theleft% = 5 'left of lettering
For j% = 0 To 7
CurrentX = theleft% + j%
CurrentY = thetop% + j%
    If j% = 7 Then ForeColor = &HFFFF00 'lt blue
        'Print "This is 3-D Look Text"
        Next
End Sub

