Attribute VB_Name = "Crayola"
'C
' R
'  A
'   Y
'    O
'     L
'      A.Bas
'By 721 & KingBe
'Created98
'This bas may be passed out please do not add nothing to it
'If you are please mail me at "FF Legends@aol.com"

'ENJoY

'CrayolaAOl is IMKEYWORD

Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
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

Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

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


Sub cRaYSendMail(SN, subject, message)
tool% = FindChildByClass(AOLsWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Call IconClick(Icon%)
Do: DoEvents
mail% = FindChildByTitle(AOLMDI(), "Write Mail")
Edit% = FindChildByClass(mail%, "_AOL_Edit")
rich% = FindChildByClass(mail%, "RICHCNTL")
Icon% = FindChildByClass(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And Edit% <> 0 And rich% <> 0 And Icon% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, SN)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Call SendMessageByString(Edit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(rich%, WM_SETTEXT, 0, message)
For GetIcon = 1 To 18
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Next GetIcon
Call IconClick(Icon%)
End Sub

Sub CrayolaStayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub CrayolaSend(chat)
room% = FindChatroom
AORich% = FindChildByClass(room%, "RICHCNTL")

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

Sub CenterMeForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub

Sub waitforok()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
Loop Until okw <> 0
   okb = FindChildByTitle(okw, "OK")
   okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
   oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub
Function UserSN()
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
Welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLUserSN = User
End Function
Function UserOnCrayola()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Welcome% = FindChildByTitle(mdi%, "Welcome,")
If Welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Sub PlaytimeUp(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub

Function CrayolaFindRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
stuff% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatroom = room%
Else:
   FindChatroom = 0
End If
End Function
Sub CrayolaFadaForm(TheForm As Form)
Dim hBrush%
    Dim FormHeight%, Red%, StepInterval%, x%, retVal%, OldMode%
    Dim FillArea As RECT
    OldMode = TheForm.ScaleMode
    TheForm.ScaleMode = 3  'Pixel
    FormHeight = TheForm.ScaleHeight
    StepInterval = FormHeight / 83
    Red = 0
    FillArea.Left = 0
    FillArea.Right = TheForm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = StepInterval
    For x = 1 To 255 / 4
        hBrush% = CreateSolidBrush(RGB(Red, Red, 255))
        retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
        retVal% = DeleteObject(hBrush)
        Red = Red + 4
        FillArea.Top = FillArea.Bottom
        FillArea.Bottom = FillArea.Bottom + StepInterval
    Next
    FillArea.Bottom = FillArea.Bottom + 200
    hBrush% = CreateSolidBrush(RGB(255, 255, 255))
    retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
    retVal% = DeleteObject(hBrush)
    TheForm.ScaleMode = OldMode
End Sub


Sub CrayolaAoL(Recipiant, message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

Call keyword("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For x = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next x

Call TimeOut(0.01)

ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function BlkBluy(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueBlack = msg
End Function
Function BlackGrey(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGrey = msg
End Function

Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Color8, Colr9, Colr10, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)
JinX4$ = RGBtoHEX(Colr4)
JinX5$ = RGBtoHEX(Colr5)
JinX6$ = RGBtoHEX(Colr6)
JinX7$ = RGBtoHEX(Colr7)
JinX8$ = RGBtoHEX(Colr8)
JinX9$ = RGBtoHEX(Colr9)
JinX10$ = RGBtoHEX(Colr10)

rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))
rednum4% = Val("&H" + Right(JinX4$, 2))
greennum4% = Val("&H" + Mid(JinX4$, 3, 2))
bluenum4% = Val("&H" + Left(JinX4$, 2))
rednum5% = Val("&H" + Right(JinX5$, 2))
greennum5% = Val("&H" + Mid(JinX5$, 3, 2))
bluenum5% = Val("&H" + Left(JinX5$, 2))
rednum6% = Val("&H" + Right(JinX6$, 2))
greennum6% = Val("&H" + Mid(JinX6$, 3, 2))
bluenum7% = Val("&H" + Left(JinX6$, 2))
rednum7% = Val("&H" + Right(JinX7$, 2))
greennum7% = Val("&H" + Mid(JinX7$, 3, 2))
bluenum7% = Val("&H" + Left(JinX7$, 2))
rednum8% = Val("&H" + Right(JinX8$, 2))
greennum8% = Val("&H" + Mid(JinX8$, 3, 2))
bluenum8% = Val("&H" + Left(JinX8$, 2))
rednum9% = Val("&H" + Right(JinX8$, 2))
greennum9% = Val("&H" + Mid(JinX8$, 3, 2))
bluenum9% = Val("&H" + Left(JinX8$, 2))
rednum10% = Val("&H" + Right(JinX10$, 2))
greennum10% = Val("&H" + Mid(JinX10$, 3, 2))
bluenum10% = Val("&H" + Left(JinX10$, 2))

FadeByColor10 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, TheText, Wavy)

End Function



Function FadeByColor9(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Color8, Colr9, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)
JinX4$ = RGBtoHEX(Colr4)
JinX5$ = RGBtoHEX(Colr5)
JinX6$ = RGBtoHEX(Colr6)
JinX7$ = RGBtoHEX(Colr7)
JinX8$ = RGBtoHEX(Colr8)
JinX9$ = RGBtoHEX(Colr9)

rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))
rednum4% = Val("&H" + Right(JinX4$, 2))
greennum4% = Val("&H" + Mid(JinX4$, 3, 2))
bluenum4% = Val("&H" + Left(JinX4$, 2))
rednum5% = Val("&H" + Right(JinX5$, 2))
greennum5% = Val("&H" + Mid(JinX5$, 3, 2))
bluenum5% = Val("&H" + Left(JinX5$, 2))
rednum6% = Val("&H" + Right(JinX6$, 2))
greennum6% = Val("&H" + Mid(JinX6$, 3, 2))
bluenum7% = Val("&H" + Left(JinX6$, 2))
rednum7% = Val("&H" + Right(JinX7$, 2))
greennum7% = Val("&H" + Mid(JinX7$, 3, 2))
bluenum7% = Val("&H" + Left(JinX7$, 2))
rednum8% = Val("&H" + Right(JinX8$, 2))
greennum8% = Val("&H" + Mid(JinX8$, 3, 2))
bluenum8% = Val("&H" + Left(JinX8$, 2))
rednum9% = Val("&H" + Right(JinX9$, 2))
greennum9% = Val("&H" + Mid(JinX9$, 3, 2))
bluenum9% = Val("&H" + Left(JinX9$, 2))

FadeByColor9 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, TheText, Wavy)

End Function



Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)
JinX4$ = RGBtoHEX(Colr4)
JinX5$ = RGBtoHEX(Colr5)

rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))
rednum4% = Val("&H" + Right(JinX4$, 2))
greennum4% = Val("&H" + Mid(JinX4$, 3, 2))
bluenum4% = Val("&H" + Left(JinX4$, 2))
rednum5% = Val("&H" + Right(JinX5$, 2))
greennum5% = Val("&H" + Mid(JinX5$, 3, 2))
bluenum5% = Val("&H" + Left(JinX5$, 2))

FadeByColor5 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, TheText, Wavy)

End Function



Function BlackGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGreen = msg
End Function

Function RGBtoHEX(RGB)
    a$ = Hex(RGB)
    B% = Len(a$)
    If B% = 5 Then a$ = "0" & a$
    If B% = 4 Then a$ = "00" & a$
    If B% = 3 Then a$ = "000" & a$
    If B% = 2 Then a$ = "0000" & a$
    If B% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function

Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Function FadeByColor6(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)
JinX4$ = RGBtoHEX(Colr4)
JinX5$ = RGBtoHEX(Colr5)
JinX6$ = RGBtoHEX(Colr6)

rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))
rednum4% = Val("&H" + Right(JinX4$, 2))
greennum4% = Val("&H" + Mid(JinX4$, 3, 2))
bluenum4% = Val("&H" + Left(JinX4$, 2))
rednum5% = Val("&H" + Right(JinX5$, 2))
greennum5% = Val("&H" + Mid(JinX5$, 3, 2))
bluenum5% = Val("&H" + Left(JinX5$, 2))
rednum6% = Val("&H" + Right(JinX6$, 2))
greennum6% = Val("&H" + Mid(JinX6$, 3, 2))
bluenum6% = Val("&H" + Left(JinX6$, 2))

FadeByColor6 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, TheText, Wavy)

End Function



Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)
JinX4$ = RGBtoHEX(Colr4)

rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))
rednum4% = Val("&H" + Right(JinX4$, 2))
greennum4% = Val("&H" + Mid(JinX4$, 3, 2))
bluenum4% = Val("&H" + Left(JinX4$, 2))

FadeByColor4 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, TheText, Wavy)

End Function



Function BdBlackBlueBlack(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
  
msg "<B>" & (msg) & "</B>"
End Function

Function YellowRedYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowRedYellow = msg
End Function


Function YellowRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowRed = msg
End Function



Sub CrayIMsOFF()
Call IMKeyword("$IM_OFF", "CRAYOLA Phader")
End Sub

Function FadeByColor8(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Color8, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)
JinX4$ = RGBtoHEX(Colr4)
JinX5$ = RGBtoHEX(Colr5)
JinX6$ = RGBtoHEX(Colr6)
JinX7$ = RGBtoHEX(Colr7)
JinX8$ = RGBtoHEX(Colr8)

rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))
rednum4% = Val("&H" + Right(JinX4$, 2))
greennum4% = Val("&H" + Mid(JinX4$, 3, 2))
bluenum4% = Val("&H" + Left(JinX4$, 2))
rednum5% = Val("&H" + Right(JinX5$, 2))
greennum5% = Val("&H" + Mid(JinX5$, 3, 2))
bluenum5% = Val("&H" + Left(JinX5$, 2))
rednum6% = Val("&H" + Right(JinX6$, 2))
greennum6% = Val("&H" + Mid(JinX6$, 3, 2))
bluenum7% = Val("&H" + Left(JinX6$, 2))
rednum7% = Val("&H" + Right(JinX7$, 2))
greennum7% = Val("&H" + Mid(JinX7$, 3, 2))
bluenum7% = Val("&H" + Left(JinX7$, 2))
rednum8% = Val("&H" + Right(JinX8$, 2))
greennum8% = Val("&H" + Mid(JinX8$, 3, 2))
bluenum8% = Val("&H" + Left(JinX8$, 2))

FadeByColor8 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, TheText, Wavy)

End Function



Sub CrayIMsON()
Call IMKeyword("$IM_ON", "CRAYOLA Phader ")
End Sub


Function BlackGreenBlack(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGreenBlack = msg
End Function


Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, Wavy As Boolean)
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, Wavy)

End Function
Function GryGreUN(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyGreen = msg
End Function
Sub WVYCRAYBluyBlk(TheText)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
SendChat (P$)
End Sub
Function BlkBluyBlk(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackBlueBlack = msg
End Function

Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This gets a color from 3 scroll bars
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)

'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)

End Function

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

Function FadeByColor7(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)
JinX3$ = RGBtoHEX(Colr3)
JinX4$ = RGBtoHEX(Colr4)
JinX5$ = RGBtoHEX(Colr5)
JinX6$ = RGBtoHEX(Colr6)
JinX7$ = RGBtoHEX(Colr7)


rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))
rednum3% = Val("&H" + Right(JinX3$, 2))
greennum3% = Val("&H" + Mid(JinX3$, 3, 2))
bluenum3% = Val("&H" + Left(JinX3$, 2))
rednum4% = Val("&H" + Right(JinX4$, 2))
greennum4% = Val("&H" + Mid(JinX4$, 3, 2))
bluenum4% = Val("&H" + Left(JinX4$, 2))
rednum5% = Val("&H" + Right(JinX5$, 2))
greennum5% = Val("&H" + Mid(JinX5$, 3, 2))
bluenum5% = Val("&H" + Left(JinX5$, 2))
rednum6% = Val("&H" + Right(JinX6$, 2))
greennum6% = Val("&H" + Mid(JinX6$, 3, 2))
bluenum7% = Val("&H" + Left(JinX6$, 2))
rednum7% = Val("&H" + Right(JinX7$, 2))
greennum7% = Val("&H" + Mid(JinX7$, 3, 2))
bluenum7% = Val("&H" + Left(JinX7$, 2))


FadeByColor8 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, TheText, Wavy)

End Function


Function FadeByColor2(Colr1, Colr2, TheText$, Wavy As Boolean)
'By 721
JinX$ = RGBtoHEX(Colr1)
JinX2$ = RGBtoHEX(Colr2)


rednum1% = Val("&H" + Right(JinX$, 2))
greennum1% = Val("&H" + Mid(JinX$, 3, 2))
bluenum1% = Val("&H" + Left(JinX$, 2))
rednum2% = Val("&H" + Right(JinX2$, 2))
greennum2% = Val("&H" + Mid(JinX2$, 3, 2))
bluenum2% = Val("&H" + Left(JinX2$, 2))



FadeByColor8 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, Wavy)

End Function



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
Sub Explode(Frm As Form, CFlag As Integer, steps As Integer)

Dim FRect As RECT
Dim FWidth, fHeight As Integer
Dim i, x, Y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush
    
    GetWindowRect Frm.hwnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    fHeight = FRect.Bottom - FRect.Top
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(Frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    For i = 1 To steps
        cx = FWidth * (i / steps)
        cy = fHeight * (i / steps)
        If CFlag Then
            x = FRect.Left + (FWidth - cx) / 2
            Y = FRect.Top + (fHeight - cy) / 2
        Else
            x = FRect.Left
            Y = FRect.Top
        End If
        Rectangle hScreen, x, Y, x + cx, Y + cy
    Next i
    If ReleaseDC(0, hScreen) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    DeleteObject (Brush)
    Frm.Show
End Sub

ub FireFade(vForm As Object)
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
Sub XAOL4_IMOff()
Call XAOL4_InstantMessage("$IM_OFF", "   (\/\_/¯\_)¯\£ÃßÝRÍÑ†H/¯(_/¯\_/\/)")
End Sub
Sub XAOL4_IMOn()
Call XAOL4_InstantMessage("$IM_ON", "   (\/\_/¯\_)¯\£ÃßÝRÍÑ†H/¯(_/¯\_/\/)")
End Sub
Function XAOL4_SpiralScroll(Txt As String)
Dim AODCOUNTER, a, thetxtlen
AODCOUNTER = 1
thetxtlen = Len(Txt)
Start:
a = a + 1
If a = thetxtlen Then GoTo last
x = Text_Looping(Txt)
Txt = x
XAOL4_ChatSend x
Pause (0.5)
AODCOUNTER = AODCOUNTER + 1
If AODCOUNTER = 4 Then
   AODCOUNTER = 2
   End If
GoTo Start
last:

End Function
Function XAOL4_UpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(die%, SW_HIDE)
x = ShowWindow(die%, SW_MINIMIZE)
Call XAOL4_SetFocus
End Function
Sub XAOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(aolmod%, SW_RESTORE)
Call XAOL4_SetFocus
End Sub
Sub XAOL4_Punter(SN As TextBox, Bombz As TextBox)
Call XAOL4_IMOff
waitforok
Do
DoEvents:
Call XAOL4_InstantMessage(SN, "99999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999")
DAWIN% = FindWindow("#32770", "America Online")
If DAWIN% Then Exit Sub: MsgBox "Sorry person isn't online!", 48, "DAMMIT!"
Bombz = Str(Val(Bombz - 1))
Loop Until Bombz <= 0
Call XAOL4_IMOn
waitforok
End Sub
Sub XAOL4_InstantMessage(person, message)
Call XAOL4_Keyword("aol://9293:" & person)
Pause (2)
Do
DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
Loop Until (im% <> 0 And aolrich% <> 0 And imsend% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, message)
For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends
AOLIcon imsend%
If im% Then Call AOLKillWindow(im%)
End Sub

Function WavYChaTRedBlue(TheText As String)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
WavYChaTRB = P$
Function WavYChaTRedGreen(TheText As String)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & t$
Next w
WavYChaTRG = P$
End Function
Function Wavy(TheText As String)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & S$ & "</sub>" & t$
Next w
Wavy = P$

End Function
