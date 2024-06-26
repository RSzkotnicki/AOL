'if you wanna steal any of this code
'make sure ya give credit where credit is due
'-----------------
'Data Types
Type RECT
 Left As Integer
 Top As Integer
 Right As Integer
 Bottom As Integer
End Type

'-----------------
'these first 5 functions are rem'd out cause my
'bas doesn't use them just in case you need to
'use them I left the dec's for them
'-----------------
'Declare Function AOLGetList% Lib "311.Dll" (ByVal index%, ByVal Buf$)
'Declare Function AOLGetCombo% Lib "311.Dll" (ByVal index%, ByVal Buf$)
'the infamous 311.dll is no longer used!
'-----------------
'Declare Function FindChildByClass% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
'Declare Function FindChildByTitle% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
'no more vbwfind.dll
'-----------------
'Declare Function agGetStringFromLPSTR Lib "APIGUIDE.DLL" (ByVal lpString As Long) As String
'apiguide.dll = LaMuH
'-----------------
'API Function's
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function CreateCompatibleDC% Lib "GDI" (ByVal hDC%)
Declare Function CreateFont% Lib "GDI" (ByVal H%, ByVal W%, ByVal E%, ByVal O%, ByVal W%, ByVal i%, ByVal U%, ByVal S%, ByVal c%, ByVal OP%, ByVal CP%, ByVal Q%, ByVal PAF%, ByVal f$)
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function CreateWindow% Lib "User" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hWndParent%, ByVal hMenu%, ByVal hInstance%, ByVal lpParam$)
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function DrawText Lib "User" (ByVal hDC As Integer, ByVal lpStr As String, ByVal nCount As Integer, lpRect As RECT, ByVal wFormat As Integer) As Integer
Declare Function EnableHardwareInput Lib "User" (ByVal bEnableInput As Integer) As Integer
Declare Function EnableWindow Lib "User" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function FindChildByTitle% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function FindChildByClass% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function FlashWindow Lib "User" (ByVal hWnd As Integer, ByVal bInvert As Integer) As Integer
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function GetDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetMenuItemCount Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer
Declare Function getparent Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetTempDrive Lib "Kernel" (ByVal cDriveLetter As Integer) As Integer
Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
Declare Function GetTopWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function IsWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function iswindowenabled Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function IsWindowVisible Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function LoadBitmap Lib "User" (ByVal hInstance%, ByVal lpBitMapName As Any) As Integer
Declare Function lstrcpy Lib "Kernel" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
Declare Function RemoveMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function sendmessagebynum Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Integer
Declare Function SendMEssageByString Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal hWnd As Integer) As Integer
Declare Function SetMenu Lib "User" (ByVal hWnd As Integer, ByVal hMenu As Integer) As Integer
Declare Function SetMenuItemBitmaps Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal hBitmapUnchecked As Integer, ByVal hBitmapChecked As Integer) As Integer
Declare Function SetParent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
Declare Function SetPixel Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long) As Long
Declare Function SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Long
Declare Function showwindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function sndPlaySound Lib "MMSystem" (ByVal lpsound As String, ByVal flag As Integer) As Integer
Declare Function StretchBlt% Lib "GDI" (ByVal hDestDC%, ByVal x%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal xSrc%, ByVal ySrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop As Long)
Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function WindowFromPoint Lib "User" (ByVal ptScreen As Any) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Integer

'API Sub's
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Sub DrawMenuBar Lib "User" (ByVal hWnd As Integer)
Declare Sub GetCursorPos Lib "User" (lpPoint As Long)
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As RECT)
Declare Sub hmemcpy Lib "Kernel" (hpvDest As Any, hpvSource As Any, ByVal cbCopy&)
Declare Sub InvertRect Lib "User" (ByVal hDC As Integer, lpRect As RECT)
Declare Sub ModifyMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpString As Long)
Declare Sub MoveWindow Lib "User" (ByVal hWnd As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub ReleaseCapture Lib "User" ()
Declare Sub Yield Lib "Kernel" ()
Declare Sub ReleaseCapture Lib "User" ()

'Important Global's
Global Const WM_USER = &H400
Global Const SWP_NOSIZE = 1
Global Const SWP_NOMOVE = &H2
Global Const SW_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'API Global's
Global Const BM_SETCHECK = WM_USER + 1
Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_GETITEMDATA = (WM_USER + 16)
Global Const CB_GETLBTEXTLEN = (WM_USER + 9)
Global Const CB_INSERTSTRING = (WM_USER + 10)
Global Const CB_SETCURSEL = (WM_USER + 14)
Global Const CB_SETEDITSEL = (WM_USER + 2)
Global Const CB_SHOWDROPDOWN = (WM_USER + 15)
Global Const EM_GETLINE = WM_USER + 20
Global Const EM_GETLINECOUNT = WM_USER + 10
Global Const EM_GETSEL = WM_USER + 0
Global Const EM_REPLACESEL = WM_USER + 18
Global Const EM_SCROLL = WM_USER + 5
Global Const EM_SETFONT = WM_USER + 19
Global Const EM_SETREADONLY = (WM_USER + 31)
Global Const EW_REBOOTSYSTEM = &H43
Global Const GW_CHILD = 5
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GWW_HINSTANCE = (-6)
Global Const HWND_NOTOPMOST = -2
Global Const HWND_TOPMOST = -1
Global Const KEY_DELETE = &H2E
Global Const LB_32GETCOUNT = &H18B
Global Const LB_32GETCURSEL = &H188
Global Const LB_32GETITEMDATA = &H199
Global Const LB_32GETTEXT = &H189
Global Const LB_32GETTEXTLEN = &H18A
Global Const LB_32SETCURSEL = &H186
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const LB_GETCURSEL = (WM_USER + 9)
Global Const LB_GETITEMDATA = (WM_USER + 26)
Global Const LB_GETITEMRECT = (WM_USER + 25)
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const LB_GETTEXTLEN = (WM_USER + 11)
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const LBN_DBLCLK = 2
Global Const MB_TASKMODAL = &H2000
Global Const MF_BITMAP = &H4
Global Const MF_BYCOMMAND = &H0
Global Const MF_DISABLED = &H2
Global Const SRCCOPY = &HCC0020
Global Const SW_HIDE = 0
Global Const SW_MAXIMIZE = 3
Global Const SW_MINIMIZE = 6
Global Const SW_NORMAL = 1
Global Const SW_RESTORE = 9
Global Const SW_SHOW = 5
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_SHOWNORMAL = 1
Global Const VK_CONTROL = &H11
Global Const VK_DELETE = &H2E
Global Const VK_DOWN = &H28
Global Const VK_HOME = &H24
Global Const VK_SPACE = &H20
Global Const VK_TAB = &H9
Global Const WM_CHAR = &H102
Global Const WM_CLEAR = &H303
Global Const WM_CLOSE = &H10
Global Const WM_COMMAND = &H111
Global Const WM_COPY = &H301
Global Const WM_GETFONT = &H31
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_MOVE = &H3
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const WM_SETCURSOR = &H20
Global Const WM_SETFONT = &H30
Global Const WM_SETTEXT = &HC
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

Sub CountMail1 ()
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
NER% = FindChildByTitle(MDI%, "New Mail")
If NER% <> 0 Then GoTo J
TOO% = FindChildByClass(Aol%, "AOL Toolbar")
NEE% = FindChildByClass(TOO%, "_AOL_Icon")
D = sendmessagebynum(NEE%, WM_LBUTTONDOWN, 0, 0)
D = sendmessagebynum(NEE%, WM_LBUTTONUP, 0, 0)
Do
c% = DoEvents()
NER% = FindChildByTitle(MDI%, "New Mail")
Loop Until NER% <> 0
TRE% = FindChildByClass(NER%, "_AOL_Tree")
Do
U = sendmessagebynum(TRE%, LB_GETCOUNT, 0, 0)
Timeout (2)
f = sendmessagebynum(TRE%, LB_GETCOUNT, 0, 0)
Loop Until U = f
J:
TRE% = FindChildByClass(NER%, "_AOL_Tree")
f = sendmessagebynum(TRE%, LB_GETCOUNT, 0, 0)
MsgBox "You have " & f & " pieces of mail!", 64, "Mail Counter"

End Sub


Sub Invert_Control (Ctrl As Control)
    
    Dim RECTANGL As RECT
    
    RECTANGL.Right = Ctrl.ScaleWidth
    RECTANGL.Bottom = Ctrl.ScaleHeight
    
    Call InvertRect(Ctrl.hDC, RECTANGL)

End Sub

Sub InviteOff ()
If online() = False Then Call ErrorMsg: Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
Call Keyword("Buddy List")
Buddy% = WaitForWin(getusersn() + "'s Buddy List Groups")
Pref% = GetWindow(getwindowbytitle(Buddy%, "Preferences"), GW_HWNDNEXT)
Do
DoEvents
click Pref%
Preferences% = getwindowbytitle(MDI%, "Buddy List Preferences")
Loop Until Preferences% <> 0
Off% = getwindowbyclass(Preferences%, "_AOL_Static")
For x = 1 To 8
    Off% = GetWindow(Off%, GW_HWNDNEXT)
Next x
x = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = GetWindow(Off%, GW_HWNDNEXT)
x = sendmessagebynum(Off%, BM_SETCHECK, True, 0)
Off% = GetWindow(Off%, GW_HWNDNEXT)
x = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = GetWindow(Off%, GW_HWNDNEXT)
x = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
AOLIcon% = GetWindow(getwindowbyclass(Preferences%, "_AOL_Edit"), GW_HWNDNEXT)
For x = 1 To 4
AOLIcon% = GetWindow(AOLIcon%, GW_HWNDNEXT)
Next x
Do
DoEvents
click (AOLIcon%)
Saved% = FindWindow("#32770", "America Online")
Loop Until Saved% <> 0
Call closewin(Saved%)
Call closewin(Buddy%)
End Sub

Sub InviteOn ()
If online() = False Then Call ErrorMsg: Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
Call Keyword("Buddy List")
Buddy% = WaitForWin(getusersn() + "'s Buddy List Groups")
Pref% = GetWindow(getwindowbytitle(Buddy%, "Preferences"), GW_HWNDNEXT)
Do
DoEvents
click Pref%
Preferences% = getwindowbytitle(MDI%, "Buddy List Preferences")
Loop Until Preferences% <> 0
Off% = getwindowbyclass(Preferences%, "_AOL_Static")
For x = 1 To 8
    Off% = GetWindow(Off%, GW_HWNDNEXT)
Next x
x = sendmessagebynum(Off%, BM_SETCHECK, True, 0)
Off% = GetWindow(Off%, GW_HWNDNEXT)
x = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = GetWindow(Off%, GW_HWNDNEXT)
x = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = GetWindow(Off%, GW_HWNDNEXT)
x = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
AOLIcon% = GetWindow(getwindowbyclass(Preferences%, "_AOL_Edit"), GW_HWNDNEXT)
For x = 1 To 4
AOLIcon% = GetWindow(AOLIcon%, GW_HWNDNEXT)
Next x
Do
DoEvents
click (AOLIcon%)
Saved% = FindWindow("#32770", "America Online")
Loop Until Saved% <> 0
Call closewin(Saved%)
Call closewin(Buddy%)

End Sub

Sub Invoke (record As String)
Keyword "aol://1391:" + record$
'AOL% = FindWindow("AOL Frame25", 0&)
'MDI% = GetWindowByClass(AOL%, "MDIClient")
'InvokeWin% = GetWindowByTitle(MDI%, "Invoke Database Record")
'If InvokeWin% = 0 Then
'    RunMenu "Invoke Database Record..."
'    InvokeWin% = WaitForWin("Invoke Database Record")
'    AOLEdit% = GetWindowByClass(InvokeWin%, "_AOL_Edit")
'    Call SetEdit(AOLEdit%, record$)
'    OK% = GetChildWin(InvokeWin%, "OK", "_AOL_Button")
'    TopWindow% = GetTopWindow(MDI%)
'    Do
'    Click OK%
'    TimeOut (.2)
'    Loop Until (GetTopWindow(MDI%) <> TopWindow%)
'Else :
'    AOLEdit% = GetWindowByClass(InvokeWin%, "_AOL_Edit")
'    Call SetEdit(AOLEdit%, record$)
'    OK% = GetChildWin(InvokeWin%, "OK", "_AOL_Button")
'    TopWindow% = GetTopWindow(MDI%)
'    Do
'    Click (OK%)
'    TimeOut (.2)
'    Loop Until (GetTopWindow(MDI%) <> TopWindow%)
'End If
End Sub

Sub Keyword (Key As String)
If online() = False Then Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
KeyWindow% = getwindowbytitle(MDI%, "Keyword")
If (KeyWindow% = 0) Then
    RunMenu "Keyword..."
    Do
    DoEvents
    KeyWindow% = getwindowbytitle(MDI%, "Keyword")
    Loop Until (KeyWindow% <> 0)
End If
DoEvents
AOLEdit% = getwindowbyclass(KeyWindow%, "_AOL_Edit")
Call setedit(AOLEdit%, Key$)
Call Enter(AOLEdit%)
End Sub

Sub keyword40 (Key As String)
If online() = False Then Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
KeyWindow% = getwindowbytitle(MDI%, "Keyword")
If (KeyWindow% = 0) Then
TOO% = FindChildByClass(Aol%, "_AOL_TOOLBAR")
ico% = FindChildByClass(TOO%, "_AOL_ICON")
ico2% = getnextwindow(ico%, 20)
click ico2%
Do
    DoEvents
    KeyWindow% = getwindowbytitle(MDI%, "Keyword")
    Loop Until (KeyWindow% <> 0)
End If
DoEvents
AOLEdit% = getwindowbyclass(KeyWindow%, "_AOL_Edit")
Call setedit(AOLEdit%, Key$)
Call Enter(AOLEdit%)

End Sub

Sub KillAd ()
Advertisement% = FindWindow("_AOL_MODAL", 0&)
Cancel1% = getwindowbytitle(Advertisement%, "Cancel")
Cancel2% = getwindowbytitle(Advertisement%, "No Thanks")
If Cancel1% <> 0 Then CancelButton% = Cancel1%
If Cancel2% <> 0 Then CancelButton% = Cancel2%
If (CancelButton% <> 0) And (Advertisement% <> GetGuest()) Then click CancelButton%
End Sub

Sub killwait ()
Aol% = FindWindow("AOL Frame25", 0&)
x = EnableWindow(Aol%, True)
If GetAOL() = 2 Then RunMenu "Exit Free Area"
If GetAOL() = 3 Or GetAOL() = 95 Then RunMenu "Exit Unlimited Use area"
End Sub

Sub AddMemberDirectory (List As ListBox, OnlineMem As Integer)
On Error Resume Next
If online() = False Then Call ErrorMsg: Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
MemRes% = getwindowbytitle(MDI%, "Member Directory Search Results")
If MemRes% = 0 Then
    SearchString$ = InputBox("Enter the string you want to search for:", "Search String")
    If OnlineMem = False Then Keyword "aol://4950:0000010000|all:" + SearchString$
    If OnlineMem = True Then Keyword "aol://4950:0000010000|all:" + SearchString$ + "|online:"
    Do
    DoEvents
    MemRes% = getwindowbytitle(MDI%, "Member Directory Search Results")
    NoneFound% = FindWindow("#32770", "America Online")
    If (MemRes% <> 0) Then Exit Do
    If (NoneFound% <> 0) Then Exit Do
    Loop
    MemberDir% = getwindowbytitle(MDI%, "Member Directory")
    If (NoneFound% <> 0) Then
	Call closewin(NoneFound%)
	Call closewin(MemberDir%)
	MsgBox "No matches were found for: " + SearchString$, 48
	Exit Sub
    End If
    If MemRes% <> 0 Then
	AOLStatic% = getwindowbyclass(MemRes%, "_AOL_Static")
	AOLList% = getwindowbyclass(MemRes%, "_AOL_Listbox")
	Do
	DoEvents
	Loop Until getapitext(AOLStatic%) <> ""
	Do
	Num1 = sendmessagebynum(AOLList%, LB_GETCOUNT, 0, 0)
	Timeout (1)
	Num2 = sendmessagebynum(AOLList%, LB_GETCOUNT, 0, 0)
	Loop Until Num1 = Num2
	More% = getwindowbyclass(MemRes%, "_AOL_Icon")
	Text$ = getapitext(AOLStatic%)
	If InStr(UCase(Text$), UCase("over ")) Then
	    Text$ = Mid(Text$, InStr(UCase(Text$), UCase("over ")) + 5)
	    If Err Then MsgBox "Error retrieving names", 64: Exit Sub
	Else :
	    Text$ = Mid(Text$, InStr(UCase(Text$), UCase("of ")) + 3)
	    If Err Then MsgBox "Error retrieving names", 64: Exit Sub
	End If
	Text$ = Left(Text$, InStr(UCase(Text$), UCase(" matching")) - 1)
	num = Val(Text$)
	Do
	DoEvents
	NumList = sendmessagebynum(AOLList%, LB_GETCOUNT, 0, 0)
	If NumList = num Then Exit Do
	If iswindowenabled(More%) Then click (More%)
	Loop
    End If
End If
MemRes% = getwindowbytitle(MDI%, "Member Directory Search Results")
AOLList% = getwindowbyclass(MemRes%, "_AOL_Listbox")
num = sendmessagebynum(AOLList%, LB_GETCOUNT, 0, 0)
For x = 0 To num - 1
    DoEvents
    i = AOLGetList(AOLList%, x, txt$)
    txt$ = Mid(txt$, InStr(txt$, Chr(9)) + 1)
    txt$ = Mid(txt$, 1, InStr(txt$, Chr(9)) - 1)
    If Err = False Then Call AddList(List, txt$)
Next x
Call closewin(MemRes%)
MemberDir% = getwindowbytitle(MDI%, "Member Directory")
Call closewin(MemberDir%)
End Sub

Sub addroom (List As ListBox)
ChatRoom% = FindChatRoom()
If ChatRoom% = 0 Then Exit Sub
AOLListBox% = getwindowbyclass(ChatRoom%, "_AOL_ListBox")
num = sendmessagebynum(AOLListBox%, LB_GETCOUNT, 0, 0)
UserSN$ = getusersn()
For Index% = 0 To num - 1
x = AOLGetList(AOLListBox%, Index%, SN$)
If (SN$ <> UserSN$) And (SN$ <> "") Then
    If Len(List.Tag) > 0 Then
	If InStr(1, SN$, List.Tag, 1) = 0 Then
	    Call AddList(List, SN$)
	End If
    Else :
	Call AddList(List, SN$)
    End If
End If
Next Index%
End Sub

Function AF_Script (ByVal Strt As String, ByVal ReplaceMe As String, ByVal ReplaceWith As String) As String
Start$ = Strt
Do While InStr(Start$, ReplaceMe$) <> 0
    x% = DoEvents()
    pos% = InStr(Start$, ReplaceMe$)
    Start$ = Left(Start$, pos% - 1) & ReplaceWith$ & Right(Start$, Len(Start$) - pos% - Len(ReplaceMe$) + 1)
    Loop
AF_Script$ = Start$
End Function

Function agGetStringFromLPSTR (lpStrings As Long) As String
   'this function was coded by
   'flow. i just re-arranged it
   Dim lpStrAddress As Long, lpStrz$
   lpStrz$ = Space$(4096)
   lpStrAddress = lpStrings&
   lpStrAddress = lstrcpy(lpStrz$, lpStrAddress)
   lpStrz$ = Trim$(lpStrz$)
   lpStrz$ = Left$(lpStrz$, Len(lpStrz$) - 1)
   agGetStringFromLPSTR = lpStrz$
End Function

Function alive (who As String)
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
Call SendEMail(who$ + ", -", "tos check", " ", True, False)
Do
DoEvents
ErrorWindow% = getwindowbytitle(MDI%, "Error")
Loop Until ErrorWindow% <> 0
view% = getwindowbyclass(ErrorWindow%, "_AOL_View")
ViewText$ = getapitext(view%)
If InStr(SpaceCase(ViewText$), SpaceCase(who$)) Then
  alive = False
Else :
  alive = True
End If
Call closewin(ErrorWindow%)
Compose% = getwindowbytitle(MDI%, "Compose Mail")
Call closewin(Compose%)
End Function

Function AnimateCaption (Object As Control, method As Integer, Caption As String, Pause As Variant)
Randomize Timer
length = Len(Caption$)
Static Used(3200) As Integer
Erase Used
Dim Text As String
Dim letter As Integer
Dim x As Integer
Dim i As Integer
Text$ = Space(length)
Select Case method
    Case 1:
	For x = 1 To length
NewOne:
	    letter = Int(Rnd * length) + 1
	    For i = 1 To length
		 If Used(i) = letter Then GoTo NewOne
	    Next i
	    Used(x) = letter
	    Mid$(Text$, letter, 1) = Mid(Caption$, letter, 1)
	    Object = Text$
	    Timeout (Pause)
	Next x
    Case 2:
	For x = length To 1 Step -1
	    Mid(Text$, x, 1) = Mid(Caption$, x, 1)
	    Object = Text$
	    Timeout (Pause)
	Next x
End Select
End Function

Function AOLGetCombo (AOLComboBox As Integer, Index, Buffers As String)
On Error Resume Next
Dim hAOLProcess As Long
Dim lBytesRead As Long
Dim sBuffer As String
Dim lAddrOfItemData As Long
Dim lAddrOfString  As Long
lAddrOfItemData = SendMessage(AOLComboBox%, CB_GETITEMDATA, Index, ByVal 0&)
lAddrOfItemData = lAddrOfItemData + (4 * 6)
Call hmemcpy(lAddrOfString, ByVal lAddrOfItemData&, 4)
lAddrOfString = lAddrOfString + 6
sBuffer = String$(1000, 0)
Call hmemcpy(ByVal sBuffer$, ByVal lAddrOfString&, Len(sBuffer))
Buffers$ = FixAPIString(sBuffer$)
End Function

Function AOLGetList (AOLListBox As Integer, Index, Buffers As String)
On Error Resume Next
Dim hAOLProcess As Long
Dim lBytesRead As Long
Dim sBuffer As String
Dim lAddrOfItemData As Long
Dim lAddrOfString  As Long
lAddrOfItemData = SendMessage(AOLListBox%, LB_GETITEMDATA, Index, 0)
lAddrOfItemData = lAddrOfItemData + (4 * 6)
Call hmemcpy(lAddrOfString, ByVal lAddrOfItemData&, 4)
lAddrOfString = lAddrOfString + 6
sBuffer = String$(1000, 0)
Call hmemcpy(ByVal sBuffer$, ByVal lAddrOfString&, Len(sBuffer))
Buffers$ = FixAPIString(sBuffer$)
End Function

Sub bustin (room As String)
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
	closewin (Full%)
	If InStr(1, Text$, "The room you requested is full.", 1) = 0 Then Exit Do
    End If
    If SpaceCase(GetRoomName()) = SpaceCase((room$)) Then Exit Do
Loop Until Abort = True
End Sub

Sub center (frm As Form)
frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
End Sub

Function chatsee () As Integer
Chat% = FindChatRoom()
view% = getwindowbyclass(Chat%, "_AOL_View")
chatsee = view%
End Function

Function ChatView () As Integer
Chat% = FindChatRoom()
view% = getwindowbyclass(Chat%, "_AOL_View")
ChatView = view%
End Function

Function ChrCode (txt As String) As String
For x = 1 To Len(txt$)
    outstring$ = outstring$ + "Chr(" + CStr(Asc(Mid(txt$, x, 1))) + ") + "
Next x
outstring$ = Trim(outstring$)
outstring$ = Mid(outstring$, 1, Len(outstring$) - 2)
ChrCode = outstring$
End Function

Sub click (send%)
DoEvents
x = sendmessagebynum(send%, WM_LBUTTONDOWN, 0, 0)
x = sendmessagebynum(send%, WM_LBUTTONUP, 0, 0)
DoEvents
End Sub

Sub closewin (hWnd As Integer)
x = sendmessagebynum(hWnd%, WM_CLOSE, 0, 0)
End Sub

Sub Combo1_KeyPress (KeyAscii As Integer, blah As String)

'    Dim sSearchText As String
'     Dim lReturn As Long
'blah.SelStart
'     If KeyAscii = 13 Then
 '       Combo1_Click
'        KeyAscii = 0
   '  Else
'        sSearchText = Left$(blah, blah) & Chr$(KeyAscii)
       ' lReturn = SendMessage(blah.hWnd, CB_FINDSTRING, -1, ByVal sSearchText)
       ' If lReturn <> CB_ERR Then
	  ' mbIgnoreListClick = True
	  ' Combo1.ListIndex = lReturn
	 '  mbIgnoreListClick = False
	 '  Combo1.Text = Combo1.List(lReturn)
	 '  Combo1.SelStart = Len(sSearchText)
	 '  Combo1.SelLength = Len(Combo1.Text)
	 '  KeyAscii = 0
       ' End If
     'End If
End Sub

Function Convert (Word As String, num As Integer) As String
Select Case num
Case 1:
    For x = 1 To Len(Word$)
	leet$ = ""
	letter$ = Mid$(Word$, x, 1)
	If letter$ = "a" Then leet$ = "å"
	If letter$ = "b" Then leet$ = "þ"
	If letter$ = "c" Then leet$ = "©"
	If letter$ = "d" Then leet$ = "d"
	If letter$ = "e" Then leet$ = "ë"
	If letter$ = "f" Then leet$ = "ƒ"
	If letter$ = "g" Then leet$ = "g"
	If letter$ = "h" Then leet$ = "h"
	If letter$ = "i" Then leet$ = "ï"
	If letter$ = "j" Then leet$ = ",j"
	If letter$ = "k" Then leet$ = "/<"
	If letter$ = "l" Then leet$ = "Ï"
	If letter$ = "m" Then leet$ = "m"
	If letter$ = "n" Then leet$ = "ñ"
	If letter$ = "o" Then leet$ = "ø"
	If letter$ = "p" Then leet$ = "p"
	If letter$ = "q" Then leet$ = "q"
	If letter$ = "r" Then leet$ = "®"
	If letter$ = "s" Then leet$ = "š"
	If letter$ = "t" Then leet$ = "†"
	If letter$ = "u" Then leet$ = "ü"
	If letter$ = "v" Then leet$ = "v"
	If letter$ = "w" Then leet$ = "vv"
	If letter$ = "x" Then leet$ = "×"
	If letter$ = "y" Then leet$ = "ý"
	If letter$ = "z" Then leet$ = "z"
	If letter$ = " " Then leet$ = " "
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
	If letter$ = "O" Then leet$ = "Ø"
	If letter$ = "P" Then leet$ = "¶"
	If letter$ = "Q" Then leet$ = "Q"
	If letter$ = "R" Then leet$ = "R"
	If letter$ = "S" Then leet$ = "Š"
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
	If letter$ = "*" Then leet$ = "™"
	If letter$ = "-" Then leet$ = "–"
	If letter$ = "/" Then leet$ = "÷"
	If letter$ = "?" Then leet$ = "¿"
	If Len(leet$) = 0 Then leet$ = letter$
	Made$ = Made$ & leet$
    Next
Case 2:
    For x = 1 To Len(Word$)
	leet$ = ""
	letter$ = Mid(Word$, x, 1)
	If UCase(letter$) = "I" Then leet$ = "1"
	If UCase(letter$) = "L" Then leet$ = "1"
	If UCase(letter$) = "O" Then leet$ = "0"
	If UCase(letter$) = "E" Then leet$ = "3"
	If UCase(letter$) = "A" Then leet$ = "4"
	If UCase(letter$) = "S" Then leet$ = "5"
	If UCase(letter$) = "T" Then leet$ = "7"
	If Len(leet$) = 0 Then leet$ = LCase(letter$)
	Made$ = Made$ + leet$
    Next x
Case 3:
    For x = 1 To Len(Word$)
	letter$ = Mid(Word$, x, 1)
	Select Case UCase(letter$)
	    Case "A": letter$ = LCase(letter$)
	    Case "E": letter$ = LCase(letter$)
	    Case "I": letter$ = LCase(letter$)
	    Case "O": letter$ = LCase(letter$)
	    Case "U": letter$ = LCase(letter$)
	    Case Else: letter$ = UCase(letter$)
	End Select
	Made$ = Made$ + letter$
    Next x
Case 4:
    For x = 1 To Len(Word$)
	letter$ = Mid$(Word$, x, 1)
	Made$ = letter$ + Made$
    Next x
Case 5:
    For x = 1 To Len(Word$)
	letter$ = Mid$(Word$, x, 1)
	Made$ = Made$ + letter$ + " "
    Next x
End Select
Convert = Made$
End Function

Function ConvertSeconds (Seconds As Integer) As String
Seconds1 = Seconds
Dim Min As Integer
While (Seconds1 >= 60)
Min = Min + 1
Seconds1 = Seconds1 - 60
Wend
sec$ = CStr(Seconds - (Min * 60))
sec$ = String(2 - Len(sec$), "0") + sec$
ConvertSeconds = Min & ":" + sec$
End Function

Function CountMail () As Integer
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
MailWin% = getwindowbytitle(MDI%, "New Mail")
If MailWin% = 0 Then
    RunMenu "Read &New Mail"
    Do
    DoEvents
    MailWin% = getwindowbytitle(MDI%, "New Mail")
    NoMail% = FindWindow("#32770", "America Online")
    If NoMail% <> 0 Then
	x = sendmessagebynum(NoMail%, WM_CLOSE, 0, 0)
	MsgBox "you have no mail", 64
	CountMail = 0
	Exit Function
    End If
    Loop Until MailWin% <> 0
    List% = getwindowbyclass(MailWin%, "_AOL_Tree")
    If GetAOL() = 95 Then
	Do
	NumMail1 = sendmessagebynum(List%, LB_32GETCOUNT, 0, 0)
	Timeout (3)
	NumMail2 = sendmessagebynum(List%, LB_32GETCOUNT, 0, 0)
	Loop Until NumMail1 = NumMail2
    Else :
	Do
	    NumMail1 = sendmessagebynum(List%, LB_GETCOUNT, 0, 0)
	    Timeout (3)
	    NumMail2 = sendmessagebynum(List%, LB_GETCOUNT, 0, 0)
	Loop Until NumMail1 = NumMail2
    End If
End If
x = showwindow(MailWin%, SW_MINIMIZE)
List% = getwindowbyclass(MailWin%, "_AOL_Tree")
If GetAOL() = 95 Then
    NumMail = sendmessagebynum(List%, LB_32GETCOUNT, 0, 0)
Else :
    NumMail = sendmessagebynum(List%, LB_GETCOUNT, 0, 0)
End If
CountMail = NumMail
End Function

Function currentroom ()
Aol = FindWindow("AOL Frame25", 0&)
bah = FindChildByClass(Aol, "_AOL_Glyph")
par% = getparent(bah)
x$ = getwintext(par%)
currentroom = x$
End Function

Function Decrypt (txt As String, PW As String) As String
If Len(txt$) = 0 Then Exit Function
If Len(PW$) = 0 Then Exit Function
pwCounter = 1
For x = 1 To Len(txt$)
CurLetter = Mid(txt$, x, 1)
AscLetter = Asc(CurLetter)
pwLetter = Mid(PW$, pwCounter, 1)
pwCounter = pwCounter + 1
If pwCounter > Len(PW$) Then pwCounter = 1
ascPW = Asc(pwLetter)
Combined = AscLetter - ascPW
If Combined < 1 Then
OutLetter = Combined + 255
OutLett = Chr(OutLetter)
Else
OutLett = Chr(Combined)
End If
Text$ = Text$ + OutLett
DoEvents
Next x
Decrypt = Text$
End Function

Sub DeleteSent (what As String)
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
RunMenu "Check Mail You've &Sent"
OutGoingMail% = WaitForWin("Outgoing Mail")
Tree% = getwindowbyclass(OutGoingMail%, "_AOL_Tree")
DelButton% = getwindowbytitle(OutGoingMail%, "Delete")
click DelButton%
If GetAOL() = 95 Then
    Do
    NumMail1 = sendmessagebynum(Tree%, LB_32GETCOUNT, 0, 0)
    Timeout (2)
    NumMail2 = sendmessagebynum(Tree%, LB_32GETCOUNT, 0, 0)
    Loop Until NumMail1 = NumMail2
Else :
    Do
    NumMail1 = sendmessagebynum(Tree%, LB_GETCOUNT, 0, 0)
    Timeout (2)
    NumMail2 = sendmessagebynum(Tree%, LB_GETCOUNT, 0, 0)
    Loop Until NumMail1 = NumMail2
End If
DelButton% = getwindowbytitle(OutGoingMail%, "Delete")
If GetAOL() = 95 Then
    NumMail = sendmessagebynum(Tree%, LB_32GETCOUNT, 0, 0)
Else :
    NumMail = sendmessagebynum(Tree%, LB_GETCOUNT, 0, 0)
End If
For x = 0 To NumMail
    Text$ = Space(256)
    i = SendMEssageByString(Tree%, LB_GETTEXT, x, Text$)
    If InStr(SpaceCase(Text$), SpaceCase(what$)) Then
	i = sendmessagebynum(Tree%, LB_SETCURSEL, x, 0)
	click DelButton%
    End If
Next x
Call closewin(OutGoingMail%)
End Sub

Sub do3d (Obj As Control, Style%, Thick%)
On Error Resume Next
Obj.Parent.AutoRedraw = True
    If Thick <= 0 Then Thick = 1
    If Thick > 8 Then Thick = 8
    OldMode = Obj.Parent.ScaleMode
    OldWidth = Obj.Parent.DrawWidth
    Obj.Parent.ScaleMode = 3
    Obj.Parent.DrawWidth = 1
    ObjHeight = Obj.Height
    ObjWidth = Obj.Width
    ObjLeft = Obj.Left
    ObjTop = Obj.Top
    
    Select Case Style
	Case 1:
	    TLshade = QBColor(8)
	    BRshade = QBColor(15)
	Case 2:
	    TLshade = QBColor(15)
	    BRshade = QBColor(8)
	Case 3:
	    TLshade = RGB(0, 0, 255)
	    BRshade = QBColor(1)
    End Select
	For i = 1 To Thick
	    CurLeft = ObjLeft - i
	    CurTop = ObjTop - i
	    CurWide = ObjWidth + (i * 2) - 1
	    CurHigh = ObjHeight + (i * 2) - 1
	    Obj.Parent.Line (CurLeft, CurTop)-Step(CurWide, 0), TLshade
	    Obj.Parent.Line -Step(0, CurHigh), BRshade
	    Obj.Parent.Line -Step(-CurWide, 0), BRshade
	    Obj.Parent.Line -Step(0, -CurHigh), TLshade
	Next i
	If Thick > 2 Then
	    CurLeft = ObjLeft - Thick - 1
	    CurTop = ObjTop - Thick - 1
	    CurWide = ObjWidth + ((Thick + 1) * 2) - 1
	    CurHigh = ObjHeight + ((Thick + 1) * 2) - 1
	    Obj.Parent.Line (CurLeft, CurTop)-Step(CurWide, 0), QBColor(0)
	    Obj.Parent.Line -Step(0, CurHigh), QBColor(0)
	    Obj.Parent.Line -Step(-CurWide, 0), QBColor(0)
	    Obj.Parent.Line -Step(0, -CurHigh), QBColor(0)
	End If
    Obj.Parent.ScaleMode = OldMode
    Obj.Parent.DrawWidth = OldWidth
End Sub

Sub editgoto (Key As String, URL As String)
'Call WriteINI(GetAOLDir() + "\GOTO.INI", "[WAOLGOTO]", "GOTO1", Key$)
'Call WriteINI(GetAOLDir() + "\GOTO.INI", "[WAOLGOTO]", "KEYWORD1", URL$)
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
click Save%
Do
DoEvents
MODAL% = FindWindow("_AOL_Modal", "Favorite Places")
Loop Until MODAL% = 0
End Sub

Function Encrypt (txt As String, PW As String) As String
If txt$ = "" Then Exit Function
If PW$ = "" Then Exit Function
pwCounter = 1
For x = 1 To Len(txt$)
letter = Mid(txt$, x, 1)
AscLetter = Asc(letter)
pwLetter = Mid(PW$, pwCounter, 1)
pwCounter = pwCounter + 1
If pwCounter > Len(PW$) Then pwCounter = 1
ascPW = Asc(pwLetter)
LetterPW = AscLetter + ascPW
If LetterPW > 255 Then
OutLetter = LetterPW - 255
Lett = Chr(OutLetter)
Else
Lett = Chr(LetterPW)
End If
Text$ = Text$ + Lett
DoEvents
Next x
Encrypt = Text$
End Function

Sub Enter (ByVal hWnd As Integer)
x = sendmessagebynum(hWnd%, WM_CHAR, 13, 0)
End Sub

Sub ErrorMsg ()
Randomize Timer
Select Case Int(Rnd * 3)
Case 0:
    Msg$ = "please sign on before using this feature!"
Case 1:
    Msg$ = "sign on first!"
Case 2:
    Msg$ = "not signed on."
End Select
MsgBox Msg$, 48
End Sub

Sub Explode (hWnd As Form, CFlag As Integer, steps As Integer)
Dim FRect As RECT
Dim fWidth, fHeight As Integer
Dim i, x, Y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush
GetWindowRect hWnd.hWnd, FRect
fWidth = (FRect.Right - FRect.Left)
fHeight = FRect.Bottom - FRect.Top
hScreen = GetDC(0)
Brush = CreateSolidBrush(0)
OldBrush = SelectObject(hScreen, Brush)
Select Case CFlag
Case False:
    For i = 1 To steps
	cx = fWidth * (i / steps)
	cy = fHeight * (i / steps)
	x = FRect.Left
	Y = FRect.Top
	Rectangle hScreen, x, Y, x + cx, Y + cy
    Next i
Case True:
    For i = 1 To steps
	cx = fWidth * (i / steps)
	cy = fHeight * (i / steps)
	x = FRect.Left + (fWidth - cx) / 2
	Y = FRect.Top + (fHeight - cy) / 2
	Rectangle hScreen, x, Y, x + cx, Y + cy
    Next i
End Select
If ReleaseDC(0, hScreen) = 0 Then MsgBox "Unable to Release Device Context", 16, "Device Error"
DeleteObject (Brush)
x = showwindow(hWnd.hWnd, SW_NORMAL)
End Sub

Function extension (exe As String, ext As String) As String
If InStr(exe$, ".") <> 0 Then extension = exe$
If InStr(exe$, ".") = 0 Then extension = exe$ + "." + ext$
End Function

Sub falldown (frm As Form, steps As Integer)
On Error Resume Next
BgColor = frm.BackColor
frm.BackColor = RGB(0, 0, 0)
For x = 0 To frm.Count - 1
frm.Controls(x).Visible = False
Next x
AddX = True
AddY = True
frm.Show
x = ((Screen.Width - frm.Width) - frm.Left) / steps
Y = ((Screen.Height - frm.Height) - frm.Top) / steps
Do
    frm.Move frm.Left + x, frm.Top + Y
Loop Until (frm.Left >= (Screen.Width - frm.Width)) Or (frm.Top >= (Screen.Height - frm.Height))
frm.Left = Screen.Width - frm.Width
frm.Top = Screen.Height - frm.Height
frm.BackColor = BgColor
For x = 0 To frm.Count - 1
frm.Controls(x).Visible = True
Next x
End Sub

Sub FastIM (who As String, messa As String)
If online() = False Then Call ErrorMsg: Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
RunMenu "Send an Instant Message"
IM% = WaitForWin("Send Instant Message")
AOLEdit% = getwindowbyclass(IM%, "_AOL_Edit")
If GetAOL() = 2 Then Message% = getnextwindow(AOLEdit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then Message% = getwindowbyclass(IM%, "RICHCNTL")
Call setedit(AOLEdit%, LCase$(who$))
Call setedit(Message%, Mess$)
send% = getnextwindow(Message%, 1)
Timeout (.2)
Call closewin(IM%)
If InStr(1, who$, "$im_", 1) Then Call waitforok
End Sub

Sub FavePlace (method As Integer, Description As String, URL As String)
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
Select Case method
Case 1:
    If Aol% = 0 Then Exit Sub
    If GetAOL() = 2 Then
	AOLToolbar% = getwindowbyclass(Aol%, "AOL Toolbar")
	AOLIcon% = getnextwindow(GetWindow(AOLToolbar%, GW_CHILD), 19)
	click AOLIcon%
    End If
    If GetAOL() = 3 Or GetAOL() = 95 Then
	AOLToolbar% = getwindowbyclass(Aol%, "AOL Toolbar")
	AOLIcon% = getnextwindow(GetWindow(AOLToolbar%, GW_CHILD), 14)
	click AOLIcon%
    End If
    Do
	DoEvents
	FavoritePlaces% = getwindowbytitle(MDI%, "Favorite Places")
    Loop Until (FavoritePlaces% <> 0)
    If (FavoritePlaces% <> 0) Then
	AOLTree% = getwindowbyclass(FavoritePlaces%, "_AOL_Tree")
	num = sendmessagebynum(AOLTree%, LB_GETCOUNT, 0, 0)
	For x = 0 To num - 1
	    length = sendmessagebynum(AOLTree%, LB_GETTEXTLEN, x, 0)
	    Text$ = Space(length)
	    i = SendMEssageByString(AOLTree%, LB_GETTEXT, x, Text$)
	    Text$ = FixAPIString(Text$)
	    If SpaceCase(Text$) = SpaceCase(Description$) Then
		i = sendmessagebynum(AOLTree%, LB_SETCURSEL, x, 0)
		Exit Sub
	    End If
	Next x
	AddPlace% = getwindowbytitle(FavoritePlaces%, "Add Favorite Place")
	click AddPlace%
	Do
	DoEvents
	AddFavPlace% = getwindowbytitle(MDI%, "Add Favorite Place")
	Loop Until AddFavPlace% <> 0
	OK% = getwindowbytitle(AddFavPlace%, "OK")
	EnterDES% = GetWindow(getwindowbytitle(AddFavPlace%, "Enter the Place's Description:"), GW_HWNDNEXT)
	EnterURL% = GetWindow(getwindowbytitle(AddFavPlace%, "Enter the Internet Address:"), GW_HWNDNEXT)
	Call setedit(EnterDES%, Description$)
	Call setedit(EnterURL%, URL$)
	click OK%
    End If
Case 2:
    FavoritePlaces% = getwindowbytitle(MDI%, "Favorite Places")
    If FavoritePlaces% = 0 Then
	If GetAOL() = 2 Then
	    AOLToolbar% = getwindowbyclass(Aol%, "AOL Toolbar")
	    AOLIcon% = getnextwindow(GetWindow(AOLToolbar%, GW_CHILD), 19)
	    click AOLIcon%
	End If
	If GetAOL() = 3 Or GetAOL() = 95 Then
	    AOLToolbar% = getwindowbyclass(Aol%, "AOL Toolbar")
	    AOLIcon% = getnextwindow(GetWindow(AOLToolbar%, GW_CHILD), 14)
	    click AOLIcon%
	End If
    End If
    If (FavoritePlaces% <> 0) Then
	AOLTree% = getwindowbyclass(FavoritePlaces%, "_AOL_Tree")
	num = sendmessagebynum(AOLTree%, LB_GETCOUNT, 0, 0)
	For x = 0 To num - 1
	    length = sendmessagebynum(AOLTree%, LB_GETTEXTLEN, x, 0)
	    Text$ = Space(length)
	    i = SendMEssageByString(AOLTree%, LB_GETTEXT, x, Text$)
	    Text$ = FixAPIString(Text$)
	    If SpaceCase(Text$) = SpaceCase(Description$) Then
		i = sendmessagebynum(AOLTree%, LB_SETCURSEL, x, 0)
	    End If
	Next x
	Connect% = getwindowbytitle(FavoritePlaces%, "Connect")
	click Connect%
    End If
End Select
End Sub

Function File_IfileExists (ByVal sFileName As String) As Integer
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

Function FillIn (txt As String, Spaces As Integer, With As String) As String
txt$ = Trim$(txt$)
With$ = Trim$(With$)
many = Spaces - Len(txt$)
Things$ = String(many, With$)
FillIn = RemoveSpace(Things$ & txt$)
End Function

Function FindChatRoom () As Integer
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
Child% = GetWindow(GetWindow(MDI%, GW_CHILD), GW_HWNDFIRST)
Do
    AOLList% = getwindowbyclass(Child%, "_AOL_ListBox")
    AOLIcon% = getwindowbyclass(Child%, "_AOL_Icon")
    AOLEdit% = getwindowbyclass(Child%, "_AOL_Edit")
    AOLView% = getwindowbyclass(Child%, "_AOL_View")
    If (AOLList% <> 0) And (AOLView% <> 0) And (AOLEdit% <> 0) And (AOLIcon% <> 0) Then Exit Do
    Child% = GetWindow(Child%, GW_HWNDNEXT)
Loop Until Child% = 0

FindChatRoom = Child%
End Function

Function FixAPIString (sText As String) As String
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then FixAPIString = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then FixAPIString = Trim(sText$)
End Function

Function FixCaps (Text As String) As String
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

Function ForwardMail (Index As Integer, ForwardWin As Integer) As Integer
AOLVer = GetAOL()
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
NewMail% = getwindowbytitle(MDI%, "New Mail")
AOLTree% = getwindowbyclass(NewMail%, "_AOL_Tree")
Read_Mail% = getwindowbytitle(NewMail%, "Read")
If AOLVer = 95 Then
    x = sendmessagebynum(AOLTree%, LB_32SETCURSEL, Index%, 0)
Else :
    x = sendmessagebynum(AOLTree%, LB_SETCURSEL, Index%, 0)
End If
click Read_Mail%
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
    Timeout (.1)
    RunMenu "Stop Incoming Text"
    Call click(ForwardIcon%)
    SendWin% = GetWindow(MDI%, GW_CHILD)
    SendWin% = GetWindow(SendWin%, GW_HWNDFIRST)
    Do
    send% = GetChildWin(SendWin%, "Send Now", "_AOL_Icon")
    SendLater% = GetChildWin(SendWin%, "Send Later", "_AOL_Icon")
    If (send% <> 0) And (SendLater% <> 0) Then Exit Do
    SendWin% = GetWindow(SendWin%, GW_HWNDNEXT)
    Loop Until SendWin% = 0
Loop Until (SendWin% <> 0)
Child% = GetWindow(MDI%, GW_CHILD)
Child% = GetWindow(Child%, GW_HWNDFIRST)
Do
send% = GetChildWin(Child%, "Send Now", "_AOL_Icon")
SendLater% = GetChildWin(Child%, "Send Later", "_AOL_Icon")
If (send% <> 0) And (SendLater% <> 0) Then
    If (Child% <> SendWin%) Then Call closewin(Child%)
End If
Child% = GetWindow(Child%, GW_HWNDNEXT)
Loop Until Child% = 0
ForwardMail = SendWin%
End Function

Function ForwardMail2 (Index As Integer, ForwardWin As Integer) As Integer     'I coded 100% of this bas
AOLVer = GetAOL()
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
NewMail% = getwindowbytitle(MDI%, "New Mail")
AOLTree% = getwindowbyclass(NewMail%, "_AOL_Tree")
Read_Mail% = getwindowbytitle(NewMail%, "Read")
click Read_Mail%
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
    Timeout (.1)
    RunMenu "Stop Incoming Text"
    Call click(ForwardIcon%)
    SendWin% = GetWindow(MDI%, GW_CHILD)
    SendWin% = GetWindow(SendWin%, GW_HWNDFIRST)
    Do
    send% = GetChildWin(SendWin%, "Send Now", "_AOL_Icon")
    SendLater% = GetChildWin(SendWin%, "Send Later", "_AOL_Icon")
    If (send% <> 0) And (SendLater% <> 0) Then Exit Do
    SendWin% = GetWindow(SendWin%, GW_HWNDNEXT)
    Loop Until SendWin% = 0
Loop Until (SendWin% <> 0)
Child% = GetWindow(MDI%, GW_CHILD)
Child% = GetWindow(Child%, GW_HWNDFIRST)
Do
send% = GetChildWin(Child%, "Send Now", "_AOL_Icon")
SendLater% = GetChildWin(Child%, "Send Later", "_AOL_Icon")
If (send% <> 0) And (SendLater% <> 0) Then
    If (Child% <> SendWin%) Then Call closewin(Child%)
End If
Child% = GetWindow(Child%, GW_HWNDNEXT)
Loop Until Child% = 0
ForwardMail2 = SendWin%

End Function

Function Generate () As String
    Randomize Timer
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

Function GenerateAscii (Text As String) As String
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

Function GetAOL () As Integer
Aol% = FindWindow("AOL Frame25", 0&)
Menu% = GetMenu(Aol%)
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
    Else :
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

Function GetAOLDir () As String
On Error Resume Next
Aol% = FindWindow("AOL Frame25", 0&)
If Aol% = 0 Then Exit Function
Text$ = Space(256)
x = GetModuleFileName(GetWindowWord(Aol%, GWW_HINSTANCE), Text$, 255)
Text$ = FixAPIString(Trim(Text$))
For x = Len(Text$) To 1 Step -1
Char$ = Mid(Text$, x, 1)
If Char$ = "\" Then Exit For
Next x
Text$ = Mid(Text$, 1, x - 1)
GetAOLDir = Text$
End Function

Function getapitext (hWnd As Integer) As String
    x = sendmessagebynum(hWnd%, WM_GETTEXTLENGTH, 0, 0)
    Text$ = Space(x + 1)
    x = SendMEssageByString(hWnd%, WM_GETTEXT, x + 1, Text$)
    getapitext = FixAPIString(Text$)
End Function

Function GetChatText () As String
On Error Resume Next
If FindChatRoom() = 0 Then Exit Function
ChatText$ = getapitext(ChatView())
For x = Len(ChatText$) To 1 Step -1
If Mid(ChatText$, x, 1) = Chr(13) Then Exit For
Next x
ChatText$ = Mid(ChatText$, x, Len(ChatText$))
GetChatText = ChatText$
End Function

Function GetChildWin (Parent As Integer, Caption As String, Class As String) As Integer
Win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
text1$ = getapitext(Win%)
text2$ = GetClass(Win%)
If InStr(1, text1$, Caption$, 1) And InStr(1, text2$, Class$, 1) Then Exit Do
Win% = GetWindow(Win%, GW_HWNDNEXT)
Loop Until Win% = 0
GetChildWin = Win%
End Function

Function GetClass (hWnd As Integer) As String
Text$ = Space(1000)
x = GetClassName(hWnd%, Text$, 1000)
Text$ = FixAPIString(Text$)
GetClass = Text$
End Function

Function GetGuest () As Integer
MODAL% = FindWindow("_AOL_Modal", 0&)
ScreenName% = getnextwindow(getwindowbytitle(MODAL%, "Screen Name:"), 1)
Password% = getnextwindow(getwindowbytitle(MODAL%, "Password:"), 1)
If (ScreenName% <> 0) And (Password% <> 0) Then GetGuest = MODAL%
End Function

Function GetIMText (Sender As String, Message As String) As Integer
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
IM% = getwindowbytitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    GetIMText = True
    Title$ = getapitext(IM%)
    Sender$ = Trim$(Mid$(Title$, InStr(Title$, ":") + 1))
    If GetAOL() = 2 Then view% = getwindowbyclass(IM%, "_AOL_View")
    If GetAOL() = 3 Or GetAOL() = 95 Then view% = getwindowbyclass(IM%, "RICHCNTL")
    Text$ = getapitext(view%)
    Call closewin(IM%)
    Where = InStr(Text$, ":")
    Message$ = Mid$(Text$, Where + 2)
End If
End Function

Function GetINI (INIFile As String, AppName As String, KeyName As String) As String
   Dim RetStr As String
   RetStr = Space(30000)
   KeyName$ = LCase(KeyName$)
   INIText$ = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), INIFile$))
   GetINI = Replace(INIText$, Chr(30), Chr(13) + Chr(10))
End Function

Function GetMouseWindow () As Integer
Dim ptCursor As Long
Call GetCursorPos(ptCursor)
hWndOver% = WindowFromPoint(ptCursor)
GetMouseWindow = hWndOver%
End Function

Function getnextwindow (hWnd As Integer, num As Integer) As Integer
NexthWnd% = hWnd%
For x = 1 To num
NexthWnd% = GetWindow(NexthWnd%, GW_HWNDNEXT)
Next x
getnextwindow = NexthWnd%
End Function

Sub getnum (og%, A)
Do
If A = 0 Then Exit Sub
B = 1 + B
og% = GetWindow(og%, GW_HWNDNEXT)
Loop Until B >= A - 1
End Sub

Function GetProgramEXE (hWnd As Integer) As String
Text$ = Space(255)
x = GetModuleFileName(GetWindowWord(hWnd%, GWW_HINSTANCE), Text$, 255)
Text$ = FixAPIString(Trim(Text$))
GetProgramEXE = LCase(Text$)
End Function

Function GetRoomName () As String
On Error Resume Next
GetRoomName = getapitext(FindChatRoom())
End Function

Function GetRoomURL () As String
x = SetFocusAPI(FindChatRoom())
RunMenu "Add to Favorite Places"
If GetAOL() = 2 Then
    Aol% = FindWindow("AOL Frame25", 0&)
    MDI% = getwindowbyclass(Aol%, "MDIClient")
    AOLToolbar% = getwindowbyclass(Aol%, "AOL Toolbar")
    AOLIcon% = getnextwindow(GetWindow(AOLToolbar%, GW_CHILD), 19)
    click AOLIcon%
End If
Go:
Do
    DoEvents
    MODAL% = FindWindow("#32770", "Name Conflict")
    FavoritePlaces% = getwindowbytitle(MDI%, "Favorite Places")
Loop Until (MODAL% <> 0) Or (FavoritePlaces% <> 0)
If (MODAL% <> 0) Then
    Call closewin(MODAL%)
    click AOLIcon%
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
    click Modify%
    Do
    DoEvents
    Modi% = getwindowbytitle(MDI%, Text$)
    EnterURL% = GetWindow(getwindowbytitle(Modi%, "Enter the Internet Address:"), GW_HWNDNEXT)
    If EnterURL% <> 0 Then
	GetRoomURL = getapitext(EnterURL%)
	Call closewin(Modi%)
	Call closewin(FavoritePlaces%)
	Exit Do
    End If
    Loop
End If
End Function

Function GetSignOn () As Integer
If online() = True Then Exit Function
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
Goodbye% = getwindowbytitle(MDI%, "Goodbye From America Online!")
Welcome% = getwindowbytitle(MDI%, "Welcome")
If Goodbye% <> 0 Then SignOn% = Goodbye%
If Welcome% <> 0 Then SignOn% = Welcome%
If IsWindowVisible(SignOn%) <> 0 Then
  GetSignOn = SignOn%
Else :
  GetSignOn = 0
End If
End Function

Function GetSystemInfo (WinDir As String, SystemDir As String)
txt$ = Space(256)
x = GetWindowsDirectory(txt$, 256)
WinDir$ = LCase(FixAPIString(txt$))
txt$ = Space(256)
x = GetSystemDirectory(txt$, 256)
SystemDir$ = LCase(FixAPIString(txt$))
End Function

Function GetTextLen (hWnd As Integer) As Integer
GetTextLen = sendmessagebynum(hWnd%, WM_GETTEXTLENGTH, 0, 0)
End Function

Function getusersn () As String
On Error Resume Next
If online() = False Then Exit Function
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
Welcome% = getwindowbytitle(MDI%, "Welcome, ")
Text$ = getapitext(Welcome%)
If Len(Text$) = 0 Then Exit Function
Text$ = Mid(Text$, InStr(Text$, ",") + 2)
Text$ = Mid(Text$, 1, InStr(Text$, "!") - 1)
getusersn = Trim(Text$)
End Function

Function GetWin (Location As String, Key As String) As String
   Dim RetStr As String
   RetStr = Space(30000)
   GetWin = Left(RetStr, GetProfileString(Location$, ByVal Key$, "", RetStr, Len(RetStr)))
End Function

Function getwindowbyclass (Parent As Integer, ByVal Class As String) As Integer
Win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = GetClass(Win%)
If SpaceCase(Text$) = SpaceCase(Class$) Then Exit Do
If FindChild = True Then
    If GetWindow(Win%, GW_CHILD) Then
	ChildWin% = getwindowbyclass(Win%, Class$)
	If ChildWin% <> 0 Then
	    Win% = ChildWin%
	    Exit Do
	End If
    End If
End If
Win% = GetWindow(Win%, GW_HWNDNEXT)
Loop Until Win% = 0
getwindowbyclass = Win%
End Function

Function getwindowbytitle (Parent As Integer, ByVal Title As String) As Integer
Win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = FixAPIString(getapitext(Win%))
If InStr(1, Text$, Title$, 1) Then Exit Do
If FindChild = True Then
    If GetWindow(Win%, GW_CHILD) Then
	ChildWin% = getwindowbytitle(Win%, Title$)
	If ChildWin% <> 0 Then
	    Win% = ChildWin%
	    Exit Do
	End If
    End If
End If
Win% = GetWindow(Win%, GW_HWNDNEXT)
Loop Until Win% = 0
getwindowbytitle = Win%
End Function

Function getwintext (hWnd As Integer) As String
lentos = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
x = SendMEssageByString(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
getwintext = Buffer$
End Function

Function GuestLogOn (SN As String, PW As String, logoff As Integer, Pause As Single) As Integer
MODAL% = FindWindow("_AOL_Modal", 0&)
ScreenName% = getnextwindow(getwindowbytitle(MODAL%, "Screen Name:"), 1)
Password% = getnextwindow(getwindowbytitle(MODAL%, "Password:"), 1)
x = SetFocusAPI(ScreenName%)
If UCase(PW$) = UCase("[sn]") Then PW$ = SN$
Call setedit(ScreenName%, SN$)
Call setedit(Password%, PW$)
OK% = getwindowbytitle(MODAL%, "OK")
Call Enter(Password%)
StartTime = Timer
Do
    DoEvents
    Call KillAd
    Invalid% = FindWindow("#32770", "America Online")
    SignOn% = GetSignOn()
    If Invalid% <> 0 Then Exit Do
    If SignOn% <> 0 Then Exit Do
    If online() <> 0 Then Exit Do
    If (Timer - StartTime) > 4 Then
	x = SetFocusAPI(ScreenName%)
	Call setedit(ScreenName%, SN$)
	x = SetFocusAPI(Password%)
	Call setedit(Password%, PW$)
	Call Enter(Password%)
	StartTime = Timer
    End If
Loop
If (Invalid% <> 0) Then
    AOLStatic% = getnextwindow(getwindowbyclass(Invalid%, "Static"), 1)
    txt$ = getapitext(AOLStatic%)
    OK% = getwindowbytitle(Invalid%, "OK")
    click OK%
    Do
    DoEvents
    Invalid% = FindWindow("#32770", "America Online")
    Loop Until (Invalid% = 0)
    DoEvents
End If
If (SignOn% <> 0) Then
    AOLStatic% = getnextwindow(getwindowbyclass(SignOn%, "_AOL_Static"), 1)
    txt$ = getapitext(AOLStatic%)
    Timeout (Pause)
End If
If InStr(UCase(txt$), UCase("Incorrect")) Or InStr(UCase(txt$), UCase("Invalid")) Then Timeout (.2): GuestLogOn = False
If InStr(UCase(txt$), UCase("signed on")) Then GuestLogOn = True
If InStr(UCase(txt$), UCase("you have been disconnected")) Then GuestLogOn = True
If online() = True Then
    GuestLogOn = True
    If logoff = True Then
	Call signoff
	Do
	DoEvents
	Loop Until GetSignOn() <> 0
	Timeout (Pause)
    End If
End If
If (online() = False) And (GetSignOn() <> 0) Then Timeout (Pause)
End Function

Sub IM_Off ()
Instantmessage "$IM_Off", "Imz OFF                        Torture ßy Tac0"
waitforok
End Sub

Sub IM_On ()
Instantmessage "$IM_ON", "Imz On                        Torture ßy Tac0"
waitforok
End Sub

Sub im_online (person$)
RunMenu "Send an Instant Message"
Do
Aol = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(Aol, "Send Instant Message")
txt% = FindChildByClass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
txt% = FindChildByClass(bah, "_AOL_Edit")
Do
rich% = FindChildByClass(bah, "RICHCNTL")
bahqw% = FindChildByTitle(bah, "Send")
DoEvents
Timeout (.001)
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then

Sendd txt%, person$
Sendd rich%, ""
Timeout (.001)
getnum rich%, 3
click rich%
Else
Sendd txt%, person$
getnum txt%, 1
Sendd txt%, ""
getnum txt%, 3
click txt%
End If
Timeout (.001)
x = sendmessagebynum(bah, WM_CLOSE, 0, 0)
End Sub

Sub Instantmessage (person$, tt$)
okw = FindWindow("#32770", "America Online")
okb = FindChildByTitle(okw, "OK")
okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)
RunMenu "Send an Instant Message"
Do
Aol = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(Aol, "Send Instant Message")
txt% = FindChildByClass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
nah = FindChildByTitle(Aol, "To:")
x = SendMEssageByString(nah, WM_SETTEXT, 0, "To: ")
txt% = FindChildByClass(bah, "_AOL_Edit")
Do
rich% = FindChildByClass(bah, "RICHCNTL")
bahqw% = FindChildByTitle(bah, "Send")
DoEvents
Timeout (.001)
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then
Sendd txt%, person$
Sendd rich%, tt$
Timeout (.001)
getnum rich%, 1
click rich%
Else
Sendd txt%, person$
getnum txt%, 1
Sendd txt%, tt$
getnum txt%, 1
click txt%
End If
Timeout (.001)
Do
Aol = FindWindow("AOL Frame25", 0&)
okw = FindWindow("#32770", "America Online")
ba1h = FindChildByTitle(Aol, "To: ")
DoEvents
Loop Until ba1h = 0 Or okw <> 0
If okw <> 0 Then killwin (bah)
End Sub

Sub killwin (Windo)
x = sendmessagebynum(Windo, WM_CLOSE, 0, 0)
End Sub

Sub LCaseControls (frm As Form)
Screen.MousePointer = 11
On Error Resume Next
For x = 0 To frm.Controls.Count - 1
frm.Controls(x).Caption = LCase(frm.Controls(x).Caption)
Next x
frm.Caption = LCase(frm.Caption)
Screen.MousePointer = 0
End Sub

Sub lobby ()
Aol% = FindWindow("AOL Frame25", 0&)
If GetAOL() = 2 Then RunMenu "Lobby"
If GetAOL() = 3 Or GetAOL() = 95 Then Keyword "Lobby"
Do
DoEvents
Loop Until InStr(GetRoomName(), "Lobby")
End Sub

Function Locate (who As String) As String
If online() = False Then Call ErrorMsg: Exit Function
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
RunMenu "Send an Instant Message"
Do
    DoEvents
    IM% = getwindowbytitle(MDI%, "Send Instant Message")
Loop Until IM% <> 0
AOLEdit% = getwindowbyclass(IM%, "_AOL_Edit")
If GetAOL() = 2 Then Message% = getnextwindow(AOLEdit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then Message% = getwindowbyclass(IM%, "RICHCNTL")
Call setedit(AOLEdit%, who$)
Call setedit(Message%, " ")
Avail% = getwindowbytitle(IM%, "Available")
If Avail% = 0 Then Avail% = getnextwindow(Message%, 1)
click Avail%
Do
    DoEvents
    Off% = FindWindow("#32770", "America Online")
    MsgStatic% = getnextwindow(getwindowbyclass(Off%, "Static"), 1)
    If (Off% <> 0) Then
	txt$ = getapitext(MsgStatic%)
	Locate = txt$
	Call closewin(Off%)
	Call closewin(IM%)
	Exit Do
    End If
Loop
Call closewin(IM%)

End Function

Function LocateOnline (who As String) As Integer
If online() = False Then Call ErrorMsg: Exit Function
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
RunMenu "Send an Instant Message"
Do
    DoEvents
    IM% = getwindowbytitle(MDI%, "Send Instant Message")
Loop Until IM% <> 0
AOLEdit% = getwindowbyclass(IM%, "_AOL_Edit")
If GetAOL() = 2 Then Message% = getnextwindow(AOLEdit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then Message% = getwindowbyclass(IM%, "RICHCNTL")
Call setedit(AOLEdit%, who$)
Call setedit(Message%, " ")
Avail% = getwindowbytitle(IM%, "Available")
If Avail% = 0 Then Avail% = getnextwindow(Message%, 1)
click Avail%
Do
    DoEvents
    Off% = FindWindow("#32770", "America Online")
    MsgStatic% = getnextwindow(getwindowbyclass(Off%, "Static"), 1)
    If (Off% <> 0) Then
	txt$ = getapitext(MsgStatic%)
	If InStr(1, txt$, "is not currently signed on.", 1) Then
	    LocateOnline = False
	Else :
	    LocateOnline = True
	End If
	Call closewin(Off%)
	Call closewin(IM%)
	Exit Do
    End If
Loop
Call closewin(IM%)
End Function

Function MakeKilla () As String
For x = 1 To 6
    Randomize Timer
    For i = 1 To 14
	mac = Int(Rnd * 3)
	If mac = 0 Then letter$ = "—"
	If mac = 1 Then letter$ = "…"
	If mac = 2 Then letter$ = "@"
	Killa$ = Killa$ & letter$
    Next
    Killa$ = Killa$ & " "
Next
MakeKilla = Killa$
End Function

Sub miniwin (hWnd%)
Dim MiniTheWindow
MiniTheWindow = showwindow(hWnd%, SW_MINIMIZE)
End Sub

Function MM_createmaillist (Lst As ListBox) As String
For i = 0 To Lst.ListCount - 1
    Final$ = Final$ & "," & Lst.List(i)
    Next i
MM_createmaillist$ = "( " & Final$ & " )"
End Function

Sub MM_FillFwd (T As String, Message As String, KeepFwd As Integer)
'If KeepFwd = False then MM_FillFwd will remove the first 5
'characters of the "Subject:" line ("Fwd: ").

Aol% = AC_AOL()
Ver% = AC_AOLVersion()
FwdWin% = FindChildByTitle(Aol%, "Fwd:")
If FwdWin% = 0 Then Debug.Print "Fwd Window NOT found!": Exit Sub
ToEd% = FindChildByClass(FwdWin%, "_AOL_EDIT")
SubjEd% = AC_GetAOLWin(FwdWin%, "_AOL_EDIT", 3)
If Ver% = 25 Then MainEd% = AC_GetAOLWin(FwdWin%, "_AOL_EDIT", 4) Else MainEd% = FindChildByClass(FwdWin%, "RICHCNTL")
SendBttn% = FindChildByClass(FwdWin%, "_AOL_ICON")
DoEvents
setedit ToEd%, T$
setedit MainEd%, CStr(Message$)
If KeepFwd% = False Then
    DoEvents
    SubjText$ = getwintext(SubjEd%)
    SubjText$ = Mid$(SubjText$, 5)
    setedit SubjEd%, SubjText$
    End If
click SendBttn%
End Sub

Function MM_FixErrorGetSN (MList As String) As String
Aol% = FindWindow("AOL Frame25", 0&)
Do Until ErrorS <> 0
    DoEvents
    ErrorS = FindChildByTitle(Aol%, "Error")
    ErrorM = FindChildByClass(ErrorS, "_AOL_View")
    Timeout (.001)
    Loop
Do Until Len(ErrorText$) <> 0
    DoEvents
    z = sendmessagebynum(ErrorM, WM_GETTEXTLENGTH, 0, 0&)
    ErrorText2$ = String$(z + 1, 0)
    G% = SendMEssageByString(ErrorM, WM_GETTEXT, 0, ErrorText2$)
    ErrorText$ = Left(ErrorText2$, G%)
    Timeout (.001)
    Loop
ErrorOK = FindChildByTitle(ErrorS, "OK")
ButtonDown = sendmessagebynum(ErrorOK, WM_LBUTTONDOWN, 0&, 0&)
ButtonUp = sendmessagebynum(ErrorOK, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do Until DashPos <> 0
    DoEvents
    DashPos = InStr(ErrorText$, "-")
    Timeout (.001)
    Loop
ErrorText$ = Left$(ErrorText$, DashPos - 2)
ErrorText$ = AF_Script(ErrorText$, "The following problems occurred while processing your request:" & Chr(13) & Chr(10) & Chr(13) & Chr(10), "")
MM_FixErrorGetSN = ErrorText$
End Function

Sub moveform (frm As Form)
ReleaseCapture
ReturnVal% = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, 2, 0)
End Sub

Function mysn ()
Aol = FindWindow%("AOL Frame25", "America  Online")
happy% = FindChildByTitle%(Aol, "Welcome,")
joy$ = getwintext((happy%))
If InStr(joy$, "!") = 0 Then Exit Function
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
mysn = yoursn
End Function

Function online () As Integer
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
Welcome% = getwindowbytitle(MDI%, "Welcome, ")
Cap$ = getapitext(Welcome%)
If InStr(Cap$, ",") <> 0 Then online = True
If InStr(Cap$, ",") = 0 Then online = False
End Function

Function Percent (Complete As Integer, total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / total * TotalOutput)
End Function

Sub PercentBar (Shape As Control, Done As Integer, total As Variant)
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

Sub playwave (file$)
x = sndPlaySound(file$, 1)
End Sub

Sub PrintText (hWnd As Integer, x As Integer, Y As Integer, Text As String)
hDC% = GetDC(hWnd%)
x = TextOut(hDC%, x, Y, Text$, Len(Text$))
End Sub

Function PrivateRoom (room As String) As Integer
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
If FindChatRoom() = 0 Then Call lobby
PublicRooms% = getwindowbytitle(MDI%, "Public Rooms")
MemberRooms% = getwindowbytitle(MDI%, "Member Rooms")
If PublicRooms% <> 0 Then RoomsList% = PublicRooms%
If MemberRooms% <> 0 Then RoomsList% = MemberRooms%
ListRooms% = GetWindow(GetWindow(getwindowbyclass(FindChatRoom(), "_AOL_ListBox"), GW_HWNDNEXT), GW_HWNDNEXT)
If RoomsList% = 0 Then
    click (ListRooms%)
    Do
    DoEvents
    PublicRooms% = getwindowbytitle(MDI%, "Public Rooms")
    MemberRooms% = getwindowbytitle(MDI%, "Member Rooms")
    Loop Until (PublicRooms% <> 0) Or (MemberRooms% <> 0)
    If PublicRooms% <> 0 Then RoomsList% = PublicRooms%
    If MemberRooms% <> 0 Then RoomsList% = MemberRooms%
    List% = getnextwindow(getwindowbytitle(RoomsList%, "Double-click to Select a Room"), 1)
    Do
    x = sendmessagebynum(List%, LB_GETCOUNT, 0, 0)
    Timeout (2)
    z = sendmessagebynum(List%, LB_GETCOUNT, 0, 0)
    Loop Until x = z
End If
If MemberRooms% <> 0 Then PRButton% = getwindowbytitle(RoomsList%, "Private Rooms")
If PublicRooms% <> 0 Then PRButton% = getwindowbytitle(RoomsList%, "Private Room")
click PRButton%
Do
DoEvents
Pr% = getwindowbytitle(MDI%, "Enter a Private Room")
Loop Until Pr% <> 0
AOLEdit% = getwindowbyclass(Pr%, "_AOL_Edit")
Call setedit(AOLEdit%, room$)
Call Enter(AOLEdit%)
Do
DoEvents
Full% = FindWindow("#32770", "America Online")
Loop Until (Full% <> 0) Or (SpaceCase(GetRoomName()) = SpaceCase(room$))
If (Full% <> 0) Then
    Call closewin(Full%)
    PrivateRoom = False
End If
If (SpaceCase(GetRoomName()) = SpaceCase(room$)) Then
    PrivateRoom = True
End If

End Function

Sub profile ()
Do
 DoEvents
   aol1 = FindWindow("AOL Frame25", 0&)
   mow = FindChildByTitle(aol1, "Member Directory")
   Keyword "memberdirectory"
   Timeout (3)
Loop Until mow <> 0

Timeout (.54)

   aol1 = FindWindow("AOL Frame25", 0&)
   mow = FindChildByTitle(aol1, "Member Directory")
   Lst% = FindChildByClass(mow, "_AOL_Icon")
   getnum Lst%, 0
   click Lst%
   
Do
 DoEvents
   aol1 = FindWindow("AOL Frame25", 0&)
   mow = FindChildByTitle(aol1, "Edit Your Online Profile")
Loop Until mow <> 0
   
   Timeout (.54)

   aol1 = FindWindow("AOL Frame25", 0&)
   mow = FindChildByTitle(aol1, "Edit Your Online Profile")
   Editt% = FindChildByClass(mow, "_AOL_Edit")
   getnum Editt%, 0
   sends Editt%, "i am dead"
   getnum Editt%, 1
   sends Editt%, "I died on " + Date + ""

   getnum Editt%, 2
   sends Editt%, "I died at " + Time + """"
   getnum Editt%, 5
   sends Editt%, "geez..those damn deltree's"
   getnum Editt%, 1
   sends Editt%, "i'm gonna watch what i download next time!"
   getnum Editt%, 2
   sends Editt%, "well..im dead  :((("
   getnum Editt%, 1
   sends Editt%, ""
   getnum Editt%, 2
   sends Editt%, "im dead, RIP " & Time & " " & Date & ""


   
   Butn% = FindChildByClass(mow, "_AOL_Button")
   getnum Butn%, 3
   click (Butn%)
   Iqon% = FindChildByClass(mow, "_AOL_Icon")
   getnum Iqon%, 0
   click (Iqon%)
   waitforok
   mow = FindChildByTitle(aol2, "Member Directory")
killwin (mow)
killwin (mow)
killwin (mow)
End Sub

Function RandomText (num As Integer) As String
Randomize Timer
For x = 1 To num
i = Int(Rnd * 26)
Select Case i
    Case 0: letter$ = "A"
    Case 1: letter$ = "B"
    Case 2: letter$ = "C"
    Case 3: letter$ = "D"
    Case 4: letter$ = "E"
    Case 5: letter$ = "F"
    Case 6: letter$ = "G"
    Case 7: letter$ = "H"
    Case 8: letter$ = "I"
    Case 9: letter$ = "J"
    Case 10: letter$ = "K"
    Case 11: letter$ = "L"
    Case 12: letter$ = "M"
    Case 13: letter$ = "N"
    Case 14: letter$ = "O"
    Case 15: letter$ = "P"
    Case 16: letter$ = "Q"
    Case 17: letter$ = "R"
    Case 18: letter$ = "S"
    Case 19: letter$ = "T"
    Case 20: letter$ = "U"
    Case 21: letter$ = "V"
    Case 22: letter$ = "W"
    Case 23: letter$ = "X"
    Case 24: letter$ = "Y"
    Case 25: letter$ = "Z"
End Select
i = Int(Rnd * 2)
txt$ = txt$ + letter$
Next x
RandomText = txt$
End Function

Function ReadINI (AppName, KeyName, FileName As String) As String
'Example: text4.text = ReadINI("DaProggy", "Lamers Name", app.path + "\Prog.ini")
Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))

End Function

Sub ReEnter ()
Call Keyword("aol://2719:2-2-" + GetRoomName())
StartTime = Timer
Do While (Timer - StartTime < 5)
    DoEvents
    Text$ = GetChatText()
    If InStr(Text$, "OnlineHost:") Then Exit Do
    Full% = FindWindow("#32770", "America Online")
    If Full% <> 0 Then
	Call closewin(Full%)
	Exit Do
    End If
Loop
Call killwait
End Sub

Sub remove (List As ListBox)
If List.ListIndex = -1 Then Exit Sub
List.RemoveItem List.ListIndex
End Sub

Function RemoveDeadFromString (Text As String, Names As String) As String
Text$ = getapitext(AOLView%)
Text$ = Mid(Text$, InStr(Text$, ":") + 5)
Text$ = Replace(Text$, Chr(13) + Chr(10), Chr(13))
Text$ = Text$ + Chr(13)
Text$ = LCase(RemoveSpace(Text$))
Names$ = RemoveSpace(Names$)
Do While InStr(Text$, "-") > 0
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
RemoveDeadFromString = Names$
End Function

Function RemoveDeadNames (List As ListBox) As Integer
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
ErrorWin% = getwindowbytitle(MDI%, "Error")
If ErrorWin% <> 0 Then
    AOLView% = getwindowbyclass(ErrorWin%, "_AOL_View")
    Text$ = getapitext(AOLView%)
Search_For_Dead:
    For x = 0 To List.ListCount - 1
	If InStr(1, RemoveSpace(Text$), RemoveSpace((List.List(x))), 1) Then List.RemoveItem x: GoTo Search_For_Dead
    Next x
    RemoveDeadNames = True
    Call closewin(ErrorWin%)
End If
End Function

Sub RemoveList (List As ListBox, who$)
On Error Resume Next
For x = 0 To List.ListCount - 1
If UCase(who$) = UCase(List.List(x)) Then List.RemoveItem x: Exit Sub
Next x
End Sub

Function RemoveSpace (txt$) As String
NoSpace$ = txt$
While InStr(NoSpace$, " ") <> 0
Where = InStr(NoSpace$, " ")
NoSpace$ = Mid(NoSpace$, 1, Where - 1) + Mid(NoSpace$, Where + 1)
Wend
RemoveSpace = NoSpace$
End Function

Function RemoveString (txt As String, Char As String) As String
NoChar$ = txt$
While InStr(NoChar$, Char$) <> 0
Where = InStr(NoChar$, Char$)
NoChar$ = Mid(NoChar$, 1, Where - 1) + Mid(NoChar$, Where + Len(Char$))
Wend
RemoveString = NoChar$
End Function

Function Replace (Object As String, what As String, With As String) As String
Text$ = Object$
Do While (InStr(1, Text$, what$, 1) > 0)
    Where = InStr(1, Text$, what$, 1)
    If (Where > 0) Then
	LeftSide$ = Mid(Text$, 1, Where - 1)
	RightSide$ = Mid(Text$, Where + Len(what$))
	Text$ = LeftSide$ + With$ + RightSide$
	Replace = Text$
    End If
Loop
Replace = Text$
End Function

Sub ResetSN (SN As String, aoldir As String, Replace As String)
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

Function RoomToString () As String
ChatRoom% = FindChatRoom()
If ChatRoom% = 0 Then Exit Function
AOLListBox% = getwindowbyclass(ChatRoom%, "_AOL_ListBox")
num = sendmessagebynum(AOLListBox%, LB_GETCOUNT, 0, 0)
UserSN$ = getusersn()
For Index% = 0 To num - 1
    x = AOLGetList(AOLListBox%, Index%, SN$)
    If (SN$ <> UserSN$) And (SN$ <> "") Then Names$ = Names$ + SN$ + ","
Next Index%
RoomToString = Names$
End Function

Sub run (ByVal menuname As String)
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
	Ret = sendmessagebynum(AOLx, WM_COMMAND, GMnID, 0)
	DoEvents
	Exit Sub
    End If
    DoEvents
    Next SCnt
Next Cnt

End Sub

Sub RunMenu (MenuCaption As String)
Aol% = FindWindow("AOL Frame25", 0&)
Menu% = GetMenu(Aol%)
ID% = SearchMenu(Menu%, MenuCaption$)
x = sendmessagebynum(Aol%, WM_COMMAND, ID%, 0)

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

Function ScanFile (FileName As String, SearchString As String) As Long
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

Function Scramble (Word$) As String
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

Function ScrambleWord (Word$)
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
Randomize Timer
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

Sub ScrollMacro (Text As String)
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    sendroom Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If Counter = 4 Then
	Timeout (2.9)
	Counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub

Function SearchMenu (mnuWnd As Integer, MenuCaption As String) As Integer
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

Sub SelectGuestLogon ()
If GetAOL() = 95 Then
    num = sendmessagebynum(AOLCombo%, CB_GETCOUNT, 0, 0)
    x = sendmessagebynum(AOLCombo%, CB_SETCURSEL, num - 2, 0)
Else :
    AOLCombo% = getwindowbyclass(GetSignOn(), "_AOL_ComboBox")
    x = sendmessagebynum(AOLCombo%, CB_SHOWDROPDOWN, True, 0)
    num = sendmessagebynum(AOLCombo%, CB_GETCOUNT, 0, 0)
    For Index% = 0 To num - 1
	x = sendmessagebynum(AOLCombo%, CB_SETCURSEL, Index%, 0)
	i = AOLGetCombo(AOLCombo%, Index%, txt$)
	If SpaceCase(txt$) = UCase("guest") Then
	    Found = True
	    Exit For
	End If
    Next Index%
    x = sendmessagebynum(AOLCombo%, CB_SHOWDROPDOWN, False, 0)
    If Found = True Then x = sendmessagebynum(AOLCombo%, CB_SETCURSEL, Index%, 0)
End If
End Sub

Sub Sendd (chatedit, sill$)
sndtext = SendMEssageByString(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub SendEMail (Names As String, Subject As String, Message As String, SendMail As Integer, WaitForSend As Integer)
On Error Resume Next
If online() = False Then Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
RunMenu "&Compose Mail"
Do
DoEvents
Compose% = getwindowbytitle(MDI%, "Compose Mail")
Loop Until Compose% <> 0
NamesEdit% = GetWindow(getwindowbytitle(Compose%, "To:"), GW_HWNDNEXT)
SubjectEdit% = GetWindow(getwindowbytitle(Compose%, "Subject:"), GW_HWNDNEXT)
If GetAOL() = 2 Then MessageEdit% = getnextwindow(SubjectEdit%, 3)
If GetAOL() = 3 Or GetAOL() = 95 Then MessageEdit% = getwindowbyclass(Compose%, "RICHCNTL")
send% = getwindowbyclass(Compose%, "_AOL_Icon")
Send_The_Mail:
Call setedit(NamesEdit%, Names$)
Call setedit(SubjectEdit%, Subject$)
Call setedit(MessageEdit%, Message$)
If SendMail = True Then click send%
If WaitForSend <> 0 Then
    Do
    DoEvents
    Compose% = getwindowbytitle(MDI%, "Compose Mail")
    Sent% = FindWindow("#32770", "America Online")
    ErrorWin% = getwindowbytitle(MDI%, "Error")
    If (Compose% = 0) Then Exit Do
    If (Sent% <> 0) Then
	Call closewin(Sent%)
	Call closewin(Compose%)
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
	Call closewin(ErrorWin%)
	GoTo Send_The_Mail
    End If
    Loop
End If
End Sub

Function sendim (ByVal who As String, ByVal Mess As String, Pause) As Integer
If online() = False Then Exit Function
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDIClient")
RunMenu "Send an Instant Message"
IM% = WaitForWin("Send Instant Message")
AOLEdit% = getwindowbyclass(IM%, "_AOL_Edit")
If GetAOL() = 2 Then Message% = getnextwindow(AOLEdit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then Message% = getwindowbyclass(IM%, "RICHCNTL")
Call setedit(AOLEdit%, LCase$(who$))
Call setedit(Message%, Mess$)
send% = getnextwindow(Message%, 1)
If Pause = True Then
    click send%
    Do
    DoEvents
	IM% = getwindowbytitle(MDI%, "Send Instant Message")
	Off% = FindWindow("#32770", "America Online")
	If Off% <> 0 Then
	    sendim = False
	    Call closewin(Off%)
	    Call closewin(IM%)
	    Exit Do
	End If
	If IM% = 0 Then Exit Do
	DoEvents
    Loop
    DoEvents
    If IM% <> 0 Then Call closewin(IM%)
    If Off% = 0 And IM% = 0 Then sendim = True
End If
End Function

Sub SendInvite (who As String, Message As String, room As String, Check As Integer)
If online() = False Then Exit Sub
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
BuddyView% = getwindowbytitle(MDI%, "Buddy List Window")
If BuddyView% = 0 Then
    Keyword "Buddy View"
    BuddyView% = WaitForWin("Buddy List Window")
End If
BuddyChatIcon% = getnextwindow(getwindowbyclass(BuddyView%, "_AOL_ListBox"), 4)
click BuddyChatIcon%
SendBuddyChat% = WaitForWin("Buddy Chat")
WhoBox% = getwindowbyclass(SendBuddyChat%, "_AOL_Edit")
MessageBox% = getnextwindow(WhoBox%, 2)
Where% = getnextwindow(MessageBox%, 6)
send% = getwindowbyclass(SendBuddyChat%, "_AOL_Icon")
Call setedit(WhoBox%, who$)
Call setedit(MessageBox%, Message$)
Call setedit(Where%, Mid(room$, 1, 75))
click send%
Do
DoEvents
i = i + 1
If i = 50 Then i = 0: click send%
If Check = True Then
    inFrom% = getwindowbytitle(MDI%, "Invitation From: " + getusersn())
End If
If Check = False Then
    SendBuddyChat% = getwindowbytitle(MDI%, "Buddy Chat")
    If SendBuddyChat% = 0 Then Exit Do
End If
Loop Until inFrom% <> 0
Call closewin(inFrom%)
End Sub

Sub sendroom (ByVal Text As String)
DoEvents
AOLEdit% = getwindowbyclass(FindChatRoom(), "_AOL_Edit")
send% = GetWindow(AOLEdit%, GW_HWNDNEXT)
Call setedit(AOLEdit%, Text$)
Call click(send%)
DoEvents
End Sub

Sub sends (chatedit, Geno$)
SendTheText = SendMEssageByString(chatedit, WM_SETTEXT, 0, Geno$)
End Sub

Sub sendtext (sill$)
Aol% = FindWindow("AOL Frame25", 0&)
chatlist% = FindChildByClass(Aol%, "_AOL_Glyph")
chatwin% = getparent(chatlist%)
chatedit% = FindChildByClass(chatwin%, "_AOL_Edit")
sndtext% = SendMEssageByString(chatedit%, WM_SETTEXT, 0, sill$)
SendNow% = sendmessagebynum(chatedit%, WM_CHAR, &HD, 0)
Timeout (.00000000001)

End Sub

Sub SendView (SN As String, txt As String)
Call setedit(ChatView(), Chr$(13) + Chr$(10) + SN$ + ":" + Chr$(9) + txt$)
End Sub

Sub SetChatPref (Arrive As Integer, Leave As Integer, Sort As Integer)
Aol% = FindWindow("AOL Frame25", 0&)
If Aol% = 0 Then Exit Sub
If GetAOL() = 2 Then RunMenu "Set Preferences"
If GetAOL() = 3 Or GetAOL() = 95 Then RunMenu "Preferences"
Pref% = WaitForWin("Preferences")
ChatPref% = GetChildWin(Pref%, "Chat", "_AOL_Icon")
click ChatPref%
Do
DoEvents
ChatPrefs% = FindWindow("_AOL_MODAL", "Chat Preferences")
Loop Until ChatPrefs% <> 0
x = sendmessagebynum(getwindowbytitle(ChatPrefs%, "Notify me when members arrive"), BM_SETCHECK, Arrive, 0)
x = sendmessagebynum(getwindowbytitle(ChatPrefs%, "Notify me when members leave"), BM_SETCHECK, Leave, 0)
x = sendmessagebynum(getwindowbytitle(ChatPrefs%, "Alphabetize the member list"), BM_SETCHECK, Sort, 0)
OK% = getwindowbytitle(ChatPrefs%, "OK")
DoEvents
click OK%
Do
DoEvents
Msg% = FindWindow("#32770", "America Online")
If Msg% <> 0 Then Call closewin(Msg%)
ChatPrefs% = FindWindow("_AOL_MODAL", "Chat Preferences")
Loop Until ChatPrefs% = 0
Call closewin(Pref%)
End Sub

Sub setedit (AOLEdit As Integer, ByVal Text As String)
x = SendMEssageByString(AOLEdit%, WM_SETTEXT, Len(Text$), Text$)
End Sub

Sub SetMailPref (ConfirmFlag As Integer, CloseFlag As Integer)
Aol% = FindWindow("AOL Frame25", 0&)
If Aol% = 0 Then Exit Sub
If (GetAOL() = 2) Then RunMenu "Set Preferences"
If (GetAOL() = 3) Or (GetAOL() = 95) Then RunMenu "Preferences"
Pref% = WaitForWin("Preferences")
Mails% = getnextwindow(getwindowbytitle(Pref%, "Mail"), 1)
If (GetAOL() = 2) Then click (Mails%)
Do
DoEvents
If (GetAOL() = 3) Or GetAOL() = 95 Then click (Mails%)
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
If (GetAOL() = 2) Then Call click(OK%)
Do
DoEvents
If (GetAOL() = 3) Or (GetAOL() = 95) Then Call click(OK%)
MailPref% = FindWindow("_AOL_MODAL", "Mail Preferences")
Loop Until (MailPref% = 0)
Call closewin(Pref%)
End Sub

Sub SetMenuBMP (frm As Form, SubMenu As Integer, mnuID As Integer, BMP As Long)
Call ModifyMenu(SubMenu%, mnuID%, MF_BITMAP Or MF_BYCOMMAND, mnuID%, BMP)
hd% = GetDC(frm.hWnd)
hMemDC = CreateCompatibleDC(hd%)
x = ReleaseDC(frm.hWnd, hd%)
x = DeleteDC(hMemDC)
End Sub

Sub SetMenuFont (frm As Form, mnuWnd As Integer, FontName As String, FontSize As Integer)
ReDim hBmp(200) As Long
DC% = frm.hDC
hMemDC = CreateCompatibleDC(DC%)
mnuCount = GetMenuItemCount(mnuWnd%)
For num = 0 To mnuCount - 1
    mnuID% = GetMenuItemID(mnuWnd%, num)
    Text$ = Space(200)
    x = GetMenuString(mnuWnd%, num, Text$, 200, WM_USER)
    Text$ = FixAPIString(Text$)
    hFont = CreateFont(FontSize, 0, 0, 0, 200, 0, 0, 0, 0, 0, 0, 0, 0, FontName$)
    hBmp(num) = CreateCompatibleBitmap(hMemDC, Len(Text$) * 6, Abs(FontSize) + 2)
    hOldBmp = SelectObject(hMemDC, hBmp(num))
    x = SelectObject(hMemDC, hFont)
    x = SetTextColor(hMemDC, frm.ForeColor)
    Call Rectangle(hMemDC, -1, -1, 201, 19)
    If Len(Text$) > 0 Then
	x = TextOut(hMemDC, 0, 0, Text$, Len(Text$))
	hBmp(num) = SelectObject(hMemDC, hOldBmp)
	Call DeleteObject(hFont)
	Call ModifyMenu(mnuWnd%, mnuID%, MF_BITMAP Or MF_BYCOMMAND, mnuID%, hBmp(num))
    End If
    SubMenu% = GetSubMenu(mnuWnd%, num)
    If (SubMenu% <> 0) Then
	Call SetMenuFont(frm, SubMenu%, FontName$, FontSize)
    End If
Next num
x = ReleaseDC(frm.hWnd, DC%)
x = DeleteDC(hMemDC)
End Sub

Sub SetPhish (SN As String, PW As String)
MODAL% = FindWindow("_AOL_MODAL", 0&)
snEdit% = getwindowbyclass(MODAL%, "_AOL_Edit")
pwEdit% = getnextwindow(snEdit%, 2)
Call setedit(snEdit%, SN$)
Call setedit(pwEdit%, PW$)
OK% = getwindowbytitle(MODAL%, "OK")
If (OK% <> 0) Then click OK%
Continue% = getwindowbytitle(MODAL%, "Continue")
If (Continue% <> 0) Then click Continue%
End Sub

Sub setwindowfocus (Windo)
x% = SetFocusAPI(Windo)
End Sub

Function SickPhrase () As String
Randomize Timer
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

Sub signoff ()
If online() = True Then
    RunMenu "Sign Off"
    If GetAOL() = 2 Then
	Do
	DoEvents
	Sign% = FindWindow("_AOL_MODAL", "America Online")
	Loop Until (Sign% <> 0)
	Yes% = getwindowbytitle(Sign%, "&Yes")
	click Yes%
    End If
End If
End Sub

Function SpaceCase (Text As String) As String
txt$ = Text$
txt$ = Trim(UCase(RemoveSpace(txt$)))
SpaceCase = txt$
End Function

Sub SpiralScroll (txt As TextBox)
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

Sub stayontop (frm As Form)
On Error Resume Next
Dim success As Integer
success% = SetWindowPos((frm.hWnd), HWND_TOPMOST, 0, 0, 0, 0, SW_FLAGS)
End Sub

Function StringTF (Boolean As Integer) As String
If Boolean = False Then
    StringTF = "False"
Else :
    StringTF = "True"
End If
End Function

Sub Text_SpiralScroll (txt As TextBox)
x = txt.Text
thastart:
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = MYSTR
sendtext "•[" + txt + "]•"
If txt.Text = x Then
Exit Sub
End If
GoTo thastart
End Sub

Sub textleft (bah$)
heh$ = Left(bah$, 1)
was$ = Right(bah$, Len(bah$) - 1)
bah$ = was$ & heh$

End Sub

Sub textright (bah$)

heh$ = Right(bah$, 1)
werd$ = Left(bah$, Len(bah$) - 1)
bah$ = heh$ & werd$


End Sub

Sub textset (hWnd As Integer, what As String)
Dim r
r = SendMEssageByString(hWnd, &HC, 0, what)

End Sub

Function AC_AOL () As Integer
DoEvents
AC_AOL% = FindWindow("AOL Frame25", 0&)
End Function

Function AC_AOLVersion ()
Aol% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(Aol%, "Welcome, " + AC_GETSN())
aol3% = FindChildByClass(Wel%, "RICHCNTL")
If aol3% = 0 Then AC_AOLVersion = 25: Exit Function
If aol3% <> 0 Then
    If getwintext(Aol%) <> "America Online" Then AC_AOLVersion = 3 Else AC_AOLVersion = 4
    End If
End Function

Function AC_GetAOLWin (Parent As Integer, ByVal ClassToFind As String, num As Integer) As Integer

If Parent% = 0 Then Debug.Print "The parent window was NOT found.": Exit Function
Init% = FindChildByClass(Parent%, ClassToFind$)
If Init% = 0 Then Debug.Print "That classname does not exist on " & Str(Parent%): Exit Function
Count% = Count% + 1
If Count% = num% Then
    AC_GetAOLWin% = Init%
    Exit Function
    End If
Waste% = Init%
Do Until Found% <> 0
    DoEvents
    Waste% = GetWindow(Waste%, GW_HWNDNEXT)
    Buf$ = String$(255, " ")
    Q% = GetClassName(Waste%, Buf$, 254)
    DoEvents
    If LCase(TrimNull(Buf$)) Like LCase(ClassToFind$) Then Count% = Count% + 1
    If Count% = num% Then
	AC_GetAOLWin% = Waste%
	Found% = Waste%
	End If
    Loop
End Function

Function AC_GETSN () As String
'This gets the user's SN from the Welcome window.
Aol = FindWindow("AOL Frame25", 0&)
Wel = FindChildByTitle(Aol, "Welcome,")
If Wel = 0 Then AC_GETSN = "Not Online": Exit Function
namelen = SendMessage(Wel, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String$(namelen, 0)
x = SendMEssageByString(Wel, WM_GETTEXT, namelen, Buffer$)
A = InStr(Buffer$, ",")
SN$ = Mid$(Buffer$, A + 2, (Len(Buffer$) - (A + 1)))
SN$ = TrimNull(SN$)
AC_GETSN$ = SN$
End Function

Function AC_Online () As Integer
Aol% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(Aol%, "Welcome, ")
Com% = FindChildByClass(Wel%, "_AOL_COMBOBOX")
If Com% <> 0 Then Wel% = 0
If Wel% = 0 Then
    MsgBox "You need to be signed onto AOL to use this feature.", 0, "Sign on!"
    AC_Online% = False
    Exit Function
    End If
If Wel% <> 0 Then AC_Online% = True
End Function

Sub AddCombo (Combo As ComboBox, txt$)
On Error Resume Next
DoEvents
For x = 0 To Combo.ListCount - 1
    If UCase$(Combo.List(x)) = UCase$(txt$) Then Exit Sub
Next
If (Len(txt$) <> 0) Then Combo.AddItem txt$
End Sub

Sub AddComboRoom (Combo As ComboBox)
ChatRoom% = FindChatRoom()
If ChatRoom% = 0 Then Exit Sub
AOLListBox% = getwindowbyclass(ChatRoom%, "_AOL_ListBox")
num = sendmessagebynum(AOLListBox%, LB_GETCOUNT, 0, 0)
UserSN$ = getusersn()
For Index% = 0 To num - 1
x = AOLGetList(AOLListBox%, Index%, SN$)
If (SN$ <> UserSN$) And (SN$ <> "") Then Call AddCombo(Combo, SN$)
Next Index%
End Sub

Sub AddList (List As ListBox, txt$)
On Error Resume Next
DoEvents
For x = 0 To List.ListCount - 1
    If UCase$(List.List(x)) = UCase$(txt$) Then Exit Sub
Next
If Len(txt$) <> 0 Then List.AddItem txt$
End Sub

Sub AddMailBox (List As ListBox)
If online() = False Then Call ErrorMsg: Exit Sub
Abort = False
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
NewMail% = getwindowbytitle(MDI%, "New Mail")
If NewMail% = 0 Then
    RunMenu "Read &New Mail"
    Do
	DoEvents
	NewMail% = getwindowbytitle(Aol%, "New Mail")
	Msg% = FindWindow("#32770", "America Online")
	If Msg% <> 0 Then Call closewin(Msg%): Exit Sub
    Loop Until NewMail% <> 0
    MailLst% = getwindowbyclass(NewMail%, "_AOL_Tree")
    Do
	NumMailz = sendmessagebynum(MailLst%, LB_GETCOUNT, 0, 0)
	Timeout (2)
	NumMail = sendmessagebynum(MailLst%, LB_GETCOUNT, 0, 0)
    Loop Until NumMail = NumMailz
End If
MDI% = getwindowbyclass(Aol%, "MDI Client")
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

Sub Timeout (duration As Single)
StartTime = Timer
Do While (Timer - StartTime) < duration
    DoEvents
Loop
End Sub

Function TOSPhrase (Phrase As Integer) As String
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

Function TrimFileName (FileName As String) As String
For x = Len(FileName$) To 1 Step -1
If Mid(FileName$, x, 1) = "\" Then Exit For
Next x
If x <> 0 Then
    TrimFileName = Mid(FileName$, x + 1)
Else :
    TrimFileName = FileName$
End If
End Function

Function TrimNull (ByVal In) As String
For x = 1 To Len(In)
    If (Mid$(In, x, 1) <> Chr$(0)) Then
    total$ = total$ + Mid$(In, x, 1)
    Else
    GoTo NullDetect
    End If
Next
NullDetect:
TrimNull = total$
End Function

Sub WaitForClose (hWnd As Integer)
Do
DoEvents
Loop Until IsWindow(hWnd%) = False
End Sub

Sub waitforok ()
Do
DoEvents
Msg% = FindWindow("#32770", "America Online")
Loop Until Msg% <> 0
Call closewin(Msg%)
End Sub

Function WaitForWin (Caption As String) As Integer
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = getwindowbyclass(Aol%, "MDI Client")
Do While Win% = 0
DoEvents
Win% = getwindowbytitle(MDI%, Caption$)
Loop
WaitForWin = Win%
End Function

Sub WriteINI (sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

Sub WriteWin (Location As String, Key As String, Value As String)
x = WriteProfileString(Location$, Key$, Value$)
End Sub

