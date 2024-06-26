Global AllDaJunk
Global FS(10)
Global kdkd$
Global damn As String
Global Abrigot
Global manyz As Integer
Global azzz As String
Global ToMail As String
Global LALaA As String
Global Minz As Integer
Global MiRUnHeX   As String
Global BurIM As Integer

Type typPerson 'used in the Who is online thingie
  sn As String
  Location As String
End Type
Global Pay As String

Declare Function SetParent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
'Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)

'Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)

Global Answer              'Correct Answer
Global DoneFlag(100)       'Determines scramble success
Global MyError             'My Error Number
Global PersonsNum          'Current Person Number
Global PersonName(100)     'Up to 100 People can play
Global PersonScor(100)     'Each players score
Global NumPlayers          'The Actual Number Playing
Global ScrambledFlag       'Has the word beed scrambled
Global Words(100)          ' Array
Global Clues(100)          'Clues Data Base Array
Declare Function getfocus Lib "User" () As Integer



Declare Function GetMenu Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function GetSubMenu Lib "user" (ByVal hmenu As Integer, ByVal nPos As Integer) As Integer
Declare Function ModifyMenu Lib "user" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpString As Any) As Integer
Declare Function GetMenuItemID Lib "user" (ByVal hmenu As Integer, ByVal nPos As Integer) As Integer
Declare Function removemenu Lib "user" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Sub DrawMenuBar Lib "user" (ByVal hWnd As Integer)
Global Const MF_SEPARATOR = &H800
Global Const MF_POPUP = &H10
Global Const MF_BYPOSITION = &H400
Global DidMen As Integer
Global IsDil As Integer
Declare Function CreateMenu% Lib "user" ()
Declare Function AppendMenu% Lib "user" (ByVal hmenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem&)
Declare Function appendmenubystring% Lib "user" Alias "AppendMenu" (ByVal hmenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)
Declare Function InsertMenu% Lib "user" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any)

Declare Function getcurrenttime& Lib "user" ()
Declare Function GetFreeSystemResources Lib "user" (ByVal fuSysResource%) As Integer
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal maxsize As Integer, ByVal filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer
Declare Sub GetScrollRange Lib "user" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Function ReleaseDC Lib "user" (ByVal hWnd%, ByVal hdc%) As Integer
Type RECT
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
End Type
Type HelpWinInfo
  wStructSize As Integer
  X As Integer
  Y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type
Declare Sub SetBKColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
Declare Sub ReleaseCapture Lib "user" ()
Declare Sub SetWindowText Lib "user" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetParent% Lib "user" (ByVal hWnd As Integer)
Declare Sub closewindow Lib "user" (ByVal hWnd As Integer)
Declare Sub MoveWindow Lib "user" (ByVal hWnd As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Declare Sub bringwindowtotop Lib "user" (ByVal hWnd As Integer)
Declare Function GetCursor Lib "user" () As Integer
Declare Function getwindowtextlength Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function EnableMenuItem Lib "user" (ByVal hmenu As Integer, ByVal wIDEnableItem As Integer, ByVal wEnable As Integer) As Integer
Declare Function DestroyMenu Lib "user" (ByVal hmenu As Integer) As Integer
Declare Function GetWindowWord Lib "user" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function GetDC Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function getwindowtask Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function EnableWindow Lib "user" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function GetActiveWindow Lib "user" () As Integer
Declare Function destroywindow Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function createwindow% Lib "user" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hWndParent%, ByVal hmenu%, ByVal hInstance%, ByVal lpParam$)
Declare Function CreatePopupMenu Lib "user" () As Integer
Declare Function SetActiveWindow Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function SetFocusAPI Lib "user" Alias "SetFocus" (ByVal hWnd As Integer) As Integer
Declare Function showwindow Lib "user" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function SndPlaySound Lib "MMSystem" (lpsound As Any, ByVal flag As Integer) As Integer
Declare Function waveOutGetNumDevs Lib "MMSystem" () As Integer
Declare Function DeleteMenu Lib "user" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function getwindow Lib "user" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function getwindowtext Lib "user" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetClassName Lib "user" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
' Function GetFocus Lib "user" () As Integer
Declare Function GetNextWindow Lib "user" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetMenuItemCount Lib "user" (ByVal hmenu As Integer) As Integer
Declare Function GetMenuState Lib "user" (ByVal hmenu As Integer, ByVal wId As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetMenuString Lib "user" (ByVal hmenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function FindWindow Lib "user" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function gettopwindow Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function SendMessageByNum& Lib "user" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function SendMessageByString& Lib "user" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function sendmessage Lib "user" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function ExitWindows Lib "user" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer
Declare Function AOLGetList% Lib "311.Dll" (ByVal Index%, ByVal buf$)
Declare Function GetNames Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByValp2$) As Integer
Declare Function agGetStringFromLPSTR$ Lib "APIGUIDE.DLL" (ByVal lpString&)
Declare Function ptGetStringFromAddress Lib "VBMSG.VBX" () As String
'Declare Function SndPlaySound Lib "MMSYSTEM.DLL" (ByVal lpszSoundName$, ByVal wFlags%)
Declare Function FindChildByTitle% Lib "vbwfind.dll" (ByVal PARENT%, ByVal title$)
Declare Function FindChildByClass% Lib "vbwfind.dll" (ByVal PARENT%, ByVal title$)
Global Const WM_USER = &H400
Global Const LB_GETTEXTLEN = (WM_USER + 11)
Global Const LB_DELETESTRING = (WM_USER + 3)
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const LB_FINDSTRING = (WM_USER + 16)
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const GW_HWNDNEXT = 2
Global Const GW_CHILD = 5
Global Const WM_CLOSE = &H10
Global Const WM_DESTROY = &H2
Global Const WM_SETTEXT = &HC
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_CHAR = &H102
Global Const WM_COMMAND = &H111
Global Const MB_ICONSTOP = 16
Global Const MB_OK = 0
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Global Const MF_BYCOMMAND = &H0
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Global Const WM_gettext = &HD
Global Const CB_ADDSTRING = (WM_USER + 3)
Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_DELETESTRING = (WM_USER + 4)
Global Program_Title
Global da_num#
Global daaf%
Global entr
Global proG_STAT$
Global werd(11) As String
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Sub addlistchk (itm As String, lst As ListBox)
 Sccc = boo
If itm = Sccc Then Exit Sub
If lst.ListCount = 0 Then lst.AddItem itm: Exit Sub

Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub

Sub Addmenu ()
aol% = FindWindow("AOLFrame25", 0&)
ab1 = GetMenu(aol%)

For X = 1 To 15
D1 = removemenu(ab1, 7, MF_BYCOMMAND)
Call DrawMenuBar(aol%)
DoEvents
Next X
Call DrawMenuBar(aol%)

'These are all the sub-menus
DoEvents
i = CreateMenu()
  c = appendmenubystring(i, 0, 601, "TO&Ser")
  c = appendmenubystring(i, 0, 602, "&Multi Task Chat")
  c = appendmenubystring(i, 0, 603, "MaSS &IMer")
  c = appendmenubystring(i, 0, 604, "IM &Punter")
  c = appendmenubystring(i, 0, 605, "M&ass Mailer")
  c = appendmenubystring(i, 0, 609, "&Bust In!")
  c = appendmenubystring(i, 0, 610, "&CD Player")
  c = appendmenubystring(i, 0, 606, "&Greetz")
  c = appendmenubystring(i, 0, 607, "RooM &Killer")
  c = appendmenubystring(i, 0, 608, "&Lamer Tools")
  c = appendmenubystring(i, 0, 621, "&Scrambler")
  c = appendmenubystring(i, 0, 622, "Mass &Phisher ")
 ' C = appendmenubystring(i, 0, 628, "M Task &AOL")
aol% = FindWindow("AOL Frame25", 0&)
ab = GetMenu(aol%)
TheMenu = ab

'This is the the shit on the Menu Bar....
X = appendmenubystring(ab, MF_POPUP, i, "<-FaLLeN &KiNGDoM->")
Call DrawMenuBar(aol%)
End Sub

Sub AddRoom (lst As ListBox)
aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient")
List% = FindChildByClass(aol%, "_AOL_Listbox")
room% = GetParent(List%)
For Index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo end_addr
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
addlistchk namez$, lst
Next Index%
end_addr:

End Sub

Function AOhWnd ()
AOhWnd = FindWindow("AOL Frame25", "America  Online")
DoEvents

End Function

Sub AOLClick (E1 As Integer)
If proG_STAT$ = "OFF" Then
Exit Sub
End If

do_wn = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
Pause .008
u_p = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AolClicker (E1 As Integer)
'Clicks an AOL button with the given handle as E1

Exit Sub


do_wn = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
Pause .008
u_p = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
End Sub

Function AOLGetVer () As Integer
Dim TB As Integer, TBCount As Integer

aol = GetAOL()
TB = FindChildByClass(aol, "AOL Toolbar")
TBCount = WinCount(TB)
 
 If TBCount = 19 Then AOLGetVer = 25
 If TBCount = 18 Then AOLGetVer = 30

End Function

Function aolhwnd ()
a = FindWindow("AOL Frame25", 0&)
aolhwnd = a
End Function

Sub ASLstFromStr (lst As ListBox, Strng As String)
Debug.Print Strng
Cma = InStr(Strng, "«š ")
Debug.Print "XCMA: " & Cma
If Cma Then
    dasn = Mid$(Strng, 1, Cma - 1)
    lst.AddItem dasn
Debug.Print "SuP"
Else
    lst.AddItem Strng
    Debug.Print "ELNOWORKA"
    Exit Sub
End If
Do
cma2 = InStr(Cma + 3, Strng, "«š ")
If cma2 Then
    stpplc = Len(Strng) - (Cma + 3) - (Len(Strng) - cma2)
    dasn = Mid$(Strng, Cma + 3, stpplc)
    lst.AddItem dasn
ElseIf cma2 = False Then
    asn = Mid$(Strng, Cma + 3)
    If asn = "" Then Exit Sub
    lst.AddItem asn
End If
Cma = cma2
Loop Until Cma = False

End Sub

Function ASStrfromLst (lstname As ListBox) As String
On Error GoTo daeeror
For i = 1 To lstname.ListCount
    DaSTr = DaSTr & lstname.List(i - 1) & "«š "
Next i
ASStrfromLst = DaSTr
Exit Function
daeeror:
ASStrfromLst = ""

End Function

Sub Blast (sn As String)
Randomize Timer
n = Int(Rnd * 3) + 1
l0778 = Len(sn)
i = 1
c = 0
'For i = 1 To Len(sn)
Do
  letter = UCase(Mid(sn, i, 1))
  Randomize
  n = Int(Rnd * 3) + 1
  If letter = "A" Then
    If n = 1 Then
        werd(i) = "A is for the Anal sex he has with his boyfriend"
    ElseIf n = 2 Then
        werd(i) = "A is for Ass he has in his dreams"
    ElseIf n = 3 Then
        werd(i) = "A is for Asshole his mom sucks"
    End If
  ElseIf letter = "B" Then
    If n = 1 Then
        werd(i) = "B Is for BiGDoG that fucked him up"
    ElseIf n = 2 Then
        werd(i) = "B is for light Bulb he jerk himself with"
    ElseIf n = 3 Then
        werd(i) = "B is for Blowjob his father gave him"
    End If
  ElseIf letter = "C" Then
    If n = 1 Then
        werd(i) = "C is for Cum his family eats for dinner "
    ElseIf n = 2 Then
        werd(i) = "C is for the Cliff he fell from"
    ElseIf n = 3 Then
        werd(i) = "C is for his Computer infected with a virus"
  End If
  ElseIf letter = "D" Then
    If n = 1 Then
        werd(i) = "D is for shit he uses as Deodorant"
    ElseIf n = 2 Then
        werd(i) = "D is for Dick he lost"
    ElseIf n = 3 Then
        werd(i) = "D is for Doorway his fat mom got stuck in"
    End If
  ElseIf letter = "E" Then
    If n = 1 Then
        werd(i) = "E is for Eagle that stole his dick"
    ElseIf n = 2 Then
        werd(i) = "E is for Entertainment he gets from fucking his mom"
    ElseIf n = 3 Then
        werd(i) = "E is for Enter key that he sucked off"
    End If
  ElseIf letter = "F" Then
    If n = 1 Then
        werd(i) = "F is for Fucker - his dad's job"
    ElseIf n = 2 Then
        werd(i) = "F is for his Faggot family"
    ElseIf n = 3 Then
        werd(i) = "F is for Fucking his sister daily"
    End If
  ElseIf letter = "G" Then
    If n = 1 Then
        werd(i) = "G is for Grim his mom has while fucking him"
    ElseIf n = 2 Then
        werd(i) = "G is for Goner he is"
    ElseIf n = 3 Then
        werd(i) = "G is for is Gay sex he has with his dog"
    End If
  ElseIf letter = "H" Then
    If n = 1 Then
        werd(i) = "H is for Hell where his mom lives"
    ElseIf n = 2 Then
        werd(i) = "H is for Head someone stole from his dick"
    ElseIf n = 3 Then
        werd(i) = "H is for HIV he has"
    End If
  ElseIf letter = "I" Then
    If n = 1 Then
        werd(i) = "I is for Instructions he needs to fuck someone"
    ElseIf n = 2 Then
        werd(i) = "I is for Illness his mom got"
    ElseIf n = 3 Then
        werd(i) = "I is for Itch he got on his dick"
    End If
  ElseIf letter = "J" Then
    If n = 1 Then
        werd(i) = "J is for Jeans he puts in his ass"
    ElseIf n = 2 Then
        werd(i) = "J is for Jelly he he instead of cum"
    ElseIf n = 3 Then
        werd(i) = "J is for Jerk he is"
    End If
  ElseIf letter = "K" Then
    If n = 1 Then
        werd(i) = "K is for Killing his mother"
    ElseIf n = 2 Then
        werd(i) = "K is for Keyword: Kids where he goes everyday"
    ElseIf n = 3 Then
        werd(i) = "K is for Kindergarten where his mom works as a slut"
    End If
  ElseIf letter = "L" Then
    If n = 1 Then
        werd(i) = "L is for Ladder he fell from"
    ElseIf n = 2 Then
        werd(i) = "L is for Lamer he is"
    ElseIf n = 3 Then
        werd(i) = "L is for Love he gets from TOS Staff"
    End If
  ElseIf letter = "M" Then
    If n = 1 Then
        werd(i) = "M is for Morning when he lost his dick"
    ElseIf n = 2 Then
        werd(i) = "M is for Motherfucker he is"
    ElseIf n = 3 Then
        werd(i) = "M is for Making his life a hell"
    End If
  ElseIf letter = "N" Then
    If n = 1 Then
        werd(i) = "N is for Nerd his boyfriend is"
    ElseIf n = 2 Then
        werd(i) = "N is for Nobody he is"
    ElseIf n = 3 Then
        werd(i) = "N is for Nail he swallowed"
    End If
  ElseIf letter = "O" Then
    If n = 1 Then
        werd(i) = "O is for Oustrich he always wanted to fuck"
    ElseIf n = 2 Then
        werd(i) = "O is for Ops he never gets on IRC"
    ElseIf n = 3 Then
        werd(i) = "O is for Owning a macintosh"
    End If
  ElseIf letter = "P" Then
    If n = 1 Then
        werd(i) = "P is for Pervert his dad is"
    ElseIf n = 2 Then
        werd(i) = "P is for Pool where he got killed"
    ElseIf n = 3 Then
        werd(i) = "P is for Prison he sits for jerking off Steve Case"
    End If
  ElseIf letter = "Q" Then
    If n = 1 Then
        werd(i) = "Q is for Question he asked his mom when he saw her dick"
    ElseIf n = 2 Then
        werd(i) = "Q is for Quarter he stole from Steve Case"
    ElseIf n = 3 Then
        werd(i) = "Q is for Quart of cum he ate today"
    End If
  ElseIf letter = "R" Then
    If n = 1 Then
        werd(i) = "R is for Railroad where his mom made a suicide"
    ElseIf n = 2 Then
        werd(i) = "R is for Restaurant where he works as a slut"
    ElseIf n = 3 Then
        werd(i) = "R is for Room where his mom fucks him"
    End If
  ElseIf letter = "S" Then
    If n = 1 Then
        werd(i) = "S is for Steve Case who sucks his dick"
    ElseIf n = 2 Then
        werd(i) = "S is for Sex lessons his mom taught him"
    ElseIf n = 3 Then
        werd(i) = "S is for Story his mom told him in a bed"
    End If
  ElseIf letter = "T" Then
    If n = 1 Then
        werd(i) = "T is for TOSAdvisor who is his boyfriend"
    ElseIf n = 2 Then
        werd(i) = "T is for Tiny dick he got"
    ElseIf n = 3 Then
        werd(i) = "T is for Toaster where he cooked his dick with"
    End If
  ElseIf letter = "U" Then
    If n = 1 Then
        werd(i) = "U is for Ugly face his girlfriend has"
    ElseIf n = 2 Then
        werd(i) = "U is for Using AOHell"
    ElseIf n = 3 Then
        werd(i) = "U is for Ugly face his girlfriend has"
    End If
  ElseIf letter = "V" Then
    If n = 1 Then
        werd(i) = "V is for Vampire that ate his dick"
    ElseIf n = 2 Then
        werd(i) = "V is for Virii his girlfriend has"
    ElseIf n = 3 Then
        werd(i) = "V is for Violating TOS for sucking a dick out of turn"
    End If
  ElseIf letter = "W" Then
    If n = 1 Then
        werd(i) = "W is for Wuss he is"
    ElseIf n = 2 Then
        werd(i) = "W is for Winning a prize for a smallest dick ever"
    ElseIf n = 3 Then
        werd(i) = "W is for WAOL 1.0 he uses"
    End If
  ElseIf letter = "X" Then
    If n = 1 Then
        werd(i) = "X is for X-Ray he took when he broke his dick"
    ElseIf n = 2 Then
        werd(i) = "X is for Xerox copies of pictures of his naked mom he makes"
    ElseIf n = 3 Then
        werd(i) = "X is for X-Ray he uses to get high"
    End If
  ElseIf letter = "Y" Then
    If n = 1 Then
        werd(i) = "Y is for Yawn his girlfriend makes when she see his tiny dick"
    ElseIf n = 2 Then
        werd(i) = "Y is for Yogurt he puts into his mom's ass"
    ElseIf n = 3 Then
        werd(i) = "Y is for Year of fucking his dog"
    End If
  ElseIf letter = "Z" Then
    If n = 1 Then
        werd(i) = "Z is for Zoo where he was born"
    ElseIf n = 2 Then
        werd(i) = "Z is for Zebra who fucked him up"
    ElseIf n = 3 Then
        werd(i) = "Z is for Zillion times he jerked off"
    End If
  ElseIf letter = "1" Then
    If n = 1 Then
        werd(i) = "1 is for the length of his dick"
    ElseIf n = 2 Then
        werd(i) = "1 is for his current age "
    ElseIf n = 3 Then
        werd(i) = "1 is for amount of balls he got"
    End If
  ElseIf letter = "2" Then
    If n = 1 Then
        werd(i) = "2 is for the number of dicks his dad has"
    ElseIf n = 2 Then
        werd(i) = "2 is for his life savings"
    ElseIf n = 3 Then
        werd(i) = "2 is how many times he fucks his mom a day "
    End If
  ElseIf letter = "3" Then
    If n = 1 Then
        werd(i) = "3 is for his test score"
    ElseIf n = 2 Then
        werd(i) = "3 is for how many cents he charges at night"
    ElseIf n = 3 Then
        werd(i) = "3 is for how many times he asked Steve Case out"
    End If
  ElseIf letter = "4" Then
    If n = 1 Then
        werd(i) = "4 is for his life savings"
    ElseIf n = 2 Then
        werd(i) = "4 is for how many Macs he ordered"
    ElseIf n = 3 Then
        werd(i) = "4 is for how many virii he got"
    End If
  ElseIf letter = "5" Then
    If n = 1 Then
        werd(i) = "5 is for how many times he forgot to use a condom"
    ElseIf n = 2 Then
        werd(i) = "5 is for how many times his mom sucked his dick"
    ElseIf n = 3 Then
        werd(i) = "5 is for how many weeks he was jerking off"
    End If
  ElseIf letter = "6" Then
    If n = 1 Then
        werd(i) = "6 is for his computer's speed in Mhz"
    ElseIf n = 2 Then
        werd(i) = "6 is for his mom's age was when she got pregnant"
    ElseIf n = 3 Then
        werd(i) = "6 is for how many times he got TOSed today"
    End If
  ElseIf letter = "7" Then
    If n = 1 Then
        werd(i) = "7 is for how many programs he decompiled"
    ElseIf n = 2 Then
        werd(i) = "7 is for how many people he ripped off"
    ElseIf n = 3 Then
        werd(i) = "7 is for number of parents he has"
    End If
  ElseIf letter = "8" Then
    If n = 1 Then
        werd(i) = "8 is for how many cockroaches he ate"
    ElseIf n = 2 Then
        werd(i) = "8 is for his current age"
    ElseIf n = 3 Then
        werd(i) = "8 is for how much he thinks 5+5 is"
    End If
  ElseIf letter = "9" Then
    If n = 1 Then
        werd(i) = "9 is for how many times he ate shit"
    ElseIf n = 2 Then
        werd(i) = "9 is for how long he was searching for his dick"
    ElseIf n = 3 Then
        werd(i) = "9 is for how many blowjobs he gives a day"
    End If
  ElseIf letter = "0" Then
    If n = 1 Then
        werd(i) = "0 is for how many eyes his dad has"
    ElseIf n = 2 Then
        werd(i) = "0 is for how much money he has"
    ElseIf n = 3 Then
        werd(i) = "0 is for how many dicks he got"
    End If
  End If
    i = i + 1
    c = c + 1
    'If l0777 = l0778 Then Exit Do
Loop Until c = Len(sn)
t1 = werd(1)
t2 = werd(2)
t3 = werd(3)
t4 = werd(4)
t5 = werd(5)
t6 = werd(6)
t7 = werd(7)
t8 = werd(8)
t9 = werd(9)
t10 = werd(10)


'here



End Sub

Sub bustinpriv ()
electric1:
If proG_STAT$ = "OFF" Then
Exit Sub
End If

Let num_try = num_try + 1
a = FindWindow("AOL Frame25", 0&)
RunMenu 2, 5

Do
DoEvents
fcr = FindChildByTitle(a, "Keyword")
DoEvents
Loop Until fcr <> 0

E1 = FindChildByClass(fcr, "_AOL_Icon")
i1 = FindChildByClass(fcr, "_AOL_Edit")
send_to = SendMessageByString(i1, WM_SETTEXT, 0, "aol://2719:2-2-api")
tta = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
ttr = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
waitforok
If OK_note = "NO" Then
Exit Sub
End If
Pause .2
GoTo electric1

End Sub

Function center (Wha)
cr = 60
c2 = Len(Wha)
kk = cr - c2
ko = Int(kk / 2)
dd = ko
center = " " & Space(dd - 1) & Wha

End Function

Sub changewav (wav As String)
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub

Sub Chat (Text As String)
If Text = "" Then Exit Sub
aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient")
List% = FindChildByClass(aol%, "_AOL_Listbox")
room% = GetParent(List%)
If room% = 0 Then Exit Sub
view% = FindChildByClass(room%, "_AOL_View")
talk% = FindChildByClass(room%, "_AOL_Edit")
send% = getwindow(talk%, GW_HWNDNEXT)
X = SendMessageByString(talk%, WM_SETTEXT, 0, Text)
Pause .5
X = sendmessage(send%, WM_CHAR, 13, 0)
End Sub

Sub ChatTexter (wordz As String)
aol% = FindWindow("AOL Frame25", 0&)
List% = FindChildByClass(aol%, "_AOL_Listbox")
room% = GetParent(List%)
talk% = FindChildByClass(room%, "_AOL_Edit")
X = SendMessageByString(talk%, WM_SETTEXT, 0, wordz)
X = SendMessageByNum(talk%, WM_CHAR, 13, 0)

End Sub

Function check (number$) As Integer
check = 1
For i% = 1 To Len(number$)
    If Asc(Mid$(number$, i%, 1)) < 48 Or Asc(Mid$(number$, i%, 1)) > 58 Then check = 0
Next i%
End Function

Sub Click (btn)
    X = SetFocusAPI(btn)
    X = SetActiveWindow(btn)
    SD% = sendmessage(btn, WM_KEYDOWN, VK_SPACE, 0&)
    SU% = sendmessage(btn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub ClickAOLMenu (Menu_String As String, Top_Position As String)
Dim Top_Position_Num As Integer
Dim buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim aol As Integer
Dim Menu_Handle As Integer


Top_Position_Num = -1
aol% = FindWindow("AOL Frame25", 0&)
Menu_Handle = GetMenu(aol%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, buffer$, Len(Top_Position) + 1, &H400)
    Trim_Buffer = TrimNull(buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, buffer$, Len(Menu_String) + 1, &H400)
    Trim_Buffer = TrimNull(buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
Loop
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = SendMessageByNum(aol, &H111, Get_ID%, 0&)

End Sub

Sub ClickButton (hWnd As Integer)
Dim r
r = SendMessageByNum(GetParent(hWnd), &H111, (GetWindowWord(hWnd, (-12))), ByVal CLng(hWnd))
    DoEvents

End Sub

Sub ClickKW25 ()
aol% = FindWindow("AOL Frame25", 0&)
Tool% = FindChildByClass(aol%, "Aol Toolbar")
Icon1% = FindChildByClass(Tool%, "_Aol_Icon")
Icon2% = getwindow(Icon1%, 2)
Icon3% = getwindow(Icon2%, 2)
Icon4% = getwindow(Icon3%, 2)
Icon5% = getwindow(Icon4%, 2)
Icon6% = getwindow(Icon5%, 2)
Icon7% = getwindow(Icon6%, 2)
Icon8% = getwindow(Icon7%, 2)
Icon9% = getwindow(Icon8%, 2)
Icon10% = getwindow(Icon9%, 2)
Icon11% = getwindow(Icon10%, 2)
Icon12% = getwindow(Icon11%, 2)
Icon13% = getwindow(Icon12%, 2)
z% = sendmessage(Icon13%, WM_LBUTTONDOWN, 0&, 0&)
q% = sendmessage(Icon13%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub ClickKW30 ()
aol% = FindWindow("AOL Frame25", 0&)
Tool% = FindChildByClass(aol%, "Aol Toolbar")
Icon1% = FindChildByClass(Tool%, "_Aol_Icon")
Icon2% = getwindow(Icon1%, 2)
Icon3% = getwindow(Icon2%, 2)
Icon4% = getwindow(Icon3%, 2)
Icon5% = getwindow(Icon4%, 2)
Icon6% = getwindow(Icon5%, 2)
Icon7% = getwindow(Icon6%, 2)
Icon8% = getwindow(Icon7%, 2)
Icon9% = getwindow(Icon8%, 2)
Icon10% = getwindow(Icon9%, 2)
Icon11% = getwindow(Icon10%, 2)
Icon12% = getwindow(Icon11%, 2)
Icon13% = getwindow(Icon12%, 2)
Icon14% = getwindow(Icon13%, 2)
Icon15% = getwindow(Icon14%, 2)
Icon16% = getwindow(Icon15%, 2)
Icon17% = getwindow(Icon16%, 2)
Icon18% = getwindow(Icon17%, 2)
Icon19% = getwindow(Icon18%, 2)
z% = sendmessage(Icon18%, WM_LBUTTONDOWN, 0&, 0&)
q% = sendmessage(Icon18%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub ClickListBox (hWnd As Integer)
Dim r
r = SendMessageByNum(hWnd, &H203, 0, 0&)

End Sub

Sub ClickToolBar (Wha)
aol = GetAOL()
TB = FindChildByClass(aol, "AOL Toolbar")
Firz = FindChildByClass(TB, "_AOL_icon")

OiK = FindChildByClass(TB, "_AOL_icon")
For b = 1 To Wha - 1
If Wha = 1 Then Exit For
OiK = GetNextWindow(OiK, GW_HWNDNEXT)
Next b
Call Click(OiK)


End Sub

Function Countmailz ()
AO% = FindWindow("AOL Frame25", 0&)
arf = FindChildByTitle(AO%, "New Mail")
Hand% = FindChildByClass(arf, "_AOL_TREE")
Mails = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
End Function

Sub eMail (t$, CC$, subj$, Mssg$, Attch$, Snd%)
Dim r, X%, C1%, c2%, C3%, C4%, C5%, C6%, C7%, C8%, C9%, C10%, C11%, C12%, C13%, C14%, C15%, C16%, C17%
 RunMenu 3, 1
X% = WaitForWindow("Compose Mail", "AOL Child")
C1% = getwindow(X%, 5)  '|Send|
c2% = getwindow(C1%, 2) '<Send>
C3% = getwindow(c2%, 2) '|Send Later|
C4% = getwindow(C3%, 2) '<Send Later>
C5% = getwindow(C4%, 2) '|Attach|
C6% = getwindow(C5%, 2) '<Attach>
C7% = getwindow(C6%, 2) '|Address Book|
C8% = getwindow(C7%, 2) '<Address Book>
C9% = getwindow(C8%, 2) '|To:|
C10% = getwindow(C9%, 2) '<To:>
C11% = getwindow(C10%, 2) '|CC:|
C12% = getwindow(C11%, 2) '<CC:>
C13% = getwindow(C12%, 2) '|Subject:|
C14% = getwindow(C13%, 2) '<Subject:>
C15% = getwindow(C14%, 2) '|File:|
C16% = getwindow(C15%, 2) '|(filename)|
C17% = getwindow(C16%, 2) '<Message>
TextSet C10%, CStr(t$)
TextSet C12%, CStr(CC$)
TextSet C14%, CStr(subj$)
TextSet C17%, CStr(Mssg$)
If Attch$ <> "" Then
    ClickButton (C6%): DoEvents
    z% = WaitForWindow("Attach File", "*")
    TextSet FindChildByClass(z%, "Edit"), CStr(Attch$)
    ClickButton (FindChildByTitle(z%, "OK")): DoEvents
End If
If Snd% = 1 Then
    ClickButton (c2%)
End If

End Sub

Sub enableaolwins ()
Dim bb As Integer
Dim dis_win As Integer
CessPit = EnableWindow(aolhwnd(), 1)

fc = FindChildByClass(aolhwnd(), "AOL Child")
req = EnableWindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = GetNextWindow(faa, 2)
res = EnableWindow(faa, 1)
DoEvents
Loop Until faf = faa
End Sub

Function Encrypt2 ()
AO% = FindWindow("AOL Frame25", 0&)
Cmpse% = FindChildByTitle%(AO%, "Compose Mail")
Edt% = FindChildByClass%(Cmpse%, "_AOL_Edit")
a% = getwindow(Edt%, GW_HWNDNEXT): CcC% = getwindow(a%, GW_HWNDNEXT)
b% = getwindow(CcC%, GW_HWNDNEXT): Subject% = getwindow(b%, GW_HWNDNEXT)
Sub1% = getwindow(Subject%, GW_HWNDNEXT)
'Sub2% = GetWindowtext(SubGW_HWNDNEXT)
Whole$ = ""
For X1 = 1 To 60
 Randomize
   valteger = Int(17 * Rnd + 1)
    If valteger = 1 Then
    vv = "š"
    ElseIf valteger = 2 Then : Whole$ = Whole$ + "¶"
    ElseIf valteger = 3 Then : Whole$ = Whole$ + "¢"
    ElseIf valteger = 4 Then : Whole$ = Whole$ + "±"
    ElseIf valteger = 5 Then : Whole$ = Whole$ + "¥"
    ElseIf valteger = 6 Then : Whole$ = Whole$ + "†"
    ElseIf valteger = 7 Then : Whole$ = Whole$ + "ø"
    ElseIf valteger = 8 Then : Whole$ = Whole$ + "‡"
    ElseIf valteger = 9 Then : Whole$ = Whole$ + "µ"
    ElseIf valteger = 10 Then : Whole$ = Whole$ + "þ"
    ElseIf valteger = 11 Then : Whole$ = Whole$ + "‰"
    ElseIf valteger = 12 Then : Whole$ = Whole$ + "¤"
    ElseIf valteger = 13 Then : Whole$ = Whole$ + "£"
    ElseIf valteger = 14 Then : Whole$ = Whole$ + "÷"
    ElseIf valteger = 15 Then : Whole$ = Whole$ + "º"
    ElseIf valteger = 16 Then : Whole$ = Whole$ + "©"
    ElseIf valteger = 17 Then : Whole$ = Whole$ + "*"
    End If
Next
X% = SendMessageByString(Subject%, WM_SETTEXT, 0, Whole$)
End Function

Sub Explode (frm As Form, CFlag As Integer)
Const STEPS = 150 'Lower Number Draws Faster, Higher Number Slower
Dim FRect As RECT
Dim FWidth, FHeight As Integer
Dim i, X, Y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush

' If CFlag = True, then explode from center of form, otherwise
' explode from upper left corner.
   ' GetWindowRect frm.hWnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    FHeight = FRect.Bottom - FRect.Top
    
' Create brush with Form's background color.
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    
' Draw rectangles in larger sizes filling in the area to be occupied
' by the form.
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
    
' Release the device context to free memory.
' Make the Form visible

    If ReleaseDC(0, hScreen) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    DeleteObject (Brush)
    frm.Show

End Sub

Function findAOLchildbytitle (titletext As String) As Integer
Dim X%
Dim ChildWnd As Integer
Dim MDIhWnd%
Dim AOLChildhWnd%
Dim RetClsName As String * 255
  
MDIhWnd% = getwindow(FindWindow("AOL Frame25", 0&), GW_CHILD)
Do
  X% = GetClassName(MDIhWnd%, RetClsName$, 254)
  If InStr(RetClsName$, "MDIClient") Then AOLChildhWnd% = MDIhWnd%
  MDIhWnd% = getwindow(MDIhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While MDIhWnd% <> 0
If titletext = "MDIClient" Then findAOLchildbytitle = AOLChildhWnd%
ChildWnd = getwindow(AOLChildhWnd%, GW_CHILD)
Do
  If InStr(windowcaption(ChildWnd), titletext) <> 0 Then
      findAOLchildbytitle = ChildWnd
      Exit Do
  End If
  ChildWnd = getwindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
End Function

Function findchatroom ()
Dim bb As Integer
Dim dis_win As Integer

dis_win = FindChildByClass(aolhwnd(), "AOL Child")

begin_find_chat:

bb = FindChildByClass(dis_win, "_AOL_Listbox")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Send")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "List" & Chr(13) & "Rooms")
    If bb <> 0 Then Let countt = countt + 1
bb = FindChildByTitle(dis_win, "Center" & Chr(13) & "Stage")
    If bb <> 0 Then Let countt = countt + 1
bb = FindChildByTitle(dis_win, "Chat" & Chr(13) & "Preferences")
    If bb <> 0 Then Let countt = countt + 1
If countt = 5 Then
  findchatroom = dis_win
  Exit Function
End If
Let countt = 0
dis_win = GetNextWindow(dis_win, 2)
If dis_win = getwindow(dis_win, GW_HWNDLAST) Then
   Exit Function
End If

GoTo begin_find_chat
End Function

Function FindChatwnd () As Integer
  Dim MDIhWnd%
  Dim AOLChildhWnd%
  Dim ChildWnd As Integer
  Dim ControlWnd As Integer
  Dim ChatWnd As Integer
  Dim TargetsFound As Integer
  Dim RetClsName As String * 255
  Dim X%
MDIhWnd% = getwindow(FindWindow("AOL Frame25", 0&), GW_CHILD)
Do
  X% = GetClassName(MDIhWnd%, RetClsName$, 254)
    If InStr(RetClsName$, "MDIClient") Then
      AOLChildhWnd% = MDIhWnd% 'Child window found!
    End If
  MDIhWnd% = getwindow(MDIhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While MDIhWnd% <> 0
ChildWnd = getwindow(AOLChildhWnd%, GW_CHILD)
Do
  ControlWnd = getwindow(ChildWnd, GW_CHILD)
  Do
    X% = GetClassName(ControlWnd, RetClsName$, 254)

    
    If InStr(RetClsName$, "_AOL_Edit") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_View") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_Listbox") Then
      TargetsFound = TargetsFound + 1:
    End If
    ControlWnd = getwindow(ControlWnd, GW_HWNDNEXT)
    DoEvents
  Loop While ControlWnd <> 0

  If TargetsFound = 3 Then ChatWnd = ChildWnd: Exit Do

  
  ChildWnd = getwindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0

FindChatwnd = ChatWnd

End Function

Function FindCht ()
aol = FindWindow("AOL Frame25", "America  Online")
lstrms = FindChildByTitle(aol, "List Rooms")
If lstrms = 0 Then FindCht = 0: Exit Function
FindCht = GetParent(lstrms)
End Function

Function findcomposemail ()
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
If dis_win = getwindow(dis_win, GW_HWNDLAST) Then
   findtocomposemail = 0
   Exit Function
End If
GoTo begin_find_composemail
End Function

Function findfwdwnd () As Integer
Dim hWnds() As Integer
Dim r, X%, i, t, g%
DoEvents
X = 0
g% = getwindow(AOhWnd(), 5)
DoEvents
Do
If FindChildByTitle(FWD%, "Send") <> 0 Then X = X + 1
DoEvents
If FindChildByTitle(g%, "Send Now") <> 0 Then X = X + 1
DoEvents
If FindChildByTitle(g%, "Address" & Chr(13) & "Book") <> 0 Then X = X + 1
If X = 3 Then GoTo Founddddd
DoEvents
g% = getwindow(g%, 2)
DoEvents
X = 0
Loop While g% <> 0


Exit Function
Founddddd:
DoEvents
findfwdwnd = g%
Exit Function

End Function

'
Function FindMailWnd () As Integer
Dim hWnds() As Integer
Dim r, X%, i, t, g%
X = 0
DoEvents
g% = getwindow(AOhWnd(), 5)
DoEvents
Do
If FindChildByTitle(g%, "Read") <> 0 Then X = X + 1
DoEvents
If FindChildByTitle(g%, "Ignore") <> 0 Then X = X + 1
DoEvents
If FindChildByTitle(g%, "Keep As New") <> 0 Then X = X + 1
DoEvents
If FindChildByTitle(g%, "Delete") <> 0 Then X = X + 1
DoEvents
If FindChildByClass(g%, "_AOL_Tree") <> 0 Then X = X + 1
If X = 5 Then GoTo Foundd
DoEvents
g% = getwindow(g%, 2)
DoEvents
X = 0
Loop While g% <> 0


Exit Function
Foundd:
DoEvents
FindMailWnd = g%
Exit Function

End Function

Function findoldim ()
Dim bb As Integer
Dim dis_win As Integer

dis_win = FindChildByClass(aolhwnd(), "AOL Child")

begin_find_oldim:

bb = FindChildByTitle(dis_win, "Respond")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Cancel")
    If bb <> 0 Then Let countt = countt + 1
If countt = 2 Then
  findoldim = dis_win
  Exit Function
End If
Let countt = 0
dis_win = GetNextWindow(dis_win, 2)
If dis_win = getwindow(dis_win, GW_HWNDLAST) Then
   findoldim = 0
   Exit Function
End If

GoTo begin_find_oldim

End Function

'
Function FindReadWnd () As Integer
Dim hWnds() As Integer
Dim r, X%, i, t, g%
X = 0
DoEvents
g% = getwindow(AOhWnd(), 5)
DoEvents
Do
If FindChildByTitle(g%, "Reply") <> 0 Then X = X + 1
DoEvents
If FindChildByTitle(g%, "Forward") <> 0 Then X = X + 1
DoEvents
If FindChildByTitle(g%, "Reply to All") <> 0 Then X = X + 1
DoEvents
If X = 3 Then GoTo Founddd
DoEvents
g% = getwindow(g%, 2)
DoEvents
X = 0
Loop While g% <> 0


Exit Function
Founddd:
DoEvents
FindReadWnd = g%
Exit Function


End Function

Function Findsn ()
'Finds the user's Screen Name...they must be signed on!

Dim dis_win2 As Integer
a = FindWindow("AOL Frame25", 0&)
dis_win2 = FindChildByClass(a, "AOL Child")

begin_find_SN:

bb$ = windowcaption(dis_win2)
    If Left(bb$, 9) = "Welcome, " Then Let countt = countt + 1
If countt = 1 Then
  val1 = InStr(bb$, " ")
  val2 = InStr(bb$, "!")
  Let sn$ = Mid$(bb$, val1 + 1, val2 - val1 - 1)
  Findsn = Trim(sn$) '_win
  Exit Function
End If
Let countt = 0
dis_win2 = GetNextWindow(dis_win2, 2)
If dis_win2 = getwindow(dis_win2, GW_HWNDLAST) Then
   Findsn = 0
   Exit Function
End If

GoTo begin_find_SN

End Function

Function findtoim ()
Dim bb As Integer
Dim dis_win As Integer

dis_win = FindChildByClass(aolhwnd(), "AOL Child")

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
If dis_win = getwindow(dis_win, GW_HWNDLAST) Then
   findtoim = 0
   Exit Function
End If

GoTo begin_find_To_im
End Function

Function FindWindowLike (hWndArray() As Integer, ByVal hWndStart As Integer, WindowText As String, Classname As String, ID) As Integer

End Function

Function GetAll (sib As Integer, class$, t$)
On Error Resume Next
buf$ = String$(255, 0)
First = getwindow(sib, 0)
X = GetClassName(First, buf$, 255)
If class$ = TrimNull(buf$) Then GetAll = First: Exit Function

PrevhWnd = First
Do
buf$ = String$(255, 0)
ThishWnd = getwindow(PrevhWnd, 2)
X = GetClassName(ThishWnd, buf$, 255)
Debug.Print buf$
'------this is wierd
On Error Resume Next
AppActivate "America  Online"
t$ = dsafsdf
X = SetParent(listWnd, dsafsdf)
'--end wierd shit
If class$ = TrimNull(buf$) Then
    GetAll = ThishWnd: Exit Function
End If
PrevhWnd = ThishWnd
Loop While ThishWnd <> 0

GetAll = 0
End Function

Function GetAOL ()
GetAOL = FindWindow("AOL Frame25", "America  Online")
End Function

Function getcap ()
Dim babe As Integer
aolhwnd1 = FindWindow("aol frame25", "america  online")
X = gettopwindow(aolhwnd1)
X = gettopwindow(X) 'active aol window
babe = getwindowtextlength(X)
Dim buffer As String * 100
z = getwindowtext(X, buffer, babe)
'buffer is window text

End Function

Function GetChatWindow ()
aol = FindWindow("AOL Frame25", "America  Online")
CHT = getwindow(aol, GW_CHILD)
listrooms = FindChildByTitle(aol, "List" & Chr(13) & "Rooms")
centerstage = FindChildByTitle(aol, "Center" & Chr(13) & "Stage")
pcstudio = FindChildByTitle(aol, "PC" & Chr(13) & "Studio")
chatpref = FindChildByTitle(aol, "Chat" & Chr(13) & "Preferences")
parcont = FindChildByTitle(aol, "Parental" & Chr(13) & "Control")
pplinrm = FindChildByClass(aol, "_AOL_LISTBOX")
If listrooms <> 0 And GetParent(listrooms) = GetParent(centerstage) And GetParent(centerstage) = GetParent(pcstudio) And GetParent(pcstudio) = GetParent(chatpref) And GetParent(chatpref) = GetParent(parcont) And GetParent(parcont) = GetParent(pplinrm) Then GetChatWindow = GetParent(listrooms):       Exit Function
GetChatWindow = 0
End Function

Function GetControl () As String
ActivehWnd = getfocus()
Dim buffer As String
buffer = String$(255, 0)
X = getwindowtext(ActivehWnd, buffer, 255)
GetControl = TrimNull(buffer)

End Function

Function GetCPUType () As String
'Example: text9.text = "Your system's CPU type is: " & sGetCPUType
Dim lWinFlags As Long

    lWinFlags = GetWinFlags()

    If lWinFlags And WF_CPU486 Then
        GetCPUType = "486"
        ElseIf lWinFlags And WF_CPU386 Then
            GetCPUType = "386"
            ElseIf lWinFlags And WF_CPU286 Then
                GetCPUType = "286"
                Else
                    GetCPUType = "Other"
    End If

End Function

Sub getcurs ()
aol% = FindWindow("AOL Frame25", 0&)
X = ExitWindows(aol%, 0)
DoEvents
DoEvents
DoEvents
SendKeys "n"

End Sub

Function GetFreeGDI () As String
'Example: text5.text = "Free GDI Resources: " & sGetFreeGDI
    GetFreeGDI = Format$(GetFreeSystemResources(GFSR_GDIRESOURCES)) + "%"

End Function

Function GetFreeSys () As String
'Example: text3.text = "Free System Resources: " & sGetFreeSys
    GetFreeSys = Format$(GetFreeSystemResources(GFSR_SYSTEMRESOURCES)) + "%"

End Function

Function GetFreeUser () As String
'Example: text4.text = "Free User Resources: " & sGetFreeUser
    GetFreeUser = Format$(GetFreeSystemResources(GFSR_USERRESOURCES)) + "%"

End Function

Sub getlisttext (lst As ListBox)
On Error Resume Next
Do Until xx = lst.ListCount
Let da_new$ = lst.List(xx)
If Trim(da_new$) <> "" Then Let namms = namms & da_new$ & "," Else Exit Sub
Let xx = xx + 1
Loop

End Sub

Function GetUser ()

''Find the welcome window
'Welcome = FindChildbyTitle(FindWindow("AOL Frame25", "America  Online"), "Welcome, ")

'Get the caption of it
'TheCaption = WindowCaption(Welcome)

'Extract the user's screen name

'On Error Resume Next
 'GetUser = Mid$(TheCaption, 10, (InStr(TheCaption, "!") - 10))
End Function

Function GetWinDir () As String
buffer$ = String$(255, 0)
X = GetWindowDirectory(buffer$, 255)
Trm$ = TrimNull(buffer$)
If Right$(Trm$, 1) <> "\" Then Trm$ = Trm$ + "\"
GetWinDir = Trm$

End Function

Function GetWindowAct () As String
Focus = getfocus()
ActivehWnd = GetParent(Focus)
Dim buffer As String
buffer = String$(255, 0)
i = getwindowtext(ActivehWnd, buffer, 255)
d = buffer
GetWindowAct = TrimNull(buffer)

End Function

Function GetWindowFromClass (PARENT As Integer, class$) As Integer
lst = getwindow(PARENT, 5)
This = getwindow(PARENT, 0)
In$ = String$(255, 0)
X = GetClassName(This, In$, 255)
Out$ = TrimNull(In$)
If class$ = Out$ Then GoTo Found
Do
This = getwindow(PARENT, 2)
In$ = String$(255, 0)
X = GetClassName(This, In$, 255)
Out$ = TrimNull(In$)
If class$ = Out$ Then GoTo Found
Loop Until This = lst
GetWindowFromClass = 0
Exit Function
Found:
GetWindowFromClass = This

End Function

Function GetWindowhWnd () As Integer
Focus = getfocus()
GetWindowhWnd = GetParent(Focus)

End Function

Function GetWinVer () As String
'Example: text2.text = "Window version: " & sGetWinVer
Dim lVer As Long, iWinVer As Integer
    lVer = GetVersion()
    iWinVer = CInt(lVer And &HFFFF&)
    GetWinVer = Format$(iWinVer And &HFF) + "." + Format$(CInt(iWinVer / 256))

End Function

Function GetYours () As String
ActivehWnd = getfocus()
Dim buffer As String
buffer = String$(255, 0)
i = getwindowtext(ActivehWnd, buffer, 255)
d = buffer
GetYours = TrimNull(buffer)

End Function

Function IFileExists (ByVal sFileName As String) As Integer
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

Sub IM_Off ()
Dim nambox As Integer
Dim IMwnd As Integer
Dim txtbx As Integer
Dim btnsend As Integer
aol = FindWindow(0&, "America  Online")
sfdlk = SetFocusAPI(aol)
X% = RunMenuByString("Send an Instant Message", "Mem&bers")
Do
For i = 1 To 25
DoEvents
Next i
IMwnd = FindChildByTitle(aol, "Send Instant Message")
Loop Until IMwnd
DoEvents
nambox = FindChildByClass(IMwnd, "_AOL_Edit")
'Call HR_Settext(nambox, "$IM_OFF")
mailopenssss = SendMessageByString(nambox, WM_SETTEXT, 0, "$IM_OFF")
txtbx = GetNextWindow(nambox, GW_HWNDNEXT)
mailopenq% = SendMessageByString(txtbx, WM_SETTEXT, 0, "WaR Rules")
'Call HR_Settext(txtbx, "v^*^v TWiSTeD SoULz v^*^v")
btnsend = FindChildByTitle(IMwnd, "Send")
Call Click(btnsend)
IMssss% = GetParent(btnsend)
  cl_ose = SendMessageByNum(IMssss%, WM_CLOSE, 0, 0&)

Do
    DoEvents
    msger% = GetActiveWindow()
    OKbuton% = FindChildByTitle(msger%, "OK")
Loop Until OKbuton%
Call Click(OKbuton%)
End Sub

Sub IM_On ()
Dim nambox As Integer
Dim IMwnd As Integer
Dim txtbx As Integer
Dim btnsend As Integer
aol = FindWindow(0&, "America  Online")
sfdlk = SetFocusAPI(aol)
X% = RunMenuByString("Send an Instant Message", "Mem&bers")
Do
For i = 1 To 25
DoEvents
Next i
IMwnd = FindChildByTitle(aol, "Send Instant Message")
Loop Until IMwnd
DoEvents
nambox = FindChildByClass(IMwnd, "_AOL_Edit")
'Call HR_Settext(nambox, "$IM_OFF")
mailopenssss = SendMessageByString(nambox, WM_SETTEXT, 0, "$IM_ON")
txtbx = GetNextWindow(nambox, GW_HWNDNEXT)
mailopenq% = SendMessageByString(txtbx, WM_SETTEXT, 0, "WaR Rules")
'Call HR_Settext(txtbx, "v^*^v TWiSTeD SoULz v^*^v")
btnsend = FindChildByTitle(IMwnd, "Send")
IMssss% = GetParent(btnsend)
  cl_ose = SendMessageByNum(IMssss%, WM_CLOSE, 0, 0&)
Call Click(btnsend)
Do
    DoEvents
    msger% = GetActiveWindow()
    OKbuton% = FindChildByTitle(msger%, "OK")
Loop Until OKbuton%
Call Click(OKbuton%)
End Sub

 Sub IMoff ()

X% = RunMenuByString("Send an Instant Message", "Mem&bers")
aol% = FindWindow("AOL Frame25", 0&)
Do
DoEvents
IMwnd% = FindChildByTitle(aol%, "Available?")
Loop Until IMwnd% <> 0

Edt% = FindChildByClass(IMwnd%, "_AOL_Edit")
X% = SendMessageByString(Edt%, WM_SETTEXT, 0, "$IM_OFF")
Edt2% = GetNextWindow(Edt%, 2)
X% = SendMessageByString(Edt2%, WM_SETTEXT, 0, "Syndicate Rulez!")
Butt% = FindChildByClass(IMwnd%, "_AOL_Button")
AOLClick (Butt%)
Do
DoEvents
Act% = GetActiveWindow()
msg% = FindChildByTitle(Act%, "OK")
Loop Until msg% <> 0
AOLClick (msg%)
CLR = sendmessage(IMwnd%, WM_CLOSE, 0, 0)
End Sub

Sub IMOn ()

X% = RunMenuByString("Send an Instant Message", "Mem&bers")
aol% = FindWindow("AOL Frame25", 0&)
Do
DoEvents
IMwnd% = FindChildByTitle(aol%, "Available?")
Loop Until IMwnd% <> 0

Edt% = FindChildByClass(IMwnd%, "_AOL_Edit")
X% = SendMessageByString(Edt%, WM_SETTEXT, 0, "$IM_On")
Edt2% = GetNextWindow(Edt%, 2)
X% = SendMessageByString(Edt2%, WM_SETTEXT, 0, "Syndicate Rulez!")
Butt% = FindChildByClass(IMwnd%, "_AOL_Button")
AOLClick (Butt%)
Do
DoEvents
Act% = GetActiveWindow()
msg% = FindChildByTitle(Act%, "OK")
Loop Until msg% <> 0
AOLClick (msg%)
CLR = sendmessage(IMwnd%, WM_CLOSE, 0, 0)
End Sub

Function IMtxt ()
   Dim NumChars As Integer
    Dim CRText As String '

    NumChars = SendMessageByNum(ViewWnd, &HE, 0, 0&)
    CRText = Space$(NumChars) 'load up CRText to NumChars
    '
    'Fill up CRText w/ IM text
    Dim X As Integer
    X = SendMessageByString(ViewWnd, &HD, NumChars, CRText)
    
    If Value = 1 Then
        txtim = Left$(CRText, X)
    Else
      txtim = Right$(Left$(CRText, X), 128)
    End If


End Function

Sub iobox (info$, titletext$)
errortitle$ = titletext$
errorstr$ = info$
'errorform.Show 1
End Sub

Function ison () As Integer
aol = FindWindow("AOL Frame25", "America  Online")
wlcm = FindChildByTitle(aol, "Welcome, ")
If wlcm <> 0 Then ison = True
If wlcm = 0 Then
    ison = False
    MsgBox "You must be signed on to AOL to do this.", 48, "You are not signed on!": Exit Function
End If
End Function

Function IsUsingMaster () As Integer
Dim TopMenuPos As Integer
Dim X As Integer, mnuWnd As Integer, subMnu As Integer, MenuID

mnuWnd = GetMenu(GetAOL())
subMnu = GetSubMenu(mnuWnd, TopMenuPos)
MenuID = GetMenuItemID(subMnu, 0)
Dim MenuString As String * 255

Const MF_BYCOMMAND = &H0
X = GetMenuString(mnuWnd, MenuID, MenuString, 254, MF_BYCOMMAND)
    
    If InStr(MenuString, "&New") <> 0 Then
        IsUsingMaster = False
    Else
'Master.AOL menu
        IsUsingMaster = True
    End If
End Function

Sub KillWait ()
a = FindWindow("AOL Frame25", 0&)
z% = RunMenuByString("&About America Online", "&Help")
Timeout (.5)
X = SendMessageByNum(FindWindow("_AOL_Modal", 0&), WM_CLOSE, 0, 0&)

End Sub

Function KTEncrypt (ByVal password, ByVal Strng, force%)
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
  If Len(Strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(Strng, 4) + Right$(Strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      Strng = Mid$(Strng, 5, Len(Strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, Strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(Strng, look + 1)) - 1)
          Strng = Left$(Strng, look - 1) + Addin$ + Mid$(Strng, look + 2)
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
    Strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + Strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(Strng)

    'Alter character code
    ToChange = Asc(Mid$(Strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(Strng, Looper, 1) = Chr$(ToChange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, Strng, Chr$(1))
      If look > 0 Then
        Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(Strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(Strng, Chr$(0))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(Strng, Chr$(10))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(Strng, Chr$(13))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(Strng, Chr$(26))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    Strng = Chr$(1) + "KT" + Chr$(1) + Strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(Strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      Strng = Mid$(Strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = Strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function

Sub letterexplode (lbl As Label, maxsize As Integer, spd As Integer)
lbl.Visible = True
Do
For i = 1 To spd
DoEvents
Next i
lbl.FontSize = lbl.FontSize + 5
'lbl.Top = lbl.Top + 10
'lbl.Left = lbl.Left + 10
Loop Until lbl.FontSize >= maxsize
'MsgBox lbl.FontSize
End Sub

Sub letterimplode (lbl As Label, minsize As Integer, spd As Integer)
Do
For i = 1 To 3
DoEvents
Next i
lbl.FontSize = lbl.FontSize - 1
lbl.ForeColor = &HFFFFFF
Timeout .000001
'lbl.Top = lbl.Top + 10
lbl.ForeColor = &HFF0000
Timeout .0000001

'lbl.Left = lbl.Left + 10
Loop Until lbl.FontSize <= minsize
'lbl.Visible = False
'MsgBox lbl.FontSize
End Sub

Sub lstfromstr (lst As ListBox, Strng As String)
Cma = InStr(Strng, ",")
If Cma Then
    dasn = Mid$(Strng, 1, Cma - 1)
    lst.AddItem dasn
Else
    lst.AddItem Strng
    Exit Sub
End If
Do
cma2 = InStr(Cma + 1, Strng, ",")
If cma2 Then
    stpplc = Len(Strng) - (Cma + 1) - (Len(Strng) - cma2)
    dasn = Mid$(Strng, Cma + 1, stpplc)
    lst.AddItem dasn
ElseIf cma2 = False Then
    asn = Mid$(Strng, Cma + 1)
    If asn = "" Then Exit Sub
    lst.AddItem asn
End If
Cma = cma2
Loop Until Cma = False
End Sub

Function MailCount ()
 AO% = FindWindow(0&, "America  Online")
 Hand% = FindChildByClass(AO%, "_AOL_TREE")
 buffer = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
 If buffer > 1 Then
   MsgBox "You have " & buffer & " messages in your Mailbox...", 0
 End If
 If buffer = 1 Then
   MsgBox "You have " & buffer & " message in your Mailbox...", 0
 End If
 If buffer < 1 Then
   MsgBox "You have no messages in your Mailbox...", 0
 End If
 MailCount = buffer
End Function

Function mailopen ()
If FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail") = False Then
Call RunMenu(3, 2)
Do
DoEvents
Box = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail")
Loop Until Box
List = FindChildByClass(Box, "_AOL_Tree")
Do
DoEvents
mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)

Call Timeout(.5)
mailnum2 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(.5)
mailnum3 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
Else
    Box = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "NewMail ")
    List = FindChildByClass(Box, "_AOL_Tree")
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
End If
mailopen = mailnum
End Function

Sub Mailtotxt ()
nof = InputBox("What do you want to name the mail.txt file?", "Name of File?")
MsgBox "Open the Box you want a txt of, and press OK"
aol = FindWindow("AOL Frame25", 0&)
X = SetFocusAPI(aol)
Box = GetWindowhWnd()
'MCAP = WindowCaption(Box)
tree% = FindChildByClass(Box, "_AOL_Tree")
 Open nof For Output As #1
Print #1, "                 Mortal Terror By: Type O"
Print #1, "                 Mail.txt of your mailbox" & Chr(13) & Chr(10)

For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1
Grr$ = String$(255, 0)
lgh = sendmessage(tree%, lb_gettextlength, i, Grr$)
Debug.Print lgh
gest = SendMessageByString(tree%, LB_GETTEXT, i, Grr$)
 TempText = TempText & Chr$(13) & Chr(10) & i + 1 & ":  " & Grr$
            Next i

Print #1, TempText

TempText = ""

  Close #1

End Sub

Function mover ()


End Function

Function MsgBoxText () As String
Dim TophWnd%
Dim X%
Dim BabyhWnd%
Dim RetClsName As String * 255
Dim MsgText As String, MsgLen As Integer
TophWnd% = GetActiveWindow()
BabyhWnd% = getwindow(TophWnd%, GW_CHILD)
X% = GetClassName(TophWnd%, RetClsName$, 254)
Do
  X% = GetClassName(BabyhWnd%, RetClsName$, 254)
  If InStr(UCase$(RetClsName$), "STATIC") Then
      BabyhWnd% = getwindow(BabyhWnd%, GW_HWNDNEXT)
      MsgLen = getwindowtextlength(BabyhWnd%)
      MsgText = String$(MsgLen, " ")
      X = getwindowtext(BabyhWnd%, MsgText, MsgLen)
      MsgBoxText = Trim$(MsgText)
      Exit Do
  End If
  BabyhWnd% = getwindow(BabyhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While BabyhWnd% <> 0
End Function

Sub notload ()
aol% = FindWindow("AOL Frame25", 0&)
If aol% = 0 Then
MsgBox "You are not signed on to AOL!"
Else
End If

End Sub

Function opnmail ()
daloopplc:
mbx = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail")
If mbx Then
    opnmail = True
    Exit Function
End If
retval = MsgBox("Your mail box was not open.  Open it then press Retry, or to cancel press Cancel.", 5, "Open your New Mail box")
    If retval = 2 Then
        opnmail = False
        Exit Function
    End If
GoTo daloopplc
End Function

Sub Pause (duratn As Integer)

Let curent = Timer

Do Until Timer - curent >= duratn
DoEvents
Loop

End Sub

Sub PlaySound (Xsound As String)
Debug.Print "Xsound  " & Xsound
Dim X%
X% = SndPlaySound(Xsound, 1)

End Sub

Sub PlayWav (da_file As String)
Waver = SndPlaySound(ByVal CStr(da_file), 1)

End Sub

Function putdash$ (txt$)
For i% = 1 To Len(txt$)
    temp$ = temp$ + Mid$(txt$, i%, 1)
    If i% Mod 4 = 0 And i% <> 16 Then temp$ = temp$ + "-"
Next i%
putdash$ = temp$
End Function

Function r_backwards (strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
r_backwards = newsent$

End Function

Function r_elite (strin As String)
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
If nextchr$ = "H" Then Let nextchr$ = "|-|"
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
If nextchr$ = " " Then Let nextchr$ = "."
Let newsent$ = newsent$ + nextchr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
r_elite = newsent$

End Function

Function r_hacker (strin As String)
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
If nextchr$ = " " Then Let nextchr$ = "."
Let newsent$ = newsent$ + nextchr$
Loop
r_hacker = newsent$

End Function

Function r_same (strr As String)
Let r_same = Trim(strr)

End Function

Function r_spaced (strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + "."
Let newsent$ = newsent$ + nextchr$
Loop
r_spaced = newsent$

End Function

Function ReadINI (AppName, KeyName, filename As String) As String
'Example: text4.text = ReadINI("DaProggy", "Lamers Name", app.path + "\Prog.ini")
Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))

End Function

Sub Remove_Menu ()
aol% = FindWindow("AOL Frame25", 0&)
ab = GetMenu(aol%)
d = removemenu(ab, 10, MF_BYPOSITION)
d = removemenu(ab, 7, MF_BYPOSITION)
d = removemenu(ab, 8, MF_BYPOSITION)
d = removemenu(ab, 9, MF_BYPOSITION)
d = removemenu(ab, 11, MF_BYPOSITION)

End Sub

Sub renamehost (sn As String)

Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , sn

Close #1
End Sub

Sub resetNew (sn As String, pth As String)
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

Sub RunMenu (horz, vert)
Dim f, gi, sm, m, a As Integer
a = FindWindow("AOL Frame25", 0&)
m = GetMenu(a)
sm = GetSubMenu(m, horz)
gi = GetMenuItemID(sm, vert)
f = SendMessageByNum(a, WM_COMMAND, gi, 0)

End Sub

Function RunMenuByString (Menu_String As String, Top_Position As String)
Top_Position_Num = -1
aol% = FindWindow("AOL Frame25", 0&)
Menu_Handle = GetMenu(aol%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, buffer$, Len(Top_Position) + 1, MF_BYPOSITION)
    Trim_Buffer = trimnull2(buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, buffer$, Len(Menu_String) + 1, MF_BYPOSITION)
    Trim_Buffer = trimnull2(buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
Loop
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = SendMessageByNum(aol%, WM_COMMAND, Get_ID%, 0&)
End Function

Function SendAction (ByVal ChatWnd As Integer, ByVal SendString As String) As Integer

End Function

Sub SendAIM (who As String, what As String)

z% = RunMenuByString("Send an Instant Message", "Mem&bers")
 Timeout (.5)
  aol% = FindWindow("AOL Frame25", 0&)
   Find% = FindChildByTitle(aol%, "Send Instant Message")
    sn% = FindChildByClass(Find%, "_AOL_Edit")
     X = SendMessageByString(sn%, WM_SETTEXT, 0, who)
    mesg% = FindChildByClass(Find%, "RICHCNTL")
   X = SendMessageByString(mesg%, WM_SETTEXT, 0, what)
  send% = getwindow(mesg%, GW_HWNDNEXT)
 X = SendMessageByNum(send%, WM_LBUTTONDOWN, 0, 0&) 'Click send
X = SendMessageByNum(send%, WM_LBUTTONUP, 0, 0&)
    
End Sub

Sub SendChat (chat_str As String)
If proG_STAT$ = "OFF" Then
Exit Sub
End If

DoEvents
Ya = FindChildByClass(findchatroom(), "_AOL_EDIT")
sb = FindChildByTitle(findchatroom(), "Send")
xa = SendMessageByString(Ya, WM_SETTEXT, 0, chat_str)
ta = SendMessageByNum(sb, WM_LBUTTONDOWN, 0, 0&)
Pause .01
wa = SendMessageByNum(sb, WM_LBUTTONUP, 0, 0&)
DoEvents
End Sub

Function sendchatroomtext (ByVal SendString As String) As Integer
'This is easy to do with the FindChatWnd function made
Dim ChatWnd As Integer 'holds chat window handle
Dim ControlWnd As Integer 'hold the current control hWnd
Dim X% 'Used because sendMessage is a function, not a sub
Dim RetClsName As String * 255 'holds the class name of ControlWnd
'the three needed button hWnds
Dim EdithWnd As Integer, ViewhWnd As Integer, ButtonhWnd As Integer

'Just get the ControlhWnds from the ChathWnd
ChatWnd = FindChatwnd()
If ChatWnd <> 0 Then 'Chat window located
  ControlWnd = getwindow(ChatWnd, GW_CHILD)
  Do
    X% = GetClassName(ControlWnd, RetClsName$, 254)
    'Test the control handles
    If InStr(RetClsName$, "_AOL_Edit") Then
        EdithWnd = ControlWnd
    ElseIf InStr(RetClsName$, "_AOL_View") Then
        ViewhWnd = ControlWnd
    ElseIf InStr(RetClsName$, "_AOL_Button") Then
        ButtonhWnd = ControlWnd
    End If
    'Find the next control handle
    ControlWnd = getwindow(ControlWnd, GW_HWNDNEXT)
    DoEvents
  Loop While ControlWnd <> 0
  'one final error check, and then the text will be sent
  If EdithWnd And ViewhWnd And ButtonhWnd Then 'make sure controls found
    'E$ is the Text you want to send to the chat room
    X = SendMessageByString(EdithWnd%, WM_SETTEXT, 0, SendString)
    'This next line send the RETURN
    X = SendMessageByNum(EdithWnd%, WM_CHAR, 13, 0)
    'Tell the procedure from which this function was called that text was sent
    sendchatroomtext = True
  End If
End If

End Function

Sub sendchattxt (CHT, txt2say)
txtbx = FindChildByClass(CHT, "_AOL_Edit")
'Call HR_Settext(txtbx, txt2say)
X = SendMessageByNum(txtbx, WM_CHAR, 13, 0)
End Sub

Sub sendclick (Handle)
X% = sendmessage(Handle, WM_LBUTTONDOWN, 0, 0&)
X% = sendmessage(Handle, WM_LBUTTONUP, 0, 0&)
End Sub

Sub sendtext (sill$)
aol% = FindWindow("AOL Frame25", 0&)
chatlist% = FindChildByClass(aol%, "_AOL_Listbox")
chatwin% = GetParent(chatlist%)
chatview% = FindChildByClass(chatwin%, "_AOL_View")
chatedit% = FindChildByClass(chatwin%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, sill$)
SendNow% = SendMessageByNum(chatedit%, WM_CHAR, &HD, 0)
'Call Click(ChatSend%)
Timeout (.1)

End Sub

Sub sendto (SendString$)
Dim aol%, aoledit%
aol% = FindWindow("AOL FRAME25", 0&)
aoledit% = FindChildByClass(aol%, "_AOL_EDIT")
X = SendMessageByString(aoledit%, WM_SETTEXT, 0&, SendString$)
DoEvents
X = SendMessageByNum(aoledit%, WM_CHAR, 13, 0&)
DoEvents
End Sub

Sub settabstops ()
ReDim TabPlace(4)
TabPlace(0) = 32
TabPlace(1) = 64
TabPlace(2) = 96
TabPlace(3) = 118
End Sub

Sub showaolwins ()
fc = FindChildByClass(aolhwnd(), "AOL Child")
req = showwindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = GetNextWindow(faa, 2)
res = showwindow(faa, 1)
DoEvents
Loop Until faf = faa


End Sub

Sub stayontop (frm As Form)
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function StrFromLst (lstname As ListBox) As String
On Error GoTo daerror
For i = 1 To lstname.ListCount
    DaSTr = DaSTr & lstname.List(i - 1) & ","
Next i
StrFromLst = DaSTr
Exit Function
daerror:
StrFromLst = ""
End Function

Function stripchar (info$, num%) As String
For i% = 1 To Len(info$)
    If Mid$(info$, i%, 1) <> Chr$(num) Then temp$ = temp$ + Mid$(info$, i%, 1)
Next i%
stripchar = LTrim$(RTrim$(temp$))
End Function

Function stripspaces (info$) As String
For i% = 1 To Len(info$)
    If InStr(Chr$(0) + " " + Chr$(10) + Chr$(13), Mid$(info$, i%, 1)) = 0 Then temp$ = temp$ + Mid$(info$, i%, 1)
Next i%
stripspaces = UCase$(LTrim$(RTrim$(temp$)))
End Function

 Sub TextSend (Hands$)
'
Dim Rets
Dim aol As Variant
Dim Room1 As Variant
Dim Room2 As Variant
Dim Room3 As Variant
Dim Room4 As Variant
Dim FindChat
aol = FindWindow("AOL FRAME25", 0&)
Room1 = FindChildByTitle(aol, "List Rooms")
Room2 = FindChildByTitle(aol, "AOL Live!")
Room3 = FindChildByTitle(aol, "The Plaza")
Room4 = FindChildByTitle(aol, "Member Directory")
If Room1 <> 0 And GetParent(Room1) = GetParent(Room2) And GetParent(Room2) = GetParent(Room3) And GetParent(Room3) = GetParent(Room4) And GetParent(Room4) = GetParent(Room1) Then
FindChat = GetParent(Room4)
Dim a
Dim b%
a = FindChat
b% = FindChildByClass(a, "_AOL_EDIT")
Rets = SendMessageByString(b%, &HC, 0, Hands$)
hfasj = SendMessageByNum(b%, WM_CHAR, 13, 0)
'
'to use type TextSend("text to send goes here")
End If
'
End Sub

Sub TextSet (hWnd As Integer, what As String)
Dim r
r = SendMessageByString(hWnd, &HC, 0, what)

End Sub

Sub Timeout (duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop
End Sub

Sub ToLBR (t$)
On Error Resume Next
AppActivate "America  Online"
listWnd = GetWindowFromClass(getfocus(), "AOL Toolbar")
t$ = aweffd
X = SetParent(listWnd, aweffd)
End Sub

Sub Top (frm As Form)
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Sub

Function trim_null (wstr As String)
wstr = Trim(wstr)
Do Until xx = Len(wstr)
Let xx = xx + 1
Let this_chr = Asc(Mid$(wstr, xx, 1))
If this_chr > 31 And this_chr <> 256 Then Let wordd = wordd & Mid$(wstr, xx, 1)
Loop
trim_null = wordd

End Function

Function TrimNull (innn$) As String
For X = 1 To Len(innn$)
    If (Mid$(innn$, X, 1) <> Chr$(0)) Then
        total$ = total$ + Mid$(innn$, X, 1)
    Else
        GoTo NullDetect
    End If
Next
NullDetect:
TrimNull = total$

End Function

Function trimnull2 (ass) As String
For X = 1 To Len(ass)
    If (Mid$(ass, X, 1) <> Chr$(0)) Then
        total$ = total$ + Mid$(ass, X, 1)
    Else
        GoTo NullDetect2
    End If
Next
NullDetect2:
trimnull2 = total$
End Function

Sub TurnIMs (Setting)
 RunMenu 4, 3
DoEvents
Clipboard.Clear
Clipboard.SetText "$im_" + Setting
z = getfocus()
DoEvents
Do
ll = getfocus()
DoEvents
Loop Until ll <> z
SendKeys "^V", True
Clipboard.Clear
Clipboard.SetText "Zero is KiNG!"
SendKeys "{TAB}", True
SendKeys "^V", True
SendKeys "{TAB}", True
SendKeys " ", True
Do
pi = GetControl()
DoEvents
Loop Until pi = "OK"
SendKeys " ", True
SendKeys "%-", True
SendKeys "C", True

End Sub

Sub vbmsgg ()
ChatText$ = agGetStringFromLPSTR$(lParam)
End Sub

Function verify (number$) As String
Do
  DoEvents
  total& = 0
  Mid$(number$, 13, 4) = LTrim$(RTrim$(Str$(Val(Mid$(number$, 13, 4)) + 1)))
  For l% = 1 To 16 Step 2
      num% = Val(Mid$(number$, l%, 1)) * 2
      If num% > 9 Then num% = num% - 9
      total& = total& + num%
  Next l%
  For l% = 2 To 16 Step 2
      num% = Val(Mid$(number$, l%, 1))
      total& = total& + num%
  Next l%
If total& Mod 10 = 0 Then
   verify = number$
   Exit Do
End If
Loop
End Function

Sub viewcng (Text As String)
aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient")
HWNDEDT% = FindChildByClass(MDI%, "_AOL_Edit")
AOLSEND% = FindChildByTitle(MDI%, "Send")
LSTBOX% = FindChildByClass(aol%, "_AOL_ListBox")
PARENT% = GetParent(LSTBOX%)
aolview% = FindChildByClass(PARENT%, "_AOL_View")
X = SendMessageByString(aolview%, WM_SETTEXT, 0, Chr$(13) & Chr$(10) & Text)  '& Chr$(13)

End Sub

Sub waitForchange ()
Dim old As Integer
Dim boy As Integer
old = getfocus()
boy = getfocus()
Do Until boy <> old
X = DoEvents()
boy = getfocus()
Loop

End Sub

Sub waitformail ()
Dim timr As Long
Dim begin As Double
Dim ending As Double
Dim dfc As Double
timr = getcurrenttime()
begin = Time
Do
listWnd = GetWindowFromClass(getfocus(), "_AOL_Tree") 'Auto "New Mail" Fill Detector
GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
kewl = lpmaxpos%
Pause (1.5)
listWnd = GetWindowFromClass(getfocus(), "_AOL_Tree")       'Auto "New Mail" FillDetector
GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
If kewl = lpmaxpos% Then
    listWnd = GetWindowFromClass(getfocus(), "_AOL_Tree")       'Auto "New Mail" Fill Detector
    GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
    kewl = lpmaxpos%
    Pause (1.5)
    If kewl = lpmaxpos% Then Exit Sub
    End If
 Loop
        'If lstlp = lpmaxpos% Then
        '  Call Pause(7)
        '    If lstlp = lpmaxpos% Then
        '    Exit Do
        '    End If
        '  Else
        '  lstlp = lpmaxpos%
        'End If
'If Cou >= 1400 Then Exit Do   'Change 1600 to more or less
'If GetCurrentTime() >= Timr + (30 * 1000) Then Exit Do
'Call Pause(.0001)             'This number determines the
X = DoEvents()                'number of loops with identical
                              'lpMaxPos's to assume the box
                              'is no longer being updated.

End Sub

Sub waitforok ()

Do
DoEvents
okw = FindWindow("#32770", "America Online")
DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function WaitForWindow (txt As String, clss As String) As Integer
Dim r
Dim hWnds() As Integer

Do
r = FindWindowLike(hWnds(), 0, txt, clss, Null)
DoEvents
Loop While r = 0
WaitForWindow = hWnds(1)

End Function

Sub waitmail ()
Do
Box = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail")
Timeout (.1)
Loop Until Box <> 0
List = FindChildByClass(Box, "_AOL_Tree")
Do
DoEvents
mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(.5)
mailnum2 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(.5)
mailnum3 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)


End Sub

Sub WaitWindow ()
A0L% = FindWindow("AOL Frame25", "America  Online")
Dim grl
Dim boy
End Sub

Function WinCount (ByVal hWnd)
 Dim ChildWin As Integer, FirstWin As Integer, LastWin As Integer, Cary As Integer, NextWini As Integer
 ChildWin = getwindow(hWnd, GW_CHILD)
 FirstWin = getwindow(ChildWin, GW_HWNDFIRST)
 LastWin = getwindow(ChildWin, GW_HWNDLAST)

  If ChildWin = 0 Then
       Cary = 0
  Else
       Cary = 1
  End If

 Do Until FirstWin = LastWin
      NextWini = getwindow(FirstWin, GW_HWNDNEXT)
      FirstWin = NextWini
      Cary = Cary + 1
      DoEvents
 Loop

If Cary = 0 Then
     WinCount = 0
Else
     WinCount = Cary
End If
End Function

Function windowcaption (hWndd As Integer)
Dim WindowText As String * 255
Dim getWinText As Integer
getWinText = getwindowtext(hWndd, WindowText, 255)
windowcaption = (WindowText)
End Function

Sub WriteINI (sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

