Attribute VB_Name = "Module1"
'~Cyrix here... I Have NO Clue on who wrote this..
'Its not a great Bas.. But if your a beginner.. Use
'This one..
'If you know who wrote this.. Please email me
'WaxOnline@Bigfoot.com
      '                 Thanks           ~Cyrix

Global Stopit
Declare Function message Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function WindowFromPoint Lib "User" (ByVal ptScreen As Any) As Integer
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer
Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Declare Sub GetCursorPos Lib "User" (lpPoint As Long)
Declare Function EnableWindow Lib "User" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Sub Drawmenubar Lib "User" (ByVal hWnd As Integer)
Declare Function GetSystemDirectory Lib "Kernel" (ByVal p$, ByVal s%) As Integer
Declare Function AOLGetList% Lib "311.Dll" (ByVal index%, ByVal Buf$)
Declare Function AOLGetcombo% Lib "311.Dll" (ByVal index%, ByVal Buf$)
Declare Function TrackPopupMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nReserved As Integer, ByVal hWnd As Integer, lpReserved As Any) As Integer
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function deletemenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer

'Menu Constants
Global Const SWP_NoActivate = &H10
Global Const MF_Popup = &H10
Global Const MF_String = &H0
Global Const MF_Enabled = &H0
Global Const MF_Separator = &H800
Global Const TPM_LEFTALIGN = &H0
Global Const TPM_RIGHTBUTTON = &H2
Declare Sub Drawmenubar Lib "User" (ByVal hWnd As Integer)
Declare Function TrackPopupMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nReserved As Integer, ByVal hWnd As Integer, lpReserved As Any) As Integer
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function deletemenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function ExitWindows Lib "User" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Global Const LB_GETTEXT = &H189
Global Const WM_GETDLGCODE = &H87
Global Const WM_CTLCOLOR = &H19
Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Global Const LBN_DBLCLK = 2
Const c003E = 16 ' &H10%
Global Const WM_USER = &H400
Global Const LB_GETSEL = (WM_USER + 8)
Declare Function sndplaysound Lib "MMSystem" (ByVal lpWavName$, ByVal Flags%) As Integer '
Global Const WM_MOVE = &H3
Global Const SWP_NOREPOSITION = &H200
Declare Sub Movewindow Lib "User" (ByVal hWnd As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Global Const WM_CLOSE = &H10
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const WM_ACTIVATE = &H6

Global Const NUMBOXES = 5
Global Const SAVEFILE = 1, LOADFILE = 2
Global Const REPLACEFILE = 1, READFILE = 2, ADDTOFILE = 3
Global Const RANDOMFILE = 4, BINARYFILE = 5

' Define a data typeto hold a record:
' Define global variables to hold the file number and record number
' of the current data file.
' Default file name to show in dialog boxes.
Global Const Err_DeviceUnavailable = 68
Global Const Err_DiskNotReady = 71, Err_FileAlreadyExists = 58
Global Const Err_TooManyFiles = 67, Err_RenameAcrossDisks = 74
Global Const Err_Path_FileAccessError = 75, Err_DeviceIO = 57
Global Const Err_DiskFull = 61, Err_BadFileName = 64
Global Const Err_BadFileNameOrNumber = 52, Err_FileNotFound = 53
Global Const Err_PathDoesNotExist = 76, Err_BadFileMode = 54
Global Const Err_FileAlreadyOpen = 55, Err_InputPastEndOfFile = 62
Global Const MB_EXCLAIM = 48, MB_STOP = 16


Global Const WM_KILLFOCUS = &H8
Global Const WM_SETFOCUS = &H7
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const WM_SIZE = &H5
Global Const SW_HIDE = 0
Global Const SW_Shownormal = 1
Global Const SW_Show = 5
Global Const WM_LBUTTONDBLCLK = &H203
Global Const SW_NORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_MAXIMIZE = 3
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_MINIMIZE = 6
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_RESTORE = 9
Global Const WM_FONTCHANGE = &H1D
Global Const WM_SETFONT = &H30

Type RECT
  Left As Integer
  top As Integer
  Right As Integer
  Bottom As Integer
End Type
Type ConvertPointAPI
    xy As Long
    End Type
Type MODEL
  usVersion         As Integer
  fl                As Long
  pctlproc          As Long
  fsClassStyle      As Integer
  flWndStyle        As Long
  cbCtlExtra        As Integer
  idBmpPalette      As Integer
  npszDefCtlName    As Integer
  npszClassName     As Integer
  npszParentClassName As Integer
  npproplist        As Integer
  npeventlist       As Integer
  nDefProp          As String * 1
  nDefEvent         As String * 1
  nValueProp        As String * 1
  usCtlVersion      As Integer
End Type
Type HelpWinInfo
  wStructSize As Integer
  x As Integer
  Y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type



'                   API Subs and Functions
'                   ----------------------

'Subs and Functions for "User"
Declare Function FillRect Lib "User" (ByVal hDC As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
Dim hBrush%
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function InsertMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Sub Drawmenubar Lib "User" (ByVal hWnd As Integer)
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer

Declare Function extfn2668 Lib "vbwfind.dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn16A8 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn1868 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function extfn2630 Lib "vbwfind.dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1AD0 Lib "User" Alias "SetWindowText" (ByVal p1%, ByVal p2$) As Integer

Declare Function FlashWindow Lib "User" (ByVal hWnd As Integer, ByVal bInvert As Integer) As Integer
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Sub Drawmenubar Lib "User" (ByVal hWnd As Integer)
Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal Y As Integer)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Function getfreesystemresources Lib "User" (ByVal fuSysResource%) As Integer
Declare Function GetKeyboardType Lib "Keyboard" (ByVal nTypeFlag As Integer) As Integer
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function ExitWindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)
Declare Function SetParent% Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer)
Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)
'Declare Function sendmessage& Lib "User" (ByVal hWnd%, ByVal wmsg%, ByVal wparam%, ByVal lParam As Any)
Declare Function CreateMenu% Lib "User" ()
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function AppendMenuByString% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)
Declare Function InsertMenu% Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any)
Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)
Declare Function WinHelpByNum% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&)
Declare Function GetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer)
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetActiveWindow% Lib "User" ()
Declare Function SetActiveWindow% Lib "User" (ByVal hWnd%)
Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function IsWindowVisible% Lib "User" (ByVal hWnd%)
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function GetCursor% Lib "User" ()
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Global Const BM_SETCHECK = WM_USER + 1
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetTopWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)
Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wId%, ByVal wFlags%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDeskTopWindow Lib "User" () As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal LParam&)
Declare Function SendMessageLong Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal hMsg As Integer, ByVal wParam As Integer, ByVal LParam As Any) As Long

'Subs and Functions for "Kernel"
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function getfreespace Lib "Kernel" (ByVal wFlags%) As Long
Declare Function GetSystegenrectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal newstring As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFilename$) As Integer
Declare Function agGetStringFromLPSTR$ Lib "APIGuide.dll" (ByVal lpString&)

'Subs and Functions for "GDI"
Declare Sub SetBKColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer

'Subs and Functions for "MMSystem"
Declare Function mciSendString& Lib "MMSystem" (ByVal cmd$, ByVal Returnstr As Any, ByVal returnlen%, ByVal hcallback%)

'Subs and Functions for "VBWFind.Dll"
Declare Function FindChild% Lib "vbwfind.dll" (ByVal hWnd%, ByVal Title$)

'Subs and Functions for "APIGuide.dll"
Declare Sub agCopyData Lib "APIGuide.dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "APIGuide.dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Sub agDWordTo2Integers Lib "APIGuide.dll" (ByVal l&, lw%, lh%)
Declare Sub agOutp Lib "APIGuide.dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpw Lib "APIGuide.dll" (ByVal portid%, ByVal outval%)
Declare Function agGetControlHwnd% Lib "APIGuide.dll" (hctl As Control)
Declare Function agGetInstance% Lib "APIGuide.dll" ()
Declare Function agGetAddressForObject& Lib "APIGuide.dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "APIGuide.dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "APIGuide.dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "APIGuide.dll" Alias "agGetAddressForObject" (ByVal lpString$)
Declare Function agGetAddressForVBString& Lib "APIGuide.dll" (vbstring$)
Declare Function agGetControlName$ Lib "APIGuide.dll" (ByVal hWnd%)
Declare Function agXPixelsToTwips& Lib "APIGuide.dll" (ByVal pixels%)
Declare Function agYPixelsToTwips& Lib "APIGuide.dll" (ByVal pixels%)
Declare Function agXTwipsToPixels% Lib "APIGuide.dll" (ByVal twips&)
Declare Function agYTwipsToPixels% Lib "APIGuide.dll" (ByVal twips&)
Declare Function agDeviceCapabilities& Lib "APIGuide.dll" (ByVal hlib%, ByVal lpszDevice$, ByVal lpszPort$, ByVal fwCapability%, ByVal lpszOutput&, ByVal lpdm&)
Declare Function agDeviceMode% Lib "APIGuide.dll" (ByVal hWnd%, ByVal hModule%, ByVal lpszDevice$, ByVal lpszOutput$)
Declare Function agExtDeviceMode% Lib "APIGuide.dll" (ByVal hWnd%, ByVal hDriver%, ByVal lpdmOutput&, ByVal lpszDevice$, ByVal lpszPort$, ByVal lpdmInput&, ByVal lpszProfile&, ByVal fwMode%)
Declare Function agInp% Lib "APIGuide.dll" (ByVal portid%)
Declare Function agInpw% Lib "APIGuide.dll" (ByVal portid%)
Declare Function agHugeOffset& Lib "APIGuide.dll" (ByVal addr&, ByVal offset&)
Declare Function agVBGetVersion% Lib "APIGuide.dll" ()
Declare Function agVBSendControlMsg& Lib "APIGuide.dll" (Ctl As Control, ByVal Msg%, ByVal wp%, ByVal lp&)
Declare Function agVBSetControlFlags& Lib "APIGuide.dll" (Ctl As Control, ByVal mask&, ByVal value&)
Declare Function dwVBSetControlFlags& Lib "APIGuide.dll" (Ctl As Control, ByVal mask&, ByVal value&)


'Subs and Functions for "VBMsg.Vbx"
Declare Sub ptGetTypeFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptCopyTypeToAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptSetControlModel Lib "VBMsg.Vbx" (Ctl As Control, lpm As MODEL)
Declare Function ptGetVariableAddress Lib "VBMsg.Vbx" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (ByVal s As String) As Long
Declare Function ptGetLongAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (l As Long) As Long
Declare Function ptGetIntegerAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (i As Integer) As Long
Declare Function ptGetIntegerFromAddress Lib "VBMsg.Vbx" (ByVal i As Long) As Integer
Declare Function ptGetLongFromAddress Lib "VBMsg.Vbx" (ByVal l As Long) As Long
Declare Function ptGetStringFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, ByVal cbBytes As Integer) As String
Declare Function ptMakelParam Lib "VBMsg.Vbx" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Declare Function ptLoWord Lib "VBMsg.Vbx" (ByVal LParam As Long) As Integer
Declare Function ptHiWord Lib "VBMsg.Vbx" (ByVal LParam As Long) As Integer
Declare Function ptMakeUShort Lib "VBMsg.Vbx" (ByVal LongVal As Long) As Integer
Declare Function ptConvertUShort Lib "VBMsg.Vbx" (ByVal ushortVal As Integer) As Long
Declare Function ptMessagetoText Lib "VBMsg.Vbx" (ByVal uMsgID As Long, ByVal bFlag As Integer) As String
Declare Function ptRecreateControlHwnd Lib "VBMsg.Vbx" (Ctl As Control) As Long
Declare Function ptGetControlModel Lib "VBMsg.Vbx" (Ctl As Control, lpm As MODEL) As Long
Declare Function ptGetControlName Lib "VBMsg.Vbx" (Ctl As Control) As String

'Subs and Functions for Other DLL's and VBX's
Declare Function VarPtr& Lib "VBRun300.Dll" (Param As Any)
Declare Function vbeNumChildWindow% Lib "VBStr.Dll" (ByVal win%, ByVal iNum%)

'                   Sound Constants
'                   ---------------


'                   Global Constants
'                   ----------------
'OpenFile() Flags
Global Const OF_READ = &H0
Global Const OF_WRITE = &H1
Global Const OF_READWRITE = &H2
Global Const OF_SHARE_COMPAT = &H0
Global Const OF_SHARE_EXCLUSIVE = &H10
Global Const OF_SHARE_DENY_WRITE = &H20
Global Const OF_SHARE_DENY_READ = &H30
Global Const OF_SHARE_DENY_NONE = &H40
Global Const OF_PARSE = &H100
Global Const OF_DELETE = &H200
Global Const OF_VERIFY = &H400
Global Const OF_SEARCH = &H400
Global Const OF_CANCEL = &H800
Global Const OF_CREATE = &H1000
Global Const OF_PROMPT = &H2000
Global Const OF_EXIST = &H4000
Global Const OF_REOPEN = &H8000
Global Const TF_FORCEDRIVE = &H80

'GetDriveType return values
Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4

'Global Memory Flags
Global Const GMEM_FIXED = &H0
Global Const GMEM_MOVEABLE = &H2
Global Const GMEM_NOCOMPACT = &H10
Global Const GMEM_NODISCARD = &H20
Global Const GMEM_ZEROINIT = &H40
Global Const GMEM_MODIFY = &H80
Global Const GMEM_DISCARDABLE = &H100
Global Const GMEM_NOT_BANKED = &H1000
Global Const GMEM_SHARE = &H2000
Global Const GMEM_DDESHARE = &H2000
Global Const GMEM_NOTIFY = &H4000
Global Const GMEM_LOWER = GMEM_NOT_BANKED
Global Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Global Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

'Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
Global Const GMEM_DISCARDED = &H4000
Global Const GMEM_LOCKCOUNT = &HFF

'Predefined Resource Types
Global Const RT_CURSOR = 1&
Global Const RT_BITMAP = 2&
Global Const RT_ICON = 3&
Global Const RT_MENU = 4&
Global Const RT_DIALOG = 5&
Global Const RT_STRING = 6&
Global Const RT_FONTDIR = 7&
Global Const RT_FONT = 8&
Global Const RT_ACCELERATOR = 9&
Global Const RT_RCDATA = 10&

'GetFreeSystemResources
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

'GetWinFlags
Global Const WF_PMODE = &H1
Global Const WF_CPU286 = &H2
Global Const WF_CPU386 = &H4
Global Const WF_CPU486 = &H8
Global Const WF_STANDARD = &H10
Global Const WF_WIN286 = &H10
Global Const WF_ENHANCED = &H20
Global Const WF_WIN386 = &H20
Global Const WF_CPU086 = &H40
Global Const WF_CPU186 = &H80
Global Const WF_LARGEFRAME = &H100
Global Const VK_END = &H23
Global Const WF_SMALLFRAME = &H200
Global Const WF_80x87 = &H400

'Parameter error checking
Global Const ERR_WARNING = 8
Global Const ERR_PARAM = 4
Global Const ERR_SIZE_MASK = 3
Global Const ERR_BYTE = 1
Global Const ERR_WORD = 2
Global Const ERR_DWORD = 3
Global Const ERR_BAD_VALUE = &H6001
Global Const ERR_BAD_FLAGS = &H6002
Global Const ERR_BAD_INDEX = &H6003
Global Const ERR_BAD_DVALUE = &H7004
Global Const ERR_BAD_DFLAGS = &H7005
Global Const ERR_BAD_DINDEX = &H7006
Global Const ERR_BAD_PTR = &H7007
Global Const ERR_BAD_FUNC_PTR = &H7008
Global Const ERR_BAD_SELECTOR = &H6009
Global Const ERR_BAD_STRING_PTR = &H700A
Global Const ERR_BAD_HANDLE = &H600B

'KERNEL parameter errors
Global Const ERR_BAD_HINSTANCE = &H6020
Global Const ERR_BAD_HMODULE = &H6021
Global Const ERR_BAD_GLOBAL_HANDLE = &H6022
Global Const ERR_BAD_LOCAL_HANDLE = &H6023
Global Const ERR_BAD_ATOM = &H6024
Global Const ERR_BAD_HFILE = &H6025

'USER parameter errors
Global Const ERR_BAD_HWND = &H6040
Global Const ERR_BAD_HMENU = &H6041
Global Const ERR_BAD_HCURSOR = &H6042
Global Const ERR_BAD_HICON = &H6043
Global Const ERR_BAD_HDWP = &H6044
Global Const ERR_BAD_CID = &H6045
Global Const ERR_BAD_HDRVR = &H6046

'GDI parameter errors
Global Const ERR_BAD_COORDS = &H7060
Global Const ERR_BAD_GDI_OBJECT = &H6061
Global Const ERR_BAD_HDC = &H6062
Global Const ERR_BAD_HPEN = &H6063
Global Const ERR_BAD_HFONT = &H6064
Global Const ERR_BAD_HBRUSH = &H6065
Global Const ERR_BAD_HBITMAP = &H6066
Global Const ERR_BAD_HRGN = &H6067
Global Const ERR_BAD_HPALETTE = &H6068
Global Const ERR_BAD_HMETAFILE = &H6069

'KERNEL errors
Global Const ERR_GALLOC = &H1
Global Const ERR_GREALLOC = &H2
Global Const ERR_GLOCK = &H3
Global Const ERR_LALLOC = &H4
Global Const ERR_LREALLOC = &H5
Global Const ERR_LLOCK = &H6
Global Const ERR_ALLOCRES = &H7
Global Const ERR_LOCKRES = &H8
Global Const ERR_LOADMODULE = &H9

'USER errors
Global Const ERR_CREATEDLG = &H40
Global Const ERR_CREATEDLG2 = &H41
Global Const ERR_REGISTERCLASS = &H42
Global Const ERR_DCBUSY = &H43
Global Const ERR_CREATEWND = &H44
Global Const ERR_STRUCEXTRA = &H45
Global Const ERR_LOADSTR = &H46
Global Const ERR_LOADMENU = &H47
Global Const ERR_NESTEDBEGINPAINT = &H48
Global Const ERR_BADINDEX = &H49
Global Const ERR_CREATEMENU = &H4A

'GDI errors
Global Const ERR_CREATEDC = &H80
Global Const ERR_CREATEMETA = &H81
Global Const ERR_DELOBJSELECTED = &H82
Global Const ERR_SELBITMAP = &H83

'Exit Window parameters

Global Const EW_RESTARTWindow = &H42
Global Const EW_REBOOTSYSTEM = &H43


'Stock system bitmaps
Global Const OBM_CLOSE = 32754
Global Const OBM_UPARROW = 32753
Global Const OBM_DNARROW = 32752
Global Const OBM_RGARROW = 32751
Global Const OBM_LFARROW = 32750
Global Const OBM_REDUCE = 32749
Global Const OBM_ZOOM = 32748
Global Const OBM_RESTORE = 32747
Global Const OBM_REDUCED = 32746
Global Const OBM_ZOOMD = 32745
Global Const OBM_RESTORED = 32744
Global Const OBM_UPARROWD = 32743
Global Const OBM_DNARROWD = 32742
Global Const OBM_RGARROWD = 32741
Global Const OBM_LFARROWD = 32740
Global Const OBM_MNARROW = 32739
Global Const OBM_COMBO = 32738
Global Const OBM_UPARROWI = 32737
Global Const OBM_DNARROWI = 32736
Global Const OBM_RGARROWI = 32735
Global Const OBM_LFARROWI = 32734
Global Const OBM_OLD_CLOSE = 32767
Global Const OBM_SIZE = 32766
Global Const OBM_OLD_UPARROW = 32765
Global Const OBM_OLD_DNARROW = 32764
Global Const OBM_OLD_RGARROW = 32763
Global Const OBM_OLD_LFARROW = 32762
Global Const OBM_BTSIZE = 32761
Global Const OBM_CHECK = 32760
Global Const OBM_CHECKBOXES = 32759
Global Const OBM_BTNCORNERS = 32758
Global Const OBM_OLD_REDUCE = 32757
Global Const OBM_OLD_ZOOM = 32756
Global Const OBM_OLD_RESTORE = 32755

'Stock system Icons
Global Const OCR_NORMAL = 32512
Global Const OCR_IBEAM = 32513
Global Const OCR_WAIT = 32514
Global Const OCR_CROSS = 32515
Global Const OCR_UP = 32516
Global Const OCR_SIZE = 32640
Global Const OCR_ICON = 32641
Global Const OCR_SIZENWSE = 32642
Global Const OCR_SIZENESW = 32643
Global Const OCR_SIZEWE = 32644
Global Const OCR_SIZENS = 32645
Global Const OCR_SIZEALL = 32646
Global Const OCR_ICOCUR = 32647
Global Const OIC_SAMPLE = 32512
Global Const OIC_HAND = 32513
Global Const OIC_QUES = 32514
Global Const OIC_BANG = 32515
Global Const OIC_NOTE = 32516

'Raster-ops (Binary)
Global Const R2_BLACK = 1 ' 0
Global Const R2_NOTMERGEPEN = 2 'DPon
Global Const R2_MASKNOTPEN = 3 'DPna
Global Const R2_NOTCOPYPEN = 4 'PN
Global Const R2_MASKPENNOT = 5 'PDna
Global Const R2_NOT = 6 'Dn
Global Const R2_XORPEN = 7 'DPx
Global Const R2_NOTMASKPEN = 8 'DPan
Global Const R2_MASKPEN = 9 'DPa
Global Const R2_NOTXORPEN = 10 'DPxn
Global Const R2_NOP = 11 'D
Global Const R2_MERGENOTPEN = 12 'DPno
Global Const R2_COPYPEN = 13 'P
Global Const R2_MERGEPENNOT = 14 'PDno
Global Const R2_MERGEPEN = 15 'DPo
Global Const R2_WHITE = 16 ' 1

'Raster-ops (Ternary)
Global Const SRCCOPY = &HCC0020
Global Const SRCPAINT = &HEE0086
Global Const SRCAND = &H8800C6
Global Const SRCINVERT = &H660046
Global Const SRCERASE = &H440328
Global Const NOTSRCCOPY = &H330008
Global Const NOTSRCERASE = &H1100A6
Global Const MERGECOPY = &HC000CA
Global Const MERGEPAINT = &HBB0226
Global Const PATCOPY = &HF00021
Global Const PATPAINT = &HFB0A09
Global Const PATINVERT = &H5A0049
Global Const DSTINVERT = &H550009
Global Const BLACKNESS = &H42&
Global Const WHITENESS = &HFF0062

'StretchBlt() Modes
Global Const BLACKONWHITE = 1
Global Const WHITEONBLACK = 2
Global Const COLORONCOLOR = 3

'PolyFill() Modes
Global Const ALTERNATE = 1
Global Const WINDING = 2

'Text Alignment Options
Global Const TA_NOUPDATECP = 0
Global Const TA_UPDATECP = 1
Global Const TA_LEFT = 0
Global Const TA_RIGHT = 2
Global Const TA_CENTER = 6
Global Const TA_TOP = 0
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24

'ExtTextOut flags
Global Const ETO_GRAYED = 1
Global Const ETO_OPAQUE = 2
Global Const ETO_CLIPPED = 4

'SetMapperFlags
Global Const ASPECT_FILTERING = &H1

'Metafile Functions
Global Const META_SETBKCOLOR = &H201
Global Const META_SETBKMODE = &H102
Global Const META_SETMAPMODE = &H103
Global Const META_SETROP2 = &H104
Global Const META_SETRELABS = &H105
Global Const META_SETPOLYFILLMODE = &H106
Global Const META_SETSTRETCHBLTMODE = &H107
Global Const META_SETTEXTCHAREXTRA = &H108
Global Const META_SETTEXTCOLOR = &H209
Global Const META_SETTEXTJUSTIFICATION = &H20A
Global Const META_SETWINDOWORG = &H20B
Global Const META_SETWINDOWEXT = &H20C
Global Const META_SETVIEWPORTORG = &H20D
Global Const META_SETVIEWPORTEXT = &H20E
Global Const META_OFFSETWINDOWORG = &H20F
Global Const META_SCALEWINDOWEXT = &H400
Global Const META_OFFSETVIEWPORTORG = &H211
Global Const META_SCALEVIEWPORTEXT = &H412
Global Const META_LINETO = &H213
Global Const META_MOVETO = &H214
Global Const META_EXCLUDECLIPRECT = &H415
Global Const META_INTERSECTCLIPRECT = &H416
Global Const META_ARC = &H817
Global Const META_ELLIPSE = &H418
Global Const META_FLOODFILL = &H419
Global Const META_PIE = &H81A
Global Const META_RECTANGLE = &H41B
Global Const META_ROUNDRECT = &H61C
Global Const META_PATBLT = &H61D
Global Const META_SAVEDC = &H1E
Global Const META_SETPIXEL = &H41F
Global Const META_OFFSETCLIPRGN = &H220
Global Const META_TEXTOUT = &H521
Global Const META_BITBLT = &H922
Global Const META_STRETCHBLT = &HB23
Global Const META_POLYGON = &H324
Global Const META_POLYLINE = &H325
Global Const META_ESCAPE = &H626
Global Const META_RESTOREDC = &H127
Global Const META_FILLREGION = &H228
Global Const META_FRAMEREGION = &H429
Global Const META_INVERTREGION = &H12A
Global Const META_PAINTREGION = &H12B
Global Const META_SELECTCLIPREGION = &H12C
Global Const META_SELECTOBJECT = &H12D
Global Const META_SETTEXTALIGN = &H12E
Global Const META_DRAWTEXT = &H62F
Global Const META_CHORD = &H830
Global Const META_SETMAPPERFLAGS = &H231
Global Const META_EXTTEXTOUT = &HA32
Global Const META_SETDIBTODEV = &HD33
Global Const META_SELECTPALETTE = &H234
Global Const META_REALIZEPALETTE = &H35
Global Const META_ANIMATEPALETTE = &H436
Global Const META_SETPALENTRIES = &H37
Global Const META_POLYPOLYGON = &H538
Global Const META_RESIZEPALETTE = &H139
Global Const META_DIBBITBLT = &H940
Global Const META_DIBSTRETCHBLT = &HB41
Global Const META_DIBCREATEPATTERNBRUSH = &H142
Global Const META_STRETCHDIB = &HF43
Global Const META_DELETEOBJECT = &H1F0
Global Const META_CREATEPALETTE = &HF7
Global Const META_CREATEBRUSH = &HF8
Global Const META_CREATEPATTERNBRUSH = &H1F9
Global Const META_CREATEPENINDIRECT = &H2FA
Global Const META_CREATEFONTINDIRECT = &H2FB
Global Const META_CREATEBRUSHINDIRECT = &H2FC
Global Const META_CREATEBITMAPINDIRECT = &H2FD
Global Const META_CREATEBITMAP = &H6FE
Global Const META_CREATEREGION = &H6FF

'Escape
Global Const NEWFRAME = 1
Global Const ABORTDOCCONST = 2
Global Const NEXTBAND = 3
Global Const SETCOLORTABLE = 4
Global Const GETCOLORTABLE = 5
Global Const FLUSHOUTPUT = 6
Global Const DRAFTMODE = 7
Global Const QUERYESCSUPPORT = 8
Global Const SETABORTPROCCONST = 9
Global Const STARTDOCCONST = 10
Global Const ENDDOCAPICONST = 11
Global Const GETPHYSPAGESIZE = 12
Global Const GETPRINTINGOFFSET = 13
Global Const GETSCALINGFACTOR = 14
Global Const MFCOMMENT = 15
Global Const GETPENWIDTH = 16
Global Const SETCOPYCOUNT = 17
Global Const SELECTPAPERSOURCE = 18
Global Const DEVICEDATA = 19
Global Const PASSTHROUGH = 19
Global Const GETTECHNOLGY = 20
Global Const GETTECHNOLOGY = 20
Global Const SETENDCAP = 21
Global Const SETLINEJOIN = 22
Global Const SETMITERLIMIT = 23
Global Const BANDINFO = 24
Global Const DRAWPATTERNRECT = 25
Global Const GETVECTORPENSIZE = 26
Global Const GETVECTORBRUSHSIZE = 27
Global Const ENABLEDUPLEX = 28
Global Const GETSETPAPERBINS = 29
Global Const GETSETPRINTORIENT = 30
Global Const ENUMPAPERBINS = 31
Global Const SETDIBSCALING = 32
Global Const EPSPRINTING = 33
Global Const ENUMPAPERMETRICS = 34
Global Const GETSETPAPERMETRICS = 35
Global Const POSTSCRIPT_DATA = 37
Global Const POSTSCRIPT_IGNORE = 38
Global Const GETEXTENDEDTEXTMETRICS = 256
Global Const GETEXTENTTABLE = 257
Global Const GETPAIRKERNTABLE = 258
Global Const GETTRACKKERNTABLE = 259
Global Const EXTTEXTOUTCONST = 512
Global Const ENABLERELATIVEWIDTHS = 768
Global Const ENABLEPAIRKERNING = 769
Global Const SETKERNTRACK = 770
Global Const SETALLJUSTVALUES = 771
Global Const SETCHARSET = 772
Global Const STRETCHBLTCONST = 2048
Global Const BEGIN_PATH = 4096
Global Const CLIP_TO_PATH = 4097
Global Const END_PATH = 4098
Global Const EXT_DEVICE_CAPS = 4099
Global Const RESTORE_CTM = 4100
Global Const SAVE_CTM = 4101
Global Const SET_ARC_DIRECTION = 4102
Global Const SET_BACKGROUND_COLOR = 4103
Global Const SET_POLY_MODE = 4104
Global Const SET_SCREEN_ANGLE = 4105
Global Const SET_SPREAD = 4106
Global Const TRANSFORM_CTM = 4107
Global Const SET_CLIP_BOX = 4108
Global Const SET_BOUNDS = 4109
Global Const SET_MIRROR_MODE = 4110

'Spooler Error Codes
Global Const SP_NOTREPORTED = &H4000
Global Const SP_ERROR = (-1)
Global Const SP_APPABORT = (-2)
Global Const SP_USERABORT = (-3)
Global Const SP_OUTOFDISK = (-4)
Global Const SP_OUTOFMEMORY = (-5)
Global Const PR_JOBSTATUS = &H0

'biCompression field constants for DIB
Global Const BI_RGB = 0&
Global Const BI_RLE8 = 1&
Global Const BI_RLE4 = 2&

'LOGFONT and TEXTMETRIC
Global Const OUT_DEFAULT_PRECIS = 0
Global Const OUT_STRING_PRECIS = 1
Global Const OUT_CHARACTER_PRECIS = 2
Global Const OUT_STROKE_PRECIS = 3
Global Const OUT_TT_PRECIS = 4
Global Const OUT_DEVICE_PRECIS = 5
Global Const OUT_RASTER_PRECIS = 6
Global Const OUT_TT_ONLY_PRECIS = 7
Global Const CLIP_DEFAULT_PRECIS = 0
Global Const CLIP_CHARACTER_PRECIS = 1
Global Const CLIP_STROKE_PRECIS = 2
Global Const CLIP_LH_ANGLES = &H10
Global Const CLIP_TT_ALWAYS = &H20
Global Const CLIP_EMBEDDED = &H80
Global Const DEFAULT_QUALITY = 0
Global Const DRAFT_QUALITY = 1
Global Const PROOF_QUALITY = 2
Global Const DEFAULT_PITCH = 0
Global Const FIXED_PITCH = 1
Global Const VARIABLE_PITCH = 2
Global Const TMPF_FIXED_PITCH = 1
Global Const TMPF_VECTOR = 2
Global Const TMPF_DEVICE = 8
Global Const TMPF_TRUETYPE = 4
Global Const ANSI_CHARSET = 0
Global Const DEFAULT_CHARSET = 1
Global Const SYMBOL_CHARSET = 2
Global Const SHIFTJIS_CHARSET = 128
Global Const OEM_CHARSET = 255
Global Const NTM_REGULAR = &H40&
Global Const NTM_BOLD = &H20&
Global Const NTM_ITALIC = &H1&
Global Const LF_FULLFACESIZE = 64
Global Const RASTER_FONTTYPE = 1
Global Const DEVICE_FONTTYPE = 2
Global Const TRUETYPE_FONTTYPE = 4


'Font Families
Global Const FF_DONTCARE = 0
Global Const FF_ROMAN = 16

'Times Roman, Century Schoolbook, etc.
Global Const FF_SWISS = 32

' Helvetica, Swiss, etc.
Global Const FF_MODERN = 48

' Pica, Elite, Courier, etc.
Global Const FF_SCRIPT = 64
Global Const FF_DECORATIVE = 80

'Font Weights
Global Const FW_DONTCARE = 0
Global Const FW_THIN = 100
Global Const FW_EXTRALIGHT = 200
Global Const FW_LIGHT = 300
Global Const FW_NORMAL = 400
Global Const FW_MEDIUM = 500
Global Const FW_SEMIBOLD = 600
Global Const FW_BOLD = 700
Global Const FW_EXTRABOLD = 800
Global Const FW_HEAVY = 900
Global Const FW_ULTRALIGHT = FW_EXTRALIGHT
Global Const FW_REGULAR = FW_NORMAL
Global Const FW_DEMIBOLD = FW_SEMIBOLD
Global Const FW_ULTRABOLD = FW_EXTRABOLD
Global Const FW_BLACK = FW_HEAVY

'Background Modes
Global Const TRANSPARENT = 1
Global Const OPAQUE = 2

'Mapping Modes
Global Const MM_TEXT = 1
Global Const MM_LOMETRIC = 2
Global Const MM_HIMETRIC = 3
Global Const MM_LOENGLISH = 4
Global Const MM_HIENGLISH = 5
Global Const MM_TWIPS = 6
Global Const MM_ISOTROPIC = 7
Global Const MM_ANISOTROPIC = 8

'Coordinate Modes
Global Const ABSOLUTE = 1
Global Const RELATIVE = 2

'Stock Logical Objects
Global Const WHITE_BRUSH = 0
Global Const LTGRAY_BRUSH = 1
Global Const GRAY_BRUSH = 2
Global Const DKGRAY_BRUSH = 3
Global Const BLACK_BRUSH = 4
Global Const NULL_BRUSH = 5
Global Const HOLLOW_BRUSH = NULL_BRUSH
Global Const WHITE_PEN = 6
Global Const BLACK_PEN = 7
Global Const NULL_PEN = 8
Global Const OEM_FIXED_FONT = 10
Global Const ANSI_FIXED_FONT = 11
Global Const ANSI_VAR_FONT = 12
Global Const SYSTEM_FONT = 13
Global Const DEVICE_DEFAULT_FONT = 14
Global Const DEFAULT_PALETTE = 15
Global Const SYSTEM_FIXED_FONT = 16

'Brush Styles
Global Const BS_SOLID = 0
Global Const BS_NULL = 1
Global Const BS_HOLLOW = BS_NULL
Global Const BS_HATCHED = 2
Global Const BS_PATTERN = 3
Global Const BS_INDEXED = 4
Global Const BS_DIBPATTERN = 5

'Hatch Styles
Global Const HS_HORIZONTAL = 0
Global Const HS_VERTICAL = 1
Global Const HS_FDIAGONAL = 2
Global Const HS_BDIAGONAL = 3
Global Const HS_CROSS = 4
Global Const HS_DIAGCROSS = 5

'Pen Styles
Global Const PS_SOLID = 0
Global Const PS_DASH = 1
Global Const PS_DOT = 2
Global Const PS_DASHDOT = 3
Global Const PS_DASHDOTDOT = 4
Global Const PS_NULL = 5
Global Const PS_INSIDEFRAME = 6

'Bounds Rectangle Constants
Global Const DCB_RESET = 1
Global Const DCB_ACCUMULATE = 2
Global Const DCB_DIRTY = 2
Global Const DCB_SET = 3
Global Const DCB_ENABLE = 4
Global Const DCB_DISABLE = 8

'GetDeviceCaps() Device Parameters
Global Const DRIVERVERSION = 0
Global Const TECHNOLOGY = 2
Global Const HORZSIZE = 4
Global Const VERTSIZE = 6
Global Const HORZRES = 8
Global Const VERTRES = 10
Global Const BITSPIXEL = 12
Global Const PLANES = 14
Global Const NUMBRUSHES = 16
Global Const NUMPENS = 18
Global Const NUMMARKERS = 20
Global Const NUMFONTS = 22
Global Const NUMCOLORS = 24
Global Const PDEVICESIZE = 26
Global Const CURVECAPS = 28
Global Const LINECAPS = 30
Global Const POLYGONALCAPS = 32
Global Const TEXTCAPS = 34
Global Const CLIPCAPS = 36
Global Const RASTERCAPS = 38
Global Const ASPECTX = 40
Global Const ASPECTY = 42
Global Const ASPECTXY = 44
Global Const LOGPIXELSX = 88
Global Const LOGPIXELSY = 90
Global Const SIZEPALETTE = 104
Global Const NUMRESERVED = 106
Global Const COLORRES = 108
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Function ReleaseDC Lib "User" (ByVal hWnd%, ByVal hDC%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Function getfocus% Lib "User" ()
Declare Sub GetScrollRange Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Function GetCurrentTime& Lib "User" ()
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Declare Function destroywindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Sub closewindow Lib "User" (ByVal hWnd As Integer)
Declare Function GetMenuItemCount Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function SetFocusApi% Lib "User" Alias "SetFocus" (ByVal hWnd As Integer)
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function test Lib "APIGuide.dll" Alias "AgGetStringFromLPSTR" (ByVal p1&) As String
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer

Global Const WM_MButtondown = &H207
Global Const WM_MButtonup = &H208
Global Const GW_HWNDNEXT = 2
Global Const WM_GETTEXTLENGTH = &HE
Global Const MF_BYPOSITION = &H400
Global Const WM_COMMAND = &H111
Global Const WM_CHAR = &H102
Global Const WM_GETTEXT = &HD
Global Const WM_SETTEXT = &HC
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function Getwindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, LParam As Any) As Long
Declare Function sendmessagebynum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal LParam&)
Declare Function SendMessagebystring& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal LParam$)
Declare Function DeleteBubble% Lib "bubble.dll" (ByVal wnd%)
Declare Function findchildbytitle% Lib "vbwfind.dll" (ByVal parent%, ByVal Title$)
Declare Function FindChildByClass% Lib "vbwfind.dll" (ByVal parent%, ByVal Title$)
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function getparent Lib "User" (ByVal hWnd As Integer) As Integer

Sub addroom(lst As ListBox)
Call KillGlyph
aol% = FindWindow("aol frame25", 0&)
thechat% = FindChatRoom()
TheList% = FindChildByClass(thechat%, "_AOL_ListBox")
If TheList% = 0 Then
MsgBox "Chat room not opened. Error. Aborting."
Exit Sub
End If
NumOfPeople% = sendmessagebynum(TheList%, LB_GETCOUNT, 0, 0&)
TheGlyph% = FindChildByClass(thechat%, "_AOL_Glyph")
killwin TheGlyph%
TheI:
For i% = 0 To NumOfPeople% - 1
NextName% = sendmessagebynum(TheList%, LB_SETCURSEL, i%, 0)
TheFocus% = SetFocusApi(TheList%)
DoubleClick TheList%
timeout 1
Do
TheGlyph% = FindChildByClass(findchildbytitle(FindWindow("AOL Frame25", 0&), ""), "_AOL_Glyph")
NoFreeze% = DoEvents()
Loop Until TheGlyph% <> 0
TheScreeName% = getparent(TheGlyph%)
TheSN$ = GetWinText(TheScreeName%)
killwin TheScreeName%
p = UserSN()
For x = 0 To lst.ListCount
 If lst.List(x) = TheSN$ Then
  GoTo TheSkip
 End If
Next x
If TheSN$ = p Then GoTo TheSkip
lst.AddItem Trim(TheSN$)
TheSkip:
Next i%
TheScreeName% = getparent(TheGlyph%)
If TheScreeName% <> 0 Then
Do
TheScreeName% = getparent(TheGlyph%)
killwin TheScreeName%
Loop Until TheScreeName% = 0
End If
End Sub

Sub anti45mintimer()
AOTimer% = FindWindow("_AOL_Palette", 0&)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Sub antiidle()
AOModal% = FindWindow("_AOL_Modal", 0&)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Sub AOLTitle(NewTitle$)
aol% = FindWindow("aol frame25", 0&)
Send aol%, NewTitle$
End Sub

Function AOLVersion()
aol% = FindWindow("AOL Frame25", "America  Online")
channels% = findchildbytitle(aol%, "Channels")
toolbar% = FindChildByClass(aol%, "AOL ToolBar")
KWord% = FindChildByClass(toolbar%, "Edit")
If channels% = 0 Then
AOLVersion = "2.5"
Exit Function
End If
If channels% <> 0 And KWord% = 0 Then
AOLVersion = "3.0"
Exit Function
End If
If channels% <> 0 And KWord% <> 0 Then
AOLVersion = "4.0"
Exit Function
End If
End Function

Sub Attention(thetext As String)
'G$ = WavYChaT("Surge ")
'L$ = WavYChaT(" by JoLT")
'aa$ = WavYChaT("Attention")
SendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call timeout(0.15)
SendChat (thetext)
Call timeout(0.15)
SendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call timeout(0.15)
'SendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Sub centerform(f As Form)
f.top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub

Sub centertopleft(frm As Form)
frm.top = 0
frm.Left = 0
End Sub

Sub centertopmid(frm As Form)
centerform frm
frm.top = 0
End Sub

Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub CloseAOL()
al% = FindWindow("aol frame25", 0&)
killwin al%
Do
timeout 0.001
Loop Until al% = 0
End Sub

Function ColorChatBlueBlack(thetext As String)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 4
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
CoLoRChaT = p$
End Function

Function ColorChatRedBlue(thetext)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 4
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
ColorChatRedBlue = p$
End Function

Function ColorChatRedGreen(thetext)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 4
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
ColorChatRedGreen = p$

End Function

Sub CountMail()
ao% = FindWindow(0&, "America  Online")
themb% = findchildbytitle(ao%, "" + x + "'s Online Mailbox")
If themb% <> 0 Then GoTo TheCount
runtoolbar 1
WaitForMailToLoad
Do
NoFreeze% = DoEvents()
x = UserSN()
ao% = FindWindow(0&, "America  Online")
themb% = findchildbytitle(ao%, "" + x + "'s Online Mailbox")
tabpage% = FindChildByClass(thmb%, "AOL_TabPage")
Hand% = FindChildByClass(tabpage%, "_AOL_TREE")
timeout 0.001
Loop Until Hand% <> 0
WaitForMailToLoad
TheCount:
room% = FindChatRoom()
Handy% = FindChildByClass(room%, "_AOL_Listbox")
Buffer = sendmessagebynum(Handy%, LB_GETCOUNT, 0, 0)
If Buffer > 1 And Buffer <> 550 Then
MsgBox "You have " & Buffer & " messages in your Mailbox...", 0
End If
If Buffer = 1 Then
MsgBox "You have " & Buffer & " message in your Mailbox...", 0
End If
If Buffer < 1 Then
MsgBox "You have no messages in your Mailbox.", 0
End If
If Buffer = 550 Then
MsgBox "Holy ShiT! Yer BoX Iz FuLL!"
End If
End Sub

Function Decryptor(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    If letter$ = "Œ" Then Leet$ = "a"
    If letter$ = "Æ" Then Leet$ = "b"
    If letter$ = "œ" Then Leet$ = "c"
    If letter$ = "¯" Then Leet$ = "d"
    If letter$ = "–" Then Leet$ = "e"
    If letter$ = "¡" Then Leet$ = "f"
    If letter$ = "§" Then Leet$ = "g"
    If letter$ = "ß" Then Leet$ = "h"
    If letter$ = "Ð" Then Leet$ = "i"
    If letter$ = "¶" Then Leet$ = "j"
    If letter$ = "ø" Then Leet$ = "k"
    If letter$ = "Ø" Then Leet$ = "l"
    If letter$ = "ç" Then Leet$ = "m"
    If letter$ = "¸" Then Leet$ = "n"
    If letter$ = "«" Then Leet$ = "o"
    If letter$ = "´" Then Leet$ = "p"
    If letter$ = "ˆ" Then Leet$ = "q"
    If letter$ = "…" Then Leet$ = "r"
    If letter$ = "š" Then Leet$ = "s"
    If letter$ = "¥" Then Leet$ = "t"
    If letter$ = "¿" Then Leet$ = "u"
    If letter$ = "å" Then Leet$ = "v"
    If letter$ = "æ" Then Leet$ = "w"
    If letter$ = "õ" Then Leet$ = "x"
    If letter$ = "ª" Then Leet$ = "y"
    If letter$ = "Å" Then Leet$ = "z"
    If letter$ = "Þ" Then Leet$ = "A"
    If letter$ = "^" Then Leet$ = "B"
    If letter$ = "Î" Then Leet$ = "C"
    If letter$ = "„" Then Leet$ = "D"
    If letter$ = "—" Then Leet$ = "E"
    If letter$ = "‹" Then Leet$ = "F"
    If letter$ = "†" Then Leet$ = "G"
    If letter$ = "þ" Then Leet$ = "H"
    If letter$ = "·" Then Leet$ = "I"
    If letter$ = "±" Then Leet$ = "J"
    If letter$ = "×" Then Leet$ = "K"
    If letter$ = "©" Then Leet$ = "L"
    If letter$ = "®" Then Leet$ = "M"
    If letter$ = "™" Then Leet$ = "N"
    If letter$ = "¹" Then Leet$ = "O"
    If letter$ = "²" Then Leet$ = "P"
    If letter$ = "³" Then Leet$ = "Q"
    If letter$ = "µ" Then Leet$ = "R"
    If letter$ = "°" Then Leet$ = "S"
    If letter$ = "˜" Then Leet$ = "T"
    If letter$ = "‰" Then Leet$ = "U"
    If letter$ = "¾" Then Leet$ = "V"
    If letter$ = "½" Then Leet$ = "W"
    If letter$ = "¼" Then Leet$ = "X"
    If letter$ = "ƒ" Then Leet$ = "Y"
    If letter$ = "‡" Then Leet$ = "Z"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
Decryptor = Made$
End Function

Sub DoubleClick(Button%)
Dim DoubleClickNow%
DoubleClickNow% = sendmessagebynum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub

Sub EliteTalker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If x = 1 Then Leet$ = "â"
    If x = 2 Then Leet$ = "å"
    If x = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If x = 1 Then Leet$ = "ë"
    If x = 2 Then Leet$ = "ê"
    If x = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If x = 1 Then Leet$ = "ì"
    If x = 2 Then Leet$ = "ï"
    If x = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If x = 1 Then Leet$ = "ô"
    If x = 2 Then Leet$ = "ð"
    If x = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If x = 1 Then Leet$ = "ù"
    If x = 2 Then Leet$ = "û"
    If x = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If x = 1 Then Leet$ = "Å"
    If x = 2 Then Leet$ = "Ä"
    If x = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If x = 1 Then Leet$ = "Ï"
    If x = 2 Then Leet$ = "Î"
    If x = 3 Then Leet$ = "Í"
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
SendChat (Made$)
End Sub

Function EliteText(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If x = 1 Then Leet$ = "â"
    If x = 2 Then Leet$ = "å"
    If x = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "e" Then
    If x = 1 Then Leet$ = "ë"
    If x = 2 Then Leet$ = "ê"
    If x = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If x = 1 Then Leet$ = "ì"
    If x = 2 Then Leet$ = "ï"
    If x = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If x = 1 Then Leet$ = "ô"
    If x = 2 Then Leet$ = "ð"
    If x = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If x = 1 Then Leet$ = "ù"
    If x = 2 Then Leet$ = "û"
    If x = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If x = 1 Then Leet$ = "Å"
    If x = 2 Then Leet$ = "Ä"
    If x = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If x = 1 Then Leet$ = "Ï"
    If x = 2 Then Leet$ = "Î"
    If x = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q

EliteText = Made$

End Function

Function Encrypter(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    If letter$ = "a" Then Leet$ = "Œ"
    If letter$ = "b" Then Leet$ = "Æ"
    If letter$ = "c" Then Leet$ = "œ"
    If letter$ = "d" Then Leet$ = "¯"
    If letter$ = "e" Then Leet$ = "–"
    If letter$ = "f" Then Leet$ = "¡"
    If letter$ = "g" Then Leet$ = "§"
    If letter$ = "h" Then Leet$ = "ß"
    If letter$ = "i" Then Leet$ = "Ð"
    If letter$ = "j" Then Leet$ = "¶"
    If letter$ = "k" Then Leet$ = "ø"
    If letter$ = "l" Then Leet$ = "Ø"
    If letter$ = "m" Then Leet$ = "ç"
    If letter$ = "n" Then Leet$ = "¸"
    If letter$ = "o" Then Leet$ = "«"
    If letter$ = "p" Then Leet$ = "´"
    If letter$ = "q" Then Leet$ = "ˆ"
    If letter$ = "r" Then Leet$ = "…"
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "¥"
    If letter$ = "u" Then Leet$ = "¿"
    If letter$ = "v" Then Leet$ = "å"
    If letter$ = "w" Then Leet$ = "æ"
    If letter$ = "x" Then Leet$ = "õ"
    If letter$ = "y" Then Leet$ = "ª"
    If letter$ = "z" Then Leet$ = "Å"
    If letter$ = "A" Then Leet$ = "Þ"
    If letter$ = "B" Then Leet$ = "^"
    If letter$ = "C" Then Leet$ = "Î"
    If letter$ = "D" Then Leet$ = "„"
    If letter$ = "E" Then Leet$ = "—"
    If letter$ = "F" Then Leet$ = "‹"
    If letter$ = "G" Then Leet$ = "†"
    If letter$ = "H" Then Leet$ = "þ"
    If letter$ = "I" Then Leet$ = "·"
    If letter$ = "J" Then Leet$ = "±"
    If letter$ = "K" Then Leet$ = "×"
    If letter$ = "L" Then Leet$ = "©"
    If letter$ = "M" Then Leet$ = "®"
    If letter$ = "N" Then Leet$ = "™"
    If letter$ = "O" Then Leet$ = "¹"
    If letter$ = "P" Then Leet$ = "²"
    If letter$ = "Q" Then Leet$ = "³"
    If letter$ = "R" Then Leet$ = "µ"
    If letter$ = "S" Then Leet$ = "°"
    If letter$ = "T" Then Leet$ = "˜"
    If letter$ = "U" Then Leet$ = "‰"
    If letter$ = "V" Then Leet$ = "¾"
    If letter$ = "W" Then Leet$ = "½"
    If letter$ = "X" Then Leet$ = "¼"
    If letter$ = "Y" Then Leet$ = "ƒ"
    If letter$ = "Z" Then Leet$ = "‡"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
Encrypter = Made$
End Function

Function Fader(thetext$)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 8
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    V$ = Mid$(g$, w + 4, 1)
    Q$ = Mid$(g$, w + 5, 1)
    x$ = Mid$(g$, w + 6, 1)
    Y$ = Mid$(g$, w + 7, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & T$ & "<FONT COLOR=" & Chr$(34) & "#DCDCDC" & Chr$(34) & ">" & V$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & Q$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & x$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & Y$
Next w
Fader = p$
End Function

Function FindChatRoom()
aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
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

Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function GetTxtRICHCNTL(hWindow As Integer)
aol% = FindWindow("AOL Frame25", 0&)
mofo% = FindChildByClass(aol%, ">Instant Message From: ")
Msg$ = String$(255, 0)
xx% = SendMessageLong(hWindow, WM_ProGGer, 254, Msg$)
Msg$ = Trim$(Msg$)
GetTxtRICHCNTL = Msg$
End Function

Function GetWinText(hWnd As Integer) As String
Dim LengthOfText, Buffer$, GetTheText
LengthOfText = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(LengthOfText)
GetTheText = SendMessagebystring(hWnd, WM_GETTEXT, LengthOfText + 1, Buffer$)
GetWinText = Buffer$
End Function

Sub HideAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
s = ShowWindow(aol%, 0)
End Sub

Sub IMBuddy(Recipiant, message)

aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient")
buddy% = findchildbytitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    keyword ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For l = 1 To 2
    AOIcon% = Getwindow(AOIcon%, 2)
Next l

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = findchildbytitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
x = SendMessagebystring(AOEdit%, WM_SETTEXT, 0, Recipiant)
x = SendMessagebystring(AORich%, WM_SETTEXT, 0, message)

For x = 1 To 9
    AOIcon% = Getwindow(AOIcon%, 2)
Next x

Call timeout(0.01)
ClickIcon (AOIcon%)

timeout 1
End Sub

Sub IMIgnore(TheList As ListBox)
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = findchildbytitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To TheList.ListCount
        If LCase$(TheList.List(findsn)) = LCase$(SNfromIM()) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            x = SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            x = SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub

Sub IMKeyword(Recipiant, message)

aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient")

Call keyword("aol://9293:")

Do: DoEvents
IMWin% = findchildbytitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
c = SendMessagebystring(AOEdit%, WM_SETTEXT, 0, Recipiant)
c = SendMessagebystring(AORich%, WM_SETTEXT, 0, message)

For x = 1 To 9
    AOIcon% = Getwindow(AOIcon%, 2)
Next x

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient")
IMWin% = findchildbytitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then g = SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Sub IMsOff()
Call IMKeyword("$IM_OFF", " ")
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", " ")
End Sub

Function IsUserOnline()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
welcome% = findchildbytitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function

Sub keyword(TheKeyWord As String)
aol% = FindWindow("AOL Frame25", 0&)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = Getwindow(AOIcon%, 2)
Next GetIcon

' ******************************
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below

'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = findchildbytitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

x = SendMessagebystring(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub

Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
aol% = FindWindow("AOL Frame25", 0&)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
x = SendMessage(glyph%, WM_CLOSE, 0, 0)
End Sub

Sub KillModal()
Modal% = FindWindow("_AOL_Modal", 0&)
x = SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub KillWait()

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = Getwindow(AOIcon%, 2)
Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = findchildbytitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

x = SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub

Sub killwin(Windo)
CloseTheMofo = sendmessagebynum(Windo, WM_CLOSE, 0, 0)
End Sub

Function LastLine(ByVal txt As String)
Dim x As Integer
Do
x = x + 1
Loop Until Mid(txt, Len(txt) - x, 1) = Chr(13)
LastLine = Right(txt, x)
End Function

Function MultiColor(thetext$)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 6
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    V$ = Mid$(g$, w + 4, 1)
    Q$ = Mid$(g$, w + 5, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#8A2BE2" & Chr$(34) & ">" & T$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & V$ & "<FONT COLOR=" & Chr$(34) & "#808000" & Chr$(34) & ">" & Q$
Next w
MultiColor = p$
End Function

Function nospaces(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    If letter$ = " " Then Leet$ = ""
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
nospaces = Made$
End Function

Sub PhishPhrases(phraZes As TextBox)
phraZes = RandomNumber(141)
If phraZes = "1" Then
phraZes = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf phraZes = "2" Then
phraZes = "hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
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
phraZes = "Hey, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
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
phraZes = "I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf phraZes = "29" Then
phraZes = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf phraZes = "30" Then
phraZes = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf phraZes = "31" Then
phraZes = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf phraZes = "32" Then
phraZes = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf phraZes = "33" Then
phraZes = "I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
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
phraZes = "geno AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
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
phraZes = "My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf phraZes = "82" Then
phraZes = "I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf phraZes = "83" Then
phraZes = "I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf phraZes = "84" Then
phraZes = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf phraZes = "85" Then
l012A = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "86" Then
phraZes = "Hi, I'm with FCB (Federal Credit Bureau). Due to The type of payment you are using, we need to validate it. Please reply with your credit card number, expiration date, full name, address, and phone number. Thank you. "
ElseIf phraZes = "87" Then
phraZes = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "
ElseIf phraZes = "88" Then
phraZes = "AOLBilling:    geno from the AOL TOS Staff. We have lost your billing info.  Please click respond and send your CC#, Expiration Date, and Billing Zip Code. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf phraZes = "89" Then
phraZes = "this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "90" Then
phraZes = "AOLBilling:    geno from the AOL TOS Staff. We have lost your account info. Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf phraZes = "91" Then
phraZes = "Hello I am from the AOL service department and there has been a minor problem with your accout. I see that we have lost your password information and need it as soon as possible. We cannot access any of your personal info without your password. If you could please type your full log-on password and click the send button with your response. Thank you and Enjoy America online."
ElseIf phraZes = "92" Then
phraZes = "Hello, I am from AOLnet and there has been a network overload on our computers. We have been trying fix this problem but have had no luck. All your account information has been lost. In order for AOL to fix this we need your full First Name and Last, Your phone number, Address, Zip code, Credit Card Number and experation date, as well as your Screen Name(s) and Log-On Password(s). If you fail to reply will result in account termination, thank you :-)"
ElseIf phraZes = "93" Then
phraZes = "Hello I am a member of the America Online Stratus Team. We have been experiancing some problems with your preticular account. It would help us if you gave is your Screen Name(s) and Log-on Password(s) so we can remedy this problem. Thank you AOLST. "
ElseIf phraZes = "94" Then
phraZes = "Hello , AOL Customer! I am sorry to Inform you account information Was not Correctly Processed into our Stratus System. We are now forced to ask you to respond with your personal log on password and you complete billing info (Name, Address, & CC Number Information) . Thank-you for your Time, And Continue To enjoy AOL!"
ElseIf phraZes = "95" Then
phraZes = "I work for AOL.  When you signed on we received an INCORRECT password.  Please respond with the CORRECT password now or face possible account termination for safety reasons, Thank you. =]"
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
phraZes = "Due to our last system crash on 4/5/98 we have lost all infromation concerning your account. Please respond with your Full address and Credit Card Number. NOTE: 60 mins are credited to your account because of the crash also."
ElseIf phraZes = "131" Then
phraZes = "We're sorry to report that your account information has been erased.  For you to stay online we need you to verify your password, Credit Card number, Expiration date, and last 4 digits of your phone number.  Thank you and sorry for the delay."
ElseIf phraZes = "132" Then
phraZes = "My name is Todd Stiger.  I work for the internet.  I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your credit card number and personal information.  Thanks!"
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
End Sub

Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndplaysound(SoundName$, wFlags%)
End Sub

Sub Punt(who$, strin$)
Do
IMBuddy who$, strin$
Loop Until FindChildByClass(findchildbytitle(FindWindow("aol frame25", 0&), "America Online"), "Button")
SendChat "" + who$ + " Has Been Punted!"
End Sub

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function ReadINI(AppName, KeyName, filename As String) As String
'To use this type:
'a$ = ReadINI("Secret Area", "Found", app.Path & "\ftsphshr.ini")
'If a$ = "True" Then msgbox "You already Have The Secret Area"
'If a$ = "False" or a$ = "" Then msgbox "You have yet to find the secret area!"
Dim sRet As String
    sRet = String(255, 0)
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Sub run(ByVal menuname As String)
' This will make any item in the File, Edit, Go To, etc.,
' menus do what you tell them. If the menu has a little
' line under it, then you must put & directly before you
' get to the letter that ahs the line under it. This works
' on both AOL3.0, AOL95, and AOL4.0
'  Example: run "Keyword"
'  Example: run "&Compose Mail"
'  Example: run "Read &New Mail"
'  Example: run "Mem&bers"
Dim mhWnd As Integer
Dim MnuCnt As Integer
Dim Cnt As Integer
Dim SmhWnd As Integer
Dim SMnuCnt As Integer
Dim SCnt As Integer
Dim lint As Integer
Dim lpString As String
aol% = FindWindow%("AOL Frame25", 0&)
mhWnd = GetMenu%(aol%)
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
        ret = sendmessagebynum(aol%, WM_COMMAND, GMnID, 0)
        DoEvents
        Exit Sub
    End If
    DoEvents
    Next SCnt
Next Cnt
End Sub

Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", 0&))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = sendmessagebynum(FindWindow("AOL Frame25", 0&), 273, AOLItemID, 0&)

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

Sub runtoolbar(iconum As Integer)
' This is only 1-6 (numbered by if they are pull down menus
' or not.) 1 is read mail, 2 is compose mail, 3 is print
' something. 4 is quotes. 5 is perks. 6 is weather.
' To use this type RunToolBar # (# is 1-6) AOL4.0 only.
aol% = FindWindow("aol frame25", 0&)
tool% = FindChildByClass(aol%, "AOL ToolBar")
readmail% = FindChildByClass(tool%, "_AOL_Icon")
If iconum = 1 Then iconum2% = readmail%
makemail% = Getwindow(readmail%, GW_HWNDNEXT)
If iconum = 2 Then iconum2% = makemail%
mailcent% = Getwindow(makemail%, GW_HWNDNEXT)
printstuff% = Getwindow(mailcent%, GW_HWNDNEXT)
If iconum = 3 Then iconum2% = printstuff%
myfiles% = Getwindow(printstuff%, GW_HWNDNEXT)
myaol% = Getwindow(myfiles%, GW_HWNDNEXT)
favorites% = Getwindow(myaol%, GW_HWNDNEXT)
internet% = Getwindow(favorites%, GW_HWNDNEXT)
channels% = Getwindow(internet%, GW_HWNDNEXT)
people% = Getwindow(channels%, GW_HWNDNEXT)
quotes% = Getwindow(people%, GW_HWNDNEXT)
If iconum = 4 Then iconum2% = quotes%
perks% = Getwindow(quotes%, GW_HWNDNEXT)
If iconum = 5 Then iconum2% = perks%
weather% = Getwindow(perks%, GW_HWNDNEXT)
If iconum = 6 Then iconum2% = weather%
glyph% = Getwindow(perks%, GW_HWNDNEXT)
back% = Getwindow(glyph%, GW_HWNDNEXT)
fwd% = Getwindow(back%, GW_HWNDNEXT)
closewww% = Getwindow(fwd%, GW_HWNDNEXT)
reload% = Getwindow(closewww%, GW_HWNDNEXT)
home% = Getwindow(reload%, GW_HWNDNEXT)
keywordbutton% = Getwindow(home%, GW_HWNDNEXT)
kwbox% = Getwindow(keywordbutton%, GW_HWNDNEXT)
go% = Getwindow(kwbox%, GW_HWNDNEXT)
If iconum = 7 Then iconum2% = go%
keywordbutton2% = Getwindow(go%, GW_HWNDNEXT)
kwbox2% = Getwindow(keywordbutton2%, GW_HWNDNEXT)
gobutton% = Getwindow(kwbox2%, GW_HWNDNEXT)
gobutton2% = Getwindow(gobutton%, GW_HWNDNEXT)
If iconum = 8 Then iconum2% = gobutton2%
ClickIcon iconum2%
ClickIcon iconum2%
ClickIcon iconum2%
End Sub

Sub Send(chatedit, jolt$)
x = SendMessagebystring(chatedit, WM_SETTEXT, 0, jolt$)
End Sub

Sub SendChat(Chat)
room% = FindChatRoom()
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = Getwindow(AORich%, 2)
AORich% = Getwindow(AORich%, 2)
AORich% = Getwindow(AORich%, 2)
AORich% = Getwindow(AORich%, 2)
AORich% = Getwindow(AORich%, 2)
AORich% = Getwindow(AORich%, 2)

x = SetFocusApi(AORich%)
x = SendMessagebystring(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
x = sendmessagebynum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub sendlater(Recipiants, subject, message)

aol% = FindWindow("AOL Frame25", 0&)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = Getwindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
AOMail% = findchildbytitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

x = SendMessagebystring(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = getnextwindow(AOEdit%, 2)
AOEdit% = getnextwindow(AOEdit%, 2)
AOEdit% = getnextwindow(AOEdit%, 2)
AOEdit% = getnextwindow(AOEdit%, 2)
x = SendMessagebystring(AOEdit%, WM_SETTEXT, 0, subject)
x = SendMessagebystring(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon% = Getwindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = findchildbytitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", 0&)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
AOIco% = Getwindow(AOIcon%, GW_HWNDNEXT)
Ico% = Getwindow(AOIco%, GW_HWNDNEXT)
ClickIcon (Ico%)
x = SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
x = SendMessage(AOError%, WM_CLOSE, 0, 0)
x = SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub sendmail(Recipiants, subject, message)

aol% = FindWindow("AOL Frame25", 0&)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = Getwindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
AOMail% = findchildbytitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

x = SendMessagebystring(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = getnextwindow(AOEdit%, 2)
AOEdit% = getnextwindow(AOEdit%, 2)
AOEdit% = getnextwindow(AOEdit%, 2)
AOEdit% = getnextwindow(AOEdit%, 2)
x = SendMessagebystring(AOEdit%, WM_SETTEXT, 0, subject)
x = SendMessagebystring(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon% = Getwindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = findchildbytitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", 0&)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
x = SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
x = SendMessage(AOError%, WM_CLOSE, 0, 0)
x = SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub ShowAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
s = ShowWindow(aol%, 5)
End Sub

Sub ShutDown(ByVal Xit As Integer)
Select Case Xit
    Case 1
        n = ExitWindows(67, 0) 'reboot the computer
    Case 2
        n = ExitWindows(66, 0) 'restart Windows
    Case 3
        n = ExitWindows(0, 0) 'exit Windows
End Select
End Sub

Sub SignOff()
RunMenu 3, 1
End Sub

Function SlowText(thetext$)
g$ = thetext$
A = Len(g$)
For w = 1 To A Step 1
    r$ = Mid$(g$, w, 1)
    s$ = Mid$(g$, w + 1, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><HTML>" & r$ & "</HTML>" & s$
Next w
SlowText = p$
End Function

Function SNfromIM()

aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(aol%, "MDIClient") '

IM% = findchildbytitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = findchildbytitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function

Sub stayontop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Function Text_Backwards(strin$)
Let inptxt$ = strin$
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
Text_Backwards = newsent$
End Function

Sub Text_Manipulation(who$, wut$)
aol% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(aol%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "RICHCNTL")
sndtext% = SendMessagebystring(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (who$) & ":" & Chr(9) & (wut$))
SendNow% = sendmessagebynum(ChatVew%, WM_CHAR, &HD, 0)
End Sub

Sub timeout(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub

Sub ToolBarIMPart1(ByVal menuname As String)
aol% = FindWindow("aol frame25", 0&)
tool% = FindChildByClass(aol%, "AOL ToolBar")
TheMail% = FindChildByClass(tool%, "_AOL_Icon")
WriteMail% = getnextwindow(TheMail%, 2)
MailCenter% = getnextwindow(WriteMail%, 2)
A% = getnextwindow(WriteMail%, 2)
b% = getnextwindow(MailCenter%, 2)
c% = getnextwindow(b%, 2)
d% = getnextwindow(c%, 2)
e% = getnextwindow(d%, 2)
f% = getnextwindow(e%, 2)
g% = getnextwindow(f%, 2)
h% = getnextwindow(g%, 2)
ClickIcon h%
i% = FindChildByClass(h%, "#32768")
j% = SetFocusApi(i%)
DoubleClick j%
End Sub

Sub ToolBarIMPart2()
ToolBarIMPart1 "&Instant Message"
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

Function trm(ByVal txt As String)
Dim x As Integer
Dim Y As String
Dim z As String
For x = 1 To Len(txt)
Y = Mid(txt, x, 1)
If Y = Chr(0) Then Y = ""
z = z & Y
Next x
trm = z
End Function

Sub UnUpchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
x = EnableWindow(Upp%, 1)
x = EnableWindow(aol%, 0)
End Sub

Sub Upchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
x = EnableWindow(aol%, 1)
x = EnableWindow(Upp%, 0)
End Sub

Function UserSN()
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(aol%, "MDIClient")
welcome% = findchildbytitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub WaitForMailToLoad()
x = UserSN()
Do
box = findchildbytitle(FindWindow("AOL Frame25", "America  Online"), x + "'s Online Mailbox")
timeout (0.1)
Loop Until box <> 0
List = FindChildByClass(box, "_AOL_Tree")
Do
DoEvents
mailnum = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call timeout(0.5)
mailnum2 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call timeout(0.5)
mailnum3 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
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
   
   Do
    okb = findchildbytitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)
   Loop Until okb = 0

End Sub

Function Wavy(thetext As String)

g$ = thetext
A = Len(g$)
For w = 1 To A Step 4
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & s$ & "</sub>" & T$
Next w
Wavy = p$

End Function

Sub WavYChaTBluEBlacK(thetext)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 4
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
SendChat (p$)
End Sub

Function WavyChatRedBlue(thetext As String)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 4
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavyChatRedBlue = p$
End Function

Function WavyChatRedGreen(thetext As String)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 4
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
WavyChatRedGreen = p$
End Function

Function WavyMultiColor(thetext$)
g$ = thetext
A = Len(g$)
For w = 1 To A Step 6
    r$ = Mid$(g$, w, 1)
    U$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    T$ = Mid$(g$, w + 3, 1)
    V$ = Mid$(g$, w + 4, 1)
    Q$ = Mid$(g$, w + 5, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sup>" & r$ & "</sup><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub><FONT COLOR=" & Chr$(34) & "#8A2BE2" & Chr$(34) & "><sup>" & T$ & "</sup><FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & "><sub>" & V$ & "</sub><FONT COLOR=" & Chr$(34) & "#808000" & Chr$(34) & ">" & Q$
Next w
WavyMultiColor = p$
End Function

Function WinTxt(ByVal hWnd As Integer)
Dim x As Integer
Dim Y As String
Dim z As Integer
x = SendMessage(hWnd, &HE, 0&, 0&)
Y = String(x + 1, " ")
z = SendMessagebystring(hWnd, &HD, x + 1, Y)
WinTxt = Left(Y, x)
End Function

Sub WriteINI(sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("AORacket", "Secret Area", "Found", app.path + "\aorocket.ini")
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub

