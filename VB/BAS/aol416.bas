Declare Sub releaseCapture Lib "User" ()
Declare Function Sendmessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParem As Integer, lparem As Any) As Long

Global Const SWP_NoActivate = &H10

Declare Sub Drawmenubar Lib "User" (ByVal hWnd As Integer)
Declare Function AOLGetList% Lib "311.Dll" (ByVal Index%, ByVal Buf$)
Declare Function TrackPopupMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nReserved As Integer, ByVal hWnd As Integer, lpReserved As Any) As Integer
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function DeleteMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer

Global Const MF_Popup = &H10
Global Const MF_String = &H0
Global Const MF_Enabled = &H0
Global Const MF_Separator = &H800
Global Const TPM_LEFTALIGN = &H0
Global Const TPM_RIGHTBUTTON = &H2





Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function ExitWindows Lib "User" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Global Const LB_GETTEXT = &H189
Global Const WM_GETDLGCODE = &H87
Global Const WM_CTLCOLOR = &H19
Declare Function showwindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Global Const LBN_DBLCLK = 2
Const c003E = 16 ' &H10%
Global Const WM_USER = &H400
Global Const LB_GETSEL = (WM_USER + 8)
Declare Function sndPlaySound Lib "MMSystem" (ByVal lpWavName$, ByVal FLAGS%) As Integer '
Global Const WM_MOVE = &H3
Global Const SWP_NOREPOSITION = &H200
Declare Sub Movewindow Lib "User" (ByVal hWnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Global Const WM_Close = &H10
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const WM_ACTIVATE = &H6

Global Const NUMBOXES = 5
Global Const SAVEFILE = 1, LOADFILE = 2
Global Const REPLACEFILE = 1, READFILE = 2, ADDTOFILE = 3
Global Const RANDOMFILE = 4, BINARYFILE = 5

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
Global Const SW_Hide = 0
Global Const SW_SHOWNORMAL = 1
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

Type Rect
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
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
  y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type

Declare Function findwindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer

Declare Function InsertMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer




Declare Function extfn2668 Lib "vbwfind.dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn16A8 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn1868 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function SetWindowPos Lib "User" (ByVal H%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function extfn2630 Lib "vbwfind.dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1AD0 Lib "User" Alias "SetWindowText" (ByVal p1%, ByVal p2$) As Integer

Declare Function FlashWindow Lib "User" (ByVal hWnd As Integer, ByVal bInvert As Integer) As Integer
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)

Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal y As Integer)

Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource%) As Integer
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function ExitWindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)
Declare Function SetParent% Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer)
Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)

Declare Function CreateMenu% Lib "User" ()

Declare Function AppendMenuByString% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)

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
Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wID%, ByVal wFlags%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDesktopWindow Lib "User" () As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lParam&)
Declare Function SendMessageLong Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal hMsg As Integer, ByVal wParam As Integer, ByVal lParam As Any) As Long

Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FileName As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal FileName As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFilename$) As Integer
Declare Function agGetStringFromLPSTR$ Lib "APIGuide.dll" (ByVal lpString&)

Declare Sub SetBKColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer

Declare Function MciSendString& Lib "MMSystem" (ByVal cmd$, ByVal Returnstr As Any, ByVal returnlen%, ByVal hcallback%)

Declare Function FindChild% Lib "vbwfind.dll" (ByVal hWnd%, ByVal Title$)

Declare Sub agCopyData Lib "APIGuide.dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "APIGuide.dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Sub agDWordTo2Integers Lib "APIGuide.dll" (ByVal L&, lw%, lh%)
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
Declare Function agVBSendControlMsg& Lib "APIGuide.dll" (ctl As Control, ByVal Msg%, ByVal wp%, ByVal lp&)
Declare Function agVBSetControlFlags& Lib "APIGuide.dll" (ctl As Control, ByVal mask&, ByVal Value&)
Declare Function dwVBSetControlFlags& Lib "APIGuide.dll" (ctl As Control, ByVal mask&, ByVal Value&)


Declare Sub ptGetTypeFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptCopyTypeToAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptSetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL)
Declare Function ptGetVariableAddress Lib "VBMsg.Vbx" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (ByVal S As String) As Long
Declare Function ptGetLongAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (L As Long) As Long
Declare Function ptGetIntegerAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (i As Integer) As Long
Declare Function ptGetIntegerFromAddress Lib "VBMsg.Vbx" (ByVal i As Long) As Integer
Declare Function ptGetLongFromAddress Lib "VBMsg.Vbx" (ByVal L As Long) As Long
Declare Function ptGetStringFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, ByVal cbBytes As Integer) As String
Declare Function ptMakelParam Lib "VBMsg.Vbx" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Declare Function ptLoWord Lib "VBMsg.Vbx" (ByVal lParam As Long) As Integer
Declare Function ptHiWord Lib "VBMsg.Vbx" (ByVal lParam As Long) As Integer
Declare Function ptMakeUShort Lib "VBMsg.Vbx" (ByVal LongVal As Long) As Integer
Declare Function ptConvertUShort Lib "VBMsg.Vbx" (ByVal ushortVal As Integer) As Long
Declare Function ptMessagetoText Lib "VBMsg.Vbx" (ByVal uMsgID As Long, ByVal bFlag As Integer) As String
Declare Function ptRecreateControlHwnd Lib "VBMsg.Vbx" (ctl As Control) As Long
Declare Function ptGetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL) As Long
Declare Function ptGetControlName Lib "VBMsg.Vbx" (ctl As Control) As String

Declare Function VarPtr& Lib "VBRun300.Dll" (Param As Any)
Declare Function vbeNumChildWindow% Lib "VBStr.Dll" (ByVal win%, ByVal iNum%)

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

Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4

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

Global Const GMEM_DISCARDED = &H4000
Global Const GMEM_LOCKCOUNT = &HFF

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

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

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

Global Const ERR_BAD_HINSTANCE = &H6020
Global Const ERR_BAD_HMODULE = &H6021
Global Const ERR_BAD_GLOBAL_HANDLE = &H6022
Global Const ERR_BAD_LOCAL_HANDLE = &H6023
Global Const ERR_BAD_ATOM = &H6024
Global Const ERR_BAD_HFILE = &H6025

Global Const ERR_BAD_HWND = &H6040
Global Const ERR_BAD_HMENU = &H6041
Global Const ERR_BAD_HCURSOR = &H6042
Global Const ERR_BAD_HICON = &H6043
Global Const ERR_BAD_HDWP = &H6044
Global Const ERR_BAD_CID = &H6045
Global Const ERR_BAD_HDRVR = &H6046

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

Global Const ERR_GALLOC = &H1
Global Const ERR_GREALLOC = &H2
Global Const ERR_GLOCK = &H3
Global Const ERR_LALLOC = &H4
Global Const ERR_LREALLOC = &H5
Global Const ERR_LLOCK = &H6
Global Const ERR_ALLOCRES = &H7
Global Const ERR_LOCKRES = &H8
Global Const ERR_LOADMODULE = &H9

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

Global Const ERR_CREATEDC = &H80
Global Const ERR_CREATEMETA = &H81
Global Const ERR_DELOBJSELECTED = &H82
Global Const ERR_SELBITMAP = &H83

Global Const EW_RESTARTWindow = &H42
Global Const EW_REBOOTSYSTEM = &H43

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

Global Const BLACKONWHITE = 1
Global Const WHITEONBLACK = 2
Global Const COLORONCOLOR = 3

Global Const ALTERNATE = 1
Global Const WINDING = 2

Global Const TA_NOUPDATECP = 0
Global Const TA_UPDATECP = 1
Global Const TA_LEFT = 0
Global Const TA_RIGHT = 2
Global Const TA_CENTER = 6
Global Const TA_TOP = 0
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24

Global Const ETO_GRAYED = 1
Global Const ETO_OPAQUE = 2
Global Const ETO_CLIPPED = 4

Global Const ASPECT_FILTERING = &H1

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

Global Const SP_NOTREPORTED = &H4000
Global Const SP_ERROR = (-1)
Global Const SP_APPABORT = (-2)
Global Const SP_USERABORT = (-3)
Global Const SP_OUTOFDISK = (-4)
Global Const SP_OUTOFMEMORY = (-5)
Global Const PR_JOBSTATUS = &H0

Global Const BI_RGB = 0&
Global Const BI_RLE8 = 1&
Global Const BI_RLE4 = 2&

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


Global Const FF_DONTCARE = 0
Global Const FF_ROMAN = 16

Global Const FF_SWISS = 32

Global Const FF_MODERN = 48

Global Const FF_SCRIPT = 64
Global Const FF_DECORATIVE = 80

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

'Global Const TRANSPARENT = 1
Global Const OPAQUE = 2

Global Const MM_TEXT = 1
Global Const MM_LOMETRIC = 2
Global Const MM_HIMETRIC = 3
Global Const MM_LOENGLISH = 4
Global Const MM_HIENGLISH = 5
Global Const MM_TWIPS = 6
Global Const MM_ISOTROPIC = 7
Global Const MM_ANISOTROPIC = 8

Global Const ABSOLUTE = 1
Global Const RELATIVE = 2

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

Global Const BS_SOLID = 0
Global Const BS_NULL = 1
Global Const BS_HOLLOW = BS_NULL
Global Const BS_HATCHED = 2
Global Const BS_PATTERN = 3
Global Const BS_INDEXED = 4
Global Const BS_DIBPATTERN = 5

Global Const HS_HORIZONTAL = 0
Global Const HS_VERTICAL = 1
Global Const HS_FDIAGONAL = 2
Global Const HS_BDIAGONAL = 3
Global Const HS_CROSS = 4
Global Const HS_DIAGCROSS = 5

Global Const PS_SOLID = 0
Global Const PS_DASH = 1
Global Const PS_DOT = 2
Global Const PS_DASHDOT = 3
Global Const PS_DASHDOTDOT = 4
Global Const PS_NULL = 5
Global Const PS_INSIDEFRAME = 6

Global Const DCB_RESET = 1
Global Const DCB_ACCUMULATE = 2
Global Const DCB_DIRTY = 2
Global Const DCB_SET = 3
Global Const DCB_ENABLE = 4
Global Const DCB_DISABLE = 8

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
Declare Function SetFocusAPI% Lib "User" Alias "SetFocus" (ByVal hWnd As Integer)
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function test Lib "APIGuide.dll" Alias "AgGetStringFromLPSTR" (ByVal p1&) As String


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
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function getwindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer

Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function sendmessagebystring& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function DeleteBubble% Lib "bubble.dll" (ByVal wnd%)
Declare Function findchildbytitle% Lib "vbwfind.dll" (ByVal parent%, ByVal Title$)
Declare Function findchildbyclass% Lib "vbwfind.dll" (ByVal parent%, ByVal Title$)

Declare Function getparent Lib "User" (ByVal hWnd As Integer) As Integer

Const POPUPMENU_LEFTALIGN = 0
Const POPUPMENU_CENTERALIGN = 4
Const POPUPMENU_RIGHTALIGN = 8



Const DB_OPTIONINIPATH = 1
Const PI = 3.141592654

Global Const GWW_HINSTANCE = (-6)
Global Const RDW_INVALIDATE = &H1
Global Const RDW_ERASE = &H4
Global Const RDW_ALLCHILDREN = &H80
Global Const COLOR_BACKGROUND = 1
Global Const COLOR_ACTIVECAPTION = 2


 
 
Declare Function GetModuleHandle% Lib "Kernel" (ByVal lpModuleName$)
Declare Function LoadString% Lib "User" (ByVal hInstance%, ByVal wID%, ByVal lpBuffer$, ByVal nBufferMax%)


Declare Function DeleteObject% Lib "GDI" (ByVal hObject%)

Declare Sub GetWindowRect Lib "User" (ByVal hWnd%, lpRect As Rect)
Declare Sub InflateRect Lib "User" (lpRect As Rect, ByVal x%, ByVal y%)
 
Declare Function CreateRectRgnIndirect% Lib "GDI" (lpRect As Rect)
Declare Function RedrawWindow% Lib "User" (ByVal hWnd%, lprcUpdate As Rect, ByVal hrgnUpdate%, ByVal fuRedraw%)
Declare Function FrameRgn% Lib "GDI" (ByVal hDC%, ByVal hRgn%, ByVal hBrush%, ByVal nWidth%, ByVal nHeight%)
Declare Function GetSysColor& Lib "User" (ByVal nIndex%)
'Declare Function Rectangle% Lib "GDI" (ByVal hDC%, ByVal X1%, ByVal Y1%, ByVal X2%, ByVal Y2%)



Declare Function GetCurrentTask% Lib "Kernel" ()
Declare Function GetModuleFileName% Lib "Kernel" (ByVal hModule%, ByVal lpFilename$, ByVal nSize%)

Declare Function ExtractIcon% Lib "Shell" (ByVal hInst%, ByVal FileName$, ByVal iIcon%)
Declare Function DestroyIcon% Lib "User" (ByVal hIcon%)
Declare Function GlobalSize& Lib "Kernel" (ByVal hGlobal%)
Declare Function GlobalLock& Lib "Kernel" (ByVal hGlobal%)
Declare Function GlobalUnlock% Lib "Kernel" (ByVal hGlobal%)
Declare Sub hmemcpy Lib "Kernel" (ByVal hpDest&, ByVal hpSource&, ByVal cbCopy&)


Declare Function IsWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function WinExec Lib "Kernel" (ByVal lpCmdLine As String, ByVal nCmdShow As Integer) As Integer

Declare Function LockWindowUpdate Lib "User" (ByVal hwndLock As Integer) As Integer


Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Integer
Const SWP_NOZORDER = &H4

Sub AOL4Color_BlueBlack (thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "" & R$ & "" & "" & u$ & "" & S$ & "" & "" & t$
Next W
chattext p$
End Sub

Sub AOLim (Who, Say)
'Example:
'AolIM "Progee","Hi!"
AA% = findwindow("Aol Frame25", 0&)
A% = findchildbytitle(AA%, "Buddy List Window")
B% = findchildbyclass(A%, "_Aol_Icon")
If A% = 0 Then KW "Buddy View"
Do
A% = findchildbytitle(AA%, "Buddy List Window")
B% = findchildbyclass(A%, "_Aol_Icon")
Call timeout(.001)
Loop Until A% <> 0
C% = getwindow(B%, GW_HWNDNEXT)
D% = getwindow(C%, GW_HWNDNEXT)
E% = getwindow(D%, GW_HWNDNEXT)
click D%
Do
z% = findchildbytitle(AA%, "Send Instant Message")
y% = findchildbyclass(z%, "_Aol_Edit")
Call timeout(.001)
Loop Until y% <> 0
W% = getwindow(y%, GW_HWNDNEXT)
K% = getwindow(W%, GW_HWNDNEXT)
LO% = findchildbyclass(z%, "RICHCNTL")
LON% = findchildbyclass(z%, "_Aol_Icon")
GF% = getwindow(LON%, GW_HWNDNEXT)
POL% = getwindow(GF%, GW_HWNDNEXT)
SAD% = getwindow(POL%, GW_HWNDNEXT)
mad% = getwindow(SAD%, GW_HWNDNEXT)
HAS% = getwindow(mad%, GW_HWNDNEXT)
WAS% = getwindow(HAS%, GW_HWNDNEXT)
MIG% = getwindow(WAS%, GW_HWNDNEXT)
WQA% = getwindow(MIG%, GW_HWNDNEXT)
LAB% = getwindow(WQA%, GW_HWNDNEXT)
H% = sendmessagebystring(y%, WM_SETTEXT, 0, Who)
i% = sendmessagebystring(LO%, WM_SETTEXT, 0, Say)
click LAB%
SendKeys "^{enter}"
End Sub

Function AppIcon2Pic% (Pic As PictureBox)
Dim hIcon%
Dim Rc%
Dim hInst%
hInst% = GetWindowWord%(Pic.hWnd, GWW_HINSTANCE)
hIcon% = ExtractIcon%(hInst%, ExeName$(hInst%), 0)
If hIcon% Then
AppIcon2Pic% = CopyIcon%(hIcon%, (Pic.Picture))
Rc% = DestroyIcon%(hIcon%)
End If
End Function

Sub CenterDialog (WinText As String, FMThwnd As Integer)
Dim lpDlgRect As Rect
Dim lpDskRect As Rect
Do
FMThwnd = findwindow(0&, WinText)
If FMThwnd Then Exit Do
x% = DoEvents()
Call timeout(.001)
Loop
Call GetWindowRect(FMThwnd, lpDlgRect)
wdth% = lpDlgRect.Right - lpDlgRect.Left
hght% = lpDlgRect.Bottom - lpDlgRect.Top
Call GetWindowRect(GetDesktopWindow(), lpDskRect)
Scrwdth% = lpDskRect.Right - lpDskRect.Left
Scrhght% = lpDskRect.Bottom - lpDskRect.Top
x% = (Scrwdth% - wdth%) \ 2
y% = (Scrhght% - hght%) \ 2
ZZ% = SetWindowPos(FMThwnd, 0, x%, y%, 0, 0, SWP_NOZORDER Or SWP_NOSIZE)
End Sub

Sub Centerform (FRM As Form)
'CenterForm me
x = Screen.Width / 2 - FRM.Width / 2
y = Screen.Height / 2 - FRM.Height / 2
FRM.Move x, y
End Sub

Sub chattext (Say)
DoEvents
A% = findwindow("Aol Frame25", 0&)
B% = findchildbyclass(A%, "_Aol_Listbox")
C% = getparent(B%)
D% = findchildbyclass(C%, "RICHCNTL")
E% = getparent(D%)
f% = findchildbyclass(E%, "_Aol_Static")
G% = getparent(f%)
H% = findchildbyclass(E%, "RICHCNTL")
ChatEdit2% = getwindow(H%, GW_HWNDNEXT)
Chatedi% = getwindow(ChatEdit2%, GW_HWNDNEXT)
Chated% = getwindow(Chatedi%, GW_HWNDNEXT)
Chate% = getwindow(Chated%, GW_HWNDNEXT)
Chat% = getwindow(Chate%, GW_HWNDNEXT)
chatedit% = getwindow(Chat%, GW_HWNDNEXT)
'sndtext1% = sendmessagebystring(chatedit%, WM_SETTEXT, 0, "")
sndtext% = sendmessagebystring(chatedit%, WM_SETTEXT, 0, Say)
SendNow% = SendMessageByNum(chatedit%, WM_CHAR, &HD, 0)
DoEvents
Call timeout(.35)
End Sub

Sub click (button%)
SendNow% = SendMessageByNum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(button%, WM_LBUTTONUP, &HD, 0)
End Sub

Sub ClickStart ()
'This is to click the startbutton
A% = findwindow("Shell_TrayWnd", 0&)
B% = findchildbyclass(A%, "Button")
SendNow% = SendMessageByNum(B%, WM_LBUTTONDOWN, &HD, 0)
End Sub

Function CopyIcon% (hSource%, hDest%)
Dim sizeSource&, sizeDest&
Dim fpSource&, fpDest&
Dim Rc%
CopyIcon% = False
sizeSource& = GlobalSize&(hSource%)
sizeDest& = GlobalSize&(hDest%)
If sizeDest& <> sizeSource& Then
If sizeSource& <> 288 Then
Exit Function
End If
End If
fpSource& = GlobalLock&(hSource%)
fpDest& = GlobalLock&(hDest%)
hmemcpy fpDest&, fpSource&, sizeSource&
Rc% = GlobalUnlock%(hDest)
Rc% = GlobalUnlock%(hSource)
CopyIcon% = True
End Function

Sub Dblclick (B%)
A% = SendMessageByNum(B%, WM_LBUTTONDBLCLK, &H203, 0)
End Sub

Sub Dictionary (whatwerd)
'this looks up a word in the dictionary
'Dictionary "Whatever"
Run "&Dictionary"
A% = findwindow("Aol Frame25", 0&)
Do
B% = findchildbytitle(A%, "Merriam-Webster Dictionary")
C% = findchildbyclass(B%, "_AOl_Edit")
Call timeout(.001)
Loop Until C% <> 0
cc% = findchildbyclass(B%, "_AOl_Icon")
D% = sendmessagebystring(C%, WM_SETTEXT, 0, whatwerd)
Call timeout(.2)
click cc%
End Sub

Sub Editprofile (Pname, PCity, PBDay, PMart, PHobbie, PComp, Pocc, PPer)
'Example:
'EditProfile "Name","City","Bday","Single","Hobbie","Computer","Occupation","Personal Quote"
KW "Member Directory"
A% = findwindow("Aol Frame25", 0&)
Do
B% = findchildbytitle(A%, "Member Directory")
C% = findchildbyclass(B%, "_Aol_Icon")
Call timeout(.001)
Loop Until B% <> 0
click C%
Do
D% = findchildbytitle(A%, "Edit Your Online Profile")
E% = findchildbyclass(D%, "_Aol_Edit")
G% = getwindow(E%, GW_HWNDNEXT)
H% = getwindow(G%, GW_HWNDNEXT)
i% = getwindow(H%, GW_HWNDNEXT)
M% = getwindow(i%, GW_HWNDNEXT)
N% = getwindow(M%, GW_HWNDNEXT)
o% = getwindow(N%, GW_HWNDNEXT)
p% = getwindow(o%, GW_HWNDNEXT)
R% = getwindow(p%, GW_HWNDNEXT)
t% = getwindow(R%, GW_HWNDNEXT)
v% = getwindow(t%, GW_HWNDNEXT)
Call timeout(.001)
Loop Until E% And G% And o% And t% And v% <> 0
f% = sendmessagebystring(E%, WM_SETTEXT, 0, Pname)
J% = sendmessagebystring(G%, WM_SETTEXT, 0, PCity)
K% = sendmessagebystring(H%, WM_SETTEXT, 0, PBDay)
L% = sendmessagebystring(o%, WM_SETTEXT, 0, PMart)
Q% = sendmessagebystring(p%, WM_SETTEXT, 0, PHobbie)
S% = sendmessagebystring(R%, WM_SETTEXT, 0, PComp)
u% = sendmessagebystring(t%, WM_SETTEXT, 0, Pocc)
W% = sendmessagebystring(v%, WM_SETTEXT, 0, PPer)
y% = findchildbyclass(D%, "_Aol_Icon")
Call timeout(.2)
click y%
Call timeout(.2)
Killok
Killwin (B%)
End Sub

Function ExeName$ (hInst%)
Dim Temp$
Dim NameLen%
Temp$ = String(255, Chr$(0))
NameLen% = GetModuleFileName%(hInst%, Temp$, Len(Temp$))
If NameLen% Then
ExeName$ = Left$(Temp$, NameLen%)
Else
ExeName$ = "<Unknown>"
End If

End Function

Sub FakeOH (txt, TXT2)
lonh = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(lonh, D)
chattext (txt & C$ & TXT2)
End Sub

Sub File_Kill (File$)
On Error GoTo manthatsux121
Kill File$
Exit Sub
manthatsux121:
Exit Sub
End Sub

Sub File_Rename (File$, File2$)
On Error GoTo oofosfos
Name File$ As File2$
Exit Sub
oofosfos:
Exit Sub
End Sub

Function FilesBirth (File)
FilesBirth = FileDateTime(File)
End Function

Function FilesSize (File)
'It gets size in Bytes
FilesSize = FileLen(File)
End Function

Function FindAOLChildByTitle (TitleText As String) As Integer
Dim x%
Dim ChildWnd As Integer
Dim MDIhWnd%
Dim AOLChildhWnd%
Dim RetClsName As String * 255
  
MDIhWnd% = getwindow(findwindow("AOL Frame25", 0&), GW_CHILD)
Do
  x% = GetClassName(MDIhWnd%, RetClsName$, 254)
  If InStr(RetClsName$, "MDIClient") Then AOLChildhWnd% = MDIhWnd%
  MDIhWnd% = getwindow(MDIhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While MDIhWnd% <> 0
If TitleText = "MDIClient" Then FindAOLChildByTitle = AOLChildhWnd%
ChildWnd = getwindow(AOLChildhWnd%, GW_CHILD)
Do
  If InStr(WindowCaption(ChildWnd), TitleText) <> 0 Then
      FindAOLChildByTitle = ChildWnd
      Exit Do
  End If
  ChildWnd = getwindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
End Function

Sub FormatDisk (f As Form)
Dim WFflag As Integer, FMThwnd As Integer
formhwnd = f.hWnd
FMThwnd = findwindow(0&, "Format - 3½ Floppy (A:)")
If FMThwnd > 0 Then 'format dialog already open
x% = SetActiveWindow(FMThwnd)
Exit Sub
End If
WFflag = False
FMhWnd = findwindow("WFS_Frame", 0&)
If FMhWnd = 0 Then
i% = WinExec("Winfile", 0)
FMhWnd = findwindow("WFS_Frame", 0&)
If FMhWnd = 0 Then
MsgBox "Can't find the File Manager.", 48, "Warning"
Exit Sub
End If
WFflag = True
End If
x% = LockWindowUpdate(GetDesktopWindow())
x% = PostMessage(FMhWnd, WM_COMMAND, &HCB, 0)
'Call CenterDialog("Format Disk", FMThwnd)
x% = LockWindowUpdate(0)
While IsWindow(FMThwnd)
x% = DoEvents()
If IsWindow(formhwnd) = 0 Then x% = PostMessage(FMThwnd, WM_Close, 0, 0)
Wend
If WFflag Then x% = PostMessage(FMhWnd, WM_Close, 0, 0)
If IsWindow(formhwnd) = 0 Then
End
Else
x% = SetActiveWindow(formhwnd)
End If
End Sub

Sub FRM_CloseLeft (FRM As Form)
FRM.Enabled = False
starto:
If FRM.Left > 0 - FRM.Width Then FRM.Left = FRM.Left - 100 Else GoTo itsawrap
GoTo starto
itsawrap:
Unload FRM
Exit Sub
End Sub

Sub FRM_CloseRight (FRM As Form)
FRM.Enabled = False
starto2:
If FRM.Left < Screen.Width + FRM.Width Then FRM.Left = FRM.Left + 100 Else GoTo itsawrap2
GoTo starto2
itsawrap2:
Unload FRM
Exit Sub
End Sub

Function GetLast (ByVal txt As String)
Dim x As Integer
Do
x = x + 1
Loop Until Mid(txt, Len(txt) - x, 1) = Chr(13)
GetLast = Right(txt, x)
End Function

Sub getnum (og%, A)
Do
If A = 0 Then Exit Sub
B = 1 + B
og% = getwindow(og%, GW_HWNDNEXT)
Loop Until B >= A - 1
End Sub

Function GetWinText (hWnd As Integer) As String
lentos = Sendmessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
x = sendmessagebystring(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
GetWinText = Buffer$
End Function

Sub IMOFF ()
AA% = findwindow("Aol Frame25", 0&)
A% = findchildbytitle(AA%, "Buddy List Window")
B% = findchildbyclass(A%, "_Aol_Icon")
If A% = 0 Then KW "Buddy View"
Do
A% = findchildbytitle(AA%, "Buddy List Window")
B% = findchildbyclass(A%, "_Aol_Icon")
Call timeout(.001)
Loop Until A% <> 0
C% = getwindow(B%, GW_HWNDNEXT)
D% = getwindow(C%, GW_HWNDNEXT)
E% = getwindow(D%, GW_HWNDNEXT)
click D%
Do
z% = findchildbytitle(AA%, "Send Instant Message")
y% = findchildbyclass(z%, "_Aol_Edit")
Call timeout(.001)
Loop Until y% <> 0
W% = getwindow(y%, GW_HWNDNEXT)
K% = getwindow(W%, GW_HWNDNEXT)
LO% = findchildbyclass(z%, "RICHCNTL")
LON% = findchildbyclass(z%, "_Aol_Icon")
GF% = getwindow(LON%, GW_HWNDNEXT)
POL% = getwindow(GF%, GW_HWNDNEXT)
SAD% = getwindow(POL%, GW_HWNDNEXT)
mad% = getwindow(SAD%, GW_HWNDNEXT)
HAS% = getwindow(mad%, GW_HWNDNEXT)
WAS% = getwindow(HAS%, GW_HWNDNEXT)
MIG% = getwindow(WAS%, GW_HWNDNEXT)
WQA% = getwindow(MIG%, GW_HWNDNEXT)
LAB% = getwindow(WQA%, GW_HWNDNEXT)
H% = sendmessagebystring(y%, WM_SETTEXT, 0, "$IM_OFF")
i% = sendmessagebystring(LO%, WM_SETTEXT, 0, "$IM_OFF")
click LAB%
Killok
Killwin (z%)
End Sub

Sub IMON ()
AA% = findwindow("Aol Frame25", 0&)
A% = findchildbytitle(AA%, "Buddy List Window")
B% = findchildbyclass(A%, "_Aol_Icon")
If A% = 0 Then KW "Buddy View"
Do
A% = findchildbytitle(AA%, "Buddy List Window")
B% = findchildbyclass(A%, "_Aol_Icon")
Call timeout(.001)
Loop Until A% <> 0
C% = getwindow(B%, GW_HWNDNEXT)
D% = getwindow(C%, GW_HWNDNEXT)
E% = getwindow(D%, GW_HWNDNEXT)
click D%
Do
z% = findchildbytitle(AA%, "Send Instant Message")
y% = findchildbyclass(z%, "_Aol_Edit")
Call timeout(.001)
Loop Until y% <> 0
W% = getwindow(y%, GW_HWNDNEXT)
K% = getwindow(W%, GW_HWNDNEXT)
LO% = findchildbyclass(z%, "RICHCNTL")
LON% = findchildbyclass(z%, "_Aol_Icon")
GF% = getwindow(LON%, GW_HWNDNEXT)
POL% = getwindow(GF%, GW_HWNDNEXT)
SAD% = getwindow(POL%, GW_HWNDNEXT)
mad% = getwindow(SAD%, GW_HWNDNEXT)
HAS% = getwindow(mad%, GW_HWNDNEXT)
WAS% = getwindow(HAS%, GW_HWNDNEXT)
MIG% = getwindow(WAS%, GW_HWNDNEXT)
WQA% = getwindow(MIG%, GW_HWNDNEXT)
LAB% = getwindow(WQA%, GW_HWNDNEXT)
H% = sendmessagebystring(y%, WM_SETTEXT, 0, "$IM_ON")
i% = sendmessagebystring(LO%, WM_SETTEXT, 0, "$IM_ON")
click LAB%
Killok
Killwin (z%)
End Sub

Sub IMRespond (what)
A% = findwindow("Aol Frame25", 0&)
B% = findchildbytitle(A%, ">Instant Message From:")
ZZ% = findchildbytitle(A%, "Instant Message From:")
C% = findchildbyclass(B%, "_Aol_Icon")
cc% = findchildbyclass(B%, "RICHCNTL")
DD% = getwindow(cc%, GW_HWNDNEXT)
EE% = getwindow(DD%, GW_HWNDNEXT)
ff% = getwindow(EE%, GW_HWNDNEXT)
GG% = getwindow(ff%, GW_HWNDNEXT)
HH% = getwindow(GG%, GW_HWNDNEXT)
II% = getwindow(HH%, GW_HWNDNEXT)
JJ% = getwindow(II%, GW_HWNDNEXT)
KK% = getwindow(JJ%, GW_HWNDNEXT)
LL% = getwindow(KK%, GW_HWNDNEXT)
MM% = getwindow(LL%, GW_HWNDNEXT)
NN% = getwindow(MM%, GW_HWNDNEXT)
OO% = getwindow(NN%, GW_HWNDNEXT)
PP% = getwindow(OO%, GW_HWNDNEXT)
D% = getwindow(C%, GW_HWNDNEXT)
E% = getwindow(D%, GW_HWNDNEXT)
f% = getwindow(E%, GW_HWNDNEXT)
G% = getwindow(f%, GW_HWNDNEXT)
H% = getwindow(G%, GW_HWNDNEXT)
i% = getwindow(H%, GW_HWNDNEXT)
J% = getwindow(i%, GW_HWNDNEXT)
K% = getwindow(J%, GW_HWNDNEXT)
L% = getwindow(K%, GW_HWNDNEXT)
M% = getwindow(L%, GW_HWNDNEXT)
N% = getwindow(M%, GW_HWNDNEXT)
o% = getwindow(N%, GW_HWNDNEXT)
p% = getwindow(o%, GW_HWNDNEXT)
QQ% = sendmessagebystring(PP%, WM_SETTEXT, 0, what)
click p%
End Sub

Sub Invitation (Who, Say, Where)
A% = findwindow("aol frame25", 0&)
B% = findchildbytitle(A%, "Buddy List Window")
C% = findchildbyclass(B%, "_Aol_Icon")
D% = getwindow(C%, GW_HWNDNEXT)
E% = getwindow(D%, GW_HWNDNEXT)
f% = getwindow(E%, GW_HWNDNEXT)
G% = getwindow(f%, GW_HWNDNEXT)
H% = getwindow(G%, GW_HWNDNEXT)
i% = getwindow(H%, GW_HWNDNEXT)
click i%
Do
J% = findwindow("aol frame25", 0&)
K% = findchildbytitle(J%, "Buddy Chat")
L% = findchildbyclass(K%, "_Aol_Edit")
DoEvents
Call timeout(.001)
Loop Until L% <> 0
N% = getwindow(L%, GW_HWNDNEXT)
p% = getwindow(N%, GW_HWNDNEXT)
Q% = getwindow(p%, GW_HWNDNEXT)
R% = getwindow(Q%, GW_HWNDNEXT)
S% = getwindow(R%, GW_HWNDNEXT)
t% = getwindow(S%, GW_HWNDNEXT)
v% = findchildbyclass(K%, "_Aol_Icon")
M% = sendmessagebystring(L%, WM_SETTEXT, 0, Who)
o% = sendmessagebystring(N%, WM_SETTEXT, 0, Say)
u% = sendmessagebystring(t%, WM_SETTEXT, 0, Where)
Call timeout(.1)
click v%
End Sub

Sub Kill45min ()
Do
A% = findwindow("_Aol_Palette", 0&)
B% = findchildbyclass(A%, "_Aol_Icon")
Call timeout(.001)
Loop Until B% <> 0
click (B%)
End Sub

Sub KillAdver ()
Do
A% = findwindow("Aol Frame25", 0&)
B% = findchildbyclass(A%, "_Aol_Image")
Call timeout(.001)
Loop Until B% <> 0
C% = SendMessageByNum(B%, WM_Close, 0, 0)
End Sub

Sub KillInvite ()
Do
A% = findwindow("Aol Frame25", 0&)
B% = findchildbytitle(A%, "Invitation From:")
Call timeout(.001)
Loop Until B% <> 0
Killwin (B%)
End Sub

Sub KillMailOK ()
Do
A% = findwindow("_Aol_Modal", 0&)
B% = findchildbyclass(A%, "_Aol_Icon")
Call timeout(.001)
Loop Until B% <> 0
click (B%)
End Sub

Sub KillModal ()
A% = findwindow("_Aol_Modal", 0&)
B% = findchildbyclass(A%, "_Aol_Icon")
click B%
End Sub

Sub Killok ()
Do
A% = findwindow("#32770", 0&)
B% = findchildbyclass(A%, "Button")
Call timeout(.001)
Loop Until B% <> 0
click B%
End Sub

Sub Killwait ()
Run "&About America Online"
Do
A% = findwindow("_Aol_Modal", 0&)
B% = findchildbyclass(A%, "_Aol_Glyph")
C% = getparent(B%)
D% = findchildbyclass(C%, "_Aol_Glyph")
Call timeout(.001)
Loop Until D% <> 0
E% = findchildbyclass(C%, "_Aol_Icon")
click E%
End Sub

Sub Killwin (windo)
x = SendMessageByNum(windo, WM_Close, 0, 0)
End Sub

Sub KW (Where$)
B% = findwindow("Aol Frame25", 0&)
A% = findchildbyclass(B%, "_Aol_Toolbar")
C% = findchildbyclass(A%, "_Aol_Icon")
D% = getwindow(C%, GW_HWNDNEXT)
E% = getwindow(D%, GW_HWNDNEXT)
f% = getwindow(E%, GW_HWNDNEXT)
G% = getwindow(f%, GW_HWNDNEXT)
H% = getwindow(G%, GW_HWNDNEXT)
i% = getwindow(H%, GW_HWNDNEXT)
J% = getwindow(i%, GW_HWNDNEXT)
K% = getwindow(J%, GW_HWNDNEXT)
L% = getwindow(K%, GW_HWNDNEXT)
M% = getwindow(L%, GW_HWNDNEXT)
N% = getwindow(M%, GW_HWNDNEXT)
o% = getwindow(N%, GW_HWNDNEXT)
p% = getwindow(o%, GW_HWNDNEXT)
Q% = getwindow(p%, GW_HWNDNEXT)
R% = getwindow(Q%, GW_HWNDNEXT)
S% = getwindow(R%, GW_HWNDNEXT)
t% = getwindow(S%, GW_HWNDNEXT)
u% = getwindow(t%, GW_HWNDNEXT)
v% = getwindow(u%, GW_HWNDNEXT)
W% = getwindow(u%, GW_HWNDNEXT)
y% = getwindow(W%, GW_HWNDNEXT)
click y%
Do
aol = findwindow("AOL Frame25", 0&)
bah = findchildbytitle(aol, "Keyword")
daedit = findchildbyclass(bah, "_AOL_Edit")
timeout (.001)
Loop Until daedit <> 0
daedit = findchildbyclass(bah, "_AOL_Edit")
send daedit, Where$
ico% = findchildbyclass(bah, "_AOL_Icon")
click ico%
End Sub

Sub Mail (Who, THESUBJECT, Say)
A% = findwindow("Aol Frame25", 0&)
BB% = findchildbyclass(A%, "_Aol_Toolbar")
DD% = findchildbyclass(BB%, "_Aol_Icon")
cc% = getwindow(DD%, GW_HWNDNEXT)
click (cc%)
Do
B% = findchildbytitle(A%, "Write Mail")
C% = findchildbyclass(B%, "_Aol_Edit")
Call timeout(.001)
Loop Until C% <> 0
D% = getwindow(C%, GW_HWNDNEXT)
E% = getwindow(D%, GW_HWNDNEXT)
i% = getwindow(E%, GW_HWNDNEXT)
J% = getwindow(i%, GW_HWNDNEXT)
K% = findchildbyclass(B%, "RICHCNTL")
f% = sendmessagebystring(C%, WM_SETTEXT, 0, Who)
G% = sendmessagebystring(J%, WM_SETTEXT, 0, THESUBJECT)
H% = sendmessagebystring(K%, WM_SETTEXT, 0, Say)
L% = findchildbyclass(B%, "_Aol_Icon")
ff% = getwindow(L%, GW_HWNDNEXT)
SAD% = getwindow(ff%, GW_HWNDNEXT)
mad% = getwindow(SAD%, GW_HWNDNEXT)
LAD% = getwindow(mad%, GW_HWNDNEXT)
HAD% = getwindow(LAD%, GW_HWNDNEXT)
RAD% = getwindow(HAD%, GW_HWNDNEXT)
SAID% = getwindow(RAD%, GW_HWNDNEXT)
MOP% = getwindow(SAID%, GW_HWNDNEXT)
LOP% = getwindow(MOP%, GW_HWNDNEXT)
HOP% = getwindow(LOP%, GW_HWNDNEXT)
COO% = getwindow(HOP%, GW_HWNDNEXT)
DOO% = getwindow(COO%, GW_HWNDNEXT)
GOO% = getwindow(DOO%, GW_HWNDNEXT)
MOO% = getwindow(GOO%, GW_HWNDNEXT)
SOO% = getwindow(MOO%, GW_HWNDNEXT)
FOO% = getwindow(SOO%, GW_HWNDNEXT)
LAP% = getwindow(FOO%, GW_HWNDNEXT)
RAP% = getwindow(LAP%, GW_HWNDNEXT)
Call timeout(.3)
click RAP%
End Sub

Sub Make3d (myForm As Form, MyCtl As Control)
'Place in Form_Paint; works best with Grey background
'Example:
'Make3d me,command1
myForm.ScaleMode = 3
myForm.CurrentX = MyCtl.Left - 1
myForm.CurrentY = MyCtl.Top + MyCtl.Height
myForm.Line -Step(0, -(MyCtl.Height + 1)), RGB(92, 92, 92)
myForm.Line -Step(MyCtl.Width + 1, 0), RGB(92, 92, 92)
myForm.Line -Step(0, MyCtl.Height + 1), RGB(255, 255, 255)
myForm.Line -Step(-(MyCtl.Width + 1), 0), RGB(255, 255, 255)
End Sub

Sub Playsound (Xsound As String)
Debug.Print "Xsound  " & Xsound
Dim x%
x% = sndPlaySound(Xsound, 1)
End Sub

Sub Print_Form (FRM As Form)
On Error GoTo ggoopp
FRM.PrintForm
Exit Sub
ggoopp:
Exit Sub
End Sub

Sub Printtext (File)
On Error GoTo Sux
x = Shell("c:\windows\notepad.exe /p " & File, 1)
Exit Sub
Sux:
Exit Sub
End Sub

Sub Project_FileCount ()
A% = findwindow("Project", 0&)
B% = findchildbyclass(A%, "Listbox")
C = SendMessageByNum(B%, LB_GETCOUNT, 0, 0)
D% = getparent(B%)
E = WindowCaption(D%)
If B% = 0 Then
MsgBox "You gotta have the Project window Open", 16, "Error:"
Else
MsgBox "You got " & C & " file/s in your Project, " & E, 64, "Project File Count"
End If
End Sub

Function Readini (AppName, KeyName, FileName As String) As String
Dim sRet As String
sRet = String(255, Chr(0))
Readini = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))
'Example:
'R = readini(("WAOL"),("AppPath"),("c:\windows\win.ini"))
'X = shell(R & "\waol.exe",1)
End Function

Function removespace (thetext As String)
Dim Text$
Dim theloop%
Text$ = thetext$
For theloop% = 1 To Len(thetext$)
If Mid(Text$, theloop%, 1) = " " Then
Text$ = Left$(Text$, theloop% - 1) + Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
removespace = Text$

End Function

Function RGB2HEX (R%, G%, B%)
'Converts the rgb values to the aol hexadecimal
'format.
Dim x%
Dim XX%
Dim Color%
Dim Divide
Dim Answer%
Dim Remainder%
Dim Configuring$
For x% = 1 To 3
If x% = 1 Then Color% = B%
If x% = 2 Then Color% = G%
If x% = 3 Then Color% = R%
For XX% = 1 To 2
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
Next XX%
Next x%
Configuring$ = "#" + removespace(Configuring$) + Chr(34)
RGB2HEX = Configuring$
End Function

Sub Roomname ()
A% = findwindow("Aol Frame25", 0&)
B% = findchildbyclass(A%, "_Aol_Listbox")
C% = getparent(B%)
D% = findchildbyclass(C%, "RICHCNTL")
E% = getparent(D%)
f% = findchildbyclass(E%, "_Aol_Static")
G% = getparent(f%)
H% = findchildbyclass(E%, "RICHCNTL")
ChatEdit2% = getwindow(H%, GW_HWNDNEXT)
Chatedi% = getwindow(ChatEdit2%, GW_HWNDNEXT)
Chated% = getwindow(Chatedi%, GW_HWNDNEXT)
Chate% = getwindow(Chated%, GW_HWNDNEXT)
Chat% = getwindow(Chate%, GW_HWNDNEXT)
chatedit% = getwindow(Chat%, GW_HWNDNEXT)
i% = getparent(chatedit%)
J = WindowCaption(i%)
AO = """"
AOP = "You are in " & J
chattext "  "
chattext AOP
chattext "  "
End Sub

Sub Run (ByVal menuname As String)
Dim mhWnd As Integer
Dim MnuCnt As Integer
Dim Cnt As Integer
Dim SmhWnd As Integer
Dim SMnuCnt As Integer
Dim SCnt As Integer
Dim lint As Integer
Dim lpString As String
Dim AOLx
AOLx = findwindow%("AOL Frame25", 0&)
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
ret = SendMessageByNum(AOLx, WM_COMMAND, GMnID, 0)
DoEvents
Exit Sub
End If
DoEvents
Next SCnt
Next Cnt
End Sub

Sub send (chatedit, sill$)
sndtext = sendmessagebystring(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub Stayontop (StatusFRM As Form)
Dim success%
success% = SetWindowPos(StatusFRM.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function Talk_Backwards (strin As TextBox)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
Talk_Backwards = newsent$

End Function

Function Talk_Color (strin2 As TextBox)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
dad = "#"
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "ff0000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Loop
Talk_Color = newsent2$
End Function

Function Talk_Color2 (strin2 As TextBox)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
dad = "#"
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "0000ff" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Loop
Talk_Color2 = newsent2$
End Function

Function Talk_Dots (strin As TextBox)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + "."
Let newsent$ = newsent$ + nextchr$
Loop
Talk_Dots = newsent$
End Function

Function Talk_Elite (strin As TextBox)
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

If nextchr$ = "A" Then Let nextchr$ = "/-\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "d"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "|V|"
If nextchr$ = "m" Then Let nextchr$ = "/V\"
If nextchr$ = "N" Then Let nextchr$ = "|\|"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "0"
If nextchr$ = "P" Then Let nextchr$ = "P"
If nextchr$ = "p" Then Let nextchr$ = "p"
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
Talk_Elite = newsent$

End Function

Function talk_fade (strin2 As TextBox, sred As VScrollBar, sblue As VScrollBar, sgreen As VScrollBar, ered As HScrollBar, eblue As HScrollBar, egreen As HScrollBar)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
dad = "#"
If sred > ered Then
red% = sred - ered
ElseIf ered > sred Then
red% = ered - sred
Else
red% = sred - ered
End If
If sblue > eblue Then
blue% = sblue - eblue
ElseIf eblue > sblue Then
blue% = eblue - sblue
Else
blue% = sblue - eblue
End If
If sgreen > egreen Then
green% = sgreen - egreen
ElseIf egreen > sgreen Then
green% = egreen - sgreen
Else
green% = sgreen - egreen
End If
redf = red% / lenth2%
bluef = blue% / lenth2%
greenf = green% / lenth2%

red% = red% + sred
green% = green% + sgreen
blue% = blue% + sblue
Let numspc2% = 1
Do While numspc2% <= lenth2%
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = "<font color=" + RGB2HEX(red%, green%, blue%) + " > " + nextchr2$
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
red% = red% + redf
blue% = blue% + bluef
green% = green% + greenf

Loop
talk_fade = newsent2$
End Function

Function Talk_Hacker (strin As TextBox)
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
Talk_Hacker = newsent$

End Function

Function Talk_leet (strin2 As TextBox)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<sub><font color=" & mad & dad & "ff0000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</sub><font color=" & mad & dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Loop
Talk_leet = newsent2$
End Function

Function Talk_Rainbow (strin2 As TextBox)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
dad = "#"

Do While numspc2% <= lenth2%

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "0d84c4" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Loop
Talk_Rainbow = newsent2$

End Function

Function Talk_Spaced (strin As TextBox)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
Talk_Spaced = newsent$
End Function

Function Talk_Wave (strin2 As TextBox)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
dad = "#"

Do While numspc2% <= lenth2%

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUP><font color=" & mad & dad & "000033" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUP><font color=" & mad & dad & "000066" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUB><font color=" & mad & dad & "000099" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUB><font color=" & mad & dad & "0000cc" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUP><font color=" & mad & dad & "0000ff" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUP><font color=" & mad & dad & "0000cc" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUB><font color=" & mad & dad & "000099" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUB><font color=" & mad & dad & "000066" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUP><font color=" & mad & dad & "000033" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUP><font color=" & mad & dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Loop
Talk_Wave = newsent2$



End Function

Function Talk_wave2 (strin2 As TextBox)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
dad = "#"

Do While numspc2% <= lenth2%

Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "330000" & mad & ">"
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUP><font color=" & mad & dad & "330033" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUP><font color=" & mad & dad & "330066" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUB><font color=" & mad & dad & "330099" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUB><font color=" & mad & dad & "3300cc" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUP><font color=" & mad & dad & "3300ff" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUP><font color=" & mad & dad & "3300cc" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUB><font color=" & mad & dad & "330099" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUB><font color=" & mad & dad & "330066" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<SUP><font color=" & mad & dad & "330033" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "</SUP><font color=" & mad & dad & "330000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Loop
Talk_wave2 = newsent2$

End Function

Sub Textset (hWnd As Integer, what As String)
Dim R
R = sendmessagebystring(hWnd, &HC, 0, what)
End Sub

Sub timeout (duration)
starttime = Timer
Do While Timer - starttime < duration
x = DoEvents()
Loop
End Sub

Function UserSN ()
wlcm% = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "Welcome, ")
dacap = WindowCaption(wlcm%)
If wlcm% <> 0 Then
numba% = (InStr(dacap, "!") - 10)
Pname$ = Mid$(dacap, 10, numba%)
UserSN = Pname$
Else
UserSN = "Signed Off"
End If
End Function

Sub waitchange ()
Dim Old As Integer
Dim Boy As Integer
Old = getfocus()
Boy = getfocus()
Do Until Boy <> Old
x = DoEvents()
Boy = getfocus()
Loop
End Sub

Function WindowCaption (hWndd As Integer)
Dim WindowText As String * 255
Dim GetWinText As Integer
GetWinText = GetWindowText(hWndd, WindowText, 255)
WindowCaption = (WindowText)
End Function

Function wintxt (ByVal hWnd As Integer)
Dim x As Integer
Dim y As String
Dim z As Integer
x = Sendmessage(hWnd, &HE, 0&, 0&)
y = String(x + 1, " ")
z = sendmessagebystring(hWnd, &HD, x + 1, y)
wintxt = Left(y, x)
End Function

Sub Writeini (sAppname, sKeyName, sNewString, sFileName As String)
Dim R As Integer
R = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
'Example:
'Call writeini(("My Proggie","Loaded","5",App.path & "\Myprog.ini"))
End Sub

