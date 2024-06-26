Attribute VB_Name = "Juno"
' Juno.bas by bEaV
' well..this thing prolly sux...I bet it duz..but I was asked a few x by a few people (no more
' than 5) 2 make sumthen 4 Juno since I use it 4 my mail..so I play'd around w/ the ApI
' viewer that I found n the DoS32.bas...and I guess..u get this...it's not 2 great
' I doubt many subs work...but if they don't..and u can fix um..send me em..and i"ll replace
' them and give u the credit...

' thanx

' bEaV

' AIM: Fei3, damit beav
' mail: phishme7@juno.com
' web: http://fast.to/protection (NOT VB..anti virus)

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_SPACE = &H20
Public Const WM_SETTEXT = &HC
Public Const WM_CHAR = &H102
Public Const ENTER_KEY = 13

Sub Butten(Button As Long)
Call SendMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub TimeOut(Time As Long)
Dim Current As Long
Current = Timer
Do Until Timer - Current >= Time
DoEvents
Loop
End Sub

Sub UserSN()
DaJunoWin& = FindWindow("JunoMainWndXQW21", "phishme7 - Juno")
DaJunoTxtLength& = GetWindowTextLength(DaJunoWin&)
DaJunoTxtLength2& = String$(DaJunoTxtLength&, 0)
DaJunoSN& = GetWindowText(DaJunoWin&, DaJunoTxtLength2&, (DaJunoTxtLength + 1))
If Not Right(DaJunoTxtLength2&, 7) = " - Juno" Then Exit Function
DaJunoSN2& = Mid$(DaJunoTxtLength2&, 1, (DaJunoTxtLength& - 7))
UserSN = DaJunoSN2&
End Sub

Sub SignOn(name$, PW$)
DaJunoWin& = FindWindow("#32770", "Welcome to Juno (Version 2.0)")
DaSNCmb& = FindWindowEx(DaJunoWin&, 0, "ComboBox", vbNullString)
Call SendMessageByString(DaSNCmb, 0, WM_SETTEXT, name$)
DaJunoTxt& = FindWindowEx(DaJunoWin&, 0, "Edit", vbNullString)
Call SendMessageByString(DaJunoTxt, 0, WM_SETTEXT, PW$)
DaJunoBtn& = FindWindowEx(DaJunoWin&, 0, "Button", vbNullString)
DaJunoBtn2& = FindWindowEx(DaJunoWin&, DaJunoBtn&, "Button", vbNullString)
Butten (DaJunoBtn2)
End Sub

Sub CheckForNewMailMsg(Yes As Boolean)
DaJunoWin& = FindWindow("JunoMainWndXQW21", " " + UserSN + " - Juno")
DaJuno32770& = FindWindowEx(DaJunoWin&, 0, "#32770", vbNullString)
DaJunoBtn& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn&, "Button", vbNullString)
DaJunoBtn3& = FindWindowEx(DaJuno32770&, DaJunoBtn2&, "Button", vbNullString)
If Yes Then
Call SendMessageLong(DaJunoBtn3, WM_CHAR, ENTER_KEY, 0&)
Else
Exit Sub
End If
End Sub

Sub MailListBox()
DaJunoWin& = FindWindow("JunoMainWndXQW21", "" + UserSN + " - Juno")
DaJunoAfx& = FindWindowEx(DaJunoWin&, 0, "Afx:400000:8", vbNullString)
DaJuno32770& = FindWindowEx(DaJunoAfx&, 0, "#32770", vbNullString)
DaJunoStatic& = FindWindowEx(DaJuno32770&, 0, "Static", vbNullString)
DaJunoStatic2& = FindWindowEx(DaJuno32770&, DaJunoStatic&, "Static", vbNullString)
DaJunoStatic3& = FindWindowEx(DaJuno32770&, DaJunoStatic2&, "Static", vbNullString)
DaJunoStatic4& = FindWindowEx(DaJuno32770&, DaJunoStatic3&, "Static", vbNullString)
DaJunoStatic5& = FindWindowEx(DaJuno32770&, DaJunoStatic4&, "Static", vbNullString)
DaJunoStatic6& = FindWindowEx(DaJuno32770&, DaJunoStatic5&, "Static", vbNullString)
DaJunoListBox& = FindWindowEx(DaJuno32770&, DaJunoStatic2&, "Static", vbNullString)
End Sub

Function MailBox()
DaJunoWin& = FindWindow("JunoMainWndXQW21", "" + UserSN + " - Juno")
DaJunoAfx& = FindWindowEx(DaJunoWin&, 0, "Afx:400000:8", vbNullString)
DaJuno32770& = FindWindowEx(DaJunoAfx&, 0, "#32770", vbNullString)
DaJunoCmb& = FindWindowEx(DaJuno32770&, 0, "ComboBox", vbNullString)
End Function

Function ReplyMail(subject$, message$)
DaJunoWin& = FindWindow("JunoMainWndXQW21", "" + UserSN + " - Juno")
DaJunoAfx& = FindWindowEx(DaJunoWin&, 0, "Afx:400000:8", vbNullString)
DaJuno32770& = FindWindowEx(DaJunoAfx&, 0, "#32770", vbNullString)
DaJunoBtn& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
Butten (DaJunoBtn)
TimeOut 0.5
DaJunoBtn2& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn3& = FindWindowEx(DaJuno32770&, DaJunoBtn2&, "Button", vbNullString)
DaJunoBtn4& = FindWindowEx(DaJuno32770&, DaJunoBtn3&, "Button", vbNullString)
DaJunoBtn5& = FindWindowEx(DaJuno32770&, DaJunoBtn4&, "Button", vbNullString)
DaJunoBtn6& = FindWindowEx(DaJuno32770&, Dajunobtn7&, "Button", vbNullString)
Dajunobtn7& = FindWindowEx(DaJuno32770&, DaJunoBtn6&, "Button", vbNullString)
Butten (Dajunobtn7)
TimeOut 0.5
DaJunoEdit& = FindWindowEx(DaJuno32770&, 0, "Edit", vbNullString)
DaJunoEdit2& = FindWindowEx(DaJuno32770&, DaJunoEdit&, "Edit", vbNullString)
DaJunoTxt& = FindWindowEx(DaJuno32770&, DaJunoEdit2&, "Edit", vbNullString)
Call SendMessageByString(DaJunoTxt, 0, WM_SETTEXT, subject$)
DaJunoTxt2& = FindWindowEx(DaJuno32770&, 0, "RICHEDIT", vbNullString)
Call SendMessageByString(DaJunoTxt2, 0, WM_SETTEXT, message$)
DaJunoBtn8& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn9& = FindWindowEx(DaJuno32770&, DaJunoBtn8&, "Button", vbNullString)
DaJunoBtn10& = FindWindowEx(DaJuno32770&, DaJunoBtn9&, "Button", vbNullString)
DaJunoSendBtn& = FindWindowEx(DaJuno32770&, DaJunoBtn10&, "Button", vbNullString)
Butten (DaJunoSendBtn)
DaJunoBtn11& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoSendBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn11&, "Button", vbNullString)
Butten (DaJunoSendBtn2)
End Function

Function ForwardMail(who$, subject$)
DaJunoWin& = FindWindow("JunoMainWndXQW21", "" + UserSN + " - Juno")
DaJunoAfx& = FindWindowEx(DaJunoWin&, 0, "Afx:400000:8", vbNullString)
DaJuno32770& = FindWindowEx(DaJunoAfx&, 0, "#32770", vbNullString)
DaJunoBtn& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn&, "Button", vbNullString)
Butten (DaJunoBtn2)
TimeOut 0.5
DaJunoEdit& = FindWindowEx(DaJuno32770&, 0, "Edit", vbNullString)
Call SendMessageByString(DaJunoEdit, 0, WM_SETTEXT, who$)
DaJunoEdit2& = FindWindowEx(DaJuno32770&, 0, "Edit", vbNullString)
DaJunoEdit3& = FindWindowEx(DaJuno32770&, DaJunoEdit2&, "Edit", vbNullString)
DaJunoEdit4& = FindWindowEx(DaJuno32770&, DaJunoEdit3&, "Edit", vbNullString)
Call SendMessageByString(DaJunoEdit4, 0, WM_SETTEXT, subject$)
DaJunoBtn3& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn4& = FindWindowEx(DaJuno32770&, DaJunoBtn3&, "Button", vbNullString)
DaJunoBtn5& = FindWindowEx(DaJuno32770&, DaJunoBtn4&, "Button", vbNullString)
DaJunoForwardBtn& = FindWindowEx(DaJuno32770&, DaJunoBtn5&, "Button", vbNullString)
Butten (DaJunoForwardBtn)
DaJunoBtn6& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoForwardBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn6&, "Button", vbNullString)
Butten (DaJunoForwardBtn2)
End Function

Sub DeleteMail()
DaJunoWin& = FindWindow("JunoMainWndXQW21", "" + UserSN + " - Juno")
DaJunoAfx& = FindWindowEx(DaJunoWin&, 0, "Afx:400000:8", vbNullString)
DaJuno32770& = FindWindowEx(DaJunoAfx&, 0, "#32770", vbNullString)
DaJunoBtn& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn&, "Button", vbNullString)
DaJunoDeleteBtn& = FindWindowEx(DaJuno32770&, DaJunoBtn2&, "Button", vbNullString)
Butten (DaJunoDeleteBtn)
End Sub

Function GetNewMail()
DaJunoWin& = FindWindow("JunoMainWndXQW21", "" + UserSN + " - Juno")
DaJunoAfx& = FindWindowEx(DaJunoWin&, 0, "Afx:400000:8", vbNullString)
DaJuno32770& = FindWindowEx(DaJunoAfx&, 0, "#32770", vbNullString)
DaJunoBtn& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn&, "Button", vbNullString)
DaJunoBtn3& = FindWindowEx(DaJuno32770&, DaJunoBtn2&, "Button", vbNullString)
DaJunoBtn4& = FindWindowEx(DaJuno32770&, DaJunoBtn3&, "Button", vbNullString)
DaJunoBtn5& = FindWindowEx(DaJuno32770&, DaJunoBtn4&, "Button", vbNullString)
DaJunoBtn6& = FindWindowEx(DaJuno32770&, DaJunoBtn5&, "Button", vbNullString)
DaJunoGetNewMailBtn& = FindWindowEx(DaJuno32770&, DaJunoBtn6&, "Button", vbNullString)
Butten (DaJunoGetNewMailBtn)
End Function

Function SendMail(who$, subject$, message$)
DaJunoWin& = FindWindow("JunoMainWndXQW21", "" + UserSN + " - Juno")
DaJunoAfx& = FindWindowEx(DaJunoWin&, 0, "Afx:400000:8", vbNullString)
Butten (DaJunoAfx)
DaJuno32770& = FindWindowEx(DaJunoAfx&, 0, "#32770", vbNullString)
DaJunoEdit& = FindWindowEx(DaJuno32770&, 0, "Edit", vbNullString)
Call SendMessageByString(DaJunoEdit, 0, WM_SETTEXT, who$)
DaJunoEdit2& = FindWindowEx(DaJuno32770&, DaJunoEdit&, "Edit", vbNullString)
DaJunoEdit3& = FindWindowEx(DaJuno32770&, DaJunoEdit2&, "Edit", vbNullString)
Call SendMessageByString(DaJunoEdit3, 0, WM_SETTEXT, subject$)
DaJunoRichEdit& = FindWindowEx(OurParent&, 0, "RICHEDIT", vbNullString)
Call SendMessageByString(DaJunoRichEdit, 0, WM_SETTEXT, message$)
DaJunoBtn& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn&, "Button", vbNullString)
DaJunoBtn3& = FindWindowEx(DaJuno32770&, DaJunoBtn2&, "Button", vbNullString)
DaJunoSendBtn& = FindWindowEx(DaJuno32770&, DaJunoBtn3&, "Button", vbNullString)
Butten (DaJunoSendBtn)
TimeOut 0.5
DaJunoBtn4& = FindWindowEx(DaJuno32770&, 0, "Button", vbNullString)
DaJunoSendBtn2& = FindWindowEx(DaJuno32770&, DaJunoBtn4&, "Button", vbNullString)
Butten (DaJunoSendBtn2)
End Function
