Attribute VB_Name = "Basic"
'This .bas file is Copyright 1998 Overmind Software
'This .bas file was created Octobert 4, 98 by KoA
'This .bas file was last modified October 4, 98
'Visit Overmind at http://www.overmind.net!
'You can email me with questions at
'tru_inc@email.msn.com.

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global inipath
Declare Function sndPlaySound Lib "MMSystem" (ByVal lpWavName$, ByVal Flags%) As Integer
Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer


Sub Pause(totaltime)
'To pause in VB do this:
'Call Pause(Interval)
'Ex. Call Pause(5)
TheStart = Timer
Do While Timer - TheStart < totaltime
    x = DoEvents()
Loop
End Sub

Function FileExists(ByVal strPathName As String) As Integer
'To check if a file exists put:
'If FileExists(Dir) = True Then
'Msgbox "This file exists!"
'Exit Sub
'End If
'If FileExists(Dir) = False Then
'Msgbox "This file does not exist!"
'Exit Sub
'End If
'Ex. If FileExists("C:\Windows\Win.ini") = True Then
'Msgbox "Good Win.ini still exists!"
'Exit Sub
'End If
'If FileExists("C:\Windows\Win.ini") = False Then
'Msgbox "Uh oh! Win.ini has been deleted!"
'Exit Sub
'End If

Dim intFileNum As Integer
On Error Resume Next
If Right$(strPathName, 1) = "\" Then
strPathName = Left$(strPathName, Len(strPathName) - 1)
End If
intFileNum = FreeFile
Open strPathName For Input As intFileNum
FileExists = IIf(Err, False, True)
Close intFileNum
Err = 0
End Function
Sub ExeOpen(File)
'To open another .exe put:
'Call ExeOpen(Dir)
'Ex. Call ExeOpen("C:\AOL30\waol.exe")
Dim x
x = Shell(File)
End Sub


Sub PlayWav(ByVal File$)
'To playwav put:
'Call PlayWav(Directory)
'Ex. Call PlayWav("C:\Windows\Media\robotz.wav")
Dim x As Integer
Dim wFlags As Integer
Dim NoFreeze As Integer

x% = sndPlaySound(File$, wFlags%): NoFreeze% = DoEvents()
End Sub

Sub CenterForm(f As Form)
'To center form put:
'CenterForm Formname
'Ex. CenterForm Form1
    f.Left = (Screen.Width - f.Width) \ 2
    f.Top = (Screen.Height - f.Height) \ 2
End Sub
Public Function GetFromINI(AppName$, KeyName$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
'If you want to write to an .ini file do this:
'W% = WritePrivateProfileString("basic", "vb32", "cool", basic.ini")
'Where it says basic is where you would see "[Basic]" in an .ini file.
'Where it says vb32 is where you would see "vb32=" in an .ini file.
'Where it says cool would make "vb32=" = "vb32=cool" in an .ini file.
'Where it says basic.ini that is where the file would be located.

'To get info from an .ini file do this:
'Cool% = GetFromINI("basic", "vb32", "basic.ini")
'Where it says basic is where you would see "[Basic]" in an .ini file.
'Where it says vb32 is where you would see "vb32=" in an .ini file.
'Where it says basic.ini that is where the file would be located.
End Function
