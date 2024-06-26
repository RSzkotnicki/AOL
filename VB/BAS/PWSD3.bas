Attribute VB_Name = "PWSD"
' PWSD3.bas by bEaV.  Well...I thought I would never get this far....but well..I DID...
' I thought I would end it @ 2..but I still got requests..and I added the 2 most COMMON
' requests after the relases of v1 (ya..that shitty version got stuff) and I added um...
' ..I added Scan the entire C drive..and I also added scan a directory...sooooo u can now do
' that w/ this bas.....jus MORE reasons 2 say....THAT BVPWSD.bas is the BEST PWSD bas out...
' and best 2 EVER come out.....w/ OUT steal'n my codes....if u EVER find a better PWSD bas
' prove it 2 me......I even updated the DLL I scan off of....I got a LOT of mail ask'n Y I
' don't jus put the code n and  scan that way....and my responce....SPEED SCANNING...when
' u use the DLL 2 scan..it makes the scaning speed much faster.....so thas y I do it...
' if ur PWSD use's my DLL use my about screen....put n the command, menu..ect CallAboutScreen
' and it will display the info on my DLL...thanx and enjoy!

' NOTE: 2 know if the file has it...it will return the string "True"..
' i.e. scanner% = scanfile ("c:\windows\desktop\progs\pws.exe", "password")
' if scanner% = true then
' msgbox "its a PWS"
' end if

' and if ur 1 of these peepz who steal codes and copy and paste..then FUCK U!!!!!!!!!!  I spent
' the hard time riting all this code below...so don't steal...if u want a sub..MAIL ME AND I
' mite let u have 1...

' webpage: http://fast.to/protection
' Contact me on AIM at: Fei3
' mail me at phishme7@juno.com

'                                          dwspy32.dll
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)

'                                          bvscans.dll
Declare Function BVScan Lib "bvscans.dll" (FileName As String, SearchString As String)
Declare Function BVAbout Lib "bvscans.dll" ()

'                                          tkscans.dll
Declare Sub TKSInitalize Lib "tkscans.dll" ()
Declare Sub TKSAbout Lib "tkscans.dll" ()
Declare Function TKSScanFile Lib "tkscans.dll" (VBFile As String, VBSearch As String, ByVal VBCase As Integer) As Integer

'                                          kernel32
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

'                                          user32
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'                                           consts
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const HWND_TOP = 0

'                                           comdlg32.dll
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
 Private Type OPENFILENAME
 lStructSize As Long
 hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   FLAGS As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
   End Type
   Global r%
Global NewString$
Global FileName$
Global Appname$

   



Sub TrunksDLLScanFile(file, StringofSearch, IngnoreCase)
  scanner% = TKSScanFile("File", "StringofSearch", IngnoreCase)
End Sub

Sub SelectFile()
'                         ______________________________
'                        |                              |
'                        |                              |
'                        |     Thanks to:               |
'                        |        ieox                  |
'                        |  http://ryan.tierranet.com/  |
'                        |                              |
'                        |______________________________|
' ieox is the man!
 Dim OpenFile As OPENFILENAME
 Dim lReturn As Long
 OpenFile.lStructSize = Len(OpenFile)
' OpenFile.hwndOwner = Me.hwnd
 OpenFile.hInstance = App.hInstance
 OpenFile.lpstrFilter = "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
 OpenFile.nFilterIndex = 1
 OpenFile.lpstrFile = String(257, 0)
 OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
 OpenFile.lpstrFileTitle = OpenFile.lpstrFile
 OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
 OpenFile.lpstrFileTitle = OpenFile.lpstrFile
 OpenFile.nMaxFileTitle = OpenFile.nMaxFile
 OpenFile.lpstrInitialDir = App.Path
 OpenFile.lpstrTitle = "http://ryan.tierranet.com"
 OpenFile.FLAGS = 0
 lReturn = GetOpenFileName(OpenFile)
 If lReturn = 0 Then
  MsgBox "Cancel"
 Else
   MsgBox "Selected: " & Trim(OpenFile.lpstrFile)
 End If
End Sub

Function IFilexcists(ByVal sFileName As String) As Integer
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IFileexcists = False
        Else
            IFilexcists = True
    End If

End Function

Sub DeleteFile(Title1, Title2, file, Title3)
messagebox% = MsgBox("are you sure you wish to delete " & LCase$(Text1) & " ?", vbYesNo, "title2")
If messagebox% = vbYes Then
Kill file
MsgBox "file deleted", 64, "title3"
End If
End Sub

Sub TrunksDLLScanMultiFiles(Lst As ListBox, StringofSearch$, IngnoreCase)
For X = 0 To Lst.ListCount = -1
  scanner% = TKSScanFile("File(x)", "StringofSearch$", IngnoreCase)
Next X
End Sub

Sub Strings_U_Can_Use()
' win.ini
' system.ini
' config.sys
' himem.sys
' pw
' sn
' pw=
' sn=
' pw =
' sn =
' screename
' screenname
' screenname=
' screename=
' password
' password=
' screenname =
' screename =
' password =
' @yahoo.com
' @hotmail.com
' @beer.com
' @mailexcite.com
' @juno.com
' @rocketmail.com
' @usa.net
' @netadress.com
' kill
' deltree
' pws
' passwordstealer
' password stealer
' deltree.exe
' c:\*.*
' format
' sendkeys
' mail
' AOL mail
' &Sent Mail
' Aol Frame25
' SN:
' PW:
' password:
' scnreenname:
' screename:
' screen name:
' PWSTEAL
' Compose Mail
' phish
' enough 2 get ya started???
End Sub
Sub CopyFile(FileName, CopyTo)
If FileName = "" Then Exit Sub
If CopyTo = "" Then Exit Sub
If Not IFileExists(FileName) Then Exit Sub
On Error GoTo AnErrOccured
If InStr(Right(FileName, 4), ".") = 0 Then Exit Sub
If InStr(Right(CopyTo, 4), ".") = 0 Then Exit Sub
FileCopy FileName, CopyTo
Exit Sub
AnErrOccured:
MsgBox "An Unexpected Error Occured!", 16, "Error"
End Sub

Sub FileChangeTo(Choice, file)
' Choices are Normal, Read only, Hidden, System, and archive
If Not IFileExists(file) Then Exit Sub
If LCase$(Choice) = "normal" Then
SetAttr file, ATTR_NORMAL
ElseIf LCase(Choice) = "readonly" Then: SetAttr file$, ATTR_READONLY
ElseIf LCase(Choice) = "hidden" Then: SetAttr file$, ATTR_HIDDEN
ElseIf LCase(Choice) = "system" Then: SetAttr file$, ATTR_SYSTEM
ElseIf LCase(Choice) = "archive" Then: SetAttr file$, ATTR_ARCHIVE
End If
NoFreeze% = DoEvents()
End Sub
Function IfDirExists(TheDirectory)
Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(Dir$(TheDirectory))
If Err Or Check = 0 Then
    IfDirExists = False
Else
    IfDirExists = True
End If
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function


Sub DeleteDirectory(DriveName)
If Not IfDirExists(DriveName) Then Exit Sub
RmDir DriveName
End Sub

Sub MakeDirectory(DriveName)
MkDir DriveName
End Sub

Sub ReNameFile(file, NewName)
Name file$ As NewName
NoFreeze% = DoEvents()
End Sub

Sub RenameDirectory(OldName, NewName)
If Not IfDirExists(OldName) Then Exit Sub
Name OldName As NewName
NoFreeze% = DoEvents()
End Sub

Sub Codes_4_This_Bas()

' ScanFile ("" + Text1.Text), password, "True"
' will scan the file n textbox 1..4 the string "password"..edit how u need

' ScanMultiFiles ("" + list1.List), password, "True"
' will scan all files n listbox 1 4 the string "password"

' selectfile ' selects a file...use'n comman dialog.dll

' deletefile ("Title of msgbox 1"), Title of msgbox2, "title of msgbox3"
' deletes a file

' deletedrive ("C:\windows\desktop\funnyshit\") ' will delete the file(or folder) "funny shit"
' deletes a drive

' scanentire_c_drive ("Drive1", "Dir1", "File1", "password")
' will scan the entire c 4 the string "password"

' scanselecteddirectory ("Dir1", "File1", "password")
' will scan the selected directory 4 the string "password"

' ScanWinDOTini
' will scan the win.ini (will also give option 2 clean it if there is sumthen!)

' if AOLUse("C:\windows\desktop\pws.exe") = true then
' will check 2 c if the file is code'd 4 AOL use

' CallAboutScreen
' will show the about screen from the DLL!

' figure the rest out on ur own!!!
End Sub

Sub DeleteFile2(Title1, Title2, file, Title3)
' this delete file is better...more confusing..but worx better..deletes all trace of file
' better than stander "kill"
       messagebox% = MsgBox("are you sure you wish to delete " & LCase$(file) & " ?", vbYesNo, "title2")
If messagebox% = vbYes Then
       Dim Block1 As String, Block2 As String, Blocks As Long
       Dim FileHandle As Integer, iLoop As Long, Offset As Long
       Const BLOCKSIZE = 4096
       Block1 = String(BLOCKSIZE, "X")
       Block2 = String(BLOCKSIZE, " ")
       FileHandle = FreeFile
       Open file For Binary As FileHandle
       Blocks = (LOF(FileHandle) \ BLOCKSIZE) + 1
              For iLoop = 1 To Blocks
                     Offset = Seek(FileHandle)
                     Put FileHandle, , Block1
                     Put FileHandle, Offset, Block2
              Next iLoop
       Close FileHandle
       Kill file
       messagebox "File deleted", , "Title3"
End Sub
Sub FileSize(file)
Dim lFileSize As Long
lFileSize = FileLen("file")
End Sub

Public Function EditFileSizeString(FileSize As String) As String
       Dim temp As String
       Dim i As Integer
       Dim StrLen As Integer
       Dim pos As Integer
       temp = ""
       i = Len(FileSize) / 3
       StrLen = Len(FileSize) Mod 3
              If (StrLen = 0) Then
                     StrLen = 3
              End If
       pos = 1
              While (i > 0)
                     temp = temp & Mid(FileSize, pos, StrLen) & ","
                     pos = pos + StrLen
                     i = i - 1
                     StrLen = 3
              Wend
       FileSize = temp & Mid(FileSize, pos, 3)
       EditFileSize = FileSize
End Function

Function GetFromINI(Appname$, Keyname$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(Appname$, ByVal Keyname$, "", RetStr, Len(RetStr), FileName$))
End Function

Function Rite2INI(Appname$, Keyname$, NewString%, FileName$)
Dim r As Integer
r = WritePrivateProfileString(Appname$, Keyname$, NewString$, FileName$)
End Function

Sub CleanWinDOTini()
FileChangeTo "normal", "C:\windows\win.ini"
Call Rite2INI("Windows", "Load", " ", "C:\windows\win.ini")
 messagebox% = MsgBox("Would u like ur win.ini 2 b 'read only' again?", vbYesNo + vbhelp, "Win.ini")
  If messagebox% = vbhelp Then
  MsgBox "Wut this duz is make it so NOTHING can b rote down 2 the win.ini..NOTHING can b saved 2 it..that means...nothing can load from ur win.ini! If a file is on win.ini..then u probaly got a pws..wut this duz it make the file load when windows loads...and can steal ur PW..if u make it read only..nuth'n can b rote 2 it!"
End If
 If messagebox% = vbYes Then
 FileChangeTo "read only", "C:\windows\win.ini"
 MsgBox "It'z read only now!", , "WHOA YAH!"
End If
End Sub

Sub ScanWinDOTini(where2load$)
windotini% = GetFromINI("Windows", "load", "C:\windows\win.ini")
If windotini% = "" Then
MsgBox "Ur win.ini is clean from fagets who mite steal ur PW", vbInformation, "CLEAN!!!!!!!!"
End If
If windotini% <> 1 Then
MsgBox "U don't have a load section to ur win.ini!!!", vbCritical, "No load!"
End If
where2load$ = windotini% ' if a filename is found..this will put it on the text/label/ ect.
messagebox% = MsgBox("Ur win.ini has the file " + windotini% + ", should this b  earesd?", vbYesNo, "Clean or NO")
If messagebox% = vbYes Then
CleanWinDOTini
MsgBox "File " + windotini% + " haz been earesed from ur win.ini!", vbInformation, "Win.ini is now clean!"
End If
If messagebox% = vbNo Then
MsgBox "k..I'll leave it alone..u better check that out tho..u dunno wut ur leave'n there......", vbExclamation, "It'z still there"
End Sub

Sub PercentBar(Picture As Control, FilesScanned As Integer, TotalFiles As Variant)
'                         ______________________________
'                        |                              |
'                        |                              |
'                        |     Thanks to:               |
'                        |        Mission               |
'                        | icyhotmission@juno.com       |
'                        |                              |
'                        |______________________________|
' sub edited by bEaV 2 fit this scenraio
On Error Resume Next
Picture.AutoRedraw = True
Picture.FillStyle = 0
Picture.DrawStyle = 0
X = Done / Total * Shape.Width
Picture.Line (0, 0)-(Picture.Width, Picture.Height), RGB(255, 255, 255), BF
Picture.Line (0, 0)-(X - 10, Picture.Height), RGB(0, 0, 255), BF
Picture.CurrentX = (Picture.Width / 2) - 100
Picture.CurrentY = (Picture.Height / 2) - 125
Picture.ForeColor = RGB(255, 0, 0)
Picture.Print Percent(FilesScanned, TotalFiles, 100) & "%"
End Sub

Sub GetFileWithOCX(filetype, Title1, where2)
On Error Resume Next
CMDialog1.InitDir = App.Path
CMDialog1.FLAGS = &H1000& Or &H4& Or &H800& Or &H4000&
CMDialog1.DefaultExt = "filetype"
CMDialog1.DialogTitle = "title1"
CMDialog1.Filter = "(*.filetype) filetype|*.filetype|"
CMDialog1.MaxFileSize = 99999
CMDialog1.FileName = ""
CMDialog1.CancelError = True
CMDialog1.Action = 1
 Open CMDialog1.FileName For Input As #1
 Text1 = LCase$(CMDialog1.FileName)
 If where2 = "" Then
 where2 = "no file selected"
 End If
 Label2.Enabled = True
 Close #1
 MsgBox CMDialog1.FileName - App.Path
End Sub

Sub ScanFile(FileName As String, seachstring As String)
scanner% = BVScan("FileName", "searchstring")
End Sub

Sub ScanMultiFiles(Lst As ListBox, SearchString)
For X = 0 To Lst.ListCount - 1
 BVScan (Lst(X)), "seachstring"
Next X
End Sub

Sub ScanEntire_C_Drive(drive As Drivebox, Dir As DriveListBox, file As FileListBox, SearchString$)
drive = file
For X = 0 To file.ListCount - 1
 ScanFile% = ScanFile("file(x)", "searchstring$")
 Next X
For c = 0 To Dir.liscount - 1
 entirec2% = ScanFile("file(c)", "searchstring$")
 Next c
results% = entirec% + entirec2%
End Sub

Sub ScanSelectedDirectory(Dir As DriveListBox, file As FileListBox, SearchString$)
Dir = file
For d = 0 To file.ListCount - 1
 Directory% = ScanFile("file(d)", "Searchstring$")
 Next d
End Sub

Sub CallAboutScreen()
' use this if ur PWSD has a about screen..and u r USING MY SCANNING ENGINE
Call BVAbout
End Sub



Sub DeleteFile2NOMSG(file, Title1)
' this delete file is better...more confusing..but worx better..deletes all trace of file
' better than stander "kill"
       Dim Block1 As String, Block2 As String, Blocks As Long
       Dim FileHandle As Integer, iLoop As Long, Offset As Long
       Const BLOCKSIZE = 4096
       Block1 = String(BLOCKSIZE, "X")
       Block2 = String(BLOCKSIZE, " ")
       FileHandle = FreeFile
       Open file For Binary As FileHandle
       Blocks = (LOF(FileHandle) \ BLOCKSIZE) + 1
              For iLoop = 1 To Blocks
                     Offset = Seek(FileHandle)
                     Put FileHandle, , Block1
                     Put FileHandle, Offset, Block2
              Next iLoop
       Close FileHandle
       Kill file
       MsgBox "File deleted", , "Title1"
End Sub

Sub DeleteFileNOMSG(file, Title1)
Kill file
MsgBox "file deleted", 64, "title1"
End Sub

Sub AOLUse(file)
heh% = ScanFile("file", "AOL Frame25")
If heh% = True Then
MsgBox " " + file + " is code'd 4 AOL Use" ' u may edit this how u need
End If
If heh% = False Then
MsgBox " " + file + " is not code'd 4 AOL Use" ' u may edit this who u need
End If
End If
End Sub
