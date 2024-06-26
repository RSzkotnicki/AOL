Attribute VB_Name = "Startup"
Option Explicit

#If Win16 Then
    Declare Function OSGetPrivateProfileString Lib "KERNEL" Alias "GetPrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$) As Integer
#Else
    Declare Function OSGetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$) As Integer
#End If

Public Const PLATFORM_WIN16 = 0
Public Const PLATFORM_WIN32 = 1
Public Const PLATFORM_1632 = 2

Public Const API_DECL32_CONVERT = 0
Public Const API_DECL32_TODO = 1

Public goUpgrade As clsUpgradeMgr

#If Win32 Then
    '---- In the Win32 Version, we used a resource file for the bitmaps.  The starting bitmap was resourceID #102
    Public Const ADD_START = 101
#Else
    '---- In Win16, we used a control array of image controls, starting at index 0
    Public Const ADD_START = -1
#End If

Sub Main()
    'goUpgrade is a global object variable.
    'All the magic happens inside the Upgrade Manager class module.
    'See the Class_Initialize event to see how the application starts.
    'By putting all the logic inside the class module, we can replace the UI but
    'keep the upgrade logic hidden inside the class.
    
    Set goUpgrade = New clsUpgradeMgr
End Sub

