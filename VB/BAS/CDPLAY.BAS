Attribute VB_Name = "Module1"
Declare Function mciSendString Lib "mmsystem" (ByVal MCI_COMMAND As String, ByVal ReturnString As String, ByVal ReturnLength As Integer, ByVal Handle As Integer) As Long
Declare Function mciGetErrorString Lib "mmsystem" (ByVal MCI_ERROR As Long, ByVal ErororString As String, ByVal ReturnLenght As Integer) As Integer

Public Muted As Boolean
Public RandomMode As Boolean
Public ShuffleMode As Boolean
Public SMatch As Boolean
Public ShuffleTrak As Integer
Public Trakarray(30) As Integer
Public ProgramMode As Boolean
Public ProgArray(30) As Integer
Public ProgTrak As Integer
Public startpos, endpos As String
