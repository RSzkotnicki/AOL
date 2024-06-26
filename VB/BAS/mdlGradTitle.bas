Attribute VB_Name = "mdlGradTitle"
Option Explicit

Public Const GT_HOW = "LtoR"
'Public Const GT_HOW = "TtoB"
    
    ' Values for GT_HOW are:
     ' TtoB Is Specified Color to Black Going Down
     ' BlueLtoR is fading Left to Right Select Color
        ' to Black
    ' Just Uncomment the one you want and
    ' Comment the other
    
    
' Color values for the Title Bar, They are
' RGB so each is 0 to 255
Public Const GT_RED = 0  ' The Red Value
Public Const GT_GREEN = 0  ' The Green Value
Public Const GT_BLUE = 255 ' The Blue Value


' Don't Comment Out any of the lines below here!!!!!
Public Const GT_SPACERVAL = 40

Public Type RECT
       Left As Long
       Top As Long
       Right As Long
       Bottom As Long
End Type

Public Type POINTAPI
       x As Long
       y As Long
End Type

Public Const COLOR_ACTIVECAPTION = 2
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const PLANES = 14 ' Number of planes
Public Const BITSPIXEL = 12 ' Number of bits per pixel


Public Declare Function GetWindowRect Lib "user32" _
       (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetSystemMetrics Lib "user32" _
       (ByVal nIndex As Long) As Long

Public Declare Function DrawFocusRect Lib "user32" _
       (ByVal hDC As Long, lpRect As RECT) As Long

Public Declare Function ClientToScreen Lib "user32" _
       (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function GetDC Lib "user32" _
       (ByVal hwnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32" _
       (ByVal hwnd As Long, ByVal hDC As Long) As Long


Declare Function CreateSolidBrush Lib "gdi32" _
       (ByVal crColor As Long) As Long

Declare Function DeleteObject Lib "gdi32" _
       (ByVal hObject As Long) As Long

Declare Function GetDeviceCaps Lib "gdi32" _
       (ByVal hDC As Long, ByVal nIndex As Long) As Long

Declare Function FillRect Lib "user32" _
       (ByVal hDC As Long, lpRect As RECT, _
       ByVal hBrush As Long) As Long

        
Public tpoint As POINTAPI
Public temp As POINTAPI
Public dpoint As POINTAPI
Public fbox As RECT
Public tbox As RECT
Public oldbox As RECT
Public TwipsPerPixelX
Public TwipsPerPixelY


Public Sub MakeGrad(PicBoxName As PictureBox, Orientation%, RStart%, GStart%, BStart%, RInc%, GInc%, BInc%)
    Dim x As Integer, y As Integer, z As Integer, Cycles As Integer
    Dim R%, G%, B%
    R% = RStart%: G% = GStart%: B% = BStart%
    If Orientation% = 0 Then
        Cycles = PicBoxName.ScaleHeight \ 100
    Else
        Cycles = PicBoxName.ScaleWidth \ 100
    End If
    
    For z = 1 To 100
        x = x + 1
        Select Case Orientation
            Case 0: 'Top to Bottom
                If x > PicBoxName.ScaleHeight Then Exit For
                PicBoxName.Line (0, x)-(PicBoxName.Width, x + Cycles - 1), RGB(R%, G%, B%), BF
            Case 1: 'Left to Right
                If x > PicBoxName.ScaleWidth Then Exit For
                PicBoxName.Line (x, 0)-(x + Cycles - 1, PicBoxName.Height), RGB(R%, G%, B%), BF
        End Select
        x = x + Cycles
        R% = R% + RInc%: G% = G% + GInc%: B% = B% + BInc%
        If R% > 255 Then R% = 255
        If R% < 0 Then R% = 0
        If G% > 255 Then G% = 255
        If G% < 0 Then G% = 0
        If B% > 255 Then B% = 255
        If B% < 0 Then B% = 0
    Next z
End Sub
