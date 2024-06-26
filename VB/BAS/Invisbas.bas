Attribute VB_Name = "Hider3"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
       Public Const SRCCOPY = &HCC0020
       Public Const SRCAND = &H8800C6
       Public Const SRCINVERT = &H660046

Declare Function GetDesktopWindow Lib "user32" () As Long
