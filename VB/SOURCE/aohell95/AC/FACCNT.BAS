  Declare Function SetWindowPos% Lib "user" (ByVal win%, ByVal awin%, ByVal xp%, ByVal yp%, ByVal xs%, ByVal ys%, ByVal FLAGS%)
  Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
  Declare Function FindWindow% Lib "user" (ByVal junk&, ByVal Title$)
  Declare Function FindChildByTitle% Lib "ole02.dll" (ByVal parent%, ByVal Title$)
  Declare Function FindChildByClass% Lib "ole02.dll" (ByVal parent%, ByVal Title$)
  Declare Function VBEnumChildWindows% Lib "ole01.dll" (ByVal win%, ByVal iNum%)
  Global StopCreate As Integer
  Global CheckingNum As String

  Global Cert11
  Global Cert12

  Global Cert21
  Global Cert22

  Global Cert31
  Global Cert32

  Global Cert41
  Global Cert42

  Global Cert51
  Global Cert52

  Global Cert61
  Global Cert62

