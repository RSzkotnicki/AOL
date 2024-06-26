  VERSION 2.00
  Begin Form AccntStatus
    BackColor       =   &H8000000F&
    BorderStyle     =   3   'Fixed Double
    Caption         =   "Account Creation Status"
    ClientHeight    =   855
    ClientLeft      =   3000
    ClientTop       =   1200
    ClientWidth     =   3735
    ControlBox      =   0   'False
    Height          =   1260
    Left            =   2940
    LinkTopic       =   "Form1"
    MaxButton       =   0   'False
    MinButton       =   0   'False
    ScaleHeight     =   855
    ScaleWidth      =   3735
    Top             =   855
    Width           =   3855
    Begin SSCommand Command3D1
      BevelWidth      =   3
      Caption         =   "&Emergency Stop"
      Font3D          =   3 'Inset w/light shading
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3735
    End
    Begin SSPanel Panel3D1
      BevelWidth      =   3
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
    End
  End
  Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer


  Const SWP_NOMOVE = 2
  Const SWP_NOSIZE = 1
  Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2

  Sub Command3D1_Click ()
    GoTo ErrBeg

    Err1:
    End
    Exit Sub

    ErrBeg:
    On Error GoTo Err1
    StopCreate = 1
    AppActivate "America  Online"
    Panel3D1.Caption = "Cancel request acknowledged."

  End Sub

  Sub Form_Load ()
    Dim Success As Integer

    Success% = SetWindowPos%(AccntStatus.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

  End Sub

