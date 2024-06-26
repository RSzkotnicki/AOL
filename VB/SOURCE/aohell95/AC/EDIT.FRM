  VERSION 2.00
  Begin Form Edit1
    BackColor       =   &H8000000F&
    BorderStyle     =   3   'Fixed Double
    Caption         =   "Edit Certificate Numbers"
    ClientHeight    =   4020
    ClientLeft      =   2670
    ClientTop       =   1575
    ClientWidth     =   4335
    Height          =   4425
    Icon            =   EDIT.FRX:0000
    Left            =   2610
    LinkTopic       =   "Form3"
    MaxButton       =   0   'False
    MinButton       =   0   'False
    ScaleHeight     =   4020
    ScaleWidth      =   4335
    Top             =   1230
    Width           =   4455
    Begin Init Init1
      Application     =   ""
      Filename        =   ""
      Height          =   420
      Left            =   3840
      Parameter       =   ""
      Top             =   3360
      Value           =   ""
      Width           =   420
    End
    Begin CommonDialog CMDialog1
      DialogTitle     =   "Save Cert. File"
      Filter          =   "Cert Files (*.CRT)|*.CRT"
      Left            =   120
      Top             =   3480
    End
    Begin SSCommand Command3D2
      BevelWidth      =   3
      Caption         =   "&Cancel"
      Font3D          =   3 'Inset w/light shading
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
    End
    Begin SSCommand Command3D1
      BevelWidth      =   3
      Caption         =   "&Save"
      Font3D          =   3 'Inset w/light shading
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
    End
    Begin TextBox Text12
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   2400
      TabIndex        =   20
      Top             =   3000
      Width           =   1575
    End
    Begin TextBox Text11
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   480
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
    End
    Begin TextBox Text10
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   2400
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
    End
    Begin TextBox Text9
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Top             =   2520
      Width           =   1575
    End
    Begin TextBox Text8
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
    End
    Begin TextBox Text7
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   480
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
    End
    Begin TextBox Text6
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
    End
    Begin TextBox Text5
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
    End
    Begin TextBox Text4
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
    End
    Begin TextBox Text3
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
    End
    Begin TextBox Text2
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Width           =   1575
    End
    Begin TextBox Text1
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   600
      Width           =   1575
    End
    Begin Label Label8
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "Cert.  Password"
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   -1    'True
      Height          =   225
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   1290
    End
    Begin Label Label7
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "Certifcate Number"
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   -1    'True
      Height          =   225
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   1530
    End
    Begin Label Label6
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "6."
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   150
    End
    Begin Label Label5
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "5."
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   150
    End
    Begin Label Label4
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "4."
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   150
    End
    Begin Label Label3
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "3."
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   150
    End
    Begin Label Label2
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "2."
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   150
    End
    Begin Label Label1
      AutoSize        =   -1    'True
      BackColor       =   &H8000000F&
      Caption         =   "1."
      FontBold        =   -1    'True
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   150
    End
  End
  Declare Function SetWindowPos Lib "User" (ByVal H%, ByVal hb%, ByVal X%, ByVal Y%, ByVal CX%, ByVal CY%, ByVal F%) As Integer


  Const SWP_NOMOVE = 2
  Const SWP_NOSIZE = 1
  Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2

  Sub Command3D1_Click ()
    GoTo begin2

    titfuck2:
    Close #1
    Exit Sub


    begin2:
    On Error GoTo titfuck2

    CMDIALOG1.Action = 2
    Shit = CMDIALOG1.Filetitle
    xfilnam$ = CMDIALOG1.Filename

    Open xfilnam$ For Output As 1
    Write #1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text
    Close #1
    init1.Filename = "aohell.ini"
    init1.Application = "Options"
    init1.Parameter = "DefCert"
    init1.Value = xfilenam$


    Cert11 = Text1.Text
    Cert12 = Text2.Text

    Cert21 = Text3.Text
    Cert22 = Text4.Text

    Cert31 = Text5.Text
    Cert32 = Text6.Text

    Cert41 = Text7.Text
    Cert42 = Text8.Text

    Cert51 = Text9.Text
    Cert52 = Text10.Text

    Cert61 = Text11.Text
    Cert62 = Text12.Text
    If FAccount.Option3D1.Value = True Then
      FAccount.Text6.Text = Cert11
      FAccount.Text7.Text = Cert12
    End If
    If FAccount.Option3D2.Value = True Then
      FAccount.Text6.Text = Cert21
      FAccount.Text7.Text = Cert22
    End If
    If FAccount.Option3D3.Value = True Then
      FAccount.Text6.Text = Cert31
      FAccount.Text7.Text = Cert32
    End If
    If FAccount.Option3D4.Value = True Then
      FAccount.Text6.Text = Cert41
      FAccount.Text7.Text = Cert42
    End If
    If FAccount.Option3D5.Value = True Then
      FAccount.Text6.Text = Cert51
      FAccount.Text7.Text = Cert52
    End If
    If FAccount.Option3D6.Value = True Then
      FAccount.Text6.Text = Cert61
      FAccount.Text7.Text = Cert62
    End If


    Unload Me


  End Sub

  Sub Command3D2_Click ()
    Unload Me

  End Sub

  Sub Form_Load ()
    Dim Success As Integer

    Success% = SetWindowPos%(Edit1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)


    Text1.Text = Cert11
    Text2.Text = Cert12

    Text3.Text = Cert21
    Text4.Text = Cert22

    Text5.Text = Cert31
    Text6.Text = Cert32

    Text7.Text = Cert41
    Text8.Text = Cert42

    Text9.Text = Cert51
    Text10.Text = Cert52

    Text11.Text = Cert61
    Text12.Text = Cert62


  End Sub

