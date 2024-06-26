  VERSION 2.00
  Begin Form FAccount
    BackColor       =   &H8000000F&
    Caption         =   "AOHell Fake Account Creator v1.0 Beta 5"
    ClientHeight    =   5160
    ClientLeft      =   1845
    ClientTop       =   1020
    ClientWidth     =   6525
    Height          =   5565
    Icon            =   ACCNT.FRX:0000
    Left            =   1785
    LinkTopic       =   "Form1"
    MaxButton       =   0   'False
    ScaleHeight     =   5160
    ScaleWidth      =   6525
    Top             =   675
    Width           =   6645
    Begin Ctl3D Ctl3D1
      ClassList       =   ACCNT.FRX:0302
      Left            =   6960
      Top             =   0
    End
    Begin SSFrame Frame3D1
      FontBold        =   0 'False
      FontItalic      =   0 'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0 'False
      FontUnderline   =   0 'False
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   -120
      TabIndex        =   15
      Top             =   -120
      Width           =   6975
      Begin Init Init1
        Application     =   ""
        Filename        =   ""
        Height          =   420
        Left            =   1560
        Parameter       =   ""
        Top             =   5040
        Value           =   ""
        Width           =   420
      End
      Begin CommonDialog CMDialog1
        DialogTitle     =   "Open Cert. File"
        Filter          =   "Cert Files (*.CRT)|*.CRT"
        Left            =   840
        Top             =   5040
      End
      Begin SSOption Option3D6
        Caption         =   "Cert 6"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "MS Sans Serif"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   255
        Left            =   5880
        TabIndex        =   46
        TabStop         =   0   'False
        Top             =   4920
        Width           =   735
      End
      Begin SSOption Option3D5
        Caption         =   "Cert 5"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "MS Sans Serif"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   255
        Left            =   5880
        TabIndex        =   45
        TabStop         =   0   'False
        Top             =   4680
        Width           =   735
      End
      Begin SSCommand Command3D2
        BevelWidth      =   3
        Caption         =   "&Edit Certs"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        Left            =   5280
        TabIndex        =   44
        Top             =   2880
        Width           =   1215
      End
      Begin SSCommand Command3D1
        BevelWidth      =   3
        Caption         =   "&Load Certs"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        Left            =   5280
        TabIndex        =   43
        Top             =   2400
        Width           =   1215
      End
      Begin SSOption Option3D4
        Caption         =   "Cert 4"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "MS Sans Serif"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   255
        Left            =   5880
        TabIndex        =   18
        TabStop         =   0   'False
        Top             =   4440
        Width           =   735
      End
      Begin SSOption Option3D1
        Caption         =   "Cert 1"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "MS Sans Serif"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   255
        Left            =   5160
        TabIndex        =   42
        Top             =   4440
        Value           =   -1  'True
        Width           =   735
      End
      Begin SSCommand Command3D4
        BevelWidth      =   3
        Caption         =   "Guest Fix"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        Left            =   5280
        TabIndex        =   41
        Top             =   1920
        Width           =   1215
      End
      Begin SSOption Option3D3
        Caption         =   "Cert 3"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "MS Sans Serif"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   255
        Left            =   5160
        TabIndex        =   40
        TabStop         =   0   'False
        Top             =   4920
        Width           =   735
      End
      Begin SSOption Option3D2
        Caption         =   "Cert 2"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "MS Sans Serif"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   255
        Left            =   5160
        TabIndex        =   39
        TabStop         =   0   'False
        Top             =   4680
        Width           =   735
      End
      Begin VBMsg VBMsg1
        Height          =   420
        Left            =   6720
        MessageCount    =   ACCNT.FRX:0304
        MessageList     =   ACCNT.FRX:0306
        MessageTypes    =   0   'Selected Messages
        PostDefault     =   0   'False
        Top             =   3360
        Width           =   420
      End
      Begin SSCommand Command3D7
        BevelWidth      =   3
        Caption         =   "&Help"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        HelpContextID   =   36
        Left            =   5280
        TabIndex        =   34
        Top             =   3960
        Width           =   1215
      End
      Begin SSCheck Check3D1
        Caption         =   "&Slow for Realism"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   255
        Left            =   3600
        TabIndex        =   38
        Top             =   3720
        Width           =   1455
      End
      Begin SSCommand Command3D3
        BevelWidth      =   3
        Caption         =   "&New User"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        Left            =   5280
        TabIndex        =   37
        Top             =   1440
        Width           =   1215
      End
      Begin TextBox Text7
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "Arial"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   315
        Left            =   2040
        TabIndex        =   12
        Top             =   4680
        Width           =   1455
      End
      Begin TextBox Text6
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "Arial"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   315
        Left            =   2040
        TabIndex        =   11
        Top             =   4200
        Width           =   1455
      End
      Begin SSCommand Command3D8
        BevelWidth      =   3
        Caption         =   "&Exit"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        Left            =   5280
        TabIndex        =   1
        Top             =   3360
        Width           =   1215
      End
      Begin SSCommand Command3D6
        BevelWidth      =   3
        Caption         =   "&Start"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        Left            =   5280
        TabIndex        =   10
        Top             =   360
        Width           =   1215
      End
      Begin SSCommand Command3D5
        BevelWidth      =   3
        Caption         =   "Generate &All"
        Font3D          =   3   'Inset w/light shading
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   375
        Left            =   5280
        TabIndex        =   33
        Top             =   840
        Width           =   1215
      End
      Begin SSFrame Frame3D4
        Caption         =   "Name and Address Info"
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   -1  'True
        ForeColor       =   &H00000000&
        Height          =   1935
        Left            =   240
        TabIndex        =   21
        Top             =   1680
        Width           =   4815
        Begin TextBox Text19
          Height          =   285
          Left            =   6240
          TabIndex        =   24
          Text            =   "Last 4"
          Top             =   1320
          Visible         =   0 'False
          Width           =   855
        End
        Begin TextBox Text18
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   1320
          MaxLength       =   8
          TabIndex        =   7
          Top             =   1440
          Width           =   855
        End
        Begin TextBox Text17
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   720
          MaxLength       =   3
          TabIndex        =   6
          Top             =   1440
          Width           =   375
        End
        Begin TextBox Text20
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   3840
          TabIndex        =   9
          Top             =   1440
          Width           =   735
        End
        Begin TextBox Text16
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   2880
          MaxLength       =   2
          TabIndex        =   8
          Top             =   1440
          Width           =   375
        End
        Begin TextBox Text15
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   3360
          TabIndex        =   5
          Top             =   960
          Width           =   1215
        End
        Begin TextBox Text14
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   6240
          TabIndex        =   23
          Text            =   "Stree add"
          Top             =   1200
          Visible         =   0 'False
          Width           =   1095
        End
        Begin TextBox Text13
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   840
          TabIndex        =   4
          Top             =   960
          Width           =   1935
        End
        Begin TextBox Text12
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   3240
          TabIndex        =   3
          Top             =   480
          Width           =   1335
        End
        Begin TextBox Text11
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   285
          Left            =   1080
          TabIndex        =   2
          Top             =   480
          Width           =   975
        End
        Begin Label Label11
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "-"
          FontBold        =   -1    'True
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   195
          Left            =   1200
          TabIndex        =   32
          Top             =   1440
          Width           =   60
        End
        Begin Label Label10
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "Phone:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   9
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   225
          Left            =   120
          TabIndex        =   31
          Top             =   1440
          Width           =   495
        End
        Begin Label Label9
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "Zip:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   210
          Left            =   3480
          TabIndex        =   30
          Top             =   1440
          Width           =   285
        End
        Begin Label Label8
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "State:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   210
          Left            =   2400
          TabIndex        =   29
          Top             =   1440
          Width           =   390
        End
        Begin Label Label7
          BackColor       =   &H8000000F&
          Caption         =   "City:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   255
          Left            =   2880
          TabIndex        =   28
          Top             =   960
          Width           =   375
        End
        Begin Label Label6
          BackColor       =   &H8000000F&
          Caption         =   "Address:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   255
          Left            =   120
          TabIndex        =   27
          Top             =   960
          Width           =   615
        End
        Begin Label Label5
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "Last Name:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   9
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   225
          Left            =   2280
          TabIndex        =   26
          Top             =   480
          Width           =   825
        End
        Begin Label Label4
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "First Name:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   210
          Left            =   120
          TabIndex        =   25
          Top             =   480
          Width           =   810
        End
      End
      Begin SSFrame Frame3D2
        Caption         =   "Payment Options"
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   -1  'True
        ForeColor       =   &H00000000&
        Height          =   1335
        Left            =   240
        TabIndex        =   16
        Top             =   240
        Width           =   4815
        Begin TextBox Text3
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Arial"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   315
          Left            =   2280
          TabIndex        =   36
          Top             =   840
          Width           =   2055
        End
        Begin OptionButton Option1
          Caption         =   "Use this &Credit Card number:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "MS Sans Serif"
          FontSize        =   8.25
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   255
          Left            =   6240
          TabIndex        =   17
          Top             =   1560
          Value           =   -1    'True
          Width           =   2295
        End
        Begin TextBox Text1
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Arial"
          FontSize        =   9.75
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   360
          Left            =   2280
          TabIndex        =   0
          Top             =   360
          Width           =   2055
        End
        Begin Label Label3
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "Transit Number:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   9.75
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   225
          Left            =   360
          TabIndex        =   19
          Top             =   360
          Width           =   1290
        End
        Begin Label Label2
          AutoSize        =   -1    'True
          BackColor       =   &H8000000F&
          Caption         =   "Bank Name:"
          FontBold        =   0 'False
          FontItalic      =   0 'False
          FontName        =   "Times New Roman"
          FontSize        =   9.75
          FontStrikethru  =   0 'False
          FontUnderline   =   0 'False
          Height          =   225
          Left            =   360
          TabIndex        =   35
          Top             =   840
          Width           =   945
        End
      End
      Begin Label Label15
        BackColor       =   &H8000000F&
        Caption         =   "The sign-up number and password listed to the left is a permanent cert. You may use another."
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   8.25
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   1095
        Left            =   3600
        TabIndex        =   13
        Top             =   4080
        Width           =   1455
      End
      Begin Label Label14
        AutoSize        =   -1  'True
        BackColor       =   &H8000000F&
        Caption         =   "Cert.   Password:"
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   225
        Left            =   240
        TabIndex        =   14
        Top             =   4680
        Width           =   1395
      End
      Begin Label Label13
        AutoSize        =   -1  'True
        BackColor       =   &H8000000F&
        Caption         =   "Certficate Number:"
        FontBold        =   -1  'True
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   225
        Left            =   240
        TabIndex        =   20
        Top             =   4200
        Width           =   1590
      End
      Begin Label Label12
        AutoSize        =   -1  'True
        BackColor       =   &H8000000F&
        Caption         =   "Tip:  Hit ""Generate All"" for fast creation"
        FontBold        =   0   'False
        FontItalic      =   0   'False
        FontName        =   "Times New Roman"
        FontSize        =   9.75
        FontStrikethru  =   0   'False
        FontUnderline   =   0   'False
        Height          =   225
        Left            =   240
        TabIndex        =   22
        Top             =   3720
        Width           =   3165
      End
    End
  End
  Declare Function SetWindowPos Lib "User" (ByVal H%, ByVal hb%, ByVal X%, ByVal Y%, ByVal CX%, ByVal CY%, ByVal F%) As Integer


  Const SWP_NOMOVE = 2
  Const SWP_NOSIZE = 1
  Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2

  Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer

  Sub Command3D1_Click ()
    GoTo begin1

    titfuck1:
    Close #1
    Exit Sub


    begin1:
    On Error GoTo titfuck1

    CMDIALOG1.Action = 1
    xfilnam1$ = CMDIALOG1.Filename


    Open xfilnam1$ For Input As 1
    Input #1, Cert11, Cert12, Cert21, Cert22, Cert31, Cert32, Cert41, Cert42, Cert51, Cert52, Cert61, Cert62
    Close #1
    init1.Filename = "aohell.ini"
    init1.Application = "Options"
    init1.Parameter = "DefCert"
    init1.Value = xfilenam$
    If Option3D1.Value = True Then
      Text6.Text = Cert11
      Text7.Text = Cert12
    End If
    If Option3D2.Value = True Then
      Text6.Text = Cert21
      Text7.Text = Cert22
    End If
    If Option3D3.Value = True Then
      Text6.Text = Cert31
      Text7.Text = Cert32
    End If
    If Option3D4.Value = True Then
      Text6.Text = Cert41
      Text7.Text = Cert42
    End If
    If Option3D5.Value = True Then
      Text6.Text = Cert51
      Text7.Text = Cert52
    End If
    If Option3D6.Value = True Then
      Text6.Text = Cert61
      Text7.Text = Cert62
    End If



  End Sub

  Sub Command3D2_Click ()
    Edit1.Show

  End Sub

  Sub Command3D3_Click ()
    ResetSN.Show
    FAccount.Hide

  End Sub

  Sub Command3D4_Click ()
    FAccount.Hide
    ResetSN2.Show



  End Sub

  Sub Command3D5_Click ()



    If Option1.Value = True Then
    'Call Command3D1_Click
    End If


    Randomize
    upperb = 15
    lowerb = 1

    Transit = Int((upperb - lowerb + 1) * Rnd + lowerb)
    If Transit = 1 Then
      Text1.Text = "311093120"
      Text3.Text = "Bank of America"
    End If
    If Transit = 2 Then
      Text1.Text = "063112155"
      Text3.Text = "1st Natn. Bank of FL"
    End If
    If Transit = 3 Then
      Text1.Text = "322270521"
      Text3.Text = "American Savings"
    End If
    If Transit = 4 Then
      Text1.Text = "31100209"
      Text3.Text = "CitiBank Deleware"
    End If
    If Transit = 5 Then
      Text1.Text = "122016066"
      Text3.Text = "City  National Bank"
    End If
    If Transit = 6 Then
      Text1.Text = "322070213"
      Text3.Text = "Coast Federal Bank"
    End If
    If Transit = 7 Then
      Text1.Text = "322270369"
      Text3.Text = "Fidelity Federal"
    End If
    If Transit = 8 Then
      Text1.Text = "122000218"
      Text3.Text = "First Interstate Bank"
    End If
    If Transit = 9 Then
      Text1.Text = "322270039"
      Text3.Text = "Great Western Bank"
    End If
    If Transit = 10 Then
      Text1.Text = "111910555"
      Text3.Text = "Nations Bank"
    End If
    If Transit = 11 Then
      Text1.Text = "114900025"
      Text3.Text = "NCNB"
    End If
    If Transit = 12 Then
      Text1.Text = "122000247"
      Text3.Text = "Wells Fargo Bank"
    End If
    If Transit = 13 Then
      Text1.Text = "061113279"
      Text3.Text = "Melon Bank"
    End If
    If Transit = 14 Then
      Text1.Text = "063104972"
      Text3.Text = "First National"
    End If
    If Transit = 15 Then
      Text1.Text = "41203824"
      Text3.Text = "Norwest Bank"
    End If

  'Female First Name

    upperbound = 2
    lowerbound = 1

    A = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

    rndanswer = A
  ' Text14.Text = A

    If A = 2 Then

      upperbound1 = 113
      lowerbound1 = 1

      B = Int((upperbound1 - lowerbound1 + 1) * Rnd + lowerbound)

      rndanswer2 = B

      If B = 1 Then Text11.Text = "Maria"
      If B = 2 Then Text11.Text = "Jane"
      If B = 3 Then Text11.Text = "Stacy"
      If B = 4 Then Text11.Text = "Merideth"
      If B = 5 Then Text11.Text = "Beth"
      If B = 6 Then Text11.Text = "Angela"
      If B = 7 Then Text11.Text = "Mary"
      If B = 8 Then Text11.Text = "Tiffany"
      If B = 9 Then Text11.Text = "Pamela"
      If B = 10 Then Text11.Text = "April"
      If B = 11 Then Text11.Text = "Louise"
      If B = 12 Then Text11.Text = "Virginia"
      If B = 13 Then Text11.Text = "Nicole"
      If B = 14 Then Text11.Text = "Carolyn"
      If B = 15 Then Text11.Text = "Julia"
      If B = 16 Then Text11.Text = "Trina"
      If B = 17 Then Text11.Text = "Robin"
      If B = 18 Then Text11.Text = "Karen"
      If B = 19 Then Text11.Text = "Susan"
      If B = 20 Then Text11.Text = "Kathy"
      If B = 21 Then Text11.Text = "Shirley"
      If B = 22 Then Text11.Text = "Renee"
      If B = 23 Then Text11.Text = "Olivia"
      If B = 24 Then Text11.Text = "Beth"
      If B = 25 Then Text11.Text = "Gladys"
      If B = 26 Then Text11.Text = "Dianne"
      If B = 27 Then Text11.Text = "Ethelene"
      If B = 28 Then Text11.Text = "Patricia"
      If B = 29 Then Text11.Text = "Anne"
      If B = 30 Then Text11.Text = "Kathleen"
      If B = 31 Then Text11.Text = "Carol"
      If B = 32 Then Text11.Text = "Kimberly"
      If B = 33 Then Text11.Text = "Marcia"
      If B = 34 Then Text11.Text = "Marsha"
      If B = 35 Then Text11.Text = "Sharon"
      If B = 36 Then Text11.Text = "Ella"
      If B = 37 Then Text11.Text = "Evelyn"
      If B = 38 Then Text11.Text = "Glenda"
      If B = 39 Then Text11.Text = "Katie"
      If B = 40 Then Text11.Text = "Kate"
      If B = 41 Then Text11.Text = "Ann"
      If B = 42 Then Text11.Text = "Dana"
      If B = 43 Then Text11.Text = "Shannon"
      If B = 44 Then Text11.Text = "Elizabeth"
      If B = 45 Then Text11.Text = "Lillian"
      If B = 46 Then Text11.Text = "Jo Ann"
      If B = 47 Then Text11.Text = "Karen"
      If B = 48 Then Text11.Text = "Edith"
      If B = 49 Then Text11.Text = "Jenny"
      If B = 50 Then Text11.Text = "Deborah"
      If B = 51 Then Text11.Text = "Elisa"
      If B = 52 Then Text11.Text = "Becky"
      If B = 53 Then Text11.Text = "Barbara"
      If B = 54 Then Text11.Text = "Blair"
      If B = 55 Then Text11.Text = "Cynthia"
      If B = 56 Then Text11.Text = "Debbie"
      If B = 57 Then Text11.Text = "Diane"
      If B = 58 Then Text11.Text = "Margaret"
      If B = 59 Then Text11.Text = "Racheal"
      If B = 60 Then Text11.Text = "Shari"
      If B = 61 Then Text11.Text = "Gail"
      If B = 62 Then Text11.Text = "Sarah"
      If B = 63 Then Text11.Text = "Cathy"
      If B = 64 Then Text11.Text = "Katrina"
      If B = 65 Then Text11.Text = "Marlene"
      If B = 66 Then Text11.Text = "Patricia"
      If B = 67 Then Text11.Text = "Lisa"
      If B = 68 Then Text11.Text = "Candace"
      If B = 69 Then Text11.Text = "Beatrice"
      If B = 70 Then Text11.Text = "Betty"
      If B = 71 Then Text11.Text = "Wendy"
      If B = 72 Then Text11.Text = "Denise"
      If B = 73 Then Text11.Text = "Rebecca"
      If B = 74 Then Text11.Text = "Arlene"
      If B = 75 Then Text11.Text = "Stephanie"
      If B = 76 Then Text11.Text = "Flora"
      If B = 77 Then Text11.Text = "Martha"
      If B = 78 Then Text11.Text = "Nancy"
      If B = 79 Then Text11.Text = "Tanya"
      If B = 80 Then Text11.Text = "Helen"
      If B = 81 Then Text11.Text = "Patty"
      If B = 82 Then Text11.Text = "Della"
      If B = 83 Then Text11.Text = "Ida"
      If B = 84 Then Text11.Text = "Lisa"
      If B = 85 Then Text11.Text = "Jeanne"
      If B = 86 Then Text11.Text = "Theresa"
      If B = 87 Then Text11.Text = "Vanessa"
      If B = 88 Then Text11.Text = "Rose"
      If B = 89 Then Text11.Text = "Benita"
      If B = 90 Then Text11.Text = "Lucy"
      If B = 91 Then Text11.Text = "Suzette"
      If B = 92 Then Text11.Text = "Jessica"
      If B = 93 Then Text11.Text = "Blair"
      If B = 94 Then Text11.Text = "Gina"
      If B = 95 Then Text11.Text = "Melissa"
      If B = 96 Then Text11.Text = "Nikki"
      If B = 97 Then Text11.Text = "Victoria"
      If B = 98 Then Text11.Text = "Priscilla"
      If B = 99 Then Text11.Text = "Catherine"
      If B = 100 Then Text11.Text = "Joan"
      If B = 101 Then Text11.Text = "Sue"
      If B = 102 Then Text11.Text = "Sonja"
      If B = 103 Then Text11.Text = "Stephanie"
      If B = 104 Then Text11.Text = "Cindy"
      If B = 105 Then Text11.Text = "Alison"
      If B = 106 Then Text11.Text = "Samantha"
      If B = 107 Then Text11.Text = "Kelly"
      If B = 108 Then Text11.Text = "Carroll"
      If B = 109 Then Text11.Text = "Lauren"
      If B = 110 Then Text11.Text = "Louis"
      If B = 111 Then Text11.Text = "Alicia"
      If B = 112 Then Text11.Text = "Sylvia"
      If B = 113 Then Text11.Text = "Jenn"

    End If

  'Male First Name

    If A = 1 Then

      upperbound2 = 101
      lowerbound2 = 1

      C = Int((upperbound2 - lowerbound2 + 1) * Rnd + lowerbound)

      If C = 1 Then Text11.Text = "Michael"
      If C = 2 Then Text11.Text = "Mike"
      If C = 3 Then Text11.Text = "Randy"
      If C = 4 Then Text11.Text = "Richard"
      If C = 5 Then Text11.Text = "Ronald"
      If C = 6 Then Text11.Text = "Samuel"
      If C = 7 Then Text11.Text = "Scott"
      If C = 8 Then Text11.Text = "Willie"
      If C = 9 Then Text11.Text = "Stanley"
      If C = 10 Then Text11.Text = "William"
      If C = 11 Then Text11.Text = "Burt"
      If C = 12 Then Text11.Text = "Bob"
      If C = 13 Then Text11.Text = "Steve"
      If C = 14 Then Text11.Text = "Jason"
      If C = 15 Then Text11.Text = "Dennis"
      If C = 16 Then Text11.Text = "Joe"
      If C = 17 Then Text11.Text = "Jim"
      If C = 18 Then Text11.Text = "John"
      If C = 19 Then Text11.Text = "Robert"
      If C = 20 Then Text11.Text = "Sidney"
      If C = 21 Then Text11.Text = "Walter"
      If C = 22 Then Text11.Text = "Keith"
      If C = 23 Then Text11.Text = "David"
      If C = 24 Then Text11.Text = "Gwen"
      If C = 25 Then Text11.Text = "Thomas"
      If C = 26 Then Text11.Text = "Bill"
      If C = 27 Then Text11.Text = "Don"
      If C = 28 Then Text11.Text = "James"
      If C = 29 Then Text11.Text = "Alfred"
      If C = 30 Then Text11.Text = "Andrew"
      If C = 31 Then Text11.Text = "Arnold"
      If C = 32 Then Text11.Text = "Bruce"
      If C = 33 Then Text11.Text = "Bryan"
      If C = 34 Then Text11.Text = "Clyde"
      If C = 35 Then Text11.Text = "Charles"
      If C = 36 Then Text11.Text = "Doug"
      If C = 37 Then Text11.Text = "Dwight"
      If C = 38 Then Text11.Text = "Dwayne"
      If C = 39 Then Text11.Text = "Edward"
      If C = 40 Then Text11.Text = "Eric"
      If C = 42 Then Text11.Text = "Chris"
      If C = 43 Then Text11.Text = "Mark"
      If C = 44 Then Text11.Text = "Harry"
      If C = 45 Then Text11.Text = "Jerry"
      If C = 46 Then Text11.Text = "Leroy"
      If C = 47 Then Text11.Text = "George"
      If C = 48 Then Text11.Text = "Samuel"
      If C = 49 Then Text11.Text = "Wesley"
      If C = 50 Then Text11.Text = "Terry"
      If C = 51 Then Text11.Text = "Jeff"
      If C = 52 Then Text11.Text = "Franklin"
      If C = 53 Then Text11.Text = "Melvin"
      If C = 54 Then Text11.Text = "Marcus"
      If C = 55 Then Text11.Text = "Rudy"
      If C = 56 Then Text11.Text = "Jimmy"
      If C = 57 Then Text11.Text = "Eddie"
      If C = 58 Then Text11.Text = "Lewis"
      If C = 59 Then Text11.Text = "Ralph"
      If C = 60 Then Text11.Text = "Kenneth"
      If C = 61 Then Text11.Text = "Phillip"
      If C = 62 Then Text11.Text = "Gregory"
      If C = 63 Then Text11.Text = "Howard"
      If C = 64 Then Text11.Text = "Geoff"
      If C = 65 Then Text11.Text = "Theo"
      If C = 66 Then Text11.Text = "Jeffrey"
      If C = 67 Then Text11.Text = "Paul"
      If C = 68 Then Text11.Text = "Cliff"
      If C = 69 Then Text11.Text = "Darrell"
      If C = 70 Then Text11.Text = "Christopher"
      If C = 71 Then Text11.Text = "Anthony"
      If C = 72 Then Text11.Text = "Chuck"
      If C = 73 Then Text11.Text = "Sam"
      If C = 74 Then Text11.Text = "Phil"
      If C = 75 Then Text11.Text = "Fredrick"
      If C = 76 Then Text11.Text = "Johnny"
      If C = 77 Then Text11.Text = "Hugh"
      If C = 78 Then Text11.Text = "Alvin"
      If C = 79 Then Text11.Text = "Alan"
      If C = 80 Then Text11.Text = "Buck"
      If C = 81 Then Text11.Text = "Bo"
      If C = 82 Then Text11.Text = "Dewitt"
      If C = 83 Then Text11.Text = "Tom"
      If C = 84 Then Text11.Text = "Tony"
      If C = 85 Then Text11.Text = "Lawrence"
      If C = 86 Then Text11.Text = "Roger"
      If C = 87 Then Text11.Text = "Marc"
      If C = 88 Then Text11.Text = "Cecil"
      If C = 89 Then Text11.Text = "Kevin"
      If C = 90 Then Text11.Text = "Maurice"
      If C = 91 Then Text11.Text = "Byron"
      If C = 92 Then Text11.Text = "Patrick"
      If C = 93 Then Text11.Text = "Johnie"
      If C = 94 Then Text11.Text = "Daniel"
      If C = 95 Then Text11.Text = "Edwin"
      If C = 96 Then Text11.Text = "Jesse"
      If C = 97 Then Text11.Text = "Leonard"
      If C = 98 Then Text11.Text = "Andy"
      If C = 99 Then Text11.Text = "Norman"
      If C = 100 Then Text11.Text = "Eugene"
      If C = 101 Then Text11.Text = "Jerrell"

    End If

  'Last Name

    upperbound4 = 457
    lowerbound4 = 1

    D = Int((upperbound4 - lowerbound4 + 1) * Rnd + lowerbound)

    If D = 1 Then Text12.Text = "Abbot"
    If D = 2 Then Text12.Text = "Abe"
    If D = 3 Then Text12.Text = "Abrams"
    If D = 4 Then Text12.Text = "Acker"
    If D = 5 Then Text12.Text = "Ackman"
    If D = 6 Then Text12.Text = "Adams"
    If D = 7 Then Text12.Text = "Adam"
    If D = 8 Then Text12.Text = "Adamson"
    If D = 9 Then Text12.Text = "Addison"
    If D = 10 Then Text12.Text = "Adkins"
    If D = 11 Then Text12.Text = "Adler"
    If D = 12 Then Text12.Text = "Allen"
    If D = 13 Then Text12.Text = "Alford"
    If D = 14 Then Text12.Text = "Allison"
    If D = 15 Then Text12.Text = "Alley"
    If D = 16 Then Text12.Text = "Allgood"
    If D = 17 Then Text12.Text = "Anderson"
    If D = 18 Then Text12.Text = "Amos"
    If D = 19 Then Text12.Text = "Applegate"
    If D = 20 Then Text12.Text = "Apple"
    If D = 21 Then Text12.Text = "Arnold"
    If D = 22 Then Text12.Text = "Ashby"
    If D = 23 Then Text12.Text = "Austin"
    If D = 24 Then Text12.Text = "Austin"
    If D = 25 Then Text12.Text = "Baker"
    If D = 26 Then Text12.Text = "Baldwin"
    If D = 27 Then Text12.Text = "Banks"
    If D = 28 Then Text12.Text = "Bannerman"
    If D = 29 Then Text12.Text = "Barber"
    If D = 30 Then Text12.Text = "Barfield"
    If D = 31 Then Text12.Text = "Barkley"
    If D = 32 Then Text12.Text = "Barnett"
    If D = 33 Then Text12.Text = "Barry"
    If D = 34 Then Text12.Text = "Basinger"
    If D = 35 Then Text12.Text = "Bays"
    If D = 36 Then Text12.Text = "Beck"
    If D = 37 Then Text12.Text = "Bell"
    If D = 38 Then Text12.Text = "Benjamin"
    If D = 39 Then Text12.Text = "Bennett"
    If D = 40 Then Text12.Text = "Bentley"
    If D = 41 Then Text12.Text = "Biggs"
    If D = 42 Then Text12.Text = "Binder"
    If D = 43 Then Text12.Text = "Blackburn"
    If D = 44 Then Text12.Text = "Blackwood"
    If D = 45 Then Text12.Text = "Blake"
    If D = 46 Then Text12.Text = "Bobbitt"
    If D = 47 Then Text12.Text = "Boggs"
    If D = 48 Then Text12.Text = "Bolden"
    If D = 49 Then Text12.Text = "Borden"
    If D = 50 Then Text12.Text = "Bosworth"
    If D = 51 Then Text12.Text = "Bowden"
    If D = 52 Then Text12.Text = "Bowen"
    If D = 53 Then Text12.Text = "Bowers"
    If D = 54 Then Text12.Text = "Bowman"
    If D = 55 Then Text12.Text = "Bradley"
    If D = 56 Then Text12.Text = "Brady"
    If D = 57 Then Text12.Text = "Briggs"
    If D = 58 Then Text12.Text = "Britt"
    If D = 59 Then Text12.Text = "Brown"
    If D = 60 Then Text12.Text = "Bullock"
    If D = 61 Then Text12.Text = "Burton"
    If D = 62 Then Text12.Text = "Byrd"
    If D = 63 Then Text12.Text = "Carlon"
    If D = 64 Then Text12.Text = "Carter"
    If D = 65 Then Text12.Text = "Casey"
    If D = 66 Then Text12.Text = "Cates"
    If D = 67 Then Text12.Text = "Caudle"
    If D = 68 Then Text12.Text = "Chambers"
    If D = 69 Then Text12.Text = "Chandler"
    If D = 70 Then Text12.Text = "Chapman"
    If D = 71 Then Text12.Text = "Chavis"
    If D = 72 Then Text12.Text = "Cheek"
    If D = 73 Then Text12.Text = "Clapp"
    If D = 74 Then Text12.Text = "Clark"
    If D = 75 Then Text12.Text = "Clarke"
    If D = 76 Then Text12.Text = "Clemmons"
    If D = 77 Then Text12.Text = "Cline"
    If D = 78 Then Text12.Text = "Cobb"
    If D = 79 Then Text12.Text = "Coble"
    If D = 80 Then Text12.Text = "Cole"
    If D = 81 Then Text12.Text = "Coleman"
    If D = 82 Then Text12.Text = "Collins"
    If D = 83 Then Text12.Text = "Combs"
    If D = 84 Then Text12.Text = "Compton"
    If D = 85 Then Text12.Text = "Conner"
    If D = 86 Then Text12.Text = "Connor"
    If D = 87 Then Text12.Text = "Conway"
    If D = 88 Then Text12.Text = "Cook"
    If D = 89 Then Text12.Text = "Cowan"
    If D = 90 Then Text12.Text = "Cox"
    If D = 91 Then Text12.Text = "Crawford"
    If D = 92 Then Text12.Text = "Crutchfield"
    If D = 93 Then Text12.Text = "Cunningham"
    If D = 94 Then Text12.Text = "Daniel"
    If D = 95 Then Text12.Text = "Darnell"
    If D = 96 Then Text12.Text = "Davis"
    If D = 97 Then Text12.Text = "Day"
    If D = 98 Then Text12.Text = "Dawson"
    If D = 99 Then Text12.Text = "Dixon"
    If D = 100 Then Text12.Text = "Douglas"
    If D = 101 Then Text12.Text = "Dunn"
    If D = 102 Then Text12.Text = "Edwards"
    If D = 103 Then Text12.Text = "Ellis"
    If D = 104 Then Text12.Text = "Ellison"
    If D = 105 Then Text12.Text = "Emmett"
    If D = 106 Then Text12.Text = "Erwin"
    If D = 107 Then Text12.Text = "Evans"
    If D = 108 Then Text12.Text = "Evenson"
    If D = 109 Then Text12.Text = "Everhart"
    If D = 110 Then Text12.Text = "Farlow"
    If D = 111 Then Text12.Text = "Fann"
    If D = 112 Then Text12.Text = "Ferguson"
    If D = 113 Then Text12.Text = "Fields"
    If D = 114 Then Text12.Text = "Fish"
    If D = 115 Then Text12.Text = "Fitzgerald"
    If D = 116 Then Text12.Text = "Flowers"
    If D = 117 Then Text12.Text = "Forrest"
    If D = 118 Then Text12.Text = "Foster"
    If D = 119 Then Text12.Text = "Fowler"
    If D = 120 Then Text12.Text = "Frazier"
    If D = 121 Then Text12.Text = "Free"
    If D = 122 Then Text12.Text = "Freeman"
    If D = 123 Then Text12.Text = "Frye"
    If D = 124 Then Text12.Text = "Fulk"
    If D = 125 Then Text12.Text = "Fulton"
    If D = 126 Then Text12.Text = "Furr"
    If D = 127 Then Text12.Text = "Galloway"
    If D = 128 Then Text12.Text = "Gardner"
    If D = 129 Then Text12.Text = "Gaulden"
    If D = 130 Then Text12.Text = "Gauley"
    If D = 131 Then Text12.Text = "Gerkin"
    If D = 132 Then Text12.Text = "Getters"
    If D = 133 Then Text12.Text = "Gibbs"
    If D = 134 Then Text12.Text = "Gibson"
    If D = 135 Then Text12.Text = "Gilbert"
    If D = 136 Then Text12.Text = "Gillett"
    If D = 137 Then Text12.Text = "Glass"
    If D = 138 Then Text12.Text = "Glenn"
    If D = 139 Then Text12.Text = "Glover"
    If D = 140 Then Text12.Text = "Goble"
    If D = 141 Then Text12.Text = "Godfrey"
    If D = 142 Then Text12.Text = "Goldston"
    If D = 143 Then Text12.Text = "Goldstein"
    If D = 144 Then Text12.Text = "Good"
    If D = 145 Then Text12.Text = "Gordon"
    If D = 146 Then Text12.Text = "Gorham"
    If D = 147 Then Text12.Text = "Gorrell"
    If D = 148 Then Text12.Text = "Gossett"
    If D = 149 Then Text12.Text = "Graham"
    If D = 150 Then Text12.Text = "Grant"
    If D = 151 Then Text12.Text = "Graves"
    If D = 152 Then Text12.Text = "Gray"
    If D = 153 Then Text12.Text = "Green"
    If D = 154 Then Text12.Text = "Garfield"
    If D = 155 Then Text12.Text = "Greene"
    If D = 156 Then Text12.Text = "Gregory"
    If D = 157 Then Text12.Text = "Gross"
    If D = 158 Then Text12.Text = "Gwynn"
    If D = 159 Then Text12.Text = "Hager"
    If D = 160 Then Text12.Text = "Haines"
    If D = 161 Then Text12.Text = "Hairston"
    If D = 162 Then Text12.Text = "Ham"
    If D = 163 Then Text12.Text = "Hamilton"
    If D = 164 Then Text12.Text = "Hammond"
    If D = 165 Then Text12.Text = "Hanks"
    If D = 166 Then Text12.Text = "Hampton"
    If D = 167 Then Text12.Text = "Hamrick"
    If D = 168 Then Text12.Text = "Harden"
    If D = 169 Then Text12.Text = "Hardy"
    If D = 170 Then Text12.Text = "Hare"
    If D = 171 Then Text12.Text = "Harper"
    If D = 172 Then Text12.Text = "Harrel"
    If D = 173 Then Text12.Text = "Harris"
    If D = 174 Then Text12.Text = "Harrison"
    If D = 175 Then Text12.Text = "Hart"
    If D = 176 Then Text12.Text = "Hatfield"
    If D = 177 Then Text12.Text = "Hayes"
    If D = 178 Then Text12.Text = "Haynes"
    If D = 179 Then Text12.Text = "Helms"
    If D = 180 Then Text12.Text = "Henderson"
    If D = 181 Then Text12.Text = "Hendricks"
    If D = 182 Then Text12.Text = "Henley"
    If D = 183 Then Text12.Text = "Henson"
    If D = 184 Then Text12.Text = "Hester"
    If D = 185 Then Text12.Text = "Hatt"
    If D = 186 Then Text12.Text = "Hickman"
    If D = 187 Then Text12.Text = "Hicks"
    If D = 188 Then Text12.Text = "Higgins"
    If D = 189 Then Text12.Text = "Hill"
    If D = 190 Then Text12.Text = "Hines"
    If D = 191 Then Text12.Text = "Hinshaw"
    If D = 192 Then Text12.Text = "Hinson"
    If D = 193 Then Text12.Text = "Hodges"
    If D = 194 Then Text12.Text = "Hobson"
    If D = 195 Then Text12.Text = "Holder"
    If D = 196 Then Text12.Text = "Holl    "
    If D = 197 Then Text12.Text = "Holiday"
    If D = 198 Then Text12.Text = "Holloway"
    If D = 199 Then Text12.Text = "Holt"
    If D = 200 Then Text12.Text = "Hoover"
    If D = 201 Then Text12.Text = "Hopkins"
    If D = 202 Then Text12.Text = "Horn"
    If D = 203 Then Text12.Text = "Hoskins"
    If D = 204 Then Text12.Text = "Hubbard"
    If D = 205 Then Text12.Text = "Hudson"
    If D = 206 Then Text12.Text = "Huff"
    If D = 207 Then Text12.Text = "Huffman"
    If D = 208 Then Text12.Text = "Hughes"
    If D = 209 Then Text12.Text = "Huneycutt"
    If D = 210 Then Text12.Text = "Hurst"
    If D = 211 Then Text12.Text = "Hutchers"
    If D = 212 Then Text12.Text = "Hylton"
    If D = 213 Then Text12.Text = "Irel    "
    If D = 214 Then Text12.Text = "Irwin"
    If D = 215 Then Text12.Text = "Isley"
    If D = 216 Then Text12.Text = "Jaby"
    If D = 217 Then Text12.Text = "Jackson"
    If D = 218 Then Text12.Text = "Jacob"
    If D = 219 Then Text12.Text = "Jarrell"
    If D = 220 Then Text12.Text = "Jester"
    If D = 221 Then Text12.Text = "Johnson"
    If D = 222 Then Text12.Text = "Jones"
    If D = 223 Then Text12.Text = "Jordan"
    If D = 224 Then Text12.Text = "Joseph"
    If D = 225 Then Text12.Text = "Joyner"
    If D = 226 Then Text12.Text = "Justice"
    If D = 227 Then Text12.Text = "Keith"
    If D = 228 Then Text12.Text = "Kelley"
    If D = 229 Then Text12.Text = "Kellis"
    If D = 230 Then Text12.Text = "Kendall"
    If D = 231 Then Text12.Text = "Kennedy"
    If D = 232 Then Text12.Text = "Kenny"
    If D = 233 Then Text12.Text = "Kim"
    If D = 234 Then Text12.Text = "Kimble"
    If D = 235 Then Text12.Text = "King"
    If D = 236 Then Text12.Text = "Kinsey"
    If D = 237 Then Text12.Text = "Kirkman"
    If D = 238 Then Text12.Text = "Kiser"
    If D = 239 Then Text12.Text = "Knight"
    If D = 240 Then Text12.Text = "Knott"
    If D = 241 Then Text12.Text = "Lamb"
    If D = 242 Then Text12.Text = "Lambert"
    If D = 243 Then Text12.Text = "Lambeth"
    If D = 244 Then Text12.Text = "Lacey"
    If D = 245 Then Text12.Text = "Lane"
    If D = 246 Then Text12.Text = "Lang"
    If D = 247 Then Text12.Text = "Langley"
    If D = 248 Then Text12.Text = "Lankford"
    If D = 249 Then Text12.Text = "Lashley"
    If D = 250 Then Text12.Text = "Lawson"
    If D = 251 Then Text12.Text = "Leary"
    If D = 252 Then Text12.Text = "Leggett"
    If D = 253 Then Text12.Text = "Legrant"
    If D = 254 Then Text12.Text = "Lemonds"
    If D = 255 Then Text12.Text = "Lanahan"
    If D = 256 Then Text12.Text = "Leonard"
    If D = 257 Then Text12.Text = "Lewis"
    If D = 258 Then Text12.Text = "Lineback"
    If D = 259 Then Text12.Text = "Lindsay"
    If D = 260 Then Text12.Text = "Little"
    If D = 261 Then Text12.Text = "Lloyd"
    If D = 262 Then Text12.Text = "Livengood"
    If D = 263 Then Text12.Text = "Locke"
    If D = 264 Then Text12.Text = "Long"
    If D = 265 Then Text12.Text = "Lord"
    If D = 266 Then Text12.Text = "Love"
    If D = 267 Then Text12.Text = "Lovett"
    If D = 268 Then Text12.Text = "Lucas"
    If D = 269 Then Text12.Text = "Lowe"
    If D = 270 Then Text12.Text = "Lyons"
    If D = 271 Then Text12.Text = "Mabe"
    If D = 272 Then Text12.Text = "Mackey"
    If D = 273 Then Text12.Text = "Maddox"
    If D = 274 Then Text12.Text = "Maness"
    If D = 275 Then Text12.Text = "Mangum"
    If D = 276 Then Text12.Text = "Mann"
    If D = 277 Then Text12.Text = "Manning"
    If D = 278 Then Text12.Text = "Marion"
    If D = 279 Then Text12.Text = "Marks"
    If D = 280 Then Text12.Text = "Marsh"
    If D = 281 Then Text12.Text = "Marshall"
    If D = 282 Then Text12.Text = "Martin"
    If D = 283 Then Text12.Text = "Martinez"
    If D = 284 Then Text12.Text = "Mathews"
    If D = 285 Then Text12.Text = "Mathis"
    If D = 286 Then Text12.Text = "Matthews"
    If D = 286 Then Text12.Text = "May"
    If D = 287 Then Text12.Text = "Mayer"
    If D = 288 Then Text12.Text = "McClain"
    If D = 289 Then Text12.Text = "McConnell"
    If D = 290 Then Text12.Text = "McCormick"
    If D = 291 Then Text12.Text = "McCraw"
    If D = 292 Then Text12.Text = "McCullough"
    If D = 293 Then Text12.Text = "McDougald"
    If D = 294 Then Text12.Text = "McPherson"
    If D = 295 Then Text12.Text = "Medlin"
    If D = 296 Then Text12.Text = "Melvin"
    If D = 297 Then Text12.Text = "Mentis"
    If D = 298 Then Text12.Text = "Meyer"
    If D = 299 Then Text12.Text = "Middleton"
    If D = 300 Then Text12.Text = "Miller"
    If D = 301 Then Text12.Text = "Mitchell"
    If D = 302 Then Text12.Text = "Montgomery"
    If D = 303 Then Text12.Text = "Moore"
    If D = 304 Then Text12.Text = "Morrison"
    If D = 305 Then Text12.Text = "Morrow"
    If D = 306 Then Text12.Text = "Morton"
    If D = 307 Then Text12.Text = "Moss"
    If D = 308 Then Text12.Text = "Murphy"
    If D = 309 Then Text12.Text = "Murray"
    If D = 310 Then Text12.Text = "Neal"
    If D = 311 Then Text12.Text = "Newnam"
    If D = 312 Then Text12.Text = "Newton"
    If D = 313 Then Text12.Text = "Nichols"
    If D = 314 Then Text12.Text = "Nixon"
    If D = 315 Then Text12.Text = "Norris"
    If D = 316 Then Text12.Text = "Norton"
    If D = 317 Then Text12.Text = "Nunn"
    If D = 318 Then Text12.Text = "O'Brien"
    If D = 319 Then Text12.Text = "O'Donnel"
    If D = 320 Then Text12.Text = "Oldham"
    If D = 321 Then Text12.Text = "Olds"
    If D = 322 Then Text12.Text = "Oliver"
    If D = 323 Then Text12.Text = "Olson"
    If D = 324 Then Text12.Text = "O'Neal"
    If D = 325 Then Text12.Text = "Osborne"
    If D = 326 Then Text12.Text = "Owens"
    If D = 327 Then Text12.Text = "Padgett"
    If D = 328 Then Text12.Text = "Pace"
    If D = 329 Then Text12.Text = "Page"
    If D = 330 Then Text12.Text = "Painter"
    If D = 331 Then Text12.Text = "Parker"
    If D = 332 Then Text12.Text = "Parrish"
    If D = 333 Then Text12.Text = "Parsons"
    If D = 334 Then Text12.Text = "Patel"
    If D = 335 Then Text12.Text = "Patterson"
    If D = 336 Then Text12.Text = "Peacock"
    If D = 337 Then Text12.Text = "Pennington"
    If D = 338 Then Text12.Text = "Perkins"
    If D = 339 Then Text12.Text = "Peterson"
    If D = 340 Then Text12.Text = "Phillips"
    If D = 341 Then Text12.Text = "Pickard"
    If D = 342 Then Text12.Text = "Pickett"
    If D = 343 Then Text12.Text = "Piper"
    If D = 344 Then Text12.Text = "Pitts"
    If D = 345 Then Text12.Text = "Pleasants"
    If D = 346 Then Text12.Text = "Poole"
    If D = 347 Then Text12.Text = "Pope"
    If D = 348 Then Text12.Text = "Posey"
    If D = 349 Then Text12.Text = "Potter"
    If D = 350 Then Text12.Text = "Powers"
    If D = 351 Then Text12.Text = "Prendergast"
    If D = 352 Then Text12.Text = "Pressley"
    If D = 353 Then Text12.Text = "Preston"
    If D = 354 Then Text12.Text = "Priddy"
    If D = 355 Then Text12.Text = "Pugh"
    If D = 356 Then Text12.Text = "Purcell"
    If D = 357 Then Text12.Text = "Purvis"
    If D = 358 Then Text12.Text = "Putnam"
    If D = 359 Then Text12.Text = "Quick"
    If D = 360 Then Text12.Text = "Raines"
    If D = 361 Then Text12.Text = "Ramsey"
    If D = 362 Then Text12.Text = "Rankin"
    If D = 363 Then Text12.Text = "Reese"
    If D = 364 Then Text12.Text = "Reeves"
    If D = 365 Then Text12.Text = "Reid"
    If D = 366 Then Text12.Text = "Richards"
    If D = 367 Then Text12.Text = "Rich"
    If D = 368 Then Text12.Text = "Riches"
    If D = 369 Then Text12.Text = "Richter"
    If D = 370 Then Text12.Text = "Riley"
    If D = 371 Then Text12.Text = "Ring"
    If D = 372 Then Text12.Text = "Ripley"
    If D = 373 Then Text12.Text = "Rippy"
    If D = 374 Then Text12.Text = "Rivera"
    If D = 375 Then Text12.Text = "Roberts"
    If D = 376 Then Text12.Text = "Robertson"
    If D = 377 Then Text12.Text = "Robinson"
    If D = 378 Then Text12.Text = "Rodgers"
    If D = 379 Then Text12.Text = "Rogers"
    If D = 380 Then Text12.Text = "Roseboro"
    If D = 381 Then Text12.Text = "Rosen"
    If D = 382 Then Text12.Text = "Ross"
    If D = 383 Then Text12.Text = "Rouse"
    If D = 384 Then Text12.Text = "Royal"
    If D = 385 Then Text12.Text = "Rucker"
    If D = 386 Then Text12.Text = "Russell"
    If D = 387 Then Text12.Text = "S    s"
    If D = 388 Then Text12.Text = "Saunders"
    If D = 389 Then Text12.Text = "Sawyer"
    If D = 390 Then Text12.Text = "Scott"
    If D = 391 Then Text12.Text = "Seawell"
    If D = 392 Then Text12.Text = "Shannon"
    If D = 393 Then Text12.Text = "Sharples"
    If D = 394 Then Text12.Text = "Shaw"
    If D = 395 Then Text12.Text = "Shelton"
    If D = 396 Then Text12.Text = "Shoffner"
    If D = 397 Then Text12.Text = "Short"
    If D = 398 Then Text12.Text = "Showfety"
    If D = 399 Then Text12.Text = "Simmons"
    If D = 400 Then Text12.Text = "Sims"
    If D = 401 Then Text12.Text = "Singer"
    If D = 402 Then Text12.Text = "Small"
    If D = 403 Then Text12.Text = "Smith"
    If D = 404 Then Text12.Text = "Snyder"
    If D = 405 Then Text12.Text = "Southern"
    If D = 406 Then Text12.Text = "Spencer"
    If D = 407 Then Text12.Text = "Staley"
    If D = 408 Then Text12.Text = "Stanley"
    If D = 409 Then Text12.Text = "Stephens"
    If D = 410 Then Text12.Text = "Stephenson"
    If D = 411 Then Text12.Text = "Stevens"
    If D = 412 Then Text12.Text = "Stewart"
    If D = 413 Then Text12.Text = "Stone"
    If D = 414 Then Text12.Text = "Stricklen"
    If D = 415 Then Text12.Text = "Swann"
    If D = 416 Then Text12.Text = "Sumner"
    If D = 417 Then Text12.Text = "Swaney"
    If D = 418 Then Text12.Text = "Talbert"
    If D = 419 Then Text12.Text = "Tate"
    If D = 420 Then Text12.Text = "Taylor"
    If D = 421 Then Text12.Text = "Thomas"
    If D = 422 Then Text12.Text = "Thompson"
    If D = 423 Then Text12.Text = "Thorne"
    If D = 424 Then Text12.Text = "Townsend"
    If D = 425 Then Text12.Text = "Troxler"
    If D = 426 Then Text12.Text = "Turner"
    If D = 427 Then Text12.Text = "Tuttle"
    If D = 428 Then Text12.Text = "Tyson"
    If D = 429 Then Text12.Text = "Umfleet"
    If D = 430 Then Text12.Text = "Underwood"
    If D = 431 Then Text12.Text = "Utley"
    If D = 432 Then Text12.Text = "Vance"
    If D = 433 Then Text12.Text = "Vanstory"
    If D = 434 Then Text12.Text = "Vines"
    If D = 435 Then Text12.Text = "Vaughan"
    If D = 436 Then Text12.Text = "Vestall"
    If D = 437 Then Text12.Text = "Walker"
    If D = 438 Then Text12.Text = "Wall"
    If D = 439 Then Text12.Text = "Wallace"
    If D = 440 Then Text12.Text = "Walters"
    If D = 441 Then Text12.Text = "Ward"
    If D = 442 Then Text12.Text = "Warren"
    If D = 443 Then Text12.Text = "Washington"
    If D = 444 Then Text12.Text = "Watkins"
    If D = 445 Then Text12.Text = "Watson"
    If D = 446 Then Text12.Text = "Wayne"
    If D = 447 Then Text12.Text = "Weaver"
    If D = 448 Then Text12.Text = "Webb"
    If D = 449 Then Text12.Text = "Webster"
    If D = 450 Then Text12.Text = "Wells"
    If D = 451 Then Text12.Text = "West"
    If D = 452 Then Text12.Text = "White"
    If D = 453 Then Text12.Text = "Wilkins"
    If D = 454 Then Text12.Text = "Williams"
    If D = 455 Then Text12.Text = "Wilson"
    If D = 456 Then Text12.Text = "Young"
    If D = 457 Then Text12.Text = "Zimmerman"

  'State

    upperbound5 = 52
    lowerbound5 = 1

    E = Int((upperbound5 - lowerbound5 + 1) * Rnd + lowerbound)

    If E = 1 Then Text16.Text = "AL"
    If E = 2 Then Text16.Text = "AK"
    If E = 3 Then Text16.Text = "AZ"
    If E = 4 Then Text16.Text = "AR"
    If E = 5 Then Text16.Text = "CA"
    If E = 6 Then Text16.Text = "CO"
    If E = 7 Then Text16.Text = "CT"
    If E = 8 Then Text16.Text = "DE"
    If E = 9 Then Text16.Text = "DC"
    If E = 10 Then Text16.Text = "FL"
    If E = 11 Then Text16.Text = "GA"
    If E = 12 Then Text16.Text = "HI"
    If E = 13 Then Text16.Text = "ID"
    If E = 14 Then Text16.Text = "IL"
    If E = 15 Then Text16.Text = "IN"
    If E = 16 Then Text16.Text = "IA"
    If E = 17 Then Text16.Text = "KS"
    If E = 18 Then Text16.Text = "KY"
    If E = 19 Then Text16.Text = "LA"
    If E = 20 Then Text16.Text = "ME"
    If E = 21 Then Text16.Text = "MD"
    If E = 22 Then Text16.Text = "MA"
    If E = 23 Then Text16.Text = "MI"
    If E = 24 Then Text16.Text = "MN"
    If E = 25 Then Text16.Text = "MS"
    If E = 26 Then Text16.Text = "MO"
    If E = 27 Then Text16.Text = "MT"
    If E = 28 Then Text16.Text = "NE"
    If E = 29 Then Text16.Text = "NV"
    If E = 30 Then Text16.Text = "NH"
    If E = 31 Then Text16.Text = "NJ"
    If E = 32 Then Text16.Text = "NM"
    If E = 33 Then Text16.Text = "NY"
    If E = 34 Then Text16.Text = "NC"
    If E = 35 Then Text16.Text = "ND"
    If E = 36 Then Text16.Text = "OH"
    If E = 37 Then Text16.Text = "OK"
    If E = 38 Then Text16.Text = "OR"
    If E = 39 Then Text16.Text = "PA"
    If E = 40 Then Text16.Text = "PR"
    If E = 41 Then Text16.Text = "RI"
    If E = 42 Then Text16.Text = "SC"
    If E = 43 Then Text16.Text = "SD"
    If E = 44 Then Text16.Text = "TN"
    If E = 45 Then Text16.Text = "TX"
    If E = 46 Then Text16.Text = "UT"
    If E = 47 Then Text16.Text = "VT"
    If E = 48 Then Text16.Text = "VA"
    If E = 49 Then Text16.Text = "WA"
    If E = 50 Then Text16.Text = "WV"
    If E = 51 Then Text16.Text = "WI"
    If E = 52 Then Text16.Text = "WY"


  'City/Town

    upperbound8 = 2
    lowerbound8 = 1

    F = Int((upperbound8 - lowerbound8 + 1) * Rnd + lowerbound)


    If E = 1 Then
      If F = 1 Then Text15.Text = "Birmingham"
      If F = 2 Then Text15.Text = "Montgomery"
    End If

    If E = 2 Then Text15.Text = "Juneau"

    If E = 3 Then Text15.Text = "Phoenix"

    If E = 4 Then Text15.Text = "Little Rock"

    If E = 5 Then
      If F = 1 Then Text15.Text = "Los Angeles"
      If F = 2 Then Text15.Text = "San Francisco"
    End If

    If E = 6 Then
      If F = 1 Then Text15.Text = "Colorado Springs"
      If F = 2 Then Text15.Text = "Denver"
    End If

    If E = 7 Then
      If F = 1 Then Text15.Text = "Hartford"
      If F = 2 Then Text15.Text = "Hartford"
    End If

    If E = 8 Then
      If F = 1 Then Text15.Text = "Dover"
      If F = 2 Then Text15.Text = "Dover"
    End If

    If E = 9 Then
      If F = 1 Then Text15.Text = "Washington"
      If F = 2 Then Text15.Text = "Washington"
    End If

    If E = 10 Then
      If F = 1 Then Text15.Text = "Miami"
      If F = 2 Then Text15.Text = "Tallahassee"
    End If

    If E = 11 Then
      If F = 1 Then Text15.Text = "Atlanta"
      If F = 2 Then Text15.Text = "Fayetteville"
    End If

    If E = 12 Then
      If F = 1 Then Text15.Text = "Honolulu"
      If F = 2 Then Text15.Text = "Honolulu"
    End If

    If E = 13 Then
      If F = 1 Then Text15.Text = "Boise"
      If F = 2 Then Text15.Text = "Boise"
    End If

    If E = 14 Then
      If F = 1 Then Text15.Text = "Chicago"
      If F = 2 Then Text15.Text = "Springfield"
    End If

    If E = 15 Then
      If F = 1 Then Text15.Text = "Indianapolis"
      If F = 2 Then Text15.Text = "South Bend"
    End If

    If E = 16 Then
      If F = 1 Then Text15.Text = "Des Moines"
      If F = 2 Then Text15.Text = "Sioux City"
    End If

    If E = 17 Then
      If F = 1 Then Text15.Text = "Dodge City"
      If F = 2 Then Text15.Text = "Topeka"
    End If

    If E = 18 Then
      If F = 1 Then Text15.Text = "Frankfort"
      If F = 2 Then Text15.Text = "Louisville"
    End If

    If E = 19 Then
      If F = 1 Then Text15.Text = "Baton Rouge"
      If F = 2 Then Text15.Text = "New Orleans"
    End If

    If E = 20 Then
      If F = 1 Then Text15.Text = "Augusta"
      If F = 2 Then Text15.Text = "Augusta"
    End If

    If E = 21 Then
      If F = 1 Then Text15.Text = "Annapolis"
      If F = 2 Then Text15.Text = "Baltimore"
    End If

    If E = 22 Then
      If F = 1 Then Text15.Text = "Boston"
      If F = 2 Then Text15.Text = "Springfield"
    End If

    If E = 23 Then
      If F = 1 Then Text15.Text = "Bay City"
      If F = 2 Then Text15.Text = "Detroit"
    End If

    If E = 24 Then
      If F = 1 Then Text15.Text = "Minneapolis"
      If F = 2 Then Text15.Text = "St. Paul"
    End If

    If E = 25 Then
      If F = 1 Then Text15.Text = "Jackson"
      If F = 2 Then Text15.Text = "Jackson"
    End If

    If E = 26 Then
      If F = 1 Then Text15.Text = "Jefferson City"
      If F = 2 Then Text15.Text = "Kansas City"
    End If

    If E = 27 Then
      If F = 1 Then Text15.Text = "Helena"
      If F = 2 Then Text15.Text = "Helena"
    End If

    If E = 28 Then
      If F = 1 Then Text15.Text = "Lincoln"
      If F = 2 Then Text15.Text = "Omaha"
    End If

    If E = 29 Then
      If F = 1 Then Text15.Text = "Carson City"
      If F = 2 Then Text15.Text = "Reno"
    End If

    If E = 30 Then
      If F = 1 Then Text15.Text = "Concord"
      If F = 2 Then Text15.Text = "Concord"
    End If

    If E = 31 Then
      If F = 1 Then Text15.Text = "Atlantic City"
      If F = 2 Then Text15.Text = "Trenton"
    End If

    If E = 32 Then
      If F = 1 Then Text15.Text = "Santa Fe"
      If F = 2 Then Text15.Text = "Santa Fe"
    End If

    If E = 33 Then
      If F = 1 Then Text15.Text = "Buffalo"
      If F = 2 Then Text15.Text = "New York"
    End If

    If E = 34 Then
      If F = 1 Then Text15.Text = "Charlotte"
      If F = 2 Then Text15.Text = "Raleigh"
    End If

    If E = 35 Then
      If F = 1 Then Text15.Text = "Bismarck"
      If F = 2 Then Text15.Text = "Bismarck"
    End If

    If E = 36 Then
      If F = 1 Then Text15.Text = "Akron"
      If F = 2 Then Text15.Text = "Columbus"
    End If

    If E = 37 Then
      If F = 1 Then Text15.Text = "Oklahoma City"
      If F = 2 Then Text15.Text = "Tulsa"
    End If

    If E = 38 Then
      If F = 1 Then Text15.Text = "Salem"
      If F = 2 Then Text15.Text = "Salem"
    End If

    If E = 39 Then
      If F = 1 Then Text15.Text = "Philadelphia"
      If F = 2 Then Text15.Text = "Pittsburgh"
    End If

    If E = 40 Then
      If F = 1 Then Text15.Text = "San Juan"
      If F = 2 Then Text15.Text = "San Juan"
    End If

    If E = 41 Then
      If F = 1 Then Text15.Text = "Providence"
      If F = 2 Then Text15.Text = "Providence"
    End If

    If E = 42 Then
      If F = 1 Then Text15.Text = "Columbia"
      If F = 2 Then Text15.Text = "Myrtle Beach"
    End If

    If E = 43 Then
      If F = 1 Then Text15.Text = "Pierre"
      If F = 2 Then Text15.Text = "Pierre"
    End If

    If E = 44 Then
      If F = 1 Then Text15.Text = "Memphis"
      If F = 2 Then Text15.Text = "Nashville"
    End If

    If E = 45 Then
      If F = 1 Then Text15.Text = "Austin"
      If F = 2 Then Text15.Text = "Dallas"
    End If

    If E = 46 Then
      If F = 1 Then Text15.Text = "Salt Lake City"
      If F = 2 Then Text15.Text = "Salt Lake City"
    End If

    If E = 47 Then
      If F = 1 Then Text15.Text = "Montpelier"
      If F = 2 Then Text15.Text = "Montpelier"
    End If

    If E = 48 Then
      If F = 1 Then Text15.Text = "Arlington"
      If F = 2 Then Text15.Text = "Richmond"
    End If

    If E = 49 Then
      If F = 1 Then Text15.Text = "Seatttle"
      If F = 2 Then Text15.Text = "Walla Walla"
    End If

    If E = 50 Then
      If F = 1 Then Text15.Text = "Charleston"
      If F = 2 Then Text15.Text = "Charleston"
    End If

    If E = 51 Then
      If F = 1 Then Text15.Text = "Green Bay"
      If F = 2 Then Text15.Text = "Milwaukee"
    End If

    If E = 52 Then
      If F = 1 Then Text15.Text = "Cheyenne"
      If F = 2 Then Text15.Text = "Cheyenne"
    End If

  'Area codes

    If E = 1 Then
      If F = 1 Then Text17.Text = "205"
      If F = 2 Then Text17.Text = "334"
    End If

    If E = 2 Then Text17.Text = "907"

    If E = 3 Then Text17.Text = "602"

    If E = 4 Then Text17.Text = "501"

    If E = 5 Then
      If F = 1 Then Text17.Text = "213"
      If F = 2 Then Text17.Text = "415"
    End If

    If E = 6 Then
      If F = 1 Then Text17.Text = "719"
      If F = 2 Then Text17.Text = "303"
    End If

    If E = 7 Then
      If F = 1 Then Text17.Text = "203"
      If F = 2 Then Text17.Text = "203"
    End If

    If E = 8 Then
      If F = 1 Then Text17.Text = "302"
      If F = 2 Then Text17.Text = "302"
    End If

    If E = 9 Then
      If F = 1 Then Text17.Text = "202"
      If F = 2 Then Text17.Text = "202"
    End If

    If E = 10 Then
      If F = 1 Then Text17.Text = "305"
      If F = 2 Then Text17.Text = "904"
    End If

    If E = 11 Then
      If F = 1 Then Text17.Text = "404"
      If F = 2 Then Text17.Text = "404"
    End If

    If E = 12 Then
      If F = 1 Then Text17.Text = "334"
      If F = 2 Then Text17.Text = "334"
    End If

    If E = 13 Then
      If F = 1 Then Text17.Text = "334"
      If F = 2 Then Text17.Text = "334"
    End If

    If E = 14 Then
      If F = 1 Then Text17.Text = "312"
      If F = 2 Then Text17.Text = "217"
    End If

    If E = 15 Then
      If F = 1 Then Text17.Text = "317"
      If F = 2 Then Text17.Text = "219"
    End If

    If E = 16 Then
      If F = 1 Then Text17.Text = "515"
      If F = 2 Then Text17.Text = "712"
    End If

    If E = 17 Then
      If F = 1 Then Text17.Text = "316"
      If F = 2 Then Text17.Text = "913"
    End If

    If E = 18 Then
      If F = 1 Then Text17.Text = "502"
      If F = 2 Then Text17.Text = "502"
    End If

    If E = 19 Then
      If F = 1 Then Text17.Text = "504"
      If F = 2 Then Text17.Text = "504"
    End If

    If E = 20 Then
      If F = 1 Then Text17.Text = "207"
      If F = 2 Then Text17.Text = "207"
    End If

    If E = 21 Then
      If F = 1 Then Text17.Text = "410"
      If F = 2 Then Text17.Text = "410"
    End If

    If E = 22 Then
      If F = 1 Then Text17.Text = "617"
      If F = 2 Then Text17.Text = "413"
    End If

    If E = 23 Then
      If F = 1 Then Text17.Text = "517"
      If F = 2 Then Text17.Text = "313"
    End If

    If E = 24 Then
      If F = 1 Then Text17.Text = "612"
      If F = 2 Then Text17.Text = "612"
    End If

    If E = 25 Then
      If F = 1 Then Text17.Text = "601"
      If F = 2 Then Text17.Text = "601"
    End If

    If E = 26 Then
      If F = 1 Then Text17.Text = "314"
      If F = 2 Then Text17.Text = "816"
    End If

    If E = 27 Then
      If F = 1 Then Text17.Text = "406"
      If F = 2 Then Text17.Text = "406"
    End If

    If E = 28 Then
      If F = 1 Then Text17.Text = "402"
      If F = 2 Then Text17.Text = "402"
    End If

    If E = 29 Then
      If F = 1 Then Text17.Text = "702"
      If F = 2 Then Text17.Text = "702"
    End If

    If E = 30 Then
      If F = 1 Then Text17.Text = "603"
      If F = 2 Then Text17.Text = "603"
    End If

    If E = 31 Then
      If F = 1 Then Text17.Text = "609"
      If F = 2 Then Text17.Text = "609"
    End If

    If E = 32 Then
      If F = 1 Then Text17.Text = "505"
      If F = 2 Then Text17.Text = "505"
    End If

    If E = 33 Then
      If F = 1 Then Text17.Text = "716"
      If F = 2 Then Text17.Text = "212"
    End If

    If E = 34 Then
      If F = 1 Then Text17.Text = "704"
      If F = 2 Then Text17.Text = "919"
    End If

    If E = 35 Then
      If F = 1 Then Text17.Text = "334"
      If F = 2 Then Text17.Text = "334"
    End If

    If E = 36 Then
      If F = 1 Then Text17.Text = "216"
      If F = 2 Then Text17.Text = "614"
    End If

    If E = 37 Then
      If F = 1 Then Text17.Text = "405"
      If F = 2 Then Text17.Text = "918"
    End If

    If E = 38 Then
      If F = 1 Then Text17.Text = "503"
      If F = 2 Then Text17.Text = "503"
    End If

    If E = 39 Then
      If F = 1 Then Text17.Text = "215"
      If F = 2 Then Text17.Text = "412"
    End If

    If E = 40 Then
      If F = 1 Then Text17.Text = "809"
      If F = 2 Then Text17.Text = "809"
    End If

    If E = 41 Then
      If F = 1 Then Text17.Text = "401"
      If F = 2 Then Text17.Text = "401"
    End If

    If E = 42 Then
      If F = 1 Then Text17.Text = "803"
      If F = 2 Then Text17.Text = "803"
    End If

    If E = 43 Then
      If F = 1 Then Text17.Text = "605"
      If F = 2 Then Text17.Text = "605"
    End If

    If E = 44 Then
      If F = 1 Then Text17.Text = "901"
      If F = 2 Then Text17.Text = "615"
    End If

    If E = 45 Then
      If F = 1 Then Text17.Text = "512"
      If F = 2 Then Text17.Text = "214"
    End If

    If E = 46 Then
      If F = 1 Then Text17.Text = "801"
      If F = 2 Then Text17.Text = "801"
    End If

    If E = 47 Then
      If F = 1 Then Text17.Text = "802"
      If F = 2 Then Text17.Text = "802"
    End If

    If E = 48 Then
      If F = 1 Then Text17.Text = "703"
      If F = 2 Then Text17.Text = "804"
    End If

    If E = 49 Then
      If F = 1 Then Text17.Text = "206"
      If F = 2 Then Text17.Text = "509"
    End If

    If E = 50 Then
      If F = 1 Then Text17.Text = "304"
      If F = 2 Then Text17.Text = "304"
    End If

    If E = 51 Then
      If F = 1 Then Text17.Text = "414"
      If F = 2 Then Text17.Text = "414"
    End If

    If E = 52 Then
      If F = 1 Then Text17.Text = "307"
      If F = 2 Then Text17.Text = "307"
    End If

    E = E
    F = F

  'Phone number prefix

    upperbound13 = 970
    lowerbound13 = 200

    H = Int((upperbound13 - lowerbound13 + 1) * Rnd + lowerbound13)

    Text18.Text = H

  'Rest of phone number

    upperbound14 = 9990
    lowerbound14 = 1000

    I = Int((upperbound14 - lowerbound14 + 1) * Rnd + lowerbound14)

    text19.Text = I

  'Street number

    upperbound15 = 3000
    lowerbound15 = 50

    J = Int((upperbound15 - lowerbound15 + 1) * Rnd + lowerbound14)

    Text13.Text = J

  'Street name
    upperbound16 = 70
    lowerbound16 = 1

    K = Int((upperbound16 - lowerbound16 + 1) * Rnd + lowerbound16)


    If K = 1 Then text14.Text = "River St."
    If K = 2 Then text14.Text = "Lindley Ave."
    If K = 3 Then text14.Text = "Ray St."
    If K = 4 Then text14.Text = "Chruch St."
    If K = 5 Then text14.Text = "Carriage Ln."
    If K = 6 Then text14.Text = "Ludwick Ln."
    If K = 7 Then text14.Text = "Quate Dr."
    If K = 8 Then text14.Text = "Summit Av."
    If K = 9 Then text14.Text = "Greenbrair Rd."
    If K = 10 Then text14.Text = "Olive St."
    If K = 11 Then text14.Text = "Pinelake Dr."
    If K = 12 Then text14.Text = "Orchard St."
    If K = 13 Then text14.Text = "Smokeridge Ln."
    If K = 14 Then text14.Text = "Coleridge Dr."
    If K = 15 Then text14.Text = "Branson Rd."
    If K = 16 Then text14.Text = "Cardigan St."
    If K = 17 Then text14.Text = "Main St."
    If K = 18 Then text14.Text = "Cottage Av."
    If K = 19 Then text14.Text = "Mr. Luther King Dr."
    If K = 20 Then text14.Text = "Westworth Dr."
    If K = 21 Then text14.Text = "Hillsdale St."
    If K = 22 Then text14.Text = "New Garden Rd."
    If K = 23 Then text14.Text = "Oak Ridge Rd."
    If K = 24 Then text14.Text = "Cascade Ln."
    If K = 25 Then text14.Text = "Lakebend Way"
    If K = 26 Then text14.Text = "Deerwood Dr."
    If K = 27 Then text14.Text = "Northwood Ave."
    If K = 28 Then text14.Text = "Woodstream Rd."
    If K = 29 Then text14.Text = "Leawood Dr."
    If K = 30 Then text14.Text = "Grayston Ln."
    If K = 31 Then text14.Text = "Knox Rd."
    If K = 32 Then text14.Text = "Bears Creek Rd."
    If K = 33 Then text14.Text = "Simons Sq."
    If K = 34 Then text14.Text = "Washington St."
    If K = 35 Then text14.Text = "Market St."
    If K = 36 Then text14.Text = "Country Dr."
    If K = 37 Then text14.Text = "Grove Ln."
    If K = 38 Then text14.Text = "Hilltop Rd."
    If K = 39 Then text14.Text = "Sherwood St."
    If K = 40 Then text14.Text = "Folger Rd."
    If K = 41 Then text14.Text = "Company Mill Ln."
    If K = 42 Then text14.Text = "Nelson St."
    If K = 43 Then text14.Text = "Hedrick Dr."
    If K = 44 Then text14.Text = "Cottage Pl."
    If K = 45 Then text14.Text = "United St."
    If K = 46 Then text14.Text = "Hewitt St."
    If K = 47 Then text14.Text = "Walker Ave."
    If K = 48 Then text14.Text = "Burns St."
    If K = 49 Then text14.Text = "Raintree Ln."
    If K = 50 Then text14.Text = "Juliet Dr."
    If K = 51 Then text14.Text = "Garlston Dr."
    If K = 52 Then text14.Text = "Kingston St."
    If K = 53 Then text14.Text = "Elm St."
    If K = 54 Then text14.Text = "Quaker Ave."
    If K = 55 Then text14.Text = "Rollingwood Dr."
    If K = 56 Then text14.Text = "Tower Rd."
    If K = 57 Then text14.Text = "Lakepoint Dr."
    If K = 58 Then text14.Text = "Chapman St."
    If K = 59 Then text14.Text = "Boyd Rd."
    If K = 60 Then text14.Text = "Ross Av."
    If K = 61 Then text14.Text = "Dogwood Dr."
    If K = 62 Then text14.Text = "Fisher Park Pl."
    If K = 63 Then text14.Text = "Murphy Rd."
    If K = 64 Then text14.Text = "Summertree Ave."
    If K = 65 Then text14.Text = "Bluemont Dr."
    If K = 66 Then text14.Text = "Weslo St."
    If K = 67 Then text14.Text = "Flint St."
    If K = 68 Then text14.Text = "Stillwell Rd."
    If K = 69 Then text14.Text = "Bingham St."
    If K = 70 Then text14.Text = "Kenview Ave."


  'Zip Code

    Dim Higher As Long
    Dim Lower As Long

    If E = 1 Then
      If F = 1 Then
        Higher = 35299
        Lower = 35200
      End If
      If F = 2 Then
        Higher = 36199
        Lower = 85050
      End If
    End If

    If E = 2 Then
      If F = 1 Then
        Higher = 99850  'And
        Lower = 99801
      End If
      If F = 2 Then
        Higher = 99850  'And
        Lower = 99801
      End If
    End If

    If E = 3 Then
      If F = 1 Then
        Higher = 85050  'And
        Lower = 85000
      End If
      If F = 2 Then
        Higher = 85050  'And
        Lower = 85000
      End If
    End If

    If E = 4 Then
      If F = 1 Then
        Higher = 72297  'And
        Lower = 72200
      End If
      If F = 2 Then
        Higher = 72297  'And
        Lower = 72200
      End If
    End If

    If E = 5 Then
      If F = 1 Then
        Higher = 90101  'And
        Lower = 90000
      End If
      If F = 2 Then
        Higher = 94175  'And
        Lower = 94100
      End If
    End If

    If E = 6 Then
      If F = 1 Then
        Higher = 80997  'And
        Lower = 80900
      End If
      If F = 2 Then
        Higher = 80299  'And
        Lower = 80202
      End If
    End If

    If E = 7 Then
      If F = 1 Then
        Higher = 6199   'And
        Lower = 6100
      End If
      If F = 2 Then
        Higher = 6199   'And
        Lower = 6100
      End If
    End If

    If E = 8 Then
      If F = 1 Then
        Higher = 19901  'And
        Lower = 19901
      End If
      If F = 2 Then
        Higher = 19901  'And
        Lower = 19901
      End If
    End If

    If E = 9 Then
      If F = 1 Then
        Higher = 20599  'And
        Lower = 20500
      End If
      If F = 2 Then
        Higher = 20599  'And
        Lower = 20500
      End If
    End If

    If E = 10 Then
      If F = 1 Then
        Higher = 33199  'And
        Lower = 33100
      End If
      If F = 2 Then
        Higher = 32301  'And
        Lower = 32312
      End If
    End If

    If E = 11 Then
      If F = 1 Then
        Higher = 30399  'And
        Lower = 30303
      End If
      If F = 2 Then
        Higher = 30214  'And
        Lower = 30214
      End If
    End If

    If E = 12 Then
      If F = 1 Then
        Higher = 96860  'And
        Lower = 96800
      End If
      If F = 2 Then
        Higher = 96860  'And
        Lower = 96800
      End If
    End If

    If E = 13 Then
      If F = 1 Then
        Higher = 83788  'And
        Lower = 83700
      End If
      If F = 2 Then
        Higher = 83788  'And
        Lower = 83700
      End If
    End If

    If E = 14 Then
      If F = 1 Then
        Higher = 60699  'And
        Lower = 60600
      End If
      If F = 2 Then
        Higher = 62786  'And
        Lower = 62700
      End If
    End If

    If E = 15 Then
      If F = 1 Then
        Higher = 46291  'And
        Lower = 46200
      End If
      If F = 2 Then
        Higher = 46637  'And
        Lower = 46600
      End If
    End If

    If E = 16 Then
      If F = 1 Then
        Higher = 50960  'And
        Lower = 50053
      End If
      If F = 2 Then
        Higher = 51109  'And
        Lower = 51100
      End If
    End If

    If E = 17 Then
      If F = 1 Then
        Higher = 67801  'And
        Lower = 67801
      End If
      If F = 2 Then
        Higher = 66618  'And
        Lower = 66600
      End If
    End If

    If E = 18 Then
      If F = 1 Then
        Higher = 40601  'And
        Lower = 40601
      End If
      If F = 2 Then
        Higher = 40231  'And
        Lower = 40200
      End If
    End If

    If E = 19 Then
      If F = 1 Then
        Higher = 70825  'And
        Lower = 70800
      End If
      If F = 2 Then
        Higher = 70149  'And
        Lower = 70100
      End If
    End If

    If E = 20 Then
      If F = 1 Then
        Higher = 4330   'And
        Lower = 4330
      End If
      If F = 2 Then
        Higher = 4330   'And
        Lower = 4330
      End If
    End If

    If E = 21 Then
      If F = 1 Then
        Higher = 21402  'And
        Lower = 21400
      End If
      If F = 2 Then
        Higher = 21239  'And
        Lower = 21299
      End If
    End If

    If E = 22 Then
      If F = 1 Then
        Higher = 2132   'And
        Lower = 2100
      End If
      If F = 2 Then
        Higher = 1101   'And
        Lower = 1101
      End If
    End If

    If E = 23 Then
      If F = 1 Then
        Higher = 48706  'And
        Lower = 48706
      End If
      If F = 2 Then
        Higher = 48200  'And
        Lower = 48200
      End If
    End If

    If E = 24 Then
      If F = 1 Then
        Higher = 55488  'And
        Lower = 55400
      End If
      If F = 2 Then
        Higher = 55150  'And
        Lower = 55100
      End If
    End If

    If E = 25 Then
      If F = 1 Then
        Higher = 39216  'And
        Lower = 39200
      End If
      If F = 2 Then
        Higher = 39216  'And
        Lower = 39200
      End If
    End If

    If E = 26 Then
      If F = 1 Then
        Higher = 65199  'And
        Lower = 65101
      End If
      If F = 2 Then
        Higher = 64120  'And
        Lower = 64100
      End If
    End If

    If E = 27 Then
      If F = 1 Then
        Higher = 59601  'And
        Lower = 59601
      End If
      If F = 2 Then
        Higher = 59601  'And
        Lower = 59601
      End If
    End If

    If E = 28 Then
      If F = 1 Then
        Higher = 68529  'And
        Lower = 68500
      End If
      If F = 2 Then
        Higher = 68138  'And
        Lower = 68046
      End If
    End If

    If E = 29 Then
      If F = 1 Then
        Higher = 89706  'And
        Lower = 89701
      End If
      If F = 2 Then
        Higher = 89506  'And
        Lower = 89500
      End If
    End If

    If E = 30 Then
      If F = 1 Then
        Higher = 3301   'And
        Lower = 3300
      End If
      If F = 2 Then
        Higher = 3301   'And
        Lower = 3300
      End If
    End If

    If E = 31 Then
      If F = 1 Then
        Higher = 8401   'And
        Lower = 8400
      End If
      If F = 2 Then
        Higher = 8600   'And
        Lower = 8600
      End If
    End If

    If E = 32 Then
      If F = 1 Then
        Higher = 87501  'And
        Lower = 87501
      End If
      If F = 2 Then
        Higher = 87501  'And
        Lower = 87501
      End If
    End If

    If E = 33 Then
      If F = 1 Then
        Higher = 14216  'And
        Lower = 14200
      End If
      If F = 2 Then
        Higher = 10242  'And
        Lower = 10001
      End If
    End If

    If E = 34 Then
      If F = 1 Then
        Higher = 28296  'And
        Lower = 28200
      End If
      If F = 2 Then
        Higher = 27615  'And
        Lower = 27600
      End If
    End If

    If E = 35 Then
      If F = 1 Then
        Higher = 58501  'And
        Lower = 58501
      End If
      If F = 2 Then
        Higher = 58501  'And
        Lower = 58501
      End If
    End If

    If E = 36 Then
      If F = 1 Then
        Higher = 44329  'And
        Lower = 44300
      End If
      If F = 2 Then
        Higher = 43230  'And
        Lower = 43200
      End If
    End If

    If E = 37 Then
      If F = 1 Then
        Higher = 73135  'And
        Lower = 73100
      End If
      If F = 2 Then
        Higher = 74120  'And
        Lower = 74100
      End If
    End If

    If E = 38 Then
      If F = 1 Then
        Higher = 97306  'And
        Lower = 97301
      End If
      If F = 2 Then
        Higher = 97306  'And
        Lower = 97301
      End If
    End If

    If E = 39 Then
      If F = 1 Then
        Higher = 19126  'And
        Lower = 19100
      End If
      If F = 2 Then
        Higher = 15237  'And
        Lower = 15200
      End If
    End If

    If E = 40 Then
      If F = 1 Then
        Higher = 901    'And
        Lower = 901
      End If
      If F = 2 Then
        Higher = 901    'And
        Lower = 901
      End If
    End If

    If E = 41 Then
      If F = 1 Then
        Higher = 2912   'And
        Lower = 2900
      End If
      If F = 2 Then Higher = 2912   'And Lower = 2900
    End If

    If E = 42 Then
      If F = 1 Then
        Higher = 29227  'And
        Lower = 29200
      End If
      If F = 2 Then Higher = 29572  'And Lower = 29572
    End If

    If E = 43 Then
      If F = 1 Then
        Higher = 57501  'And
        Lower = 57501
      End If
      If F = 2 Then
        Higher = 57501  'And
        Lower = 57501
      End If
    End If

    If E = 44 Then
      If F = 1 Then
        Higher = 38166  'And
        Lower = 38100
      End If
      If F = 2 Then
        Higher = 27250  'And
        Lower = 37200
      End If
    End If

    If E = 45 Then
      If F = 1 Then
        Higher = 78759  'And
        Lower = 78700
      End If
      If F = 2 Then
        Higher = 75230  'And
        Lower = 75200
      End If
    End If

    If E = 46 Then
      If F = 1 Then
        Higher = 84124  'And
        Lower = 84100
      End If
      If F = 2 Then
        Higher = 84124  'And
        Lower = 84100
      End If
    End If

    If E = 47 Then
      If F = 1 Then
        Higher = 5604   'And
        Lower = 5602
      End If
      If F = 2 Then
        Higher = 5604   'And
        Lower = 5602
      End If
    End If

    If E = 48 Then
      If F = 1 Then
        Higher = 22214  'And
        Lower = 22200
      End If
      If F = 2 Then
        Higher = 23237  'And
        Lower = 23219
      End If
    End If

    If E = 49 Then
      If F = 1 Then
        Higher = 98134  'And
        Lower = 98100
      End If
      If F = 2 Then
        Higher = 99362  'And
        Lower = 99362
      End If
    End If

    If E = 50 Then
      If F = 1 Then
        Higher = 25314  'And
        Lower = 25300
      End If
      If F = 2 Then
        Higher = 25314  'And
        Lower = 25300
      End If
    End If

    If E = 51 Then
      If F = 1 Then
        Higher = 54304  'And
        Lower = 54300
      End If
      If F = 2 Then
        Higher = 53228  'And
        Lower = 53200
      End If
    End If

    If E = 52 Then
      If F = 1 Then
        Higher = 82002  'And
        Lower = 82001
      End If
      If F = 2 Then
        Higher = 82002  'And
        Lower = 82001
      End If
    End If

    L = Int((Higher - Lower + 1) * Rnd + Lower)

    Text20.Text = L

    DoEvents
    DoEvents
    DoEvents

    If Text20.Text < 10000 Then Text20.Text = "0" + L
    If Text20.Text < 1000 Then Text20.Text = "00" + L

    Text13.Text = Text13.Text + " " + text14.Text
    Text18.Text = Text18.Text + "-" + text19.Text

    upperbound40 = 950000000
    lowerbound40 = 100000000

    CheckNum = Int((upperbound40 - lowerbound40 + 1) * Rnd + lowerbound40)
    CheckingNum = Str(CheckNum)
    CheckingNum = Trim$(CheckingNum)

  End Sub

  Sub Command3D6_Click ()
    Dim Response%
    If Text1.Text = "" Or Text3.Text = "" Or Text3.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text20.Text = "" Or Text15.Text = "" Or Text11.Text = "" Or Text12.Text = "" Then
      MsgBox "You have not completely filled out this form.", 16, "Account Creator"
      Exit Sub
    End If

    C1 = "D"
    D3 = "ZK"
    DF = "DL"
    CD = "EI"
    C2 = "a"
    C3 = " "
    C4 = "C"
    C5 = "h"
    DK = "K"
    CD = ">"
    C6 = "r"
    C7 = "o"
    DS = "b"
    SDK = "I"
    SAA = "t"
    C8 = "n"
    C9 = "i"
    AS1 = "J"
    T4 = "i"
    AS3 = "l"
    C10 = "c"

    CName = C1 + C2 + C3 + C4 + C5 + C6 + C7 + C8 + C9 + C10


    GoTo AccountStart

    AccountErr:
    MsgBox "An error has occured while trying to create your acccount.  Sign back on and try again.  If still unsuccessful, hit the ""Generate All"" button again, and copy down all the information AOHell generates and make the account manually.", 48, "Creation Error!"
    FAccount.Show
    Unload AccntStatus
    Exit Sub

    AccountStart:
    On Error GoTo AccountErr

    Cert = FindWindow("_AOL_Modal", "")
    Wel = FindChildByTitle(Cert, "Continue")


    If Wel = 0 Then
      Response = MsgBox("It does not appear that you are connected to AOL.  You must click the ""New User"" name on AOL, then hit the ""Sign On"" button.  Once you are connected and at the screen where you enter the certificate number, hit the Start button on the AOHell Account Creator.  If you are connected then AOHell couldn't detect it -- do you want to continue? Hit No if you are not connected", 276, "Must Be Connected!")
    'Possible Response values: 6--YES, 7--NO.
      If Response = 6 Then
        GoTo Proceed1
      ElseIf Response = 7 Then
        Exit Sub

      End If

    End If


    Proceed1:

    FAccount.Hide
    AccntStatus.Show



    AppActivate "America  Online"
    DoEvents
    Call Pause(1.5)

    lint = 100

    AccntStatus.Panel3D1.Caption = "Entering Cert Number..."

    CertSend = VBEnumChildWindows(Cert, 8)
    lpstring = Text6.Text
    SendString = SendMessageByString(CertSend, WM_SETTEXT, lint, lpstring)

    PassSend = VBEnumChildWindows(Cert, 10)
    lpstring = Text7.Text
    SendString = SendMessageByString(PassSend, WM_SETTEXT, lint, lpstring)
    DoEvents
    Call Pause(1)

    If StopCreate = 1 Then
      End

    End If

    VBMsg1.SubClasshWnd = Wel
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    If StopCreate = 1 Then
      End
    End If

    DoEvents
    AccntStatus.Panel3D1.Caption = "Scrolling past bullshit..."
    Call Pause(3.5)

    Cert2 = FindWindow("_AOL_Modal", "")

    While Cert2 = 0
      DoEvents
      Cert2 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If


    Wend
    Cont1 = FindChildByTitle(Cert2, "Continue")

    If StopCreate = 1 Then
      End
    End If

    VBMsg1.SubClasshWnd = Cont1
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
    DoEvents
    If StopCreate = 1 Then
      End
    End If

    Call Pause(3.5)

    Cert3 = FindWindow("_AOL_Modal", "")

    While Cert3 = 0
      DoEvents
      Cert3 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If



    Wend
    DoEvents

    AccntStatus.Panel3D1.Caption = "Entering name and address info..."

    FName = VBEnumChildWindows(Cert3, 4)
    lpstring = Text11.Text
    SendString = SendMessageByString(FName, WM_SETTEXT, lint, lpstring)

    LName = VBEnumChildWindows(Cert3, 6)
    lpstring = Text12.Text
    SendString = SendMessageByString(LName, WM_SETTEXT, lint, lpstring)
    If StopCreate = 1 Then
      End
    End If

    Address = VBEnumChildWindows(Cert3, 8)
    lpstring = Text13.Text
    SendString = SendMessageByString(Address, WM_SETTEXT, lint, lpstring)

    City = VBEnumChildWindows(Cert3, 10)
    lpstring = Text15.Text
    SendString = SendMessageByString(City, WM_SETTEXT, lint, lpstring)

    State = VBEnumChildWindows(Cert3, 12)
    lpstring = Text16.Text
    SendString = SendMessageByString(State, WM_SETTEXT, lint, lpstring)

    ZipCode = VBEnumChildWindows(Cert3, 14)
    lpstring = Text20.Text
    SendString = SendMessageByString(ZipCode, WM_SETTEXT, lint, lpstring)

    DayPhone = VBEnumChildWindows(Cert3, 18)
    lpstring = Text17.Text + "-" + Text18.Text
    SendString = SendMessageByString(DayPhone, WM_SETTEXT, lint, lpstring)

    EvenPhone = VBEnumChildWindows(Cert3, 20)
    lpstring = Text17.Text + "-" + Text18.Text
    SendString = SendMessageByString(EvenPhone, WM_SETTEXT, lint, lpstring)
    If StopCreate = 1 Then
      End
    End If

    If Check3D1.Value = True Then
      AccntStatus.Panel3D1.Caption = "Pausing for realism..."
      DoEvents
      Call Pause(8)
    End If

    Cont2 = FindChildByTitle(Cert3, "Continue")

    If StopCreate = 1 Then
      End
    End If

    VBMsg1.SubClasshWnd = Cont2
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
    DoEvents
    AccntStatus.Panel3D1.Caption = "Scrolling past more bullshit..."

    Call Pause(3.5)
    If StopCreate = 1 Then
      End
    End If

    Cert4 = FindWindow("_AOL_Modal", "")

    While Cert4 = 0
      DoEvents
      Cert4 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If


    Wend
    DoEvents

    Cont3 = FindChildByTitle(Cert4, "Continue")

    If StopCreate = 1 Then
      End
    End If

    VBMsg1.SubClasshWnd = Cont3
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
    Call Pause(3.5)


    Cert5 = FindWindow("_AOL_Modal", "")
    If StopCreate = 1 Then
      End
    End If

    While Cert5 = 0
      DoEvents
      Cert5 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If


    Wend
    DoEvents

    Cont4 = FindChildByTitle(Cert5, "Continue")

    If StopCreate = 1 Then
      End
    End If

    VBMsg1.SubClasshWnd = Cont4
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
    AccntStatus.Panel3D1.Caption = "Preparing Checking Account info..."

    Call Pause(3)

    If StopCreate = 1 Then
      End
    End If

    SendKeys "{TAB} "

    DoEvents
    Call Pause(2)

    DoEvents

    SendKeys "{TAB 2}"
    DoEvents
    Call Pause(.5)
    SendKeys "{DOWN 4}"
    DoEvents
    Call Pause(.5)
    SendKeys "{ENTER}"

    DoEvents
    Call Pause(3.3)


  'CardNumber = VBEnumChildWindows(Cert6, 4)
  'lpstring = Text1.Text
  'SendString = SendMessageByString(CardNumber, WM_SETTEXT, lint, lpstring)
  '
  'ExpDate = VBEnumChildWindows(Cert6, 6)
  'lpstring = Text2.Text
  'SendString = SendMessageByString(ExpDate, WM_SETTEXT, lint, lpstring)

  'BankName = VBEnumChildWindows(Cert6, 8)
  'lpstring = Text3.Text
  'SendString = SendMessageByString(BankName, WM_SETTEXT, lint, lpstring)

  'FName1 = VBEnumChildWindows(Cert6, 11)
  'lpstring = Text11.Text
  'SendString = SendMessageByString(FName1, WM_SETTEXT, lint, lpstring)

  'LName1 = VBEnumChildWindows(Cert6, 13)
  'lpstring = Text12.Text
  'SendString = SendMessageByString(LName1, WM_SETTEXT, lint, lpstring)
    DoEvents
    DoEvents
    If StopCreate = 1 Then
      End
    End If

    If Check3D1.Value = True Then
      AccntStatus.Panel3D1.Caption = "Pausing for realism..."
      DoEvents
      Call Pause(8)
    End If

  '  AccntStatus.Panel3D1.Caption = "Scrolling past yet more bullshit..."
  '
  'Cont5 = FindChildByTitle(Cert6, "Continue")
  '
    If StopCreate = 1 Then
      End
    End If

  ' VBMsg1.SubClasshWnd = Cont5
  ' VBMsg1.SendMessage = WM_LBUTTONDOWN
  ' VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
  ' Call Pause(3)

    Cert7 = FindWindow("_AOL_Modal", "")

    While Cert7 = 0
      DoEvents
      Cert7 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If

    Wend
    DoEvents
  'SendKeys "{TAB 3} "
    Cont5 = FindChildByTitle(Cert7, "Use Checking Account")
  '
    If StopCreate = 1 Then
      End
    End If

    VBMsg1.SubClasshWnd = Cont5
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP


    DoEvents
    AccntStatus.Panel3D1.Caption = "Entering Checking Info..."
    Call Pause(3.3)
    DoEvents

    Cert8 = FindWindow("_AOL_Modal", "")

    While Cert8 = 0
      DoEvents
      Cert8 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If

    Wend

    CardNumber = VBEnumChildWindows(Cert8, 4)
    lpstring = Text3.Text
    SendString = SendMessageByString(CardNumber, WM_SETTEXT, lint, lpstring)
  '
    ExpDate = VBEnumChildWindows(Cert8, 6)
    lpstring = Text15.Text
    SendString = SendMessageByString(ExpDate, WM_SETTEXT, lint, lpstring)

    BankName = VBEnumChildWindows(Cert8, 8)
    lpstring = Text16.Text
    SendString = SendMessageByString(BankName, WM_SETTEXT, lint, lpstring)

    FName1 = VBEnumChildWindows(Cert8, 11)
    lpstring = Text1.Text
    SendString = SendMessageByString(FName1, WM_SETTEXT, lint, lpstring)

    CheckNum = VBEnumChildWindows(Cert8, 13)
    lpstring = CheckingNum
    SendString = SendMessageByString(CheckNum, WM_SETTEXT, lint, lpstring)
    DoEvents


    SendKeys "{ENTER}"
    DoEvents
    Call Pause(3)
    AccntStatus.Panel3D1.Caption = "Scrolling past yet more bullshit..."
    SendKeys "{ENTER}"
    DoEvents
    Call Pause(3)
    SendKeys "{ENTER}"
    Call Pause(3)

    Cert8 = FindWindow("_AOL_Modal", "")

    While Cert8 = 0
      DoEvents
      Cert8 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If

    Wend

    Accept = FindChildByTitle(Cert8, "I ACCEPT")
    VBMsg1.SubClasshWnd = Accept
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
    Call Pause(3)

    Cert8 = FindWindow("_AOL_Modal", "")

    While Cert8 = 0
      DoEvents
      Cert8 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If

    Wend
    DoEvents
    Cont7 = FindChildByTitle(Cert8, "Continue")
    VBMsg1.SubClasshWnd = Cont7
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
    Call Pause(3)

    Cert9 = FindWindow("_AOL_Modal", "")

    While Cert9 = 0
      DoEvents
      Cert9 = FindWindow("_AOL_Modal", "")

      If StopCreate = 1 Then
        End
      End If


    Wend
    DoEvents
    Cont8 = FindChildByTitle(Cert9, "Continue")
    VBMsg1.SubClasshWnd = Cont8
    VBMsg1.SendMessage = WM_LBUTTONDOWN
    VBMsg1.SendMessage = WM_LBUTTONUP
    DoEvents
    Call Pause(2)
    AccntStatus.Panel3D1.Caption = "Account created: Process complete"
    DoEvents
    Call Pause(1.5)
    MsgBox "Your account has been created. Now you just have to enter the Screen Name and password you want to use.  The AOHell Fake Account Creator has been brought to you by " + CName + ".", 64, "Ready to Make Screen Name"
    Unload AccntStatus
    End

  End Sub

  Sub Command3D7_Click ()

    SendKeys "{F1}"

  End Sub

  Sub Command3D8_Click ()
    End

  End Sub

  Sub Form_Load ()
    Dim Success As Integer

    Success% = SetWindowPos%(FAccount.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    A = "a"
    B = "b"
    C = "c"
    D = "d"
    E = "e"
    F = "f"
    G = "g"
    H = "h"
    I = "i"
    J = "j"
    K = "k"
    L = "l"
    M = "m"
    N = "n"
    O = "o"
    P = "p"
    Q = "q"
    R = "r"
    S = "s"
    T = "t"
    U = "u"
    V = "v"
    W = "w"
    X = "x"
    Y = "y"
    Z = "z"
    SP = " "
    CA = "A"
    CB = "B"
    CC = "C"
    CD = "D"
    CE = "E"
    CF = "F"
    CG = "G"
    CH = "H"
    CI = "I"
    CJ = "J"
    CK = "K"
    CL = "L"
    CM = "M"
    CN = "N"
    CO = "O"
    CP = "P"
    CQ = "Q"
    CR = "R"
    CS = "S"
    CT = "T"
    CU = "U"
    CV = "V"
    CW = "W"
    CX = "X"
    CY = "Y"
    CZ = "Z"
    FAccount.Caption = "AOHell Fake Account Creator v1.0 Beta 5"    ' + CD + A + SP + CC + H + R + O + N + I + C

    Dim DefCert As String

    init1.Filename = "aohell.ini"
    init1.Application = "Options"
    init1.Parameter = "DefCert"
    DefCert = init1.Value
    DefCert1$ = App.Path + "\" + DefCert
    GoTo BegCer

    nexy:
    DefCert2 = App.Path + "\default.crt"


    Open DefCert2 For Input As 1
    Input #1, Cert11, Cert12, Cert21, Cert22, Cert31, Cert32, Cert41, Cert42, Cert51, Cert52, Cert61, Cert62
    Close #1
    Resume nexy1

    BegCer:
    On Error GoTo nexy
    Open DefCert1$ For Input As 1
    Input #1, Cert11, Cert12, Cert21, Cert22, Cert31, Cert32, Cert41, Cert42, Cert51, Cert52, Cert61, Cert62
    Close #1

    nexy1:
    Text6.Text = Cert11
    Text7.Text = Cert12

  End Sub

  Sub Form_Resize ()
    Dim Success As Integer

    Success% = SetWindowPos%(FAccount.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

  End Sub

  Sub Option3_Click ()

  End Sub

  Sub Option3D1_Click (Value As Integer)
    If Option3D1.Value = True Then
      Text6.Text = Cert11
      Text7.Text = Cert12
    End If

  End Sub

  Sub Option3D2_Click (Value As Integer)
    If Option3D2.Value = True Then
      Text6.Text = Cert21
      Text7.Text = Cert22
    End If

  End Sub

  Sub Option3D3_Click (Value As Integer)
    If Option3D3.Value = True Then
      Text6.Text = Cert31
      Text7.Text = Cert32
    End If

  End Sub

  Sub Option3D4_Click (Value As Integer)
    If Option3D4.Value = True Then
      Text6.Text = Cert41
      Text7.Text = Cert42
    End If

  End Sub

  Sub Option3D5_Click (Value As Integer)
    Text6.Text = Cert51
    Text7.Text = Cert52

  End Sub

  Sub Option3D6_Click (Value As Integer)
    Text6.Text = Cert61
    Text7.Text = Cert62

  End Sub

  Sub Option4_Click ()

  End Sub

  Sub Pause (ByVal nSecond As Single)
    Dim t0 As Single
    t0 = Timer
    Do While Timer - t0 < nSecond
      Dim dummy As Integer
      dummy = DoEvents()
      If Timer < t0 Then
        t0 = t0 - 24 * 60 * 60
      End If
    Loop
  End Sub

