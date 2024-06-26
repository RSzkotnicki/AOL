VERSION 2.00
Begin Form CreditWiz 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Credit Wizard v1.1  -  By Da Chronic and The Rizzer"
   ClientHeight    =   6225
   ClientLeft      =   390
   ClientTop       =   675
   ClientWidth     =   8790
   FillColor       =   &H00FF0000&
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Times New Roman"
   FontSize        =   9.75
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   ForeColor       =   &H00000000&
   Height          =   6915
   Icon            =   CCGEN.FRX:0000
   Left            =   330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8790
   Top             =   45
   Width           =   8910
   Begin TextBox Text1 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Text            =   "4228"
      Top             =   5280
      Width           =   615
   End
   Begin TextBox Text2 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   39
      Text            =   "1"
      Top             =   5760
      Width           =   615
   End
   Begin CommonDialog CMDialog1 
      Filter          =   "Credit Wizard 1.1(*.TXT)|*.TXT"
      Left            =   5280
      Top             =   5400
   End
   Begin SSCommand Command3D4 
      BevelWidth      =   3
      Caption         =   "&Save to file"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   4680
      Width           =   1335
   End
   Begin SSFrame Frame3D4 
      Caption         =   "Info Generator"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   4440
      TabIndex        =   17
      Top             =   120
      Width           =   4215
      Begin SSCommand Command3D8 
         BevelWidth      =   3
         Caption         =   "&Generate"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         Top             =   1440
         Width           =   975
      End
      Begin TextBox Text19 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   1440
         Width           =   735
      End
      Begin TextBox Text18 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   1440
         Width           =   495
      End
      Begin TextBox Text17 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Top             =   1440
         Width           =   495
      End
      Begin TextBox Text20 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin TextBox Text16 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin TextBox Text15 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin TextBox Text14 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin TextBox Text13 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin TextBox Text12 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin TextBox Text11 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin Label Label12 
         BackColor       =   &H8000000F&
         Caption         =   " -"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   1440
         Width           =   375
      End
      Begin Label Label11 
         BackColor       =   &H8000000F&
         Caption         =   "Phone:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   495
      End
      Begin Label Label10 
         BackColor       =   &H8000000F&
         Caption         =   "Zip"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   35
         Top             =   1080
         Width           =   375
      End
      Begin Label Label9 
         BackColor       =   &H8000000F&
         Caption         =   "St."
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   1080
         Width           =   255
      End
      Begin Label Label8 
         BackColor       =   &H8000000F&
         Caption         =   "City:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin Label Label7 
         BackColor       =   &H8000000F&
         Caption         =   "Address:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000F&
         Caption         =   "Last name"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   165
         Left            =   2040
         TabIndex        =   31
         Top             =   360
         Width           =   780
      End
      Begin Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000F&
         Caption         =   "First name:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Serif"
         FontSize        =   6.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   165
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
   End
   Begin SSCommand Command3D7 
      BevelWidth      =   3
      Caption         =   "Info On Program"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   5760
      Width           =   2055
   End
   Begin SSCommand Command3D6 
      BevelWidth      =   3
      Caption         =   "&Extrapolate"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin SSCommand Command3D5 
      BevelWidth      =   3
      Caption         =   "&Check a Card"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin TextBox Text3 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2160
      Width           =   3255
   End
   Begin Ctl3D Ctl3D1 
      ClassList       =   CCGEN.FRX:0302
      Left            =   5880
      Top             =   5640
   End
   Begin SSCommand Command3D3 
      BevelWidth      =   3
      Caption         =   "E&xit"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
   End
   Begin SSCommand Command3D2 
      BevelWidth      =   3
      Caption         =   "&Print"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin SSCommand Command3D1 
      BevelWidth      =   3
      Caption         =   "&Generate"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Card"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin SSOption Option3D3 
         Caption         =   "&Discover"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
      End
      Begin SSOption Option3D2 
         Caption         =   "&MasterCard"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1335
      End
      Begin SSOption Option3D1 
         Caption         =   "&Visa"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Value           =   -1  'True
         Width           =   855
      End
      Begin Image Image3 
         Height          =   600
         Left            =   2880
         Picture         =   CCGEN.FRX:0304
         Top             =   360
         Width           =   960
      End
      Begin Image Image2 
         Height          =   600
         Left            =   1560
         Picture         =   CCGEN.FRX:113E
         Top             =   360
         Width           =   960
      End
      Begin Image Image1 
         Height          =   600
         Left            =   240
         Picture         =   CCGEN.FRX:1F78
         Top             =   360
         Width           =   960
      End
   End
   Begin ListBox List5 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   2550
      Left            =   5160
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
   End
   Begin CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin Label Label14 
      BackColor       =   &H8000000F&
      Caption         =   "Bank Prefix:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   5280
      Width           =   1575
   End
   Begin Label Label13 
      BackColor       =   &H8000000F&
      Caption         =   "Number of cards:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5760
      Width           =   1575
   End
   Begin Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Credit Wizard 1.1"
      FontBold        =   0   'False
      FontItalic      =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   5400
      Width           =   2295
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Numbers:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Banks:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin Menu mnuCredit 
      Caption         =   "&Credit "
      Begin Menu mnuGen 
         Caption         =   "&Generate"
         Shortcut        =   ^G
      End
      Begin Menu mnuExtro 
         Caption         =   "&Extrapolate"
         Shortcut        =   ^E
      End
      Begin Menu mnuCheck 
         Caption         =   "C&heck"
         Shortcut        =   ^H
      End
      Begin Menu mnuhyp 
         Caption         =   "-"
      End
      Begin Menu mnuPrnt 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin Menu mnuhyp2 
         Caption         =   "-"
      End
      Begin Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
  Dim numcards As Integer

  Sub command2_click ()
    Dim er As Double


    Randomize

    er# = Rnd * 10 ^ 16
    crp# = er#


    Print er#
    For x = 15 To 0 Step -1

      jack = crp# / 10 ^ x
      stuff# = Int(jack)
      Print stuff#

    Next x









  End Sub

  Sub command3_click ()
    PrintForm


  End Sub

  Sub Command3D1_Click ()
    holder = 1

    text3.Text = ""

    If text2.Text = "" Then
      MsgBox "Please enter the amount of cards you wish to generate.", 64
      GoTo endall
    End If

    If Text1.Text = "" Then
      MsgBox "Please enter bank prefix.", 64
      GoTo endall
    End If


    On Error GoTo titfuck



  'For a card with an even number of digits, double every odd digit and
  'subtract 9 if the product is greater than 10.  Add up all the even
  'digits as well as the doubled-odd digits, and the result must be a
  'multiple of 10 or it's not a valid card.  If the card has an odd
  'number of digits, perform the same addition doubling the even digits
  'instead.

    Randomize

    Static car(16)
  'list5.additem = 1
    BANK = Text1.Text

    If text2 = "" Then
      MsgBox "Please Enter in the Number of Cards you want"
      Exit Sub
    End If

    numcards = text2
    If numcards = 0 Then
      MsgBox "Please Enter in the Number of Cards you want"
      Exit Sub
    End If




    For Times = 1 To numcards

top1:
    're-initialize cars
      For xxx = 1 To 16
        car(xxx) = 0
      Next xxx

      Randomize

      cardsnum = Rnd * 10000
      cardsnum2 = Rnd * 10000
      cardsnum3 = Rnd * 10000
      cardnum = Int(cardsnum)
      CARDNUM2 = Int(cardsnum2)
      cardnum3 = Int(cardsnum3)





      Cr1 = cardnum \ 1000
      Cra = Cr1 * 1000
      car(1) = Cr1

      CRB = cardnum - Cra


      cr2 = CRB \ 100
      Crq = cr2 * 100
      car(2) = cr2

      cr3 = CRB - Crq
      crz = cr3 \ 10
      crp = crz * 10
      car(3) = crz

      Cr4 = cr3 - crp
      car(4) = Cr4



    'next




      Cb1 = CARDNUM2 \ 1000
      Cba = Cb1 * 1000
      car(5) = Cb1

      CbB = CARDNUM2 - Cba


      cb2 = CbB \ 100
      Cbq = cb2 * 100
      car(6) = cb2

      cb3 = CbB - Cbq
      cbz = cb3 \ 10
      cbp = cbz * 10
      car(7) = cbz

      Cb4 = cb3 - cbp
      car(8) = Cb4

    'next




      Cv1 = cardnum3 \ 1000
      Cva = Cv1 * 1000
      car(9) = Cv1

      CvB = cardnum3 - Cva


      cv2 = CvB \ 100
      Cvq = cv2 * 100
      car(10) = cv2

      cv3 = CvB - Cvq
      cvz = cv3 \ 10
      cvp = cvz * 10
      car(11) = cvz

      Cv4 = cv3 - cvp
      car(12) = Cv4

    'next

      Cw1 = BANK \ 1000
      Cwa = Cw1 * 1000
      car(13) = Cw1

      CwB = BANK - Cwa


      cw2 = CwB \ 100
      Cwq = cw2 * 100
      car(14) = cw2

      cw3 = CwB - Cwq
      cwz = cw3 \ 10
      cwp = cwz * 10
      car(15) = cwz

      Cw4 = cw3 - cwp
      car(16) = Cw4

    'next

      For rot = 1 To 15 Step 2


        car(rot) = car(rot) * 2
        If car(rot) > 9 Then

          car(rot) = car(rot) - 9
        End If

      Next rot
      sum = 0

      For CT = 1 To 16
        sum = sum + car(CT)
      Next CT
      shit2 = sum
      shit3 = shit2 \ 10
      shit4 = shit3 * 10
      shit5 = shit2 - shit4

      If shit5 = 0 Then
        x0 = Str$(BANK)

        x = Str$(cardnum)
        X1 = Str$(CARDNUM2)
        X2 = Str$(cardnum3)

        hl = Str$(holder)

        text3.Text = text3.Text & hl & ".)" & x0 & " " & Format$(x, "0###") & " " & Format$(X1, "0###") & " " & Format$(X2, "0###") & Chr$(13) & Chr$(10)
        holder = holder + 1





      Else
        GoTo top1
      End If
    'list1.Height = list1.Height + 195

    Next Times
    GoTo endall

titfuck:
    MsgBox "You may only generate up to 14 cards at a time."
    GoTo endall

endall:

  End Sub

  Sub Command3D2_Click ()
  '===BEGIN MAKING OF INFO TEXT===
    NewLine = Chr$(13) + Chr$(10)
    info = text11.Text + " " + text12.Text + NewLine + text13.Text + " " + text14.Text + NewLine + text15.Text + ", " + text16.Text + ", " + text20.Text + NewLine + text17.Text + "-" + text18.Text + "-" + text19.Text
  '===END MAKING OF INFO TEXT===

    jacko1 = list5.List(list5.ListIndex)
    jackass2 = Right(jacko1, 35)

    Printer.Print "                               Credit Wizard v1.1"
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print "                 " + jackass2
    Printer.Print " "
    Printer.Print
    Printer.Print text3.Text
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print info
    Printer.Print " "
    Printer.Print "                               Credit Wizard v1.1"
    Printer.EndDoc

  End Sub

  Sub Command3D3_Click ()
    End

  End Sub

  Sub Command3D4_Click ()
    GoTo begin2

titfuck2:
    Exit Sub

begin2:
    On Error GoTo titfuck2

  '===BEGIN MAKING OF INFO TEXT===
    NewLine = Chr$(13) + Chr$(10)
    info = text11.Text + " " + text12.Text + NewLine + text13.Text + " " + text14.Text + NewLine + text15.Text + ", " + text16.Text + ", " + text20.Text + NewLine + text17.Text + "-" + text18.Text + "-" + text19.Text
  '===END MAKING OF INFO TEXT===

    jacko1 = list5.List(list5.ListIndex)
    jackass2 = Right(jacko1, 35)

    done = jackass2 + NewLine + text3.Text + NewLine + info

    CMDIALOG1.Action = 2

    xfilnam$ = CMDIALOG1.Filename


    Open xfilnam$ For Output As #1

    Write #1, done

    Close #1

  End Sub

  Sub Command3D5_Click ()


    cardcheck.Show


  End Sub

  Sub Command3D6_Click ()


    extrapolate.Show


  End Sub

  Sub Command3D7_Click ()

    A = "a"
    b = "b"
    C = "c"
    d = "d"
    E = "e"
    F = "f"
    G = "g"
    H = "h"
    I = "i"
    J = "j"
    K = "k"
    l = "l"
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
    x = "x"
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


    chron = CD + A + SP + CC + H + R + O + N + I + C

    rizz = SP + A + N + d + SP + CT + H + E + SP + CR + I + Z + Z + E + R + SP

    MsgBox "Credit Wizard v1.1" + Chr(13) + " " + Chr(13) + "By: " + chron + rizz + Chr(13) + " " + Chr(13) + "Greetings to: " + Chr(13) + " " + Chr(13) + "Soul Crusher, Blue Demon, Hell Pup, Tantalos, and Toof"


  End Sub

  Sub Command3D8_Click ()
'// If you use this code in your program and do not
'// give Chronic the proper credit you are a dickless
'// terd!  :)  Be a man, give credit where credit is due


    Randomize

  'Female First Name

    upperbound = 2
    lowerbound = 1

    A = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

    rndanswer = A
  ' Text14.Text = A

    If A = 2 Then

      upperbound1 = 113
      lowerbound1 = 1

      b = Int((upperbound1 - lowerbound1 + 1) * Rnd + lowerbound)

      rndanswer2 = b

      If b = 1 Then text11.Text = "Maria"
      If b = 2 Then text11.Text = "Jane"
      If b = 3 Then text11.Text = "Stacy"
      If b = 4 Then text11.Text = "Merideth"
      If b = 5 Then text11.Text = "Beth"
      If b = 6 Then text11.Text = "Angela"
      If b = 7 Then text11.Text = "Mary"
      If b = 8 Then text11.Text = "Tiffany"
      If b = 9 Then text11.Text = "Pamela"
      If b = 10 Then text11.Text = "April"
      If b = 11 Then text11.Text = "Louise"
      If b = 12 Then text11.Text = "Virginia"
      If b = 13 Then text11.Text = "Nicole"
      If b = 14 Then text11.Text = "Carolyn"
      If b = 15 Then text11.Text = "Julia"
      If b = 16 Then text11.Text = "Trina"
      If b = 17 Then text11.Text = "Robin"
      If b = 18 Then text11.Text = "Karen"
      If b = 19 Then text11.Text = "Susan"
      If b = 20 Then text11.Text = "Kathy"
      If b = 21 Then text11.Text = "Shirley"
      If b = 22 Then text11.Text = "Renee"
      If b = 23 Then text11.Text = "Olivia"
      If b = 24 Then text11.Text = "Beth"
      If b = 25 Then text11.Text = "Gladys"
      If b = 26 Then text11.Text = "Dianne"
      If b = 27 Then text11.Text = "Ethelene"
      If b = 28 Then text11.Text = "Patricia"
      If b = 29 Then text11.Text = "Anne"
      If b = 30 Then text11.Text = "Kathleen"
      If b = 31 Then text11.Text = "Carol"
      If b = 32 Then text11.Text = "Kimberly"
      If b = 33 Then text11.Text = "Marcia"
      If b = 34 Then text11.Text = "Marsha"
      If b = 35 Then text11.Text = "Sharon"
      If b = 36 Then text11.Text = "Ella"
      If b = 37 Then text11.Text = "Evelyn"
      If b = 38 Then text11.Text = "Glenda"
      If b = 39 Then text11.Text = "Katie"
      If b = 40 Then text11.Text = "Kate"
      If b = 41 Then text11.Text = "Ann"
      If b = 42 Then text11.Text = "Dana"
      If b = 43 Then text11.Text = "Shannon"
      If b = 44 Then text11.Text = "Elizabeth"
      If b = 45 Then text11.Text = "Lillian"
      If b = 46 Then text11.Text = "Jo Ann"
      If b = 47 Then text11.Text = "Karen"
      If b = 48 Then text11.Text = "Edith"
      If b = 49 Then text11.Text = "Jenny"
      If b = 50 Then text11.Text = "Deborah"
      If b = 51 Then text11.Text = "Elisa"
      If b = 52 Then text11.Text = "Becky"
      If b = 53 Then text11.Text = "Barbara"
      If b = 54 Then text11.Text = "Blair"
      If b = 55 Then text11.Text = "Cynthia"
      If b = 56 Then text11.Text = "Debbie"
      If b = 57 Then text11.Text = "Diane"
      If b = 58 Then text11.Text = "Margaret"
      If b = 59 Then text11.Text = "Racheal"
      If b = 60 Then text11.Text = "Shari"
      If b = 61 Then text11.Text = "Gail"
      If b = 62 Then text11.Text = "Sarah"
      If b = 63 Then text11.Text = "Cathy"
      If b = 64 Then text11.Text = "Katrina"
      If b = 65 Then text11.Text = "Marlene"
      If b = 66 Then text11.Text = "Patricia"
      If b = 67 Then text11.Text = "Lisa"
      If b = 68 Then text11.Text = "Candace"
      If b = 69 Then text11.Text = "Beatrice"
      If b = 70 Then text11.Text = "Betty"
      If b = 71 Then text11.Text = "Wendy"
      If b = 72 Then text11.Text = "Denise"
      If b = 73 Then text11.Text = "Rebecca"
      If b = 74 Then text11.Text = "Arlene"
      If b = 75 Then text11.Text = "Stephanie"
      If b = 76 Then text11.Text = "Flora"
      If b = 77 Then text11.Text = "Martha"
      If b = 78 Then text11.Text = "Nancy"
      If b = 79 Then text11.Text = "Tanya"
      If b = 80 Then text11.Text = "Helen"
      If b = 81 Then text11.Text = "Patty"
      If b = 82 Then text11.Text = "Della"
      If b = 83 Then text11.Text = "Ida"
      If b = 84 Then text11.Text = "Lisa"
      If b = 85 Then text11.Text = "Jeanne"
      If b = 86 Then text11.Text = "Theresa"
      If b = 87 Then text11.Text = "Vanessa"
      If b = 88 Then text11.Text = "Rose"
      If b = 89 Then text11.Text = "Benita"
      If b = 90 Then text11.Text = "Lucy"
      If b = 91 Then text11.Text = "Suzette"
      If b = 92 Then text11.Text = "Jessica"
      If b = 93 Then text11.Text = "Blair"
      If b = 94 Then text11.Text = "Gina"
      If b = 95 Then text11.Text = "Melissa"
      If b = 96 Then text11.Text = "Nikki"
      If b = 97 Then text11.Text = "Victoria"
      If b = 98 Then text11.Text = "Priscilla"
      If b = 99 Then text11.Text = "Catherine"
      If b = 100 Then text11.Text = "Joan"
      If b = 101 Then text11.Text = "Sue"
      If b = 102 Then text11.Text = "Sonja"
      If b = 103 Then text11.Text = "Stephanie"
      If b = 104 Then text11.Text = "Cindy"
      If b = 105 Then text11.Text = "Alison"
      If b = 106 Then text11.Text = "Samantha"
      If b = 107 Then text11.Text = "Kelly"
      If b = 108 Then text11.Text = "Carroll"
      If b = 109 Then text11.Text = "Lauren"
      If b = 110 Then text11.Text = "Louis"
      If b = 111 Then text11.Text = "Alicia"
      If b = 112 Then text11.Text = "Sylvia"
      If b = 113 Then text11.Text = "Jenn"

    End If

  'Male First Name

    If A = 1 Then

      upperbound2 = 101
      lowerbound2 = 1

      C = Int((upperbound2 - lowerbound2 + 1) * Rnd + lowerbound)

      If C = 1 Then text11.Text = "Michael"
      If C = 2 Then text11.Text = "Mike"
      If C = 3 Then text11.Text = "Randy"
      If C = 4 Then text11.Text = "Richard"
      If C = 5 Then text11.Text = "Ronald"
      If C = 6 Then text11.Text = "Samuel"
      If C = 7 Then text11.Text = "Scott"
      If C = 8 Then text11.Text = "Willie"
      If C = 9 Then text11.Text = "Stanley"
      If C = 10 Then text11.Text = "William"
      If C = 11 Then text11.Text = "Burt"
      If C = 12 Then text11.Text = "Bob"
      If C = 13 Then text11.Text = "Steve"
      If C = 14 Then text11.Text = "Jason"
      If C = 15 Then text11.Text = "Dennis"
      If C = 16 Then text11.Text = "Joe"
      If C = 17 Then text11.Text = "Jim"
      If C = 18 Then text11.Text = "John"
      If C = 19 Then text11.Text = "Robert"
      If C = 20 Then text11.Text = "Sidney"
      If C = 21 Then text11.Text = "Walter"
      If C = 22 Then text11.Text = "Keith"
      If C = 23 Then text11.Text = "David"
      If C = 24 Then text11.Text = "Gwen"
      If C = 25 Then text11.Text = "Thomas"
      If C = 26 Then text11.Text = "Bill"
      If C = 27 Then text11.Text = "Don"
      If C = 28 Then text11.Text = "James"
      If C = 29 Then text11.Text = "Alfred"
      If C = 30 Then text11.Text = "Andrew"
      If C = 31 Then text11.Text = "Arnold"
      If C = 32 Then text11.Text = "Bruce"
      If C = 33 Then text11.Text = "Bryan"
      If C = 34 Then text11.Text = "Clyde"
      If C = 35 Then text11.Text = "Charles"
      If C = 36 Then text11.Text = "Doug"
      If C = 37 Then text11.Text = "Dwight"
      If C = 38 Then text11.Text = "Dwayne"
      If C = 39 Then text11.Text = "Edward"
      If C = 40 Then text11.Text = "Eric"
      If C = 42 Then text11.Text = "Chris"
      If C = 43 Then text11.Text = "Mark"
      If C = 44 Then text11.Text = "Harry"
      If C = 45 Then text11.Text = "Jerry"
      If C = 46 Then text11.Text = "Leroy"
      If C = 47 Then text11.Text = "George"
      If C = 48 Then text11.Text = "Samuel"
      If C = 49 Then text11.Text = "Wesley"
      If C = 50 Then text11.Text = "Terry"
      If C = 51 Then text11.Text = "Jeff"
      If C = 52 Then text11.Text = "Franklin"
      If C = 53 Then text11.Text = "Melvin"
      If C = 54 Then text11.Text = "Marcus"
      If C = 55 Then text11.Text = "Rudy"
      If C = 56 Then text11.Text = "Jimmy"
      If C = 57 Then text11.Text = "Eddie"
      If C = 58 Then text11.Text = "Lewis"
      If C = 59 Then text11.Text = "Ralph"
      If C = 60 Then text11.Text = "Kenneth"
      If C = 61 Then text11.Text = "Phillip"
      If C = 62 Then text11.Text = "Gregory"
      If C = 63 Then text11.Text = "Howard"
      If C = 64 Then text11.Text = "Geoff"
      If C = 65 Then text11.Text = "Theo"
      If C = 66 Then text11.Text = "Jeffrey"
      If C = 67 Then text11.Text = "Paul"
      If C = 68 Then text11.Text = "Cliff"
      If C = 69 Then text11.Text = "Darrell"
      If C = 70 Then text11.Text = "Christopher"
      If C = 71 Then text11.Text = "Anthony"
      If C = 72 Then text11.Text = "Chuck"
      If C = 73 Then text11.Text = "Sam"
      If C = 74 Then text11.Text = "Phil"
      If C = 75 Then text11.Text = "Fredrick"
      If C = 76 Then text11.Text = "Johnny"
      If C = 77 Then text11.Text = "Hugh"
      If C = 78 Then text11.Text = "Alvin"
      If C = 79 Then text11.Text = "Alan"
      If C = 80 Then text11.Text = "Buck"
      If C = 81 Then text11.Text = "Bo"
      If C = 82 Then text11.Text = "Dewitt"
      If C = 83 Then text11.Text = "Tom"
      If C = 84 Then text11.Text = "Tony"
      If C = 85 Then text11.Text = "Lawrence"
      If C = 86 Then text11.Text = "Roger"
      If C = 87 Then text11.Text = "Marc"
      If C = 88 Then text11.Text = "Cecil"
      If C = 89 Then text11.Text = "Kevin"
      If C = 90 Then text11.Text = "Maurice"
      If C = 91 Then text11.Text = "Byron"
      If C = 92 Then text11.Text = "Patrick"
      If C = 93 Then text11.Text = "Johnie"
      If C = 94 Then text11.Text = "Daniel"
      If C = 95 Then text11.Text = "Edwin"
      If C = 96 Then text11.Text = "Jesse"
      If C = 97 Then text11.Text = "Leonard"
      If C = 98 Then text11.Text = "Andy"
      If C = 99 Then text11.Text = "Norman"
      If C = 100 Then text11.Text = "Eugene"
      If C = 101 Then text11.Text = "Jerrell"

    End If

  'Last Name

    upperbound4 = 457
    lowerbound4 = 1

    d = Int((upperbound4 - lowerbound4 + 1) * Rnd + lowerbound)

    If d = 1 Then text12.Text = "Abbot"
    If d = 2 Then text12.Text = "Abe"
    If d = 3 Then text12.Text = "Abrams"
    If d = 4 Then text12.Text = "Acker"
    If d = 5 Then text12.Text = "Ackman"
    If d = 6 Then text12.Text = "Adams"
    If d = 7 Then text12.Text = "Adam"
    If d = 8 Then text12.Text = "Adamson"
    If d = 9 Then text12.Text = "Addison"
    If d = 10 Then text12.Text = "Adkins"
    If d = 11 Then text12.Text = "Adler"
    If d = 12 Then text12.Text = "Allen"
    If d = 13 Then text12.Text = "Alford"
    If d = 14 Then text12.Text = "Allison"
    If d = 15 Then text12.Text = "Alley"
    If d = 16 Then text12.Text = "Allgood"
    If d = 17 Then text12.Text = "Anderson"
    If d = 18 Then text12.Text = "Amos"
    If d = 19 Then text12.Text = "Applegate"
    If d = 20 Then text12.Text = "Apple"
    If d = 21 Then text12.Text = "Arnold"
    If d = 22 Then text12.Text = "Ashby"
    If d = 23 Then text12.Text = "Austin"
    If d = 24 Then text12.Text = "Austin"
    If d = 25 Then text12.Text = "Baker"
    If d = 26 Then text12.Text = "Baldwin"
    If d = 27 Then text12.Text = "Banks"
    If d = 28 Then text12.Text = "Bannerman"
    If d = 29 Then text12.Text = "Barber"
    If d = 30 Then text12.Text = "Barfield"
    If d = 31 Then text12.Text = "Barkley"
    If d = 32 Then text12.Text = "Barnett"
    If d = 33 Then text12.Text = "Barry"
    If d = 34 Then text12.Text = "Basinger"
    If d = 35 Then text12.Text = "Bays"
    If d = 36 Then text12.Text = "Beck"
    If d = 37 Then text12.Text = "Bell"
    If d = 38 Then text12.Text = "Benjamin"
    If d = 39 Then text12.Text = "Bennett"
    If d = 40 Then text12.Text = "Bentley"
    If d = 41 Then text12.Text = "Biggs"
    If d = 42 Then text12.Text = "Binder"
    If d = 43 Then text12.Text = "Blackburn"
    If d = 44 Then text12.Text = "Blackwood"
    If d = 45 Then text12.Text = "Blake"
    If d = 46 Then text12.Text = "Bobbitt"
    If d = 47 Then text12.Text = "Boggs"
    If d = 48 Then text12.Text = "Bolden"
    If d = 49 Then text12.Text = "Borden"
    If d = 50 Then text12.Text = "Bosworth"
    If d = 51 Then text12.Text = "Bowden"
    If d = 52 Then text12.Text = "Bowen"
    If d = 53 Then text12.Text = "Bowers"
    If d = 54 Then text12.Text = "Bowman"
    If d = 55 Then text12.Text = "Bradley"
    If d = 56 Then text12.Text = "Brady"
    If d = 57 Then text12.Text = "Briggs"
    If d = 58 Then text12.Text = "Britt"
    If d = 59 Then text12.Text = "Brown"
    If d = 60 Then text12.Text = "Bullock"
    If d = 61 Then text12.Text = "Burton"
    If d = 62 Then text12.Text = "Byrd"
    If d = 63 Then text12.Text = "Carlon"
    If d = 64 Then text12.Text = "Carter"
    If d = 65 Then text12.Text = "Casey"
    If d = 66 Then text12.Text = "Cates"
    If d = 67 Then text12.Text = "Caudle"
    If d = 68 Then text12.Text = "Chambers"
    If d = 69 Then text12.Text = "Chandler"
    If d = 70 Then text12.Text = "Chapman"
    If d = 71 Then text12.Text = "Chavis"
    If d = 72 Then text12.Text = "Cheek"
    If d = 73 Then text12.Text = "Clapp"
    If d = 74 Then text12.Text = "Clark"
    If d = 75 Then text12.Text = "Clarke"
    If d = 76 Then text12.Text = "Clemmons"
    If d = 77 Then text12.Text = "Cline"
    If d = 78 Then text12.Text = "Cobb"
    If d = 79 Then text12.Text = "Coble"
    If d = 80 Then text12.Text = "Cole"
    If d = 81 Then text12.Text = "Coleman"
    If d = 82 Then text12.Text = "Collins"
    If d = 83 Then text12.Text = "Combs"
    If d = 84 Then text12.Text = "Compton"
    If d = 85 Then text12.Text = "Conner"
    If d = 86 Then text12.Text = "Connor"
    If d = 87 Then text12.Text = "Conway"
    If d = 88 Then text12.Text = "Cook"
    If d = 89 Then text12.Text = "Cowan"
    If d = 90 Then text12.Text = "Cox"
    If d = 91 Then text12.Text = "Crawford"
    If d = 92 Then text12.Text = "Crutchfield"
    If d = 93 Then text12.Text = "Cunningham"
    If d = 94 Then text12.Text = "Daniel"
    If d = 95 Then text12.Text = "Darnell"
    If d = 96 Then text12.Text = "Davis"
    If d = 97 Then text12.Text = "Day"
    If d = 98 Then text12.Text = "Dawson"
    If d = 99 Then text12.Text = "Dixon"
    If d = 100 Then text12.Text = "Douglas"
    If d = 101 Then text12.Text = "Dunn"
    If d = 102 Then text12.Text = "Edwards"
    If d = 103 Then text12.Text = "Ellis"
    If d = 104 Then text12.Text = "Ellison"
    If d = 105 Then text12.Text = "Emmett"
    If d = 106 Then text12.Text = "Erwin"
    If d = 107 Then text12.Text = "Evans"
    If d = 108 Then text12.Text = "Evenson"
    If d = 109 Then text12.Text = "Everhart"
    If d = 110 Then text12.Text = "Farlow"
    If d = 111 Then text12.Text = "Fann"
    If d = 112 Then text12.Text = "Ferguson"
    If d = 113 Then text12.Text = "Fields"
    If d = 114 Then text12.Text = "Fish"
    If d = 115 Then text12.Text = "Fitzgerald"
    If d = 116 Then text12.Text = "Flowers"
    If d = 117 Then text12.Text = "Forrest"
    If d = 118 Then text12.Text = "Foster"
    If d = 119 Then text12.Text = "Fowler"
    If d = 120 Then text12.Text = "Frazier"
    If d = 121 Then text12.Text = "Free"
    If d = 122 Then text12.Text = "Freeman"
    If d = 123 Then text12.Text = "Frye"
    If d = 124 Then text12.Text = "Fulk"
    If d = 125 Then text12.Text = "Fulton"
    If d = 126 Then text12.Text = "Furr"
    If d = 127 Then text12.Text = "Galloway"
    If d = 128 Then text12.Text = "Gardner"
    If d = 129 Then text12.Text = "Gaulden"
    If d = 130 Then text12.Text = "Gauley"
    If d = 131 Then text12.Text = "Gerkin"
    If d = 132 Then text12.Text = "Getters"
    If d = 133 Then text12.Text = "Gibbs"
    If d = 134 Then text12.Text = "Gibson"
    If d = 135 Then text12.Text = "Gilbert"
    If d = 136 Then text12.Text = "Gillett"
    If d = 137 Then text12.Text = "Glass"
    If d = 138 Then text12.Text = "Glenn"
    If d = 139 Then text12.Text = "Glover"
    If d = 140 Then text12.Text = "Goble"
    If d = 141 Then text12.Text = "Godfrey"
    If d = 142 Then text12.Text = "Goldston"
    If d = 143 Then text12.Text = "Goldstein"
    If d = 144 Then text12.Text = "Good"
    If d = 145 Then text12.Text = "Gordon"
    If d = 146 Then text12.Text = "Gorham"
    If d = 147 Then text12.Text = "Gorrell"
    If d = 148 Then text12.Text = "Gossett"
    If d = 149 Then text12.Text = "Graham"
    If d = 150 Then text12.Text = "Grant"
    If d = 151 Then text12.Text = "Graves"
    If d = 152 Then text12.Text = "Gray"
    If d = 153 Then text12.Text = "Green"
    If d = 154 Then text12.Text = "Garfield"
    If d = 155 Then text12.Text = "Greene"
    If d = 156 Then text12.Text = "Gregory"
    If d = 157 Then text12.Text = "Gross"
    If d = 158 Then text12.Text = "Gwynn"
    If d = 159 Then text12.Text = "Hager"
    If d = 160 Then text12.Text = "Haines"
    If d = 161 Then text12.Text = "Hairston"
    If d = 162 Then text12.Text = "Ham"
    If d = 163 Then text12.Text = "Hamilton"
    If d = 164 Then text12.Text = "Hammond"
    If d = 165 Then text12.Text = "Hanks"
    If d = 166 Then text12.Text = "Hampton"
    If d = 167 Then text12.Text = "Hamrick"
    If d = 168 Then text12.Text = "Harden"
    If d = 169 Then text12.Text = "Hardy"
    If d = 170 Then text12.Text = "Hare"
    If d = 171 Then text12.Text = "Harper"
    If d = 172 Then text12.Text = "Harrel"
    If d = 173 Then text12.Text = "Harris"
    If d = 174 Then text12.Text = "Harrison"
    If d = 175 Then text12.Text = "Hart"
    If d = 176 Then text12.Text = "Hatfield"
    If d = 177 Then text12.Text = "Hayes"
    If d = 178 Then text12.Text = "Haynes"
    If d = 179 Then text12.Text = "Helms"
    If d = 180 Then text12.Text = "Henderson"
    If d = 181 Then text12.Text = "Hendricks"
    If d = 182 Then text12.Text = "Henley"
    If d = 183 Then text12.Text = "Henson"
    If d = 184 Then text12.Text = "Hester"
    If d = 185 Then text12.Text = "Hatt"
    If d = 186 Then text12.Text = "Hickman"
    If d = 187 Then text12.Text = "Hicks"
    If d = 188 Then text12.Text = "Higgins"
    If d = 189 Then text12.Text = "Hill"
    If d = 190 Then text12.Text = "Hines"
    If d = 191 Then text12.Text = "Hinshaw"
    If d = 192 Then text12.Text = "Hinson"
    If d = 193 Then text12.Text = "Hodges"
    If d = 194 Then text12.Text = "Hobson"
    If d = 195 Then text12.Text = "Holder"
    If d = 196 Then text12.Text = "Holl    "
    If d = 197 Then text12.Text = "Holiday"
    If d = 198 Then text12.Text = "Holloway"
    If d = 199 Then text12.Text = "Holt"
    If d = 200 Then text12.Text = "Hoover"
    If d = 201 Then text12.Text = "Hopkins"
    If d = 202 Then text12.Text = "Horn"
    If d = 203 Then text12.Text = "Hoskins"
    If d = 204 Then text12.Text = "Hubbard"
    If d = 205 Then text12.Text = "Hudson"
    If d = 206 Then text12.Text = "Huff"
    If d = 207 Then text12.Text = "Huffman"
    If d = 208 Then text12.Text = "Hughes"
    If d = 209 Then text12.Text = "Huneycutt"
    If d = 210 Then text12.Text = "Hurst"
    If d = 211 Then text12.Text = "Hutchers"
    If d = 212 Then text12.Text = "Hylton"
    If d = 213 Then text12.Text = "Irel    "
    If d = 214 Then text12.Text = "Irwin"
    If d = 215 Then text12.Text = "Isley"
    If d = 216 Then text12.Text = "Jaby"
    If d = 217 Then text12.Text = "Jackson"
    If d = 218 Then text12.Text = "Jacob"
    If d = 219 Then text12.Text = "Jarrell"
    If d = 220 Then text12.Text = "Jester"
    If d = 221 Then text12.Text = "Johnson"
    If d = 222 Then text12.Text = "Jones"
    If d = 223 Then text12.Text = "Jordan"
    If d = 224 Then text12.Text = "Joseph"
    If d = 225 Then text12.Text = "Joyner"
    If d = 226 Then text12.Text = "Justice"
    If d = 227 Then text12.Text = "Keith"
    If d = 228 Then text12.Text = "Kelley"
    If d = 229 Then text12.Text = "Kellis"
    If d = 230 Then text12.Text = "Kendall"
    If d = 231 Then text12.Text = "Kennedy"
    If d = 232 Then text12.Text = "Kenny"
    If d = 233 Then text12.Text = "Kim"
    If d = 234 Then text12.Text = "Kimble"
    If d = 235 Then text12.Text = "King"
    If d = 236 Then text12.Text = "Kinsey"
    If d = 237 Then text12.Text = "Kirkman"
    If d = 238 Then text12.Text = "Kiser"
    If d = 239 Then text12.Text = "Knight"
    If d = 240 Then text12.Text = "Knott"
    If d = 241 Then text12.Text = "Lamb"
    If d = 242 Then text12.Text = "Lambert"
    If d = 243 Then text12.Text = "Lambeth"
    If d = 244 Then text12.Text = "Lacey"
    If d = 245 Then text12.Text = "Lane"
    If d = 246 Then text12.Text = "Lang"
    If d = 247 Then text12.Text = "Langley"
    If d = 248 Then text12.Text = "Lankford"
    If d = 249 Then text12.Text = "Lashley"
    If d = 250 Then text12.Text = "Lawson"
    If d = 251 Then text12.Text = "Leary"
    If d = 252 Then text12.Text = "Leggett"
    If d = 253 Then text12.Text = "Legrant"
    If d = 254 Then text12.Text = "Lemonds"
    If d = 255 Then text12.Text = "Lanahan"
    If d = 256 Then text12.Text = "Leonard"
    If d = 257 Then text12.Text = "Lewis"
    If d = 258 Then text12.Text = "Lineback"
    If d = 259 Then text12.Text = "Lindsay"
    If d = 260 Then text12.Text = "Little"
    If d = 261 Then text12.Text = "Lloyd"
    If d = 262 Then text12.Text = "Livengood"
    If d = 263 Then text12.Text = "Locke"
    If d = 264 Then text12.Text = "Long"
    If d = 265 Then text12.Text = "Lord"
    If d = 266 Then text12.Text = "Love"
    If d = 267 Then text12.Text = "Lovett"
    If d = 268 Then text12.Text = "Lucas"
    If d = 269 Then text12.Text = "Lowe"
    If d = 270 Then text12.Text = "Lyons"
    If d = 271 Then text12.Text = "Mabe"
    If d = 272 Then text12.Text = "Mackey"
    If d = 273 Then text12.Text = "Maddox"
    If d = 274 Then text12.Text = "Maness"
    If d = 275 Then text12.Text = "Mangum"
    If d = 276 Then text12.Text = "Mann"
    If d = 277 Then text12.Text = "Manning"
    If d = 278 Then text12.Text = "Marion"
    If d = 279 Then text12.Text = "Marks"
    If d = 280 Then text12.Text = "Marsh"
    If d = 281 Then text12.Text = "Marshall"
    If d = 282 Then text12.Text = "Martin"
    If d = 283 Then text12.Text = "Martinez"
    If d = 284 Then text12.Text = "Mathews"
    If d = 285 Then text12.Text = "Mathis"
    If d = 286 Then text12.Text = "Matthews"
    If d = 286 Then text12.Text = "May"
    If d = 287 Then text12.Text = "Mayer"
    If d = 288 Then text12.Text = "McClain"
    If d = 289 Then text12.Text = "McConnell"
    If d = 290 Then text12.Text = "McCormick"
    If d = 291 Then text12.Text = "McCraw"
    If d = 292 Then text12.Text = "McCullough"
    If d = 293 Then text12.Text = "McDougald"
    If d = 294 Then text12.Text = "McPherson"
    If d = 295 Then text12.Text = "Medlin"
    If d = 296 Then text12.Text = "Melvin"
    If d = 297 Then text12.Text = "Mentis"
    If d = 298 Then text12.Text = "Meyer"
    If d = 299 Then text12.Text = "Middleton"
    If d = 300 Then text12.Text = "Miller"
    If d = 301 Then text12.Text = "Mitchell"
    If d = 302 Then text12.Text = "Montgomery"
    If d = 303 Then text12.Text = "Moore"
    If d = 304 Then text12.Text = "Morrison"
    If d = 305 Then text12.Text = "Morrow"
    If d = 306 Then text12.Text = "Morton"
    If d = 307 Then text12.Text = "Moss"
    If d = 308 Then text12.Text = "Murphy"
    If d = 309 Then text12.Text = "Murray"
    If d = 310 Then text12.Text = "Neal"
    If d = 311 Then text12.Text = "Newnam"
    If d = 312 Then text12.Text = "Newton"
    If d = 313 Then text12.Text = "Nichols"
    If d = 314 Then text12.Text = "Nixon"
    If d = 315 Then text12.Text = "Norris"
    If d = 316 Then text12.Text = "Norton"
    If d = 317 Then text12.Text = "Nunn"
    If d = 318 Then text12.Text = "O'Brien"
    If d = 319 Then text12.Text = "O'Donnel"
    If d = 320 Then text12.Text = "Oldham"
    If d = 321 Then text12.Text = "Olds"
    If d = 322 Then text12.Text = "Oliver"
    If d = 323 Then text12.Text = "Olson"
    If d = 324 Then text12.Text = "O'Neal"
    If d = 325 Then text12.Text = "Osborne"
    If d = 326 Then text12.Text = "Owens"
    If d = 327 Then text12.Text = "Padgett"
    If d = 328 Then text12.Text = "Pace"
    If d = 329 Then text12.Text = "Page"
    If d = 330 Then text12.Text = "Painter"
    If d = 331 Then text12.Text = "Parker"
    If d = 332 Then text12.Text = "Parrish"
    If d = 333 Then text12.Text = "Parsons"
    If d = 334 Then text12.Text = "Patel"
    If d = 335 Then text12.Text = "Patterson"
    If d = 336 Then text12.Text = "Peacock"
    If d = 337 Then text12.Text = "Pennington"
    If d = 338 Then text12.Text = "Perkins"
    If d = 339 Then text12.Text = "Peterson"
    If d = 340 Then text12.Text = "Phillips"
    If d = 341 Then text12.Text = "Pickard"
    If d = 342 Then text12.Text = "Pickett"
    If d = 343 Then text12.Text = "Piper"
    If d = 344 Then text12.Text = "Pitts"
    If d = 345 Then text12.Text = "Pleasants"
    If d = 346 Then text12.Text = "Poole"
    If d = 347 Then text12.Text = "Pope"
    If d = 348 Then text12.Text = "Posey"
    If d = 349 Then text12.Text = "Potter"
    If d = 350 Then text12.Text = "Powers"
    If d = 351 Then text12.Text = "Prendergast"
    If d = 352 Then text12.Text = "Pressley"
    If d = 353 Then text12.Text = "Preston"
    If d = 354 Then text12.Text = "Priddy"
    If d = 355 Then text12.Text = "Pugh"
    If d = 356 Then text12.Text = "Purcell"
    If d = 357 Then text12.Text = "Purvis"
    If d = 358 Then text12.Text = "Putnam"
    If d = 359 Then text12.Text = "Quick"
    If d = 360 Then text12.Text = "Raines"
    If d = 361 Then text12.Text = "Ramsey"
    If d = 362 Then text12.Text = "Rankin"
    If d = 363 Then text12.Text = "Reese"
    If d = 364 Then text12.Text = "Reeves"
    If d = 365 Then text12.Text = "Reid"
    If d = 366 Then text12.Text = "Richards"
    If d = 367 Then text12.Text = "Rich"
    If d = 368 Then text12.Text = "Riches"
    If d = 369 Then text12.Text = "Richter"
    If d = 370 Then text12.Text = "Riley"
    If d = 371 Then text12.Text = "Ring"
    If d = 372 Then text12.Text = "Ripley"
    If d = 373 Then text12.Text = "Rippy"
    If d = 374 Then text12.Text = "Rivera"
    If d = 375 Then text12.Text = "Roberts"
    If d = 376 Then text12.Text = "Robertson"
    If d = 377 Then text12.Text = "Robinson"
    If d = 378 Then text12.Text = "Rodgers"
    If d = 379 Then text12.Text = "Rogers"
    If d = 380 Then text12.Text = "Roseboro"
    If d = 381 Then text12.Text = "Rosen"
    If d = 382 Then text12.Text = "Ross"
    If d = 383 Then text12.Text = "Rouse"
    If d = 384 Then text12.Text = "Royal"
    If d = 385 Then text12.Text = "Rucker"
    If d = 386 Then text12.Text = "Russell"
    If d = 387 Then text12.Text = "S    s"
    If d = 388 Then text12.Text = "Saunders"
    If d = 389 Then text12.Text = "Sawyer"
    If d = 390 Then text12.Text = "Scott"
    If d = 391 Then text12.Text = "Seawell"
    If d = 392 Then text12.Text = "Shannon"
    If d = 393 Then text12.Text = "Sharples"
    If d = 394 Then text12.Text = "Shaw"
    If d = 395 Then text12.Text = "Shelton"
    If d = 396 Then text12.Text = "Shoffner"
    If d = 397 Then text12.Text = "Short"
    If d = 398 Then text12.Text = "Showfety"
    If d = 399 Then text12.Text = "Simmons"
    If d = 400 Then text12.Text = "Sims"
    If d = 401 Then text12.Text = "Singer"
    If d = 402 Then text12.Text = "Small"
    If d = 403 Then text12.Text = "Smith"
    If d = 404 Then text12.Text = "Snyder"
    If d = 405 Then text12.Text = "Southern"
    If d = 406 Then text12.Text = "Spencer"
    If d = 407 Then text12.Text = "Staley"
    If d = 408 Then text12.Text = "Stanley"
    If d = 409 Then text12.Text = "Stephens"
    If d = 410 Then text12.Text = "Stephenson"
    If d = 411 Then text12.Text = "Stevens"
    If d = 412 Then text12.Text = "Stewart"
    If d = 413 Then text12.Text = "Stone"
    If d = 414 Then text12.Text = "Stricklen"
    If d = 415 Then text12.Text = "Swann"
    If d = 416 Then text12.Text = "Sumner"
    If d = 417 Then text12.Text = "Swaney"
    If d = 418 Then text12.Text = "Talbert"
    If d = 419 Then text12.Text = "Tate"
    If d = 420 Then text12.Text = "Taylor"
    If d = 421 Then text12.Text = "Thomas"
    If d = 422 Then text12.Text = "Thompson"
    If d = 423 Then text12.Text = "Thorne"
    If d = 424 Then text12.Text = "Townsend"
    If d = 425 Then text12.Text = "Troxler"
    If d = 426 Then text12.Text = "Turner"
    If d = 427 Then text12.Text = "Tuttle"
    If d = 428 Then text12.Text = "Tyson"
    If d = 429 Then text12.Text = "Umfleet"
    If d = 430 Then text12.Text = "Underwood"
    If d = 431 Then text12.Text = "Utley"
    If d = 432 Then text12.Text = "Vance"
    If d = 433 Then text12.Text = "Vanstory"
    If d = 434 Then text12.Text = "Vines"
    If d = 435 Then text12.Text = "Vaughan"
    If d = 436 Then text12.Text = "Vestall"
    If d = 437 Then text12.Text = "Walker"
    If d = 438 Then text12.Text = "Wall"
    If d = 439 Then text12.Text = "Wallace"
    If d = 440 Then text12.Text = "Walters"
    If d = 441 Then text12.Text = "Ward"
    If d = 442 Then text12.Text = "Warren"
    If d = 443 Then text12.Text = "Washington"
    If d = 444 Then text12.Text = "Watkins"
    If d = 445 Then text12.Text = "Watson"
    If d = 446 Then text12.Text = "Wayne"
    If d = 447 Then text12.Text = "Weaver"
    If d = 448 Then text12.Text = "Webb"
    If d = 449 Then text12.Text = "Webster"
    If d = 450 Then text12.Text = "Wells"
    If d = 451 Then text12.Text = "West"
    If d = 452 Then text12.Text = "White"
    If d = 453 Then text12.Text = "Wilkins"
    If d = 454 Then text12.Text = "Williams"
    If d = 455 Then text12.Text = "Wilson"
    If d = 456 Then text12.Text = "Young"
    If d = 457 Then text12.Text = "Zimmerman"

  'State

    upperbound5 = 52
    lowerbound5 = 1

    E = Int((upperbound5 - lowerbound5 + 1) * Rnd + lowerbound)

    If E = 1 Then text16.Text = "AL"
    If E = 2 Then text16.Text = "AK"
    If E = 3 Then text16.Text = "AZ"
    If E = 4 Then text16.Text = "AR"
    If E = 5 Then text16.Text = "CA"
    If E = 6 Then text16.Text = "CO"
    If E = 7 Then text16.Text = "CT"
    If E = 8 Then text16.Text = "DE"
    If E = 9 Then text16.Text = "DC"
    If E = 10 Then text16.Text = "FL"
    If E = 11 Then text16.Text = "GA"
    If E = 12 Then text16.Text = "HI"
    If E = 13 Then text16.Text = "ID"
    If E = 14 Then text16.Text = "IL"
    If E = 15 Then text16.Text = "IN"
    If E = 16 Then text16.Text = "IA"
    If E = 17 Then text16.Text = "KS"
    If E = 18 Then text16.Text = "KY"
    If E = 19 Then text16.Text = "LA"
    If E = 20 Then text16.Text = "ME"
    If E = 21 Then text16.Text = "MD"
    If E = 22 Then text16.Text = "MA"
    If E = 23 Then text16.Text = "MI"
    If E = 24 Then text16.Text = "MN"
    If E = 25 Then text16.Text = "MS"
    If E = 26 Then text16.Text = "MO"
    If E = 27 Then text16.Text = "MT"
    If E = 28 Then text16.Text = "NE"
    If E = 29 Then text16.Text = "NV"
    If E = 30 Then text16.Text = "NH"
    If E = 31 Then text16.Text = "NJ"
    If E = 32 Then text16.Text = "NM"
    If E = 33 Then text16.Text = "NY"
    If E = 34 Then text16.Text = "NC"
    If E = 35 Then text16.Text = "ND"
    If E = 36 Then text16.Text = "OH"
    If E = 37 Then text16.Text = "OK"
    If E = 38 Then text16.Text = "OR"
    If E = 39 Then text16.Text = "PA"
    If E = 40 Then text16.Text = "PR"
    If E = 41 Then text16.Text = "RI"
    If E = 42 Then text16.Text = "SC"
    If E = 43 Then text16.Text = "SD"
    If E = 44 Then text16.Text = "TN"
    If E = 45 Then text16.Text = "TX"
    If E = 46 Then text16.Text = "UT"
    If E = 47 Then text16.Text = "VT"
    If E = 48 Then text16.Text = "VA"
    If E = 49 Then text16.Text = "WA"
    If E = 50 Then text16.Text = "WV"
    If E = 51 Then text16.Text = "WI"
    If E = 52 Then text16.Text = "WY"


  'City/Town

    upperbound8 = 2
    lowerbound8 = 1

    F = Int((upperbound8 - lowerbound8 + 1) * Rnd + lowerbound)


    If E = 1 Then
      If F = 1 Then text15.Text = "Birmingham"
      If F = 2 Then text15.Text = "Montgomery"
    End If

    If E = 2 Then text15.Text = "Juneau"

    If E = 3 Then text15.Text = "Phoenix"

    If E = 4 Then text15.Text = "Little Rock"

    If E = 5 Then
      If F = 1 Then text15.Text = "Los Angeles"
      If F = 2 Then text15.Text = "San Francisco"
    End If

    If E = 6 Then
      If F = 1 Then text15.Text = "Colorado Springs"
      If F = 2 Then text15.Text = "Denver"
    End If

    If E = 7 Then
      If F = 1 Then text15.Text = "Hartford"
      If F = 2 Then text15.Text = "Hartford"
    End If

    If E = 8 Then
      If F = 1 Then text15.Text = "Dover"
      If F = 2 Then text15.Text = "Dover"
    End If

    If E = 9 Then
      If F = 1 Then text15.Text = "Washington"
      If F = 2 Then text15.Text = "Washington"
    End If

    If E = 10 Then
      If F = 1 Then text15.Text = "Miami"
      If F = 2 Then text15.Text = "Tallahassee"
    End If

    If E = 11 Then
      If F = 1 Then text15.Text = "Atlanta"
      If F = 2 Then text15.Text = "Fayetteville"
    End If

    If E = 12 Then
      If F = 1 Then text15.Text = "Honolulu"
      If F = 2 Then text15.Text = "Honolulu"
    End If

    If E = 13 Then
      If F = 1 Then text15.Text = "Boise"
      If F = 2 Then text15.Text = "Boise"
    End If

    If E = 14 Then
      If F = 1 Then text15.Text = "Chicago"
      If F = 2 Then text15.Text = "Springfield"
    End If

    If E = 15 Then
      If F = 1 Then text15.Text = "Indianapolis"
      If F = 2 Then text15.Text = "South Bend"
    End If

    If E = 16 Then
      If F = 1 Then text15.Text = "Des Moines"
      If F = 2 Then text15.Text = "Sioux City"
    End If

    If E = 17 Then
      If F = 1 Then text15.Text = "Dodge City"
      If F = 2 Then text15.Text = "Topeka"
    End If

    If E = 18 Then
      If F = 1 Then text15.Text = "Frankfort"
      If F = 2 Then text15.Text = "Louisville"
    End If

    If E = 19 Then
      If F = 1 Then text15.Text = "Baton Rouge"
      If F = 2 Then text15.Text = "New Orleans"
    End If

    If E = 20 Then
      If F = 1 Then text15.Text = "Augusta"
      If F = 2 Then text15.Text = "Augusta"
    End If

    If E = 21 Then
      If F = 1 Then text15.Text = "Annapolis"
      If F = 2 Then text15.Text = "Baltimore"
    End If

    If E = 22 Then
      If F = 1 Then text15.Text = "Boston"
      If F = 2 Then text15.Text = "Springfield"
    End If

    If E = 23 Then
      If F = 1 Then text15.Text = "Bay City"
      If F = 2 Then text15.Text = "Detroit"
    End If

    If E = 24 Then
      If F = 1 Then text15.Text = "Minneapolis"
      If F = 2 Then text15.Text = "St. Paul"
    End If

    If E = 25 Then
      If F = 1 Then text15.Text = "Jackson"
      If F = 2 Then text15.Text = "Jackson"
    End If

    If E = 26 Then
      If F = 1 Then text15.Text = "Jefferson City"
      If F = 2 Then text15.Text = "Kansas City"
    End If

    If E = 27 Then
      If F = 1 Then text15.Text = "Helena"
      If F = 2 Then text15.Text = "Helena"
    End If

    If E = 28 Then
      If F = 1 Then text15.Text = "Lincoln"
      If F = 2 Then text15.Text = "Omaha"
    End If

    If E = 29 Then
      If F = 1 Then text15.Text = "Carson City"
      If F = 2 Then text15.Text = "Reno"
    End If

    If E = 30 Then
      If F = 1 Then text15.Text = "Concord"
      If F = 2 Then text15.Text = "Concord"
    End If

    If E = 31 Then
      If F = 1 Then text15.Text = "Atlantic City"
      If F = 2 Then text15.Text = "Trenton"
    End If

    If E = 32 Then
      If F = 1 Then text15.Text = "Santa Fe"
      If F = 2 Then text15.Text = "Santa Fe"
    End If

    If E = 33 Then
      If F = 1 Then text15.Text = "Buffalo"
      If F = 2 Then text15.Text = "New York"
    End If

    If E = 34 Then
      If F = 1 Then text15.Text = "Charlotte"
      If F = 2 Then text15.Text = "Raleigh"
    End If

    If E = 35 Then
      If F = 1 Then text15.Text = "Bismarck"
      If F = 2 Then text15.Text = "Bismarck"
    End If

    If E = 36 Then
      If F = 1 Then text15.Text = "Akron"
      If F = 2 Then text15.Text = "Columbus"
    End If

    If E = 37 Then
      If F = 1 Then text15.Text = "Oklahoma City"
      If F = 2 Then text15.Text = "Tulsa"
    End If

    If E = 38 Then
      If F = 1 Then text15.Text = "Salem"
      If F = 2 Then text15.Text = "Salem"
    End If

    If E = 39 Then
      If F = 1 Then text15.Text = "Philadelphia"
      If F = 2 Then text15.Text = "Pittsburgh"
    End If

    If E = 40 Then
      If F = 1 Then text15.Text = "San Juan"
      If F = 2 Then text15.Text = "San Juan"
    End If

    If E = 41 Then
      If F = 1 Then text15.Text = "Providence"
      If F = 2 Then text15.Text = "Providence"
    End If

    If E = 42 Then
      If F = 1 Then text15.Text = "Columbia"
      If F = 2 Then text15.Text = "Myrtle Beach"
    End If

    If E = 43 Then
      If F = 1 Then text15.Text = "Pierre"
      If F = 2 Then text15.Text = "Pierre"
    End If

    If E = 44 Then
      If F = 1 Then text15.Text = "Memphis"
      If F = 2 Then text15.Text = "Nashville"
    End If

    If E = 45 Then
      If F = 1 Then text15.Text = "Austin"
      If F = 2 Then text15.Text = "Dallas"
    End If

    If E = 46 Then
      If F = 1 Then text15.Text = "Salt Lake City"
      If F = 2 Then text15.Text = "Salt Lake City"
    End If

    If E = 47 Then
      If F = 1 Then text15.Text = "Montpelier"
      If F = 2 Then text15.Text = "Montpelier"
    End If

    If E = 48 Then
      If F = 1 Then text15.Text = "Arlington"
      If F = 2 Then text15.Text = "Richmond"
    End If

    If E = 49 Then
      If F = 1 Then text15.Text = "Seatttle"
      If F = 2 Then text15.Text = "Walla Walla"
    End If

    If E = 50 Then
      If F = 1 Then text15.Text = "Charleston"
      If F = 2 Then text15.Text = "Charleston"
    End If

    If E = 51 Then
      If F = 1 Then text15.Text = "Green Bay"
      If F = 2 Then text15.Text = "Milwaukee"
    End If

    If E = 52 Then
      If F = 1 Then text15.Text = "Cheyenne"
      If F = 2 Then text15.Text = "Cheyenne"
    End If

  'Area codes

    If E = 1 Then
      If F = 1 Then text17.Text = "205"
      If F = 2 Then text17.Text = "334"
    End If

    If E = 2 Then text17.Text = "907"

    If E = 3 Then text17.Text = "602"

    If E = 4 Then text17.Text = "501"

    If E = 5 Then
      If F = 1 Then text17.Text = "213"
      If F = 2 Then text17.Text = "415"
    End If

    If E = 6 Then
      If F = 1 Then text17.Text = "719"
      If F = 2 Then text17.Text = "303"
    End If

    If E = 7 Then
      If F = 1 Then text17.Text = "203"
      If F = 2 Then text17.Text = "203"
    End If

    If E = 8 Then
      If F = 1 Then text17.Text = "302"
      If F = 2 Then text17.Text = "302"
    End If

    If E = 9 Then
      If F = 1 Then text17.Text = "202"
      If F = 2 Then text17.Text = "202"
    End If

    If E = 10 Then
      If F = 1 Then text17.Text = "305"
      If F = 2 Then text17.Text = "904"
    End If

    If E = 11 Then
      If F = 1 Then text17.Text = "404"
      If F = 2 Then text17.Text = "404"
    End If

    If E = 12 Then
      If F = 1 Then text17.Text = "334"
      If F = 2 Then text17.Text = "334"
    End If

    If E = 13 Then
      If F = 1 Then text17.Text = "334"
      If F = 2 Then text17.Text = "334"
    End If

    If E = 14 Then
      If F = 1 Then text17.Text = "312"
      If F = 2 Then text17.Text = "217"
    End If

    If E = 15 Then
      If F = 1 Then text17.Text = "317"
      If F = 2 Then text17.Text = "219"
    End If

    If E = 16 Then
      If F = 1 Then text17.Text = "515"
      If F = 2 Then text17.Text = "712"
    End If

    If E = 17 Then
      If F = 1 Then text17.Text = "316"
      If F = 2 Then text17.Text = "913"
    End If

    If E = 18 Then
      If F = 1 Then text17.Text = "502"
      If F = 2 Then text17.Text = "502"
    End If

    If E = 19 Then
      If F = 1 Then text17.Text = "504"
      If F = 2 Then text17.Text = "504"
    End If

    If E = 20 Then
      If F = 1 Then text17.Text = "207"
      If F = 2 Then text17.Text = "207"
    End If

    If E = 21 Then
      If F = 1 Then text17.Text = "410"
      If F = 2 Then text17.Text = "410"
    End If

    If E = 22 Then
      If F = 1 Then text17.Text = "617"
      If F = 2 Then text17.Text = "413"
    End If

    If E = 23 Then
      If F = 1 Then text17.Text = "517"
      If F = 2 Then text17.Text = "313"
    End If

    If E = 24 Then
      If F = 1 Then text17.Text = "612"
      If F = 2 Then text17.Text = "612"
    End If

    If E = 25 Then
      If F = 1 Then text17.Text = "601"
      If F = 2 Then text17.Text = "601"
    End If

    If E = 26 Then
      If F = 1 Then text17.Text = "314"
      If F = 2 Then text17.Text = "816"
    End If

    If E = 27 Then
      If F = 1 Then text17.Text = "406"
      If F = 2 Then text17.Text = "406"
    End If

    If E = 28 Then
      If F = 1 Then text17.Text = "402"
      If F = 2 Then text17.Text = "402"
    End If

    If E = 29 Then
      If F = 1 Then text17.Text = "702"
      If F = 2 Then text17.Text = "702"
    End If

    If E = 30 Then
      If F = 1 Then text17.Text = "603"
      If F = 2 Then text17.Text = "603"
    End If

    If E = 31 Then
      If F = 1 Then text17.Text = "609"
      If F = 2 Then text17.Text = "609"
    End If

    If E = 32 Then
      If F = 1 Then text17.Text = "505"
      If F = 2 Then text17.Text = "505"
    End If

    If E = 33 Then
      If F = 1 Then text17.Text = "716"
      If F = 2 Then text17.Text = "212"
    End If

    If E = 34 Then
      If F = 1 Then text17.Text = "704"
      If F = 2 Then text17.Text = "919"
    End If

    If E = 35 Then
      If F = 1 Then text17.Text = "334"
      If F = 2 Then text17.Text = "334"
    End If

    If E = 36 Then
      If F = 1 Then text17.Text = "216"
      If F = 2 Then text17.Text = "614"
    End If

    If E = 37 Then
      If F = 1 Then text17.Text = "405"
      If F = 2 Then text17.Text = "918"
    End If

    If E = 38 Then
      If F = 1 Then text17.Text = "503"
      If F = 2 Then text17.Text = "503"
    End If

    If E = 39 Then
      If F = 1 Then text17.Text = "215"
      If F = 2 Then text17.Text = "412"
    End If

    If E = 40 Then
      If F = 1 Then text17.Text = "809"
      If F = 2 Then text17.Text = "809"
    End If

    If E = 41 Then
      If F = 1 Then text17.Text = "401"
      If F = 2 Then text17.Text = "401"
    End If

    If E = 42 Then
      If F = 1 Then text17.Text = "803"
      If F = 2 Then text17.Text = "803"
    End If

    If E = 43 Then
      If F = 1 Then text17.Text = "605"
      If F = 2 Then text17.Text = "605"
    End If

    If E = 44 Then
      If F = 1 Then text17.Text = "901"
      If F = 2 Then text17.Text = "615"
    End If

    If E = 45 Then
      If F = 1 Then text17.Text = "512"
      If F = 2 Then text17.Text = "214"
    End If

    If E = 46 Then
      If F = 1 Then text17.Text = "801"
      If F = 2 Then text17.Text = "801"
    End If

    If E = 47 Then
      If F = 1 Then text17.Text = "802"
      If F = 2 Then text17.Text = "802"
    End If

    If E = 48 Then
      If F = 1 Then text17.Text = "703"
      If F = 2 Then text17.Text = "804"
    End If

    If E = 49 Then
      If F = 1 Then text17.Text = "206"
      If F = 2 Then text17.Text = "509"
    End If

    If E = 50 Then
      If F = 1 Then text17.Text = "304"
      If F = 2 Then text17.Text = "304"
    End If

    If E = 51 Then
      If F = 1 Then text17.Text = "414"
      If F = 2 Then text17.Text = "414"
    End If

    If E = 52 Then
      If F = 1 Then text17.Text = "307"
      If F = 2 Then text17.Text = "307"
    End If

    E = E
    F = F

  'Phone number prefix

    upperbound13 = 970
    lowerbound13 = 200

    H = Int((upperbound13 - lowerbound13 + 1) * Rnd + lowerbound13)

    text18.Text = H

  'Rest of phone number

    upperbound14 = 9990
    lowerbound14 = 1000

    I = Int((upperbound14 - lowerbound14 + 1) * Rnd + lowerbound14)

    text19.Text = I

  'Street number

    upperbound15 = 3000
    lowerbound15 = 50

    J = Int((upperbound15 - lowerbound15 + 1) * Rnd + lowerbound14)

    text13.Text = J

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

    l = Int((Higher - Lower + 1) * Rnd + Lower)

    text20.Text = l

    DoEvents
    DoEvents
    DoEvents

    If text20.Text < 10000 Then text20.Text = "0" + l
    If text20.Text < 1000 Then text20.Text = "00" + l

  End Sub

  Sub Form_Load ()
'// If you use this code in your program and do not
'// give Chronic the proper credit you are a dickless
'// terd!  :)  Be a man, give credit where credit is due
  
  
  
  'Dim Success As Integer

  'Success% = SetWindowPos%(creditwiz.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)


  'text5.Visible = False

  'Blah = "4225 - Chase Manhatten Bank"
  'list5.AddItem Blah
  '
  'list5.AddItem "4013 - Citibank"
  'list5.AddItem "4060 - Associates National Bank"
  'list5.AddItem "4271 - Citibank Preferred"
  'list5.AddItem "4332 - First American Bank"
  'list5.AddItem "4436 - Security Bank and Trust"
  'list5.AddItem "4503 - Canadian Imperial Bank"
  'list5.AddItem "4564 - Bank of Queensland"
  'list5.AddItem "4726 - Wells Fargo"
  'list5.AddItem "4811 - Bank of Hawaii"
  'list5.AddItem "4833 - U.S. Bank"
  'list5.AddItem "4842 - Security Pacific Washington"
  'list5.AddItem "4921 - Hong Kong/National Bank"
  'list5.ListIndex = 0
  '

    list5.AddItem "4013 - Citibank"
    list5.AddItem "4019 - Bank of America"
    list5.AddItem "4024 - Bank of America"
    list5.AddItem "4027 - Rockwell Federal Credit Union"
    list5.AddItem "4032 - Household Bank"
    list5.AddItem "4052 - First Cincinatti"
    list5.AddItem "4060 - Associates National Bank"
    list5.AddItem "4070 - Security Pacific"
    list5.AddItem "4071 - Colonial National Bank"
    list5.AddItem "4094 - A.M.C. Federal Credit Union"
    list5.AddItem "4113 - Valley National Bank"
    list5.AddItem "4114 - Chemical Bank"
    list5.AddItem "4121 - Pennsylvania State Employees Credit Union"
    list5.AddItem "4122 - Union Trust"
    list5.AddItem "4128 - Citibank"
    list5.AddItem "4131 - State Street Bank"
    list5.AddItem "4125 - Marine Midland"
    list5.AddItem "4225 - Chase Manhatten Bank"
    list5.AddItem "4226 - Chase Manhatten Bank"
    list5.AddItem "4231 - Chase Lincoln First Classic"
    list5.AddItem "4232 - Chase Lincoln First Classic"
    list5.AddItem "4239 - Corestates"
    list5.AddItem "4241 - National Westminester Bank"
    list5.AddItem "4250 - First Chicago Bank"
    list5.AddItem "4253 - Consumers Edge"
    list5.AddItem "4254 - Security First"
    list5.AddItem "4271 - Citibank Preferred"
    list5.AddItem "4301 - Monogram Bank"
    list5.AddItem "4302 - H.H.B.C."
    list5.AddItem "4310 - B.C.F.U."
    list5.AddItem "4311 - First National Bank of Louisville"
    list5.AddItem "4317 - Gold Dome"
    list5.AddItem "4327 - First Atlanta"
    list5.AddItem "4332 - First American Bank"
    list5.AddItem "4339 - Primerica Bank"
    list5.AddItem "4342 - N.C.M.B. / Nations Bank"
    list5.AddItem "4356 - National Bank of Delaware"
    list5.AddItem "4368 - National West"
    list5.AddItem "4387 - Bank One"
    list5.AddItem "4388 - First Signature Bank & Trust"
    list5.AddItem "4401 - Gary-Wheaton Bank"
    list5.AddItem "4413 - Firstier Bank Lincoln"
    list5.AddItem "4418 - Bank of Omaha"
    list5.AddItem "4421 - Indiana National Bank"
    list5.AddItem "4428 - Bank of Hoven"
    list5.AddItem "4436 - Security Bank and Trust"
    list5.AddItem "4443 - Merril Lynch Bank/Trust"
    list5.AddItem "4447 - AmeriTrust"
    list5.AddItem "4452 - Empire Affiliates Federal Credit Union"
    list5.AddItem "4498 - Republic Savings"
    list5.AddItem "4502 - C.I.B.C."
    list5.AddItem "4503 - Canadian Imperial Bank"
    list5.AddItem "4506 - Belgium A.S.L.K"
    list5.AddItem "4510 - Royal Bank of Canada"
    list5.AddItem "4520 - Toronto Dominion of Canada"
    list5.AddItem "4537 - Bank of Nova Scotia"
    list5.AddItem "4538 - Bank of Nova Scotia"
    list5.AddItem "4539 - Barclays"
    list5.AddItem "4544 - T.S.B. Bank"
    list5.AddItem "4556 - Citibank"
    list5.AddItem "4564 - Bank of Queensland"
    list5.AddItem "4673 - First Card"
    list5.AddItem "4707 - Tompkins County Trust"
    list5.AddItem "4719 - Rocky Mountain"
    list5.AddItem "4721 - First Security"
    list5.AddItem "4722 - West Bank"
    list5.AddItem "4726 - Wells Fargo"
    list5.AddItem "4783 - AT&T's Universal Card"
    list5.AddItem "4784 - AT&T's Universal Card"
    list5.AddItem "4800 - M.B.N.A. North America"
    list5.AddItem "4811 - Bank of Hawaii"
    list5.AddItem "4819 - Macom Federal Crdit Union"
    da = "a dk owk ak dkas x sl ei a kda ls so "
    flah = "By Da Chronic and The Rizzer"
    list5.AddItem "4820 - IBM Mid America Federal Credit Union"
    list5.AddItem "4833 - U.S. Bank"
    list5.AddItem "4842 - Security Pacific Washington"
    list5.AddItem "4897 - Village Bank of Chicago"
    list5.AddItem "4921 - Hong Kong/National Bank"
    list5.ListIndex = 0
  CreditWiz.Caption = "Credit Wizard v1.1  -  " + flah
   asdf = "asldkfjlasdfja"
   blah1 = CreditWiz.Caption
   blah2 = InStr(blah1, "Da Chronic")
   dff = "asldkf"
   If blah2 = 0 Then
   MsgBox "Hacked version detected!"
   End
   End If


  End Sub

  Sub Form_Resize ()
  'Dim Success As Integer

  'Success% = SetWindowPos%(creditwiz.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

  End Sub

  Sub List5_Click ()
    jacko = list5.List(list5.ListIndex)
    jackass = Left(jacko, 4)
    ju = CInt(jackass)

    Text1.Text = ju


  End Sub

  Sub List5_KeyPress (KeyAscii As Integer)
    jacko = list5.List(list5.ListIndex)
    jackass = Left(jacko, 4)
    ju = CInt(jackass)

    Text1.Text = ju

  End Sub

  Sub mnuCheck_Click ()
    extrap1 = 1

    cardcheck.Show


  'MsgBox "This feature will be functional in a future beta release.", 16, "Generate"
  End Sub

  Sub mnuExit_Click ()
    End

  End Sub

  Sub mnuExtro_Click ()
    Call Command3D6_Click

  End Sub

  Sub mnuGen_Click ()
    Call Command3D1_Click

  End Sub

  Sub mnuPrint_Click ()
    PrintForm

  End Sub

  Sub mnuPrnt_Click ()

    Call Command3D2_Click

  End Sub

  Sub Option3D1_Click (Value As Integer)
    list5.Clear

    list5.AddItem "4013 - Citibank"
    list5.AddItem "4019 - Bank of America"
    list5.AddItem "4024 - Bank of America"
    list5.AddItem "4027 - Rockwell Federal Credit Union"
    list5.AddItem "4032 - Household Bank"
    list5.AddItem "4052 - First Cincinatti"
    list5.AddItem "4060 - Associates National Bank"
    list5.AddItem "4070 - Security Pacific"
    list5.AddItem "4071 - Colonial National Bank"
    list5.AddItem "4094 - A.M.C. Federal Credit Union"
    list5.AddItem "4113 - Valley National Bank"
    list5.AddItem "4114 - Chemical Bank"
    list5.AddItem "4121 - Pennsylvania State Employees Credit Union"
    list5.AddItem "4122 - Union Trust"
    list5.AddItem "4128 - Citibank"
    list5.AddItem "4131 - State Street Bank"
    list5.AddItem "4125 - Marine Midland"
    list5.AddItem "4225 - Chase Manhatten Bank"
    list5.AddItem "4226 - Chase Manhatten Bank"
    list5.AddItem "4231 - Chase Lincoln First Classic"
    list5.AddItem "4232 - Chase Lincoln First Classic"
    list5.AddItem "4239 - Corestates"
    list5.AddItem "4241 - National Westminester Bank"
    list5.AddItem "4250 - First Chicago Bank"
    list5.AddItem "4253 - Consumers Edge"
    list5.AddItem "4254 - Security First"
    list5.AddItem "4271 - Citibank Preferred"
    list5.AddItem "4301 - Monogram Bank"
    list5.AddItem "4302 - H.H.B.C."
    list5.AddItem "4310 - B.C.F.U."
    list5.AddItem "4311 - First National Bank of Louisville"
    list5.AddItem "4317 - Gold Dome"
    list5.AddItem "4327 - First Atlanta"
    list5.AddItem "4332 - First American Bank"
    list5.AddItem "4339 - Primerica Bank"
    list5.AddItem "4342 - N.C.M.B. / Nations Bank"
    list5.AddItem "4356 - National Bank of Delaware"
    list5.AddItem "4368 - National West"
    list5.AddItem "4387 - Bank One"
    list5.AddItem "4388 - First Signature Bank & Trust"
    list5.AddItem "4401 - Gary-Wheaton Bank"
    list5.AddItem "4413 - Firstier Bank Lincoln"
    list5.AddItem "4418 - Bank of Omaha"
    list5.AddItem "4421 - Indiana National Bank"
    list5.AddItem "4428 - Bank of Hoven"
    list5.AddItem "4436 - Security Bank and Trust"
    list5.AddItem "4443 - Merril Lynch Bank/Trust"
    list5.AddItem "4447 - AmeriTrust"
    list5.AddItem "4452 - Empire Affiliates Federal Credit Union"
    list5.AddItem "4498 - Republic Savings"
    list5.AddItem "4502 - C.I.B.C."
    list5.AddItem "4503 - Canadian Imperial Bank"
    list5.AddItem "4506 - Belgium A.S.L.K"
    list5.AddItem "4510 - Royal Bank of Canada"
    list5.AddItem "4520 - Toronto Dominion of Canada"
    list5.AddItem "4537 - Bank of Nova Scotia"
    list5.AddItem "4538 - Bank of Nova Scotia"
    list5.AddItem "4539 - Barclays"
    list5.AddItem "4544 - T.S.B. Bank"
    list5.AddItem "4556 - Citibank"
    list5.AddItem "4564 - Bank of Queensland"
    list5.AddItem "4673 - First Card"
    list5.AddItem "4707 - Tompkins County Trust"
    list5.AddItem "4719 - Rocky Mountain"
    list5.AddItem "4721 - First Security"
    list5.AddItem "4722 - West Bank"
    list5.AddItem "4726 - Wells Fargo"
    list5.AddItem "4783 - AT&T's Universal Card"
    list5.AddItem "4784 - AT&T's Universal Card"
    list5.AddItem "4800 - M.B.N.A. North America"
    list5.AddItem "4811 - Bank of Hawaii"
    list5.AddItem "4819 - Macom Federal Crdit Union"
    list5.AddItem "4820 - IBM Mid America Federal Credit Union"
    list5.AddItem "4833 - U.S. Bank"
    list5.AddItem "4842 - Security Pacific Washington"
    list5.AddItem "4897 - Village Bank of Chicago"
    list5.AddItem "4921 - Hong Kong/National Bank"
    list5.ListIndex = 0
  'list5.additem = "4225 - Chase Manhatten Bank"
  'list5.AddItem Blah

  '
  'list5.AddItem "4013 - Citibank"
  'list5.AddItem "4060 - Associates National Bank"
  'list5.AddItem "4271 - Citibank Preferred"
  'list5.AddItem "4332 - First American Bank"
  'list5.AddItem "4436 - Security Bank and Trust"
  'list5.AddItem "4503 - Canadian Imperial Bank"
  'list5.AddItem "4564 - Bank of Queensland"
  'list5.AddItem "4726 - Wells Fargo"
  'list5.AddItem "4811 - Bank of Hawaii"
  'list5.AddItem "4833 - U.S. Bank"
  'list5.AddItem "4842 - Security Pacific Washington"
  'list5.AddItem "4921 - Hong Kong/National Bank"
  'list5.ListIndex = 0
  '
  '
  End Sub

  Sub Option3D2_Click (Value As Integer)
    list5.Clear
    list5.AddItem "5100 - Southwestern States Bankcard Association"
    list5.AddItem "5110 - Universal Travel Voucher"
    list5.AddItem "5120 - Western States Bankcard Association"
    list5.AddItem "5130 - Eurocard France"
    list5.AddItem "5140 - Mountain States Bankcard Association"
    list5.AddItem "5150 - Credit Systems, Inc."
    list5.AddItem "5160 - Westpac Banking Corporation"
    list5.AddItem "5170 - Midamerica Bankcard Association"
    list5.AddItem "5172 - First Bank Card Center"
    list5.AddItem "5180 - Computer Communications of America"
    list5.AddItem "5190 - Bank of Montreal"
    list5.AddItem "5191 - Bank of Montreal"
    list5.AddItem "5201 - Mellon Bank, N.A."
    list5.AddItem "5202 - Central Trust Company, N.A."
    list5.AddItem "5204 - Security Pacific National Bank"
    list5.AddItem "5205 - Promocion Y Operacion, S.A."
    list5.AddItem "5206 - list5.additem o Nacional de Mexico"
    list5.AddItem "5207 - New England Bankcard Association, Inc."
    list5.AddItem "5208 - Million Card Service Co., Ltd."
    list5.AddItem "5209 - The Citizens & Southern National Bank"
    list5.AddItem "5210 - Kokunai Shinpan Company, Ltd."
    list5.AddItem "5211 - Chemical Bank Delaware"
    list5.AddItem "5212 - F.C.C. National Bank"
    list5.AddItem "5213 - The Bankcard Association, Inc."
    list5.AddItem "5215 - Marine Midland Bank, N.A."
    list5.AddItem "5216 - Old Kent Bank & Trust Co."
    list5.AddItem "5217 - Union Trust"
    list5.AddItem "5218 - Citibank, N.A."
    list5.AddItem "5219 - Central Finance Co., Ltd."
    list5.AddItem "5220 - Sovran Bank/Central South"
    list5.AddItem "5221 - Standard Bank of South Africa, Ltd."
    list5.AddItem "5222 - Security Bank & Trust Company"
    list5.AddItem "5223 - Trustmark National Bank"
    list5.AddItem "5224 - Midland Bank"
    list5.AddItem "5225 - First Pennsylvania Bank, N.A."
    list5.AddItem "5226 - Eurocard Ab"
    list5.AddItem "5227 - Rocky Mountain Bankcard System, Inc"
    list5.AddItem "5228 - First Union National Bank of North Carolina"
    list5.AddItem "5229 - Sunwest Bank of Albuquerque, N.A."
    list5.AddItem "5230 - Harris Trust & Savings Bank"
    list5.AddItem "5231 - Badische Beamtenbank EG"
    list5.AddItem "5232 - Eurocard Deutschland"
    list5.AddItem "5233 - list5.additem systems Association, Inc."
    list5.AddItem "5234 - Citibank <Arizona>"
    list5.AddItem "5235 - Finanacial Transaction Systems, Inc."
    list5.AddItem "5236 - First Tennessee Bank, N.A."
    list5.ListIndex = 0

  End Sub

  Sub Option3D3_Click (Value As Integer)
    list5.Clear
    list5.AddItem "6013 - MBNA"
    list5.ListIndex = 0

  End Sub

  Sub Text2_KeyDown (KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_RETURN Then
      Call Command3D1_Click
    End If

  End Sub

  Sub Text2_KeyPress (KeyAscii As Integer)
  'If KeyAscii = 13 Then Command3D1_Click

  End Sub

  Sub Text4_KeyPress (KeyAscii As Integer)
    Static decrd(17)
    Static shh(17)

    If KeyAscii = 13 Then
    '  text5.Visible = False

      On Error GoTo assmunch
      sum = 0
    'reinitialize
      For jjjj = 1 To 16
        decrd(jjjj) = 0
      Next jjjj

    '  qqq$ = Text4.Text
      xrt = Mid(qqq$, 1, 8)
      xpt = Mid(qqq$, 9, 8)

      decard = Format$(xrt, "########")
      decard2 = Str$(xpt)

      For C = 1 To 8
        shh(C) = Mid$(decard, C, 1)
      Next C


      For c2 = 9 To 16
        shh(c2) = Mid$(decard2, c2 - 7, 1)
      Next c2

      For xr = 1 To 16
        fukaol = Val(shh(xr))
        decrd(xr) = fukaol
      Next xr



      For CC = 1 To 15 Step 2
        decrd(CC) = decrd(CC) * 2
        If decrd(CC) > 9 Then decrd(CC) = decrd(CC) - 9
      Next CC


      For bbb = 1 To 16
        sum = sum + decrd(bbb)
      Next bbb



      xre = sum Mod 10


      If xre = 0 Then
      ' label4.Visible = True
      ' label4.FontSize = 12.5
      ' label4.Caption = "How Many Cards?"
      ' text5.Visible = True
      ' text5.SetFocus

      Else
      'label4.FontSize = 7
      ' label4.Visible = True
      ' label4.Caption = "Card did not fit any known algorithm"
      End If
      Exit Sub
assmunch:
      If Err = 13 Then
        MsgBox "Please Enter The Card# With No Spaces"
      Else
        MsgBox "Something Weird Happened"
      End If
      Exit Sub

    End If





  End Sub

