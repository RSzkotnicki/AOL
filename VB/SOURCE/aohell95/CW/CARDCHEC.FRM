VERSION 2.00
Begin Form CardCheck 
   BackColor       =   &H8000000F&
   Caption         =   "Check Validity"
   ClientHeight    =   1620
   ClientLeft      =   2730
   ClientTop       =   2460
   ClientWidth     =   3330
   Height          =   2025
   Icon            =   CARDCHEC.FRX:0000
   Left            =   2670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3330
   Top             =   2115
   Width           =   3450
   Begin SSCommand Command3D1 
      BevelWidth      =   3
      Caption         =   "&Check The Card"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin TextBox Text1 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin Label Label3 
      BackColor       =   &H8000000F&
      Caption         =   "Do not enter any hyphens (""-""'s)"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin Label Label2 
      BackColor       =   &H8000000F&
      Caption         =   "Card #:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End

  Sub Command3D1_Click ()
    shit$ = Text1.Text
  'label1.Visible = True
    bankname = Mid(shit$, 1, 4)
    Cardtype = Mid(shit$, 1, 1)


    THECARD = CheckValid(shit$)
  'label1.Caption = "what the fuck"

    If THECARD = True Then
      If Cardtype = "4" Then Kard = "Visa"
      If Cardtype = "3" Then Kard = "American Express"
      If Cardtype = "5" Then Kard = "Master Card"
      If Cardtype = "6" Then Kard = "Discover"
      banksname = banc(bankname)
      If banc(bankname) = "" Then banksname = " Unknown "
      MsgBox "Valid " & Kard & " of " & banksname & " Bank Issue", 64, "Valid Card"
    End If
    If THECARD = False Then
      MsgBox "Card did not fit any known algorithm", 64, "Invalid Card"
    End If
  End Sub

  Sub Form_Load ()
    Call loadbanks

  End Sub

  Sub Text1_KeyDown (KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_RETURN Then
      Call Command3D1_Click
    End If

  End Sub

  Sub Text1_KeyPress (KeyAscii As Integer)
    If KeyCode = KEY_RETURN Then
      Call Command3D1_Click
    End If

  End Sub

