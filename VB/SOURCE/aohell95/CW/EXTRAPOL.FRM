VERSION 2.00
Begin Form Extrapolate 
   BackColor       =   &H8000000F&
   Caption         =   "Extapolate Cards"
   ClientHeight    =   3810
   ClientLeft      =   2025
   ClientTop       =   1650
   ClientWidth     =   5085
   Height          =   4215
   Icon            =   EXTRAPOL.FRX:0000
   Left            =   1965
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5085
   Top             =   1305
   Width           =   5205
   Begin TextBox Text3 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin SSCommand Command3D2 
      BevelWidth      =   3
      Caption         =   "&Cancel"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin SSCommand Command3D1 
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
      Left            =   960
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin TextBox Text2 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   420
      Left            =   1680
      MaxLength       =   21
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin TextBox Text1 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   15
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   4575
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000F&
      Caption         =   "Num. of cards to extrapolate:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   2490
   End
   Begin Label Label1 
      BackColor       =   &H8000000F&
      Caption         =   "Card number to Extrapolate:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End

  Sub Command3D1_Click ()
    Static extrapcard$(9999)
    Static cardstring2(20)
    Static dacrd(20)
    Static s(25)

    cccc$ = text2.Text
    ccc = checkvalid(cccc$)
    Randomize

    If ccc = False Then
      msg = "The Card Number you entered is Invalid Extrapolate Anyway?"
      response = MsgBox(msg, 4, "Invalid Card")
      If response = 6 Then
        GoTo finlater
      Else
        text1.Text = "funk Dat"
        Exit Sub
      End If
    End If
finlater:

    For x4 = 1 To 12
      totalcard$ = totalcard$ + cardstring(x4)
    Next x4

    For times = 1 To text3.Text + 1
Repeatrnd:
      fullcard$ = totalcard$
      For x = 13 To 16
        dacrd(x) = Str(Int((Rnd * 10)))
      Next x
      For x1 = 13 To 16
        fullcard$ = fullcard$ + dacrd(x1)
      Next x1
      checkextrap = checkvalid(fullcard$)
      If checkextrap = True Then
        extrapcard$(kont) = fullcard$
        kont = kont + 1
      Else
        GoTo Repeatrnd
      End If
    Next times
    endofline = Chr$(13) + Chr$(10)

    For knt = 1 To (times - 2)
      format1 = Mid$(extrapcard$(knt), 1, 4)
      format2 = Mid$(extrapcard$(knt), 5, 4)
      format3 = Mid$(extrapcard$(knt), 9, 4)
      For xmen = 13 To 25
        s(xmen) = Mid$(extrapcard$(knt), xmen, 1)
        If s(xmen) <> " " Then format4 = format4 + s(xmen)
      Next xmen
      THEFuLLString$ = THEFuLLString$ & knt & ".) " & format1 & " " & format2 & " " & format3 & " " & format4 & endofline
      format4 = ""
    Next knt

    text1.Text = THEFuLLString$
  End Sub

  Sub Command3D2_Click ()
  'thecards$ = text2.Text
  'cantdoextrap = False
  'onecard$ = extrapolat(thecards$)
  'If cantdoextrap = False Then
  'text1.Text = onecard$
  'Else
  'MsgBox "howdy fucknut"
  'End If
    Unload Me

  End Sub

