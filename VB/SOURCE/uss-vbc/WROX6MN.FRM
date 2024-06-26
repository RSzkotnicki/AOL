VERSION 2.00
Begin Form frmMain 
   Caption         =   "Wrox 6 - Messy Office 4"
   ClientHeight    =   5820
   ClientLeft      =   2760
   ClientTop       =   1905
   ClientWidth     =   7365
   Height          =   6510
   Left            =   2700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Top             =   1275
   Width           =   7485
   WindowState     =   2  'Maximized
   Begin PictureBox picRuler 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   0
      Picture         =   WROX6MN.FRX:0000
      ScaleHeight     =   300
      ScaleWidth      =   7335
      TabIndex        =   4
      Top             =   885
      Width           =   7365
   End
   Begin PictureBox picStatus 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   WROX6MN.FRX:2CDE
      ScaleHeight     =   315
      ScaleWidth      =   7365
      TabIndex        =   3
      Top             =   5505
      Width           =   7365
   End
   Begin TextBox txtEditArea 
      BorderStyle     =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1065
      Left            =   1785
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2220
      Width           =   1305
   End
   Begin PictureBox picToolbar2 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   435
      Left            =   0
      Picture         =   WROX6MN.FRX:6C78
      ScaleHeight     =   405
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   450
      Width           =   7365
   End
   Begin PictureBox picToolbar 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   450
      Left            =   0
      Picture         =   WROX6MN.FRX:BB82
      ScaleHeight     =   420
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7365
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
   End
   Begin Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin Menu mnuView 
      Caption         =   "&View"
   End
   Begin Menu mnuInsert 
      Caption         =   "&Insert"
   End
   Begin Menu mnuFormat 
      Caption         =   "&Format"
   End
   Begin Menu mnuTools 
      Caption         =   "&Tools"
      Begin Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin Menu mnuTable 
      Caption         =   "&Table"
   End
   Begin Menu mnuWindow 
      Caption         =   "&Window"
   End
   Begin Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Option Explicit

Sub Form_Resize ()

    ' When the form is resized, the text box where the user enters the text
    ' also needs to be resized
    txtEditArea.Top = picToolbar.Height + picRuler.Height + picToolbar2.Height
    txtEditArea.Left = 0

    txtEditArea.Height = frmMain.ScaleHeight - (picToolbar.Height + picRuler.Height - picToolbar2.Height - picStatus.Height)
    txtEditArea.Width = frmMain.ScaleWidth

End Sub

Sub picRuler_Click ()

    ' Remember, this is just a try it, clicking on the ruler does nothing
    MsgBox "The current version of Wrox 6 doesn't support rulers - they just look nice", , "Dopey User Error"

End Sub

Sub picToolbar_Click ()

    ' The toolbar doesnt really work, so we'd better tell the user
    MsgBox "This is only a Try It Out, the toolbar doesn't actually work", , "Don't be daft"

End Sub

