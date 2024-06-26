VERSION 2.00
Begin Form frmOptions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Wrox 6 Options"
   ClientHeight    =   2550
   ClientLeft      =   5805
   ClientTop       =   2520
   ClientWidth     =   3705
   Height          =   2955
   Left            =   5745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3705
   Top             =   2175
   Width           =   3825
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   1935
      TabIndex        =   2
      Top             =   1785
      Width           =   1665
   End
   Begin CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   75
      TabIndex        =   1
      Top             =   1800
      Width           =   1665
   End
   Begin Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Interface Options"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   3525
      Begin CheckBox chkStatus 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status Bar On/Off"
         Height          =   240
         Left            =   360
         TabIndex        =   6
         Top             =   1230
         Width           =   2640
      End
      Begin CheckBox chkRuler 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show the ruler"
         Height          =   225
         Left            =   360
         TabIndex        =   5
         Top             =   945
         Width           =   2325
      End
      Begin CheckBox chkToolbar2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Display the second toolbar"
         Height          =   210
         Left            =   360
         TabIndex        =   4
         Top             =   675
         Width           =   2685
      End
      Begin CheckBox chkToolbar1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Display the top toolbar"
         Height          =   225
         Left            =   360
         TabIndex        =   3
         Top             =   375
         Width           =   2400
      End
   End
End
Option Explicit

Sub cmdCancel_Click ()

    ' User pressed cancel, so simply unload the form
    Unload frmOptions


End Sub

Sub cmdOK_Click ()

    ' User clicked OK, we need to change the previous form

End Sub

Sub Form_Load ()

    ' When the form loads, centre it
    frmOptions.Left = (Screen.Width - frmOptions.Width) / 2
    frmOptions.Top = (Screen.Height - frmOptions.Height) / 2


End Sub

