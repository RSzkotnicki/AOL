VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
   Height          =   2475
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   7365
   Top             =   1140
   Width           =   7485
   Begin PictureBox Picture1 
      Height          =   1695
      Left            =   150
      ScaleHeight     =   1665
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   120
      Width           =   7005
      Begin CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   480
         Left            =   5610
         TabIndex        =   1
         Top             =   1110
         Width           =   1230
      End
   End
End
Option Explicit

Sub cmdOK_Click ()

    Unload frmMain

End Sub

Sub Form_Load ()

    Do While list1.ListCount < 100

        list1.AddItem "Item " & list1.ListCount

    Loop

End Sub

