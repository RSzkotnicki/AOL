VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "MM Bot"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2700
   LinkTopic       =   "Form4"
   ScaleHeight     =   1515
   ScaleWidth      =   2700
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "5"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "/Don't MM Me"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "/MM Me"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Min(z)"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "2 Get Off"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "2 Get On"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True

End Sub

Private Sub Timer1_Timer()
Do
SendChat "BlaH AoL4 MMer by BlaH - MMing n " + Text3.text & " min(z)"
TimeOut 0.5
SendChat "Type " + Text1.text & " 2 get on"
TimeOut 0.5
SendChat "Type " + Text2.text & " 2 get off"
TimeOut 60
Text3.text = Val(Text3) - 1
If Text3.text = "0" Then
Timer1.Enabled = False
Loop Until Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Do
If LastChatLine = " " + Text1.text Then
Form2.List1.AddItem SNFromLastChatLine
SendChat "U have been added " + SNFromLastChatLine
Else
If LastChatLine = " " + Text2.text Then
Form2.List1.RemoveItem SNFromLastChatLine
SendChat "U have been removed " + SNFromLastChatLine
Else
End If
Loop Until Timer2.Enabled = False
End Sub

