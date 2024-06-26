VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MMer Example by bEaV"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "MM Bot"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   3480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   3435
      TabIndex        =   12
      Top             =   0
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Kill Dupes"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Back"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "View Mail b'n Sent"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Peepz 2 MM"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Mail"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Select Mail!"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AddMailList (List1.List)
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Command3.Visible = False
Command4.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
Command4.Visible = False
Command3.Visible = True
Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
Command5.Visible = False
Command6.Visible = True
Form1.Height = 3810
End Sub

Private Sub Command6_Click()
Command6.Visible = False
Command5.Visible = True
Form1.Height = 2625
End Sub

Private Sub Command7_Click()
KillDupes (List2.List)
End Sub

Private Sub Command8_Click()
Form4.Show
End Sub

Private Sub Form_Load()
OpenNewMails
CountMail
Label2.Caption = " " + CountMail
End Sub

Private Sub List1_Click()
List2.AddItem List1.ListIndex ' adds  all mail 2 b MMed 2 the other list box where u
' send
End Sub


Private Sub Timer1_Timer()
OpenNewMails
Do
PercentBar ("picture1"), "label1", Label2
ReadMail
ClickForward
ForwardMail ("list1.list"), "Blah AoL 4.0 MMeR mailing " + Label2.Caption & " mail(z)<br>This is mail #: " + Label1.Caption & "<br>MM brought 2 u by: " + UserSN
ClickKeepAsNew
Loop Until Timer1.Enabled = False
End Sub
