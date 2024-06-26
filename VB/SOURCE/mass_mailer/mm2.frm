VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Peepz 2 MM"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form2"
   ScaleHeight     =   2595
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "?"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MM Bot"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hide dis Win"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add the Room"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
List1.Clear
End Sub

Private Sub Command4_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command5_Click()
Form2.Hide
End Sub

Private Sub Command7_Click()
MsgBox "add room adds the room, clear clears it, add adds a person, MM Bot alows peepz 2 get on and off if they want", vbInformation, "?"
End Sub

Private Sub List1_Click()

End Sub
