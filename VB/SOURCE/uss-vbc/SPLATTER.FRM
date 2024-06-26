VERSION 2.00
Begin Form screen1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Double
   ClientHeight    =   5775
   ClientLeft      =   1185
   ClientTop       =   840
   ClientWidth     =   7335
   Height          =   6180
   Left            =   1125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7335
   Top             =   495
   Width           =   7455
   WindowState     =   2  'Maximized
End
Option Explicit

Sub Command1_Click ()
End
End Sub

Sub Form_Load ()
    Do
    Dim nXCoord As Integer
    Dim nYCoord As Integer

    Dim nRed As Integer
    Dim nGreen As Integer
    Dim nBlue As Integer

    Dim nIndex As Integer

    ' First, we show the form to make it appear on the screen, before the
    ' load event is finished.
    screen1.Show

    ' The Doevents method is described in full in chapter 14 - it lets windows get on with
    ' updating the screen and responding to our events, before our code kicks in.
    DoEvents

    ' The randomize method tells VB to get ready to produce some random numbers.
    ' See chapter 4, variables and data types for more information.
    Randomize

    For nIndex = 1 To 2000

        ' First we need to work out some random coordinates to plot the dot at
        nXCoord = Int(Rnd(1) * screen1.ScaleWidth)
        nYCoord = Int(Rnd(1) * screen1.ScaleHeight)
    
        ' Then, random values for the red, green and blue settings of the dot's colour are
        ' decided upon.
        nRed = Int(Rnd(1) * 255)
        nGreen = Int(Rnd(1) * 255)
        nBlue = Int(Rnd(1) * 255)

        ' Finally, the point is actually drawn, using the Pset routine
        PSet (nXCoord, nYCoord), RGB(nRed, nGreen, nBlue)

    Next

   Loop

End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
End
End Sub

