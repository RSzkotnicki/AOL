Attribute VB_Name = "RsZzsFade"
'Check it out yo
' this is mah first fade bas Enjoy !!!




'Pre-set 2 color fade combinations begin here
Public Sub ImGreenlightBlueBlue(text1 As TextBox, text2 As TextBox)
      a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 560 Then f = (560 - (f - 560))
        G = RGB(f, 560 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    Call sendim(text2, "<Font Face=""Lucida Blackletter"">" & msg & " ")
    
    

End Sub

Public Sub IMRedYellowRed(text1 As TextBox, text2 As TextBox)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 1000 + f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    Call sendim(text2, "<Font Face=""Lucida Blackletter"">" & msg & "")
    
End Sub


'Pre-set 3 Color fade combinations begin here


'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function


'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


'Variable color fade functions begin here

Function BlackBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackBlue = msg
End Function

Function BlackGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGreen = msg
End Function

Function BlackGrey(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 220 / a
        f = E * B
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGrey = msg
End Function

Function BlackIce(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 560 / a
        f = E * B
        If f > 350 Then f = (370 - (f - 350))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<FONT FACE=""Snap ITC"">" & msg & ""
    
End Function

Function BlackPurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackPurple = msg
End Function

Function BlackRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackRed = msg
End Function

Function Nutsack(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 600 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & " "
    
End Function

Function BlackYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackYellow = msg
End Function

Function BlueGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueGreen = msg
End Function

Function BlueGreenLightGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 400 Then f = (400 - (f - 400))
        G = RGB(400 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
   SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
   
End Function

Function BluegreenlightgreenBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(400 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Comic Sans MS"">" & msg & ""
    
End Function

Function BlueLightBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 855 Then f = (855 - (f - 855))
        G = RGB(855 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
  SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
  
End Function

Function BluePinkRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 345 Then f = (345 - (f - 345))
        G = RGB(345 + f, 345 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""

End Function

Function BluePurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BluePurple = msg
End Function

Function BlueRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueRed = msg
End Function

Function BlueYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueYellow = msg
End Function

Sub FadeFormBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    Form1.DrawStyle = vbInsideSolid
    Form1.DrawMode = vbCopyPen
    Form1.ScaleMode = vbPixels
    Form1.DrawWidth = 2
    Form1.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FadeFormGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormPurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub

Function Fucker(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 720 Then f = (720 - (f - 720))
        G = RGB(f, 720 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
    
End Function

Function FuckNutz(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 612 / a
        f = E * B
        If f > 255 Then f = (300 - (f - 277))
        G = RGB(266, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
SendChat "<Font Face=""MS Reference1"">" & msg & ""

End Function

Function Gold(text1 As String)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 789 Then f = (767 - (f - 435))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=Tempus Sans ITC>" & msg & ""
End Function

Function GreenlightBlueBlue(text1)
      a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 560 Then f = (560 - (f - 560))
        G = RGB(f, 560 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & " "
    
End Function

Function GreenRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenRed = msg
End Function

Function GreenWhiteBlue(text1)
      a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 345 Then f = (345 - (f - 345))
        G = RGB(f, 345 + f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""

End Function

Function GreenYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenYellow = msg
End Function

Function GreyBlack(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 220 / a
        f = E * B
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyBlack = msg
End Function

Function GreyBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyBlue = msg
End Function

Function GreyBlueGrey(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyBlueGrey = msg
End Function

Function GreyGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyGreen = msg
End Function

Function GreyPurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyPurple = msg
End Function

Function GreyPurpleGrey(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyPurpleGrey = msg
End Function

Function GreyRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyRed = msg
End Function

Function GreyRedGrey(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyRedGrey = msg
End Function

Function GreyYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyYellow = msg
End Function

Function GreyYellowGrey(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyYellowGrey = msg
End Function

Function Koo(text1)
   a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 500 Then f = (500 - (f - 500))
        G = RGB(500 - f, f, 500 - f)
        
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
    
End Function

Function Manual(textz)

End Function

Function PurpleBlack(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleBlack = msg
End Function

Function PurpleBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleBlue = msg
End Function

Function PurpleBluePurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleBluePurple = msg
End Function

Function PurpleBlueRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 466 Then f = (466 - (f - 466))
        G = RGB(466 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
    
End Function

Function PurpleGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleGreen = msg
End Function

Function PurpleGreenPurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleGreenPurple = msg
End Function

Function PurpleRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleRed = msg
End Function

Function PurpleRedPurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleRedPurple = msg
End Function

Function PurpleYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleYellow = msg
End Function

Function PurpleYellowPurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleYellowPurple = msg
End Function

Function purpleyellowwhite(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 346 Then f = (346 - (f - 346))
        G = RGB(500 - f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
    
End Function

Function purty(text1 As String)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
 SendChat "<Font Face=""Comic Sans MS"">" & msg & ""
End Function

Function RedBlack(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    RedBlack = msg
End Function

Function RedBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    RedBlue = msg
End Function

Function RedBlueLightBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 346 Then f = (346 - (f - 346))
        G = RGB(f, f, 346 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
   SendChat " <Font Face=""Lucida Blackletter"">" & msg & ""
   
End Function

Function RedGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    RedGreen = msg
End Function

Function RedGreenRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    RedGreenRed = msg
End Function

Function RedLightGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 490 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
   SendChat "<Font Face=""Comic Sans MS"">" & msg & ""
   
End Function

Function RedPinkBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 400 Then f = (400 - (f - 400))
        G = RGB(400 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
End Function

Function RedPurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    RedPurple = msg
End Function

Function RedPurpleRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    RedPurpleRed = msg
End Function

Function RedWhiteBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 360 Then f = (360 - (f - 360))
        G = RGB(360 - f, 360 - f, f)
        
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
    
End Function

Function RedYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
 
End Function

Function RedYellowGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 560 Then f = (560 - (f - 560))
        G = RGB(0, f, 560 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & " "
    
End Function

Function RedYellowRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 1000 + f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Lucida Blackletter"">" & msg & ""
    
End Function

Function RGB2HEX(r, G, B)
    Dim x&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For x& = 1 To 3
        If x& = 1 Then Color& = B
        If x& = 2 Then Color& = G
        If x& = 3 Then Color& = r
        For xx& = 1 To 2
            Divide = Color& / 16
            Answer& = Int(Divide)
            Remainder& = (10000 * (Divide - Answer&)) / 625
            If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
            If Remainder& = 10 Then Configuring$ = "A" + Configuring$
            If Remainder& = 11 Then Configuring$ = "B" + Configuring$
            If Remainder& = 12 Then Configuring$ = "C" + Configuring$
            If Remainder& = 13 Then Configuring$ = "D" + Configuring$
            If Remainder& = 14 Then Configuring$ = "E" + Configuring$
            If Remainder& = 15 Then Configuring$ = "F" + Configuring$
            Color& = Answer&
        Next xx&
    Next x&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Function test1(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 700 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
    SendChat "<Font Face=""Comic Sans MS"">" & msg & ""
    
End Function

Function ThreeColors(text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, WavY As Boolean)

'This code is still buggy, use at your own risk

    d = Len(text)
        If d = 0 Then GoTo TheEnd
        If d = 1 Then Fade1 = text
    For x = 2 To 500 Step 2
        If d = x Then GoTo Evens
    Next x
    For x = 3 To 501 Step 2
        If d = x Then GoTo Odds
    Next x
Evens:
    c = d \ 2
    Fade1 = Left(text, c)
    Fade2 = Right(text, c)
    GoTo TheEnd
Odds:
    c = d \ 2
    Fade1 = Left(text, c)
    Fade2 = Right(text, c + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If WavY = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If WavY = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If WavY = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If WavY = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    msg = FadeA + FadeB
    ThreeColors = msg
End Function

Function TrimSpaces(text)
    If InStr(text, " ") = 0 Then
    TrimSpaces = text
    Exit Function
    End If
    For TrimSpace = 1 To Len(text)
    thechar$ = Mid(text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
End Function

Function TwoColors(text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = C1
    C2BAK = C2
    C3BAK = C3
    C4BAK = C4
    c = 0
    o = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For x = 1 To Len(text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(text) * x) + Red1
        VAL2 = (BVAL2 / Len(text) * x) + Green1
        VAL3 = (BVAL3 / Len(text) * x) + Blue1
        
        C1 = RGB2HEX(VAL1, VAL2, VAL3)
        C2 = RGB2HEX(VAL1, VAL2, VAL3)
        C3 = RGB2HEX(VAL1, VAL2, VAL3)
        C4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: msg = msg & "<FONT COLOR=#" + C1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        
        If WavY = True Then
            If o2 = 1 Then msg = msg + "<SUB>"
            If o2 = 3 Then msg = msg + "<SUP>"
            msg = msg + Mid$(text, x, 1)
            If o2 = 1 Then msg = msg + "</SUB>"
            If o2 = 3 Then msg = msg + "</SUP>"
            If Q2 = 2 Then
                Q = 1
                Q2 = 1
                If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
                If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
                If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
                If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
            End If
        ElseIf WavY = False Then
            msg = msg + Mid$(text, x, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        End If
nc:     Next x
    C1 = C1BAK
    C2 = C2BAK
    C3 = C3BAK
    C4 = C4BAK
    TwoColors = msg
End Function

Function Weird(text1)

    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 250 Then f = (300 - (f - 230))
        G = RGB(250, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d & "<b>"
    Next B
SendChat "<Font Face=""Lucida Handwriting"">" & msg & ""

End Function

Function YellowBlue(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowBlue = msg
End Function

Function YellowBlueYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowBlueYellow = msg
End Function

Function YellowGreen(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowGreen = msg
End Function

Function YellowPurple(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowPurple = msg
End Function

Function YellowPurpleYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowPurpleYellow = msg
End Function

Function YellowRed(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 255 / a
        f = E * B
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowRed = msg
End Function

Function YellowRedYellow(text1)
    a = Len(text1)
    For B = 1 To a
        c = Left(text1, B)
        d = Right(c, 1)
        E = 510 / a
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowRedYellow = msg
End Function

