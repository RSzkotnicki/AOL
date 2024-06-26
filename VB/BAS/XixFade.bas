Attribute VB_Name = "XixFade"
'================
'\¯\    /¯/  |¯|  \¯\    /¯/
'  \  \/  /    |  |    \  \/  /
'  /  /\  \    |  |    /  /\  \
'/_/    \_\  |_|  /_/    \_\
'================
'        ßÿ XïXâs
'This bas was made using the
'TwoColor and ThreeColor Functions
'From Cryofade.Bas.
'I take absolutely no credit for those
'Functions but all the color fades
'Higher than 3 colors are made my me
'Using the TwoColor and ThreeColor
'Functions as their bases.
'I also want to thank RangA! for the
'SeeFade function code.
'This bas may be freely distributed.
'--------
'P.S.
' The wavy function may not fully werk
' In some of the higher numbers of fades.
' I tried to fix it but could not find the prob.
' I recomend that u do not try to use the
' Wavy function in the higher fades.

Function ThirteenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Red12, Green12, Blue12, Red13, Green13, Blue13, Wavy As Boolean)

Do While Len(Text) < 7
Text = Text + " "
Loop
If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 4
B = 5
c = 7
d = 8
E = 9
Do While a <> 106
    If Len(Text) = a Then
        Text = Text + "  "
    End If
    If Len(Text) = B Then
        Text = Text + " "
    End If
    If Len(Text) = c Then
        Text = Text + "     "
    End If
    If Len(Text) = d Then
        Text = Text + "    "
    End If
    If Len(Text) = E Then
        Text = Text + "   "
    End If
    a = a + 6
    B = B + 6
    c = c + 6
    d = d + 6
    E = E + 6
Loop
p = Len(Text) / 6
Thirteen1 = ThreeColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Thirteen2 = ThreeColors(Mid(Text, p + 1, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Thirteen3 = ThreeColors(Mid(Text, p + p + 1, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy)
Thirteen4 = ThreeColors(Mid(Text, p + p + p + 1, p), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy)
Thirteen5 = ThreeColors(Mid(Text, p + p + p + p + 1, p), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Wavy)
Thirteen6 = ThreeColors(Right(Text, p), Red11, Green11, Blue11, Red12, Green12, Blue12, Red13, Green13, Blue13, Wavy)
ThirteenColors = Thirteen1 + Thirteen2 + Thirteen3 + Thirteen4 + Thirteen5 + Thirteen6

End Function
Function TwelveColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Red12, Green12, Blue12, Wavy As Boolean)

Do While Len(Text) < 7
Text = Text + " "
Loop
a = 4
B = 5
c = 7
d = 8
E = 9
Do While a <> 106
    If Len(Text) = a Then
        Text = Text + "  "
    End If
    If Len(Text) = B Then
        Text = Text + " "
    End If
    If Len(Text) = c Then
        Text = Text + "     "
    End If
    If Len(Text) = d Then
        Text = Text + "    "
    End If
    If Len(Text) = E Then
        Text = Text + "   "
    End If
    a = a + 6
    B = B + 6
    c = c + 6
    d = d + 6
    E = E + 6
Loop
p = Len(Text) / 6
Twelve1 = ThreeColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Twelve2 = ThreeColors(Mid(Text, p + 1, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Twelve3 = ThreeColors(Mid(Text, p + p + 1, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy)
Twelve4 = ThreeColors(Mid(Text, p + p + p + 1, p), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy)
Twelve5 = ThreeColors(Mid(Text, p + p + p + p + 1, p), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Wavy)
Twelve6 = TwoColors(Right(Text, p), Red11, Green11, Blue11, Red12, Green12, Blue12, Wavy)
TwelveColors = Twelve1 + Twelve2 + Twelve3 + Twelve4 + Twelve5 + Twelve6
End Function

Function ElevenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Wavy As Boolean)


If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 4
B = 6
c = 7
d = 8
Do While a <> 104
    If Len(Text) = a Then
        Text = Text + " "
    End If
    If Len(Text) = B Then
        Text = Text + "    "
    End If
    If Len(Text) = c Then
        Text = Text + "   "
    End If
    If Len(Text) = d Then
        Text = Text + "  "
    End If
    a = a + 5
    B = f + 5
    c = c + 5
    d = d + 5
Loop
p = Len(Text) / 5
Eleven1 = ThreeColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Eleven2 = ThreeColors(Mid(Text, p + 1, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Eleven3 = ThreeColors(Mid(Text, p + p + 1, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy)
Eleven4 = ThreeColors(Mid(Text, p + p + p + 1, p), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy)
Eleven5 = ThreeColors(Right(Text, p), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Wavy)
ElevenColors = Eleven1 + Eleven2 + Eleven3 + Eleven4 + Eleven5

End Function

Function TenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)


If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 4
B = 6
c = 7
d = 8
Do While a <> 104
    If Len(Text) = a Then
        Text = Text + " "
    End If
    If Len(Text) = B Then
        Text = Text + "    "
    End If
    If Len(Text) = c Then
        Text = Text + "   "
    End If
    If Len(Text) = d Then
        Text = Text + "  "
    End If
    a = a + 5
    B = f + 5
    c = c + 5
    d = d + 5
Loop
p = Len(Text) / 5
Ten1 = ThreeColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Ten2 = ThreeColors(Mid(Text, p + 1, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Ten3 = ThreeColors(Mid(Text, p + p + 1, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy)
Ten4 = ThreeColors(Mid(Text, p + p + p + 1, p), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy)
Ten5 = TwoColors(Right(Text, p), Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy)
TenColors = Ten1 + Ten2 + Ten3 + Ten4 + Ten5

End Function
Function NineColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)

text2 = Text
If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 5
Do While a <> 101
    If Len(Text) = a Then
        Text = Text + " "
    End If
    a = a + 2
Loop
p = Len(Text) / 2
Nine1 = FiveColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Nine2 = FiveColors(Right(Text, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy)
NineColors = Nine1 + Nine2
If Mid(NineColors, Len(NineColors) - 5, 1) = " " And Right(text2, 1) <> " " Then
NineColors = Left(NineColors, Len(NineColors) - 6)
End If
If Mid(NineColors, Len(NineColors) - 6, 1) = " " And Right(text2, 1) <> " " Then
NineColors = Left(NineColors, Len(NineColors) - 7)
End If
End Function
Function SeeFade(R1, G1, B1, R2, B2, G2, pctre)
'i have often found that this will only work once,
'so for this reason i recomend u copy and paste
'the code into the Paint Proc of a picture box.
'This only shows 2 colors faded at a time.

On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double

Static SplitNum(3) As Double
Static DivideNum(3) As Double

Dim FadeW As Integer
Dim Loo As Integer
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100
FadeW = pctre.Width / 100
For Loo = 0 To 100

pctre.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)

Next Loo

End Function
Function EightColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)

text2 = Text
If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 4
B = 5
c = 6
d = 8
E = 9
f = 10
Do While a <> 102
    If Len(Text) = a Then
        Text = Text + "   "
    End If
    If Len(Text) = B Then
        Text = Text + "  "
    End If
    If Len(Text) = c Then
        Text = Text + " "
    End If
    If Len(Text) = d Then
        Text = Text + "      "
    End If
    If Len(Text) = E Then
        Text = Text + "     "
    End If
    If Len(Text) = f Then
        Text = Text + "    "
    End If
    a = a + 7
    B = B + 7
    c = c + 7
    d = d + 7
    E = E + 7
    f = f + 7
Loop
p = Len(Text) / 7
Eight1 = TwoColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy)
Eight2 = TwoColors(Mid(Text, p + 1, p), Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Eight3 = TwoColors(Mid(Text, p + p + 1, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy)
Eight4 = TwoColors(Mid(Text, p + p + p + 1, p), Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Eight5 = TwoColors(Mid(Text, p + p + p + p + 1, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy)
Eight6 = TwoColors(Mid(Text, p + p + p + p + p + 1, p), Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy)
Eight7 = TwoColors(Right(Text, p), Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy)
EightColors = Eight1 + Eight2 + Eight3 + Eight4 + Eight5 + Eight6 + Eight7

End Function

Function SevenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)

text2 = Text
If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 4
B = 5
c = 7
d = 8
E = 9
Do While a <> 106
    If Len(Text) = a Then
        Text = Text + "  "
    End If
    If Len(Text) = B Then
        Text = Text + " "
    End If
    If Len(Text) = c Then
        Text = Text + "     "
    End If
    If Len(Text) = d Then
        Text = Text + "    "
    End If
    If Len(Text) = E Then
        Text = Text + "   "
    End If
    a = a + 6
    B = B + 6
    c = c + 6
    d = d + 6
    E = E + 6
Loop
p = Len(Text) / 6
Seven1 = TwoColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy)
Seven2 = TwoColors(Mid(Text, p + 1, p), Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Seven3 = TwoColors(Mid(Text, p + p + 1, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy)
Seven4 = TwoColors(Mid(Text, p + p + p + 1, p), Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Seven5 = TwoColors(Mid(Text, p + p + p + p + 1, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy)
Seven6 = TwoColors(Right(Text, p), Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy)
SevenColors = Seven1 + Seven2 + Seven3 + Seven4 + Seven5 + Seven6

End Function
Function SixColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)

text2 = Text
If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 4
B = 6
c = 7
d = 8
Do While a <> 104
    If Len(Text) = a Then
        Text = Text + " "
    End If
    If Len(Text) = B Then
        Text = Text + "    "
    End If
    If Len(Text) = c Then
        Text = Text + "   "
    End If
    If Len(Text) = d Then
        Text = Text + "  "
    End If
    a = a + 5
    B = f + 5
    c = c + 5
    d = d + 5
Loop
p = Len(Text) / 5
Six1 = TwoColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy)
Six2 = TwoColors(Mid(Text, p + 1, p), Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Six3 = TwoColors(Mid(Text, p + p + 1, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy)
Six4 = TwoColors(Mid(Text, p + p + p + 1, p), Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
Six5 = TwoColors(Right(Text, p), Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy)
SixColors = Six1 + Six2 + Six3 + Six4 + Six5
End Function

Function FourColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)

text2 = Text
If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 5
f = 4
Do While a <> 101 And f <> 100
    If Len(Text) = a Then
        Text = Text + " "
    End If
    If Len(Text) = f Then
        Text = Text + "  "
    End If
    a = a + 3
    f = f + 3
Loop
p = Len(Text) / 3
Four1 = TwoColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy)
Four2 = TwoColors(Mid(Text, p + 1, p), Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Four3 = TwoColors(Right(Text, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy)
FourColors = Four1 + Four2 + Four3
If Mid(FourColors, Len(FourColors) - 5, 1) = " " And Right(text2, 1) <> " " Then
FourColors = Left(FourColors, Len(FourColors) - 6)
End If
If Mid(FourColors, Len(FourColors) - 6, 1) = " " And Right(text2, 1) <> " " Then
FourColors = Left(FourColors, Len(FourColors) - 7)
End If
End Function
























Function FiveColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)

text2 = Text
If Len(Text) < 4 Then
    Text = Text + "   "
End If
a = 5
Do While a <> 101
    If Len(Text) = a Then
        Text = Text + " "
    End If
    a = a + 2
Loop
p = Len(Text) / 2
Five1 = ThreeColors(Left(Text, p), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy)
Five2 = ThreeColors(Right(Text, p), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy)
FiveColors = Five1 + Five2
If Mid(FiveColors, Len(FiveColors) - 5, 1) = " " And Right(text2, 1) <> " " Then
FiveColors = Left(FiveColors, Len(FiveColors) - 6)
End If
If Mid(FiveColors, Len(FiveColors) - 6, 1) = " " And Right(text2, 1) <> " " Then
FiveColors = Left(FiveColors, Len(FiveColors) - 7)
End If
End Function
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
















'Variable color fade functions begin here


Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
    C1BAK = C1
    C2BAK = C2
    C3BAK = C3
    C4BAK = C4
    c = 0
    o = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For X = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(Text) * X) + Red1
        VAL2 = (BVAL2 / Len(Text) * X) + Green1
        VAL3 = (BVAL3 / Len(Text) * X) + Blue1
        
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
        
        If Wavy = True Then
            If o2 = 1 Then msg = msg + "<SUB>"
            If o2 = 3 Then msg = msg + "<SUP>"
            msg = msg + Mid$(Text, X, 1)
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
        ElseIf Wavy = False Then
            msg = msg + Mid$(Text, X, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        End If
nc:     Next X
    C1 = C1BAK
    C2 = C2BAK
    C3 = C3BAK
    C4 = C4BAK
    TwoColors = msg
End Function

Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)


    d = Len(Text)
        If d = 0 Then GoTo TheEnd
        If d = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If d = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If d = X Then GoTo Odds
    Next X
Evens:
    c = d \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c)
    GoTo TheEnd
Odds:
    c = d \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If Wavy = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If Wavy = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If Wavy = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If Wavy = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    msg = FadeA + FadeB
    ThreeColors = msg
End Function

Function RGB2HEX(R, G, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
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
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Function TrimSpaces(Text)
    If InStr(Text, " ") = 0 Then
    TrimSpaces = Text
    Exit Function
    End If
    For TrimSpace = 1 To Len(Text)
    thechar$ = Mid(Text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
End Function

