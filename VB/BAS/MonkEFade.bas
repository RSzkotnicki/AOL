Attribute VB_Name = "MonkEFade"

Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000
Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
dacolor10$ = RGBtoHEX(Colr10)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + Right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, TheText, Wavy)

End Function

Function FadeByColor2(Colr1, Colr2, TheText$, Wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, Wavy)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, Wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, Wavy)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, TheText$, Wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, TheText, Wavy)

End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, TheText$, Wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, TheText, Wavy)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, TheText$, Wavy As Boolean)

    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(TheText, frthlen)
    
    'part1
    textlen% = Len(part1)
    For i = 1 To textlen%
        TextDone$ = Left(part1, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2)
    For i = 1 To textlen%
        TextDone$ = Left(part2, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3)
    For i = 1 To textlen%
        TextDone$ = Left(part3, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4)
    For i = 1 To textlen%
        TextDone$ = Left(part4, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, TheText$, Wavy As Boolean)
'by monk-e-god
    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(TheText, fstlen% + seclen% + thrdlen + 1, frthlen%)
    part5$ = Mid(TheText, fstlen% + seclen% + thrdlen + frthlen + 1, fithlen%)
    part6$ = Mid(TheText, fstlen% + seclen% + thrdlen + frthlen + fithlen + 1, sixlen%)
    part7$ = Mid(TheText, fstlen% + seclen% + thrdlen + frthlen + fithlen + sixlen + 1, sevlen%)
    part8$ = Mid(TheText, fstlen% + seclen% + thrdlen + frthlen + fithlen + sixlen + sevlen + 1, eightlen%)
    part9$ = Right(TheText, ninelen)
    
    'part1
    textlen% = Len(part1)
    For i = 1 To textlen%
        TextDone$ = Left(part1, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2)
    For i = 1 To textlen%
        TextDone$ = Left(part2, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3)
    For i = 1 To textlen%
        TextDone$ = Left(part3, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4)
    For i = 1 To textlen%
        TextDone$ = Left(part4, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5)
    For i = 1 To textlen%
        TextDone$ = Left(part5, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6)
    For i = 1 To textlen%
        TextDone$ = Left(part6, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7)
    For i = 1 To textlen%
        TextDone$ = Left(part7, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8)
    For i = 1 To textlen%
        TextDone$ = Left(part8, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9)
    For i = 1 To textlen%
        TextDone$ = Left(part9, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function


Function InverseColor(OldColor)
dacolor$ = RGBtoHEX(OldColor)
redx% = Val("&H" + Right(dacolor$, 2))
greenx% = Val("&H" + Mid(dacolor$, 3, 2))
bluex% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - redx%
newgreen% = 255 - greenx%
newblue% = 255 - bluex%
InverseColor = RGB(newred%, newgreen%, newblue%)

End Function

Function RGBtoHEX(RGB)
'heh, I didnt make this one...
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function

Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, TheText$, Wavy As Boolean)
'by monk-e-god
    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Right(TheText, thrdlen)
    
    'part1
    textlen% = Len(part1)
    For i = 1 To textlen%
        TextDone$ = Left(part1, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2)
    For i = 1 To textlen%
        TextDone$ = Left(part2, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3)
    For i = 1 To textlen%
        TextDone$ = Left(part3, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Sub FadePreview(PreTxtMain As Control, FadedText As String, PreTxt As TextBox)
'by monk-e-god
'-FADE PREVIEW-
'To use the fadepreview you need a
'rich textbox (which requires an ocx)
'and a regular text box in which the
'HTML will be interpreted.

'example:
'HTMLbox.Text = FadeByColor4(FADE_RED, FADE_BLACK, FADE_GREY, FADE_GREEN, "Red/Black/Grey/Green Fade Preview", False)
'Call FadePreview(PreviewBox, HTMLbox.Text, InvisBox)

'now in the rich textbox, PreviewBox, you
'will see a Red to Black to Grey to Green
'fade saying "Red/Black/Grey/Green Fade Preview"

'NOTE: You cannot preview wavy fades.
'NOTE: PreTxtMain MUST be a rich textbox!

PreTxtMain.Text = ""
Dim Starts()
Dim Lengths()
Dim Colors()
Dim LastHtml%
Dim CurStart%
Dim CurLen%
Dim CurColor$
Dim NumFades%
PreTxt.Text = FadedText
NumFades = 0
LastHtml = 2
findhtml = 1
While findhtml
If NumFades = 0 Then findhtml = 0

NumFades = NumFades + 1
findhtml = InStr(findhtml + 1, PreTxt.Text, "<Font Color=#") 'InStr(LastHtml - 1, PreTxt.Text, "<Font Color=#")
If findhtml = 0 Then GoTo Blah
LastHtml = InStr(findhtml + 1, PreTxt.Text, ">")
thecolor = Mid(PreTxt.Text, findhtml + 13, 6)
htmlblue$ = Right(thecolor, 2)
htmlgreen$ = Mid(thecolor, 3, 2)
htmlred$ = Left(thecolor, 2)
vbcolor = "&H00" + htmlblue + htmlgreen + htmlred + "&"

nexthtml% = InStr(findhtml + 1, PreTxt.Text, "<Font Color=#")
CurLen = 1
Firstpart$ = Left(PreTxt.Text, findhtml - 1)
Secondpart$ = Mid(PreTxt.Text, LastHtml + 1)
PreTxt.Text = Firstpart$ + Secondpart$
CurStart = findhtml
CurColor = vbcolor
ReDim Preserve Starts(NumFades)
ReDim Preserve Lengths(NumFades)
ReDim Preserve Colors(NumFades)
Starts(NumFades) = CurStart - 1
Lengths(NumFades) = CurLen
Colors(NumFades) = CurColor

Blah:
h = h
Wend
PreTxtMain.Text = PreTxt.Text
For cc = 1 To NumFades - 1
PreTxtMain.SelStart = Starts(cc)
PreTxtMain.SelLength = Lengths(cc)
PreTxtMain.SelColor = Val(Colors(cc))
h = h
Next cc
PreTxtMain.SelLength = 0

End Sub

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, Wavy As Boolean)
'by monk-e-god
    textlen% = Len(TheText)
    fstlen% = (Int(textlen) / 2)
    part1$ = Left(TheText, fstlen)
    part2$ = Right(TheText, textlen - fstlen)
    'part1
    textlen% = Len(part1)
    For i = 1 To textlen%
        TextDone$ = Left(part1, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2)
    For i = 1 To textlen%
        TextDone$ = Left(part2, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, Wavy As Boolean)
'by monk-e-god
    textlen$ = Len(TheText)
    For i = 1 To textlen$
        TextDone$ = Left(TheText, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    FadeTwoColor = Faded$
End Function
