' FRMMAIN.FRM
Option Explicit
Declare Function extfn2C0 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn218 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Sub extsubC8 Lib "User" Alias "GetCursorPos" (p1 As Any)
Declare Function extfn138 Lib "Kernel" Alias "GetModuleFileName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1E0 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn1A8 Lib "User" Alias "GetWindowLong" (ByVal p1%, ByVal p2%) As Long
Declare Function extfn250 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn170 Lib "User" Alias "GetWindowWord" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn288 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%) As Integer
Declare Function extfn100 Lib "User" Alias "WindowFromPoint" (ByVal p1 As Any) As Integer
Const c002C =-6 ' &HFFFA%
Const c002E =-12 ' &HFFF4%
Const c0030 =-16 ' &HFFF0%
Const c0032 =-1 ' &HFFFF%
Const c003A = 3 ' &H3%

Sub Form_Load ()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    sub330 Me
End Sub

Sub sub368 ()
End
End Sub

Sub sub330 (p0044 As Form)
Dim l0046 As Integer
    l0046% = extfn288(p0044.hWnd, c0032, 0, 0, 0, 0, c003A)
End Sub

Sub Timer1_Timer ()
Dim l0048
Dim l004C As String * 100
Dim l0050 As String * 100
Dim l0052 As Integer
Dim l0054 As Integer
Dim l0058 As String * 100
Dim l005A As Integer
Dim l005C As Long
Dim l005E As Integer
Dim l0062 As String * 100
Dim l0066 As String * 100
Static m0068 As Integer
Dim l006A As Variant
      Call extsubC8(l0048)
      l0052 = extfn100(l0048)
      If  l0052 <> m0068 Then
         m0068 = l0052
         Cls
         Print "Window Handle: "; (l0052)
         Print "Focus: "; extfn2C0()
         l006A = extfn250(l0052, l004C, 100)
         Print "Window Text: " & Left(l004C, l006A)
         l006A = extfn218(l0052, l0050, 100)
         Print "Window Class Name: "; Left(l0050, l006A)
         l005C = extfn1A8(l0052, c0030)
         Print "Window Style: "; (l005C)
         l0054 = extfn1E0(l0052)
         If  l0054 <> 0 Then
            l005A = extfn170(l0052, c002E)
            Print "Window ID Number: "; (l005A)
            Print "Parent Window Handle: "; (l0054)
            l006A = extfn250(l0054, l0062, 100)
            Print "Parent Window Text: " & Left(l0062, l006A)
            l006A = extfn218(l0054, l0058, 100)
            Print "Parent Window Class Name: "; Left(l0058, l006A)
         Else
            Print "Window ID Number: N/A"
            Print "Parent Window Handle: N/A"
            Print "Parent Window Text : N/A"
            Print "Parent Window Class Name: N/A"
         End If
         l005E = extfn170(l0052, c002C)
         l006A = extfn138(l005E, l0066, 100)
         Print "Module: "; Left(l0066, l006A)
      End If
End Sub
