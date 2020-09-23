VERSION 5.00
Begin VB.Form mesh11 
   Caption         =   "Circuit Description"
   ClientHeight    =   4140
   ClientLeft      =   4860
   ClientTop       =   5430
   ClientWidth     =   3270
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mesh11.frx":0000
   LinkTopic       =   "Mesh"
   MousePointer    =   3  'I-Beam
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   3270
   WindowState     =   1  'Minimized
   Begin VB.TextBox stripcall 
      BackColor       =   &H00FFC0C0&
      Height          =   360
      Left            =   1560
      TabIndex        =   10
      Text            =   "Mesh1.stripcall"
      Top             =   0
      Width           =   1452
   End
   Begin VB.ListBox depvolt 
      Height          =   1020
      Left            =   1200
      TabIndex        =   9
      Top             =   2640
      Width           =   1212
   End
   Begin VB.ListBox depres 
      Height          =   1020
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1212
   End
   Begin VB.ListBox semic 
      Height          =   1020
      Left            =   2400
      TabIndex        =   7
      Top             =   2640
      Width           =   852
   End
   Begin VB.ListBox Volt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   972
   End
   Begin VB.ListBox Resist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   972
   End
   Begin VB.ListBox LOOPS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   852
   End
   Begin VB.ListBox LINKS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   492
   End
   Begin VB.Label Number 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5760
      TabIndex        =   6
      Top             =   2760
      Width           =   252
   End
   Begin VB.Label NODE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5160
      TabIndex        =   5
      Top             =   2760
      Width           =   372
   End
   Begin VB.Label BRANCH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4680
      TabIndex        =   4
      Top             =   2760
      Width           =   252
   End
End
Attribute VB_Name = "mesh11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MathDoc, prcall
Public Function STR2(a)
' Corrected str function
a = LTrim(Str(a))
If InStr(a, "E") Or InStr(a, "D") Then
  xp = Val(Right(a, 3))
  If xp > 0 Then
   If Len(a) < xp + 6 Then
        pose = InStr(a, "E") + InStr(a, "D")
        Select Case Len(a) - 4
         Case 1: addx = xp - 1
         Case Else: addx = xp - (Len(a) - 6)
        End Select
        For z = 3 To pose - 1
         Add = Add + Mid(a, z, 1)
        Next
        a = Left(a, 1) + Add + String(addx, "0") + Mid(a, pose)
        Print a
    End If
  Else
   If Left(a, 1) <> "-" Then
    a = "." + String(Abs(xp) - 1, "0") + Left(a, 1) + Mid(a, 3)
   Else
    a = "-0." + String(Abs(xp) - 1, "0") + Mid(a, 2, 1) + Mid(a, 4)
   End If
  End If
  a = Left(a, Len(a) - 4)
End If
STR2 = a
End Function
Public Function expand(s$(), stat)
' Remove math glitches
s(1) = plusrem(s(1))
s(2) = plusrem(s(2))
con = s(1)
ReDim vbl(10), expr(2), Var(50) As String
For pri = 1 To 2
 If s(pri) = "" Then s(pri) = "-999"
 If Left(s(pri), 1) <> "+" And Left(s(pri), 1) <> "-" Then s(pri) = "+" + s(pri)
  X = 1: Do
  nxp = Mid(s(pri), X + 1, 1)
  crp = Mid(s(pri), X, 1)
  If (crp = "+" Or crp = "-") Then
   If nxp >= "?" Then
    s(pri) = Left(s(pri), X) + "1" + Mid(s(pri), X + 1)
    X = 1
   End If
   cp = cp + 1
  ReDim p(cp), num(cp), vr(cp)
  End If
  If crp = "/" And nxp <> "+" And nxp <> "-" Then
   s(pri) = Left(s(pri), X) + "+" + Mid(s(pri), X + 1)
   X = 1
  End If
  If lsp = "0" And crp = "." Then
   s(pri) = Left(s(pri), X - 2) + Mid(s(pri), X)
  End If
  X = X + 1
  lsp = crp
 Loop While X < Len(s(pri))
Next
 ReDim num(cp + cp1)
 ReDim vr(cp + cp1)
For X = 1 To 2
cp = 0: Do: cp = cp + 1
  p1 = InStr(p(cp - 1) + 2, s(X), "+")
  p2 = InStr(p(cp - 1) + 2, s(X), "-")
  If p1 = 0 Then p1 = Len(s(X)) + 1
  If p2 = 0 Then p2 = Len(s(X)) + 1
  If p1 < p2 Then p(cp) = p1
  If p2 < p1 Then p(cp) = p2
 Loop Until p(cp) = 0
 p(cp) = Len(s(X)) + 1: p(0) = 1
 For cx = 1 To cp
 num(cx + cp1) = Val(Mid(s(X), p(cx - 1), p(cx) - p(cx - 1)))
 If num(cx + cp1) >= 0 Then offset = 1 Else offset = 0
 vps = p(cx - 1) + Len(STR2(num(cx + cp1))) + offset
 vr(cx + cp1) = Mid(s(X), vps, Abs(p(cx) - vps))
 p(cx - 1) = 0
Next cx
If cp1 = 0 Then cp1 = cp: p(cx) = 0: p(cx - 1) = 0
Next X
Select Case stat
Case Xmult, Xmult + 10
        ' expand terms:
        cp = cp + cp1
        ReDim op((cp - cp1) * cp1)
        For pri = 1 To cp1
         For sec = cp1 + 1 To cp
         bse = bse + 1
         If alorder(vr(pri)) > alorder(vr(sec)) Then
          op(bse) = vr(pri) + vr(sec)
         Else
          op(bse) = vr(sec) + vr(pri)
         End If
         If op(bse) = "JJ" Then
          op(bse) = STR2(num(pri) * num(sec) * -1)
         Else
          op(bse) = STR2(num(pri) * num(sec)) + op(bse)
         End If
        Next sec, pri
        For pri = 1 To bse
         If Left(op(pri), 1) <> "-" Then op(pri) = "+" + op(pri)
         op(0) = op(0) + op(pri)
        Next
        s(1) = op(0)
        If stat = Xmult + 10 Then
         expand = op(0)
         Else
         expand = expand(s(), Xstrip)
        End If
Case Xcplx:
        cp = cp + cp1
        For i = 1 To cp1
         Select Case vr(i)
          Case "J": TOK = i
          Case ""
          Case Else: fail = 1: Exit For
         End Select
        Next
        For i = cp1 + 1 To cp
         Select Case vr(i)
          Case "J": BOK = i
          Case ""
          Case Else: fail = 1: Exit For
         End Select
        Next
        If TOK + BOK = 0 Or fail = 1 Then GoTo EndXpand
        If BOK Then
         ' prepare conjugate
         For i = cp1 + 1 To cp
          If i <> BOK Then
           If num(i) >= 0 Then conj = conj + "+"
           conj = conj + STR2(num(i))
          Else
           If num(i) <= 0 Then conj = conj + "+" Else conj = conj + "-"
           conj = conj + STR2(Abs(num(i))) + "J"
          End If
          Next
         ' prepare topline - S(1)
         s(1) = "": s(2) = conj
         For i = 1 To cp1
           If num(i) >= 0 Then s(1) = s(1) + "+"
           s(1) = s(1) + STR2(num(i)) + vr(i)
         Next
         s(1) = expand(s(), Xmult)
         tl = expand(s(), Xstrip)
         ' prepare bottomline - s(2)
         s(2) = "": s(1) = conj
         For i = cp1 + 1 To cp
           If num(i) >= 0 Then s(2) = s(2) + "+"
           s(2) = s(2) + STR2(num(i)) + vr(i)
         Next
         s(1) = expand(s(), Xmult)
         bl = expand(s(), Xstrip)
        End If
        For i = cp1 + 1 To cp
         Sum = Sum + num(i)
        Next
        Sum = 1 / Sum
        If tl <> "" Then s(1) = tl
        If bl <> "" Then
                s(2) = STR2(1 / Val(bl))
                Else: s(2) = STR2(Sum)
        End If
        expand = expand(s(), Xmult)
Case Xstrip
        For pri = 1 To cp1 - 1
        For sec = pri + 1 To cp1
        If vr(pri) = vr(sec) Then
                num(pri) = Val(num(pri)) + Val(num(sec))
                num(sec) = 0: vr(sec) = ""
        End If
        Next sec, pri: s(2) = ""
        For pri = 1 To cp1
         If num(pri) > 0 Then s(2) = s(2) + "+"
         If num(pri) <> "0" Then s(2) = s(2) + STR2(num(pri)) + vr(pri)
        Next pri
        expand = s(2)
        If s(2) = "" Then expand = "+0"
Case Xlist:
        For i = 1 To cp1
         Select Case vr(i)
          Case "", "?", "GND"
          Case Else: tmop = tmop + vr(i) + ","
         End Select
        Next
        expand = tmop
Case Xfact:
      GoSub SVS
        lwcm = LCm(num(1), num(2))
        pri = 0: Do Until pri >= vc: pri = pri + 1
        If Right(vbl(pri), 1) = "J" And Len(vbl(pri)) > 1 Then
         vbl(pri) = Left(vbl(pri), Len(vbl(pri)) - 1)
         vc = vc + 1: vbl(vc) = "J"
        End If
        everyterm = 1
        For sec = 1 To cp + cp1
          Temp = num(sec)
      If pri = 1 Then lwcm = LCm(lwcm, Temp)
          If InStr(vr(sec), vbl(pri)) = 0 Then everyterm = 0
         Next
         If everyterm Then
          For sec = 1 To cp + cp1
           vr(sec) = strem(vr(sec), vbl(pri))
          Next
         End If
        Loop
        For sec = 1 To cp + cp1
         If sec <= cp1 Then lne = 1 Else lne = 2
         If num(sec) < 0 Then
          expr(lne) = expr(lne) + "-"
         Else: expr(lne) = expr(lne) + "+"
         End If
         expr(lne) = expr(lne) + STR2(Abs(num(sec))) + vr(sec)
        Next sec
        For sec = 1 To 2
         s(1) = ""
         s(2) = ""
         s(1) = expr(sec)
         s(2) = STR2(1 / lwcm)
         expr(sec) = expand(s(), Xmult)
        Next
        expand = expr(1) + "/" + expr(2)
Case Xphamp:
        For i = 1 To cp1
         If vr(i) = "J" Then b = b + num(i) Else a = a + num(i)
        Next i
        Amp = Sqr(a ^ 2 - b ^ 2)
        phs = Tan(b / a) * 180 / PI
        nl = Chr(13) + Chr(10)
        Select Case vr(cp1 + 1)
         Case "Both"
           expand = nl + "Amplitude=" + pref2.SI_convert(STR2(Amp)) + " Amps"
           expand = nl + "Phase=" + STR2(phs) + " Degrees"
         Case "Phase"
           expand = STR2(phs)
         Case Else
           expand = STR2(Amp)
         End Select
Case Xvar:
        con = s(1)
        GoSub SVS
        expand = "0"
        complex = RemoveIdentity(vbl)
        complex = RemoveIdentity(vr)
        variables = UBound(vr)
        For pri = 1 To variables
         found = False
         For sec = 1 To UBound(vbl)
          If vr(pri) = vbl(sec) Then found = True
         Next
         If Not found Then variables = 3
        Next
        Select Case variables
         Case 0: expand = "4"
         Case 1: If complex Then expand = "2," + vr(1) Else expand = "1," + vr(1)
         Case 2: If complex Then expand = "3," + vr(1) + "," + vr(2)
                 If semic.ListCount Then expand = "0"
        End Select
Case Xqfx: ' SFX
      divpos = cp1 + cp
        For i = cp1 + 1 To cp1 + cp
        vrib = InStr(vr(i), "/")
        If vrib Then divpos = i: vr(i) = Left(vr(i), vrib - 1)
         Do Until InStr(vr(i), pref2.DVBL) = 0
          vr(i) = strem(vr(i), pref2.DVBL)
          If vr(i) <> "" Then
           s(1) = con: s(2) = vr(i)
           vr(i) = expand(s(), Xmult + 10)
          Else
           vr(i) = con ' why?
          End If
         Loop
        Next i
        ' re-form lines
        For i = cp1 + 1 To cp1 + cp
         s(1) = num(i): s(2) = vr(i)
         If vr(i) = "" Then s(2) = "1"
         s(1) = expand(s(), Xmult + 10)
         Sign = Left(s(1), 1)
         If Sign <> "+" And Sign <> "-" Then Sign = "+" Else Sign = ""
         If i <= divpos Then
          tline = tline + Sign + s(1)
         Else: bline = bline + Sign + s(1)
         End If
        Next
        s(1) = bline: bline = expand(s(), Xstrip)
        s(1) = tline: tline = expand(s(), Xstrip)
        If numerical(bline) Then
         s(1) = tline
         Select Case Val(bline)
         Case -999: expand = tline
         Case 0: expand = "Infinity"
         Case Else: s(2) = STR2(1 / Val(bline))
                    s(2) = Mid(s(2), 1, 10)
                     expand = expand(s(), Xmult + 10)
         End Select
         Else
         expand = tline + "/" + bline
        End If
Case Xcram, Xcram + 10:
        Dim X1, Y1, m(6) As String
        GoSub SVS
        For i = 1 To vc
         con = vbl(i)
         If Right(vbl(i), 1) = "J" Then con = Left(con, Len(con) - 1)
         If con <> "" Then
          Select Case X1
           Case "": X1 = con
           Case Else: Y1 = con
          End Select
         End If
        Next
        If stat = Xcram + 10 Then expand = X1 + "," + Y1: GoTo EndXpand
         For i = 1 To cp + cp1
         stat = 0
         If Right(vr(i), 1) = "J" Then
          stat = stat + 1
          vr(i) = Left(vr(i), Len(vr(i)) - 1)
         End If
         If InStr(vr(i), X1) Then
             stat = stat + 2
             vr(i) = strem(vr(i), X1)
            ElseIf InStr(vr(i), Y1) Then vr(i) = strem(vr(i), Y1)
            Else: stat = stat + 4
         End If
        If i > cp1 Xor num(i) >= 0 Xor stat >= 4 Then
         m(stat) = m(stat) + "+"
         Else: m(stat) = m(stat) + "-"
        End If
        m(stat) = m(stat) + STR2(Abs(num(i))) + vr(i)
       Next
       For i = 0 To 5
        If m(i) = "" Then m(i) = "+0"
       Next
       ' x(RxIy+IxRy)-RtIy+ItRy=0
       ' y(RxIy+IxRy)-RxIt+IxRt=0
       s(1) = m(2): s(2) = m(1)
       expr(1) = expand(s(), Xmult + 10)
       s(1) = m(3): s(2) = m(0)
       expr(1) = expr(1) + expand(s(), Xmult + 10)
       expr(2) = expr(1)
       s(1) = expr(1): s(2) = X1
       expr(1) = expand(s(), Xmult + 10)
       s(1) = m(4): s(2) = m(1)
       expr(1) = expr(1) + invert(expand(s(), Xmult + 10))
       s(1) = m(5): s(2) = m(0)
       expr(1) = expr(1) + expand(s(), Xmult + 10)
       s(1) = expr(2): s(2) = Y1
       expr(2) = expand(s(), Xmult + 10)
       s(1) = m(2): s(2) = m(5)
       expr(2) = expr(2) + invert(expand(s(), Xmult + 10))
       s(1) = m(3): s(2) = m(4)
       expr(2) = expr(2) + expand(s(), Xmult)
       For i = 1 To 2
         s(1) = expr(i): s(2) = expr(i)
         expr(i) = expand(s(), Xfact)
         expr(i) = Left(expr(i), InStr(expr(i), "/") - 1)
         s(1) = expr(i)
         expr(i) = expand(s(), Xstrip)
          s(1) = expr(i): s(2) = expr(i)
         expr(i) = expand(s(), Xfact)
         expr(i) = Left(expr(i), InStr(expr(i), "/") - 1)
        Next
       expand = expr(1) + "=0," + expr(2) + "=0"
       
End Select
GoTo EndXpand
SVS:
s(2) = "": expr(0) = ""
For i = 1 To Volt.ListCount
 s(1) = Volt.list(i - 1)
 expr(0) = expr(0) + expand(s(), Xlist)
 s(1) = depvolt.list(i - 1)
 expr(0) = expr(0) + expand(s(), Xlist)
Next
For i = 1 To Resist.ListCount
 s(1) = Resist.list(i - 1)
 expr(0) = expr(0) + expand(s(), Xlist)
 s(1) = depres.list(i - 1)
 expr(0) = expr(0) + expand(s(), Xlist)
Next: vc = 1
For i = 1 To Len(expr(0)) - 1
 If Mid(expr(0), i, 1) <> "," Then
  vbl(vc) = vbl(vc) + Mid(expr(0), i, 1)
 Else
  vc = vc + 1: vbl(vc) = ""
 End If
Next
Return
EndXpand:
End Function

Private Function cramner(n) As String
Dim cnt, nodes, i, X, Y As Byte
Dim s(2) As String
nl = Chr(13) + Chr(10)
cnt = Val(Number.Caption)
nodes = Val(NODE.Caption)
ReDim loopz(cnt), eq(cnt) As String
For i = 1 To cnt
 loopz(i) = LOOPS.list(i - 1)
Next
ReDim V(nodes) As String
For i = 1 To nodes
 V(i) = Volt.list(i)
Next i
ReDim mx(cnt, cnt)
For X = 1 To cnt
 For Y = 1 To cnt
  mx(X, Y) = strip(Union(loopz(X), loopz(Y)))
  If X <> Y Then mx(X, Y) = invert(mx(X, Y))
Next Y, X: ' Bottom line
bl = remove0(det(cnt, mx()))
If mesh11.MathDoc Then
 ' Maths Documentation
 For Y = 1 To cnt
  For X = 1 To cnt
   s(1) = mx(X, Y)
   s(2) = "i"
   If X > 1 Then s(2) = s(2) + STR2(X)
   eq(Y) = eq(Y) + format(expand(s(), Xmult))
  Next
  MDI.HTML = MDI.HTML + remove1(eq(Y)) + "=" + remove1(volts(loopz(Y))) + nl + "<BR>"
 Next
End If
If InStr(bl, "(") Then GoTo SKIPSIMP1
For i = 1 To Len(bl) - 1
  C = Mid(bl, i, 1)
  C2 = Mid(bl, i + 1, 1)
  If (C = "+" Or C = "-") And Asc(C2) > 63 Then GoTo SKIPSIMP1
 Next
SKIPSIMP1:
For i = 1 To cnt:
vS = volts(loopz(i))
If vS = "" Then
vS = "0"
ElseIf Asc(Left(vS, 1)) > 57 Then vS = "+1" + vS
End If
mx(n, i) = strip(vS)
Next:
tl = remove0(det(cnt, mx()))
If InStr(tl, "(") Then GoTo SKIPSIMP2
    For i = 1 To Len(tl) - 1
    C = Mid(tl, i, 1)
    C2 = Mid(tl, i + 1, 1)
    If (C = "+" Or C = "-") And Asc(C2) > 63 Then GoTo SKIPSIMP2
    Next
    tl = strip(tl)
SKIPSIMP2:

If Not numerical(tl) Or Not numerical(bl) Then
     s(1) = tl: s(2) = bl
    op = expand(s(), Xfact)
    tl = Left(op, InStr(op, "/") - 1)
    bl = Mid(op, InStr(op, "/") + 1)
    
End If
foundit:
s(1) = tl: tl = expand(s(), Xstrip)
s(1) = bl: bl = expand(s(), Xstrip)
cramner = remove1(tl) + "/" + remove1(bl)
End Function
Private Function det(m, ma())
Dim s(2) As String
If m = 1 Then
 det = ma(1, 1)
Else
 ReDim mx(m, m)
 DoEvents
 For matrix = 1 To m
  RealY = 1
  For i = 2 To m
   For j = 1 To m
    If matrix <> j Then
     mx(i - 1, RealY) = ma(i, j)
     RealY = RealY + 1
    End If
  Next j, i
  Sign = ma(1, matrix)
  If matrix Mod 2 Then Sign = invert(Sign)
  s(1) = Sign
  s(2) = det(m - 1, mx)
  ans = ans + expand(s(), Xmult)
 Next matrix
 det = ans
End If
End Function
Public Function invert(a) As String
Dim i As Byte
If Left(a, 1) <> "+" And Left(a, 1) <> "-" Then a = "+" + a
For i = 2 To Len(a)
la = Mid(a, i - 1, 1)
If (la = "+" Or la = "-") And Mid(a, i, 1) >= "A" Then a = Left(a, i - 1) + "1" + Mid(a, i)
Next
For i = 1 To Len(a)
If Mid(a, i, 1) = "(" And InStr(i, a, "(") > i Then i = InStr(i, a, "(")
If Mid(a, i, 1) = "+" Then op = op + "-": GoTo skip
If Mid(a, i, 1) = "-" Then op = op + "+": GoTo skip
op = op + Mid(a, i, 1)
skip: Next
invert = op
End Function

Private Function remove0(a) As String
Dim lp, i, K As Byte
For lp = 1 To 5
For i = 2 To Len(a)
p = Mid(a, i - 1, 1)
If Mid(a, i, 1) = "0" And (p = "+" Or p = "-" Or p = ")") Then
K = i: Do: K = K + 1:
p = Mid(a, K, 1)
If p = "(" Then K = InStr(K, a, ")")
Loop Until p = "+" Or p = "-" Or K >= Len(a)
If K = Len(a) Then
a = Left(a, i - 2)
Else
a = Left(a, i - 2) + Mid(a, K)
End If
End If: Next
endremove0:
Next lp
remove0 = a
End Function
Public Function remove1(a) As String
Dim i As Byte
If Len(a) < 2 Then GoTo endremove1
If Len(a) > 2 Then
If Left(a, 1) = "1" And Asc(Mid(a, 2, 1)) > 63 Then a = Mid(a, 2)
If Left(a, 1) = "+" Then a = Mid(a, 2)
Else
If Left(a, 1) = "1" And Asc(Mid(a, 2, 1)) > 63 Then remove1 = Mid(a, 2): GoTo very_end
End If
For i = 2 To Len(a) - 1
p2 = Mid(a, i + 1, 1)
p = Mid(a, i - 1, 1)
If Mid(a, i, 1) = "+" And p = "/" Then GoTo SKIPREM
If Mid(a, i, 1) = "1" Then
If p = "(" And Asc(p2) > 63 Then GoTo SKIPREM:
If (p = "+" Or p = "-") And (p2 = "(" Or Asc(p2) > 63) Then GoTo SKIPREM:
If p = ")" And Str(Val(p2)) <> p Then GoTo SKIPREM
End If
op = op + Mid(a, i, 1)
SKIPREM:
Next
endremove1:
If Len(a) > 2 Then
remove1 = Left(a, 1) + op + Right(a, 1)
Else: remove1 = a
End If
very_end:
End Function
Private Function rlt(a) As String
For i = Len(a) - 1 To 1 Step -1
 If InStr(i + 1, a, "+") Or InStr(i + 1, a, "-") Then
  op = Mid(a, i, 1) + op
 End If
Next: rlt = op
End Function
Public Function rev(a) As String
 ' I didn't use strReverse, for
 ' vb5 compatibility
 strRev = ""
 For i = Len(a) To 1 Step -1
  strRev = strRev + Mid(a, i, 1)
 Next
 rev = strRev
End Function
Private Function strip(a) As String
Dim s(2) As String
s(1) = a
strip = expand(s(), Xstrip)
End Function
Private Function Union(m1, m2) As String
Dim i, pri, sec, branches As Byte
branches = mesh11.LINKS.ListCount
ReDim link(branches), z(branches)
For i = 1 To branches
link(i) = LINKS.list(i - 1)
z(i) = Resist.list(i - 1)
Next
m1 = m1 + Left(m1, 1)
m2 = m2 + Left(m2, 1)
For pri = 1 To Len(m1) - 1
For sec = 1 To Len(m2) - 1
p1 = Mid(m1, pri, 2)
If p1 = Mid(m2, sec, 2) Or rev(p1) = Mid(m2, sec, 2) Then
For i = 1 To branches
If p1 = link(i) Or rev(p1) = link(i) Then
op = op + "+" + z(i)
End If: Next i
End If
Next sec, pri
m1 = Left(m1, Len(m1) - 1)
m2 = Left(m2, Len(m2) - 1)
Union = op
End Function
Private Function volts(ByVal m) As String
Dim i, Ref As Byte
m = rev(m)
z = Val(NODE.Caption)
ReDim V(z)
For i = 1 To z
  V(i) = Volt.list(i - 1)
Next: Volt = "+0"

For i = 1 To Len(m)
 If V(Asc(Mid(m, i, 1)) - 64) = "GND" Then gnd = 1
Next
If gnd = 0 Then GoTo endvolt

For i = 1 To Len(m)
  nv = V(Asc(Mid(m, i, 1)) - 64)
  If nv <> "GND" And nv <> "?" And nv <> "+0" Then
    If Ref = 0 Then
     op = nv: Ref = Asc(Mid(m, i, 1)) - 64
    Else
     inv = invert(nv)
     op = op + inv
   End If
  End If
Next
For i = 1 To Len(m)
 If LeftN(Asc(Mid(m, i, 1)) - 64) < LeftN(Ref) Then op = invert(op): Exit For
Next i
If op = "" Then op = "0"
volts = op
endvolt:
End Function
Private Function branch_current(bds) As String
Dim cnt, pass, bubble, templ, temph, j As Byte
Dim s(2) As String
cnt = mesh11.LOOPS.ListCount
ReDim wght(cnt) As Double
ReDim Order(cnt) As Byte
For j = 0 To cnt - 1
 cwghs = volts(LOOPS.list(j))
 Order(j) = j
 wght(j) = Val(cwghs)
 If wght(j) = 0 Then
   For i = 1 To Len(cwghs)
    wght(j) = wght(j) + Asc(Mid(cwghs, i, 1))
   Next
 End If
Next ' bubble sort (biggest first)
For pass = 1 To cnt
  For bubble = 1 To cnt - 1
   If wght(bubble) > wght(bubble - 1) Then
    templ = wght(bubble - 1)
    temph = wght(bubble)
    wght(bubble) = templ
    wght(bubble - 1) = temph
    templ = Order(bubble - 1)
    temph = Order(bubble)
    Order(bubble - 1) = temph
    Order(bubble) = templ
   End If
Next bubble, pass
For j = 0 To cnt - 1
 loopy = mesh11.LOOPS.list(Order(j)) + Left(mesh11.LOOPS.list(Order(j)), 1)
 If InStr(loopy, bds) Or InStr(loopy, rev(bds)) Then
 ' add to total current (curr)
   If curr = "" Then
     curr = cramner(Order(j) + 1)
     If MathDoc Then MathDoc = False
   Else
     C = cramner(Order(j) + 1)
     tlc = Left(C, InStr(C, "/") - 1)
     blc = Mid(C, Len(tlc) + 2)
     If tlc = "0" Then curr = "Infinity": Exit For
     If blc = "0" Then curr = "0": Exit For
     tlb = Left(curr, InStr(curr, "/") - 1)
     blb = Mid(curr, Len(tlb) + 2)
     'curr = branch_current(bds)
     s(1) = tlc + invert(tlb)
     s(1) = expand(s(), Xstrip)
     s(2) = blb
     curr = expand(s(), Xfact)
     MathDoc = True
   End If
 End If
Next
branch_current = curr
End Function
Public Function Analyse()
Dim volttot, restot, oloopT, oresT, odepT, olinkT As Byte
Dim it, rlp, j, link, found, f As Byte
Dim alpha, clockwise As Boolean
' put Numeric Semiconductor support here.
' this code is repeated in pref2hy.picture2.case6
' it detects alpha chars in either
' dep or indep lists
volttot = mesh11.Volt.ListCount - 1
restot = mesh11.Resist.ListCount - 1
ReDim dep(volttot + restot + 1): ' Alpha = False
ReDim dpp(volttot + restot + 1) As Byte
ReDim pis(volttot + restot + 1): ' pre iteration states
ReDim oloop(LOOPS.ListCount): oloopT = LOOPS.ListCount
ReDim ores(Resist.ListCount): oresT = Resist.ListCount
ReDim odep(depres.ListCount): odepT = depres.ListCount
ReDim olink(LINKS.ListCount): olinkT = LINKS.ListCount
For i = 0 To restot
  pis(i) = Resist.list(i)
  If Not numerical(Resist.list(i)) Then alpha = True
  If depres.list(i) <> "" Then
    dpc = dpc + 1
    dep(dpc) = depres.list(i)
    dpp(dpc) = i
    If Not numerical(Left(dep(dpc), Len(dep(dpc)) - 3)) Then alpha = True
  End If
Next
For i = 0 To volttot
  pis(i + 1) = Resist.list(i)
  If Not numerical(mesh11.Volt.list(i)) Then alpha = True
  If mesh11.depvolt.list(i) <> "" Then
    dpc = dpc + 1
    dep(dpc) = depvolt.list(i)
    dpp(dpc) = i + restot + 1
    If Not numerical(Left(dep(dpc), Len(dep(dpc)) - 3)) Then alpha = True
  End If
Next
If Not alpha Then
 ' iterative ASCS here
 mesh11.MathDoc = False
 For it = 1 To 5
   For i = dpc To 1 Step -1
     cbran = Right(dep(i), 2)
     PrT = branch_current(cbran)
     tline = Left(PrT, InStr(PrT, "/") - 1)
     bline = Mid(PrT, InStr(PrT, "/") + 1)
     If Left(Right(dep(i), 3), 1) = "/" Then swap = tline: tline = bline: bline = swap
     vll = Val(tline) * Val(Left(dep(dpc), Len(dep(dpc)) - 3))
     If Val(bline) = 0 Then
      vll = vll / 0.001
      Else: vll = vll / Val(bline)
     End If
     vls = STR2(vll)
     If Left(vls, 1) <> "-" Then vls = "+" + vls
     If dpp(i) > oresT Then
       If it > 1 Then Volt.list(dpp(i) - restot - 1) = rlt(Volt.list(dpp(i) - restot - 1))
       Volt.list(dpp(i) - restot - 1) = Volt.list(dpp(i) - restot - 1) + vls
     Else
       If it > 1 Then Resist.list(i) = rlt(Resist.list(i))
       Resist.list(i) = Resist.list(i) + vls
     End If
     ' store loops
     For rlp = 0 To LOOPS.ListCount - 1
     oloop(rlp) = LOOPS.list(rlp)
     Next
     ' store resist
     For rlp = 0 To Resist.ListCount - 1
     ores(rlp) = Resist.list(rlp)
     Next
     ' store depres
     For rlp = 0 To depres.ListCount - 1
     odep(rlp) = depres.list(rlp)
     Next
     ' store links
     For rlp = 0 To LINKS.ListCount - 1
     olink(rlp) = LINKS.list(rlp)
     Next
     ' reasoning: put all the list data away in arrays
     '            before they are hacked to pieces by
     '            the semic loops.
     '   (1)      if a reverse pn jn is found,
     '            the list data is altered.
     '   (2)      if no paths remain the data is restored
     '            if a path does remain...
     '            restore resistances & voltages only <<?
     ' pn jns
     For j = 0 To semic.ListCount - 1
       
       n1 = Left(semic.list(j), 1): Max = 1
       n2 = Right(semic.list(j), 1)
       For z = 1 To LOOPS.ListCount
        If InStr(LOOPS.list(z - 1), n1) And InStr(LOOPS.list(z - 1), n2) And calc(z) > calc(Max) Then Max = z
       Next
       For z = 1 To Len(LOOPS.list(Max - 1))
        Temp = Mid(LOOPS.list(Max - 1), z, 1)
        If Temp <> n1 And Temp <> n2 Then n3 = Temp: Exit For
       Next
       op = "": dr = 1
       For zy = 0 To 9: oop = op
        For zx = (dr - 1) * -9.5 To (dr + 1) * 9.5 Step dr
         If zx = LeftN(Asc(n1) - 64) And zy = upperN(Asc(n1) - 64) Then op = op + n1
         If zx = LeftN(Asc(n2) - 64) And zy = upperN(Asc(n2) - 64) Then op = op + n2
         If zx = LeftN(Asc(n3) - 64) And zy = upperN(Asc(n3) - 64) Then op = op + n3
        Next zx
       If oop <> op Then dr = dr * -1
       Next zy
       ' rectify op
        op = Mid(op, InStr(op, n1)) + Left(op, InStr(op, n1) - 1)
       PrT = branch_current(semic.list(j))
       bline = Mid(PrT, InStr(PrT, "/") + 1)
       If Val(PrT) < 0 Xor Val(bline) < 0 Xor op <> n1 + n2 + n3 Then
         For lnk = 0 To LINKS.ListCount - 1
           If LINKS.list(lnk) = semic.list(j) Or LINKS.list(lnk) = rev(semic.list(j)) Then
            LINKS.RemoveItem lnk: Exit For
           End If
         Next
         link = semic.list(j)
         Resist.RemoveItem dpp(j)
         depres.RemoveItem dpp(j)
         found = 0
         For f = 0 To LOOPS.ListCount - 1
           Temp = LOOPS.list(f)
           Temp = Temp + Left(Temp, 1)
           If InStr(Temp, link) Or InStr(Temp, rev(link)) Then LOOPS.RemoveItem f: Temp = ""
           If InStr(Temp, Right(prcall, 2)) Or InStr(rev(Temp), Right(prcall, 2)) Then found = 1
         Next
         If found = 0 Then
          Analyse = "0/1"
          ' restore loops
          LOOPS.Clear
          For rlp = 0 To oloopT - 1
           LOOPS.AddItem oloop(rlp)
          Next
          ' restore resist
          Resist.Clear
          For rlp = 0 To oresT - 1
          Resist.AddItem ores(rlp)
          Next
          ' restore depres
          depres.Clear
          For rlp = 0 To odepT - 1
           depres.AddItem odep(rlp)
          Next
          ' restore links
          LINKS.Clear
          For rlp = 0 To olinkT - 1
           LINKS.AddItem olink(rlp)
          Next
          GoTo leave
         End If
         ' restore all values to pre-iterative-states
         ' only if a diode is o/c'd !!!
         For f = 0 To restot
           mesh11.Resist.list(f) = pis(f)
         Next
         For f = restot + 1 To restot + voltot + 1
           mesh11.Volt.list(f) = pis(f)
         Next
       End If
  Next j, i, it
End If
cbran = Right(prcall, 2)
Analyse = branch_current(cbran)
leave:
End Function
Public Function numerical(a) As Boolean
Dim i, ask As Byte
numerical = True
If a = "?" Or a = "GND" Then GoTo endnumer
For i = 1 To Len(a)
ask = Asc(Mid(a, i, 1))
If ask > 63 Or ask < 43 Then numerical = False: GoTo endnumer
Next
endnumer:
End Function
Private Function LCm(ByVal a, ByVal b)
a = Abs(Int(a)): b = Abs(Int(b))
If a < b Then tp = b: b = a: a = tp
' xmod addition a>b
aMb = a
If b > 1 Then
 Do Until aMb - b <= b
  aMb = aMb - b
 Loop
 aMb = Abs(aMb - b)
End If
If a = b Then aMb = b
If b <= 1 Then
 LCm = 1
ElseIf aMb = b Then
 LCm = b
Else
 LCm = LCm(b, aMb)
End If
End Function

Private Function calc(n) As Double
C = cramner(n)
tline = Val(Left(C, InStr(C, "/") - 1))
bline = Val(Mid(C, InStr(C, "/") + 1))
If bline = 0 Then bline = 0.001
calc = tline / bline
End Function
Function alorder(a) As Integer
Dim i As Byte
Dim op As Integer

For i = 1 To Len(a)
op = op + Asc(Mid(a, i, 1))
Next
alorder = op
End Function
Public Function strem(st, re) As String
pose = InStr(st, re)
op = ""
If pose = 0 Then op = st: st = ""
For i = 1 To Len(st)
 If i < pose Or i > pose + Len(re) - 1 Then op = op + Mid(st, i, 1)
Next
strem = op
End Function
Function plusrem(a)
lp = Left(a, 1)
For ln = Len(a) To 1 Step -1
If Asc(Mid(a, ln)) >= Asc("0") Then Exit For
Next
For i = 2 To ln + 1
 cp = Mid(a, i, 1)
 If (cp = "+" Or cp = "-") And (lp = "+" Or lp = "-") Then
  If cp = lp Then cp = "+" Else cp = "-"
  lp = ""
 End If
 'If lp = "/" And cp <> "+" And cp <> "-" Then lp = "/+"
 op = op + lp
 lp = cp
Next
plusrem = op
End Function
Public Function RemoveIdentity(ByRef list)
pri = 1
Do While pri <= UBound(list)
 po1 = InStr(list(pri), "[")
 po2 = InStr(list(pri), "]")
 If po1 * po2 Then list(pri) = Left(list(pri), po1 - 1) + Mid(list(pri), po2 + 1)
 If Right(list(i), 1) = "J" Then
  list(i) = Left(list(i), Len(list(i)) - 1)
  RemoveIdentity = True
 End If
 If list(pri) = "" Then
  Item = pri
  GoSub RemoveItem
  pri = 0
 End If
 pri = pri + 1
Loop
pri = 1: sec = 1
Do While pri <= UBound(list)
 While sec <= UBound(list)
  If pri <> sec And InStr(list(sec), list(pri)) Then
   Item = sec
   GoSub RemoveItem
   pri = 1: sec = 1
  End If
  sec = sec + 1
 Wend
pri = pri + 1
sec = 1
Loop
Exit Function
RemoveItem:
ul = UBound(list)
For z = Item + 1 To ul
 list(z - 1) = list(z)
Next
ReDim Preserve list(ul - 1)
Return
End Function
