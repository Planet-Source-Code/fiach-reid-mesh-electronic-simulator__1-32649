VERSION 5.00
Begin VB.Form pref2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Graph"
   ClientHeight    =   4665
   ClientLeft      =   3660
   ClientTop       =   2925
   ClientWidth     =   6030
   Icon            =   "pref2.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   6030
   Begin VB.TextBox Text1 
      Height          =   252
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "pref2.frx":27A2
      Top             =   120
      Visible         =   0   'False
      Width           =   7605
   End
End
Attribute VB_Name = "pref2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FxVer, DVBL, ASCSN, dy, Ref, freq, Ctyp, Ccbran, CUplim, Clowlim, Coldval
Public Function AnalyseRequest(location)
 Dim buffer As String
 File = App.Path + "\Mesh21.htm"
 Ref = location
 
 Picture2_MouseUp 6, 1, 0, 0, 0
 buffer = Space(FileLen(File))
 Open File For Binary As #1
  Get #1, , buffer
 Close #1
 AnalyseRequest = buffer
End Function
Public Sub DrawGraph(X, Y, w, h, obj)
' Copyright 1999 F. Reid
' obj=1 =>printer
' obj=0 =>Stdout
frmProgress.Show
' ********** Graph preparation *********
     ReDim pnt(101)
     Min = 9999999: Max = Min * -1
     llim = Val(CUplim)
     Ulim = Val(Clowlim)
     If llim > Ulim Then swap = Ulim: Ulim = llim: llim = swap
     stp = Abs(Ulim - llim) / 100
     frmProgress.Progress.Visible = True
     For i = 0 To 100
      pnt(i) = fx(Round(llim + i * stp, 10))
      'If Picture2(31).Visible Then
      ' pnt(i) = (fx(llim + i * stp + 0.00001) - pnt(i)) / 0.00001
      'End If
      If pnt(i) > Max Then Max = pnt(i)
      If pnt(i) < Min Then Min = pnt(i)
      If i Mod 5 = 0 Then
       frmProgress.Progress.Value = i: DoEvents
      End If
      Next i
' **************************************
Select Case True
 Case Max <> Min: scy = -h / (Max - Min)
 Case Max = Min And Max <> 0: scy = -h * 10 / Max
 Case Else: scy = -h
End Select
If obj Then
 Printer.PSet (X, (pnt(0) - Min) * scy + Y)
 Else
 PSet (X, (pnt(0) - Min) * scy + Y)
End If
frmProgress.Progress.Value = 0
DoEvents
frmProgress.Progress.Visible = False
If obj Then
 Printer.DrawWidth = 3
 Else
 Me.DrawWidth = 3
End If
For i = 1 To 100
 cx = X + i * w / 100
 cy = (pnt(i) - Min) * scy + Y
 If obj Then
  Printer.PSet (cx, cy + h)
  Else
  Me.PSet (cx, cy + h), RGB(255, 0, 0)
 End If
Next i
If llim <= 0 And Ulim >= 0 Then
 'Vertical axis
 range = Abs(Max - Min)
 For ex = -12 To 9 'femto to Giga
  If 5 * 10 ^ ex > range Then Exit For
 Next: ex = ex - 1
 xpost = llim / (Ulim - llim) * w * -1 + X
 DrawWidth = 1
 If obj Then
  Printer.Line (xpost, Y)-Step(0, h)
  Else
   Me.Line (xpost, Y)-Step(0, h)
 End If
cx = Min:
Do Until cx >= Max
 cx = cx + 10 ^ ex
 valu = Int(cx / 10 ^ ex) * 10 ^ ex
 tick = (valu - Min) * scy + Y
 If valu = 0 Then
 ' QuickDraw Horixontal Axis
  ypost = tick
   range = Abs(Ulim - llim)
   For ex2 = -12 To 9 'femto to Giga
    If 5 * 10 ^ ex2 > range Then Exit For
   Next
   ex2 = ex2 - 1
   If obj Then
    Printer.Line (X, ypost + h)-Step(w, 0)
   Else
    Me.Line (X, ypost + h)-Step(w, 0)
   End If
   hcx = llim: DrawWidth = 1
   Do Until hcx >= Ulim
    hcx = hcx + 10 ^ ex2
    hvalu = Int(hcx / 10 ^ ex2) * 10 ^ ex2
    tick = (hvalu - llim) * w / (Ulim - llim)
    If obj Then
     Printer.Line (tick + X, ypost + (h - 50))-Step(0, 100)
     Printer.Print convert_SI(hvalu)
    Else
     Me.Line (tick + X, ypost + (h - 50))-Step(0, 100)
    End If
   Loop
  tick = ypost
 End If
  If obj Then
  Printer.Line (xpost - 50, tick + h)-Step(100, 0)
  If (cx - 10 ^ ex) <> Min Then Printer.Print convert_SI(valu)
 Else
  Me.Line (xpost - 50, tick + h)-Step(100, 0)
 End If
Loop
End If
Me.Width = w
Me.Height = h
On Error Resume Next
Kill App.Path + "\graph.bmp"
On Error GoTo 0
'SavePicture Me.Image, App.Path + "\Graph.bmp"
' DEBUG ONLY!!!!
Me.Show

'Me.Cls
ic = Chr(34)
nl = Chr(13) + Chr(10)
HTML = "<center>"
HTML = HTML + "<IMG src=" + ic + App.Path + "\Graph.bmp" + ic + ">" + nl
HTML = HTML + "</center><BR><a>" + convert_SI(10 ^ ex2) + "a<BR>" + nl
HTML = HTML + convert_SI(10 ^ ex) + DVBL + "<BR></a>" + nl
'If fx(llim) > 0 Xor fx(Ulim) > 0 Then
' T2(1).Text = LTrim(CUplim)
'End If
If InStr(tempfn, ",") Then tempfn = Left(tempfn, InStr(tempfn, ",") - 1)
frmProgress.Hide
End Sub


Private Sub Form_Activate()
'HTML.Offline = 0
frmProgress.Progress.Value = 0
frmProgress.Progress.Visible = False
'Me.Height = Backdrop.Height
End Sub
Private Sub Form_Load()
 'HTML.Offline = 0
 'RichOutput
End Sub

Private Sub Form_Unload(Cancel As Integer)
canvas(DD.Cform).changed = True
MDI.Refresh
End Sub
Public Sub RichOutput()
Open App.Path + "\Mesh21.htm" For Output As #1
Print #1, Text1.Text
Close
HTML.Navigate App.Path + "\Mesh21.htm"
' cropping?
End Sub

Private Sub Picture2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim s(2) As String
'Picture2(Index).BorderStyle = 0
Select Case Index
Case 0 'delete
    'If Picture2(29).Visible Then
     'Line (240, 2760)-Step(5532, 1692), RGB(255, 255, 255), BF
     'Line (240, 2760 + 846)-Step(5532, 0), RGB(0, 0, 0)
     'Line (240 + 2766, 2760)-Step(0, 1692), RGB(0, 0, 0)
    'Else
     canvas(DD.Cform).changed = True
     canvas(DD.Cform).datalist.list(Val(Ref)) = "-1"
     canvas(DD.Cform).ReNew
     MDI.Refresh
     pref2.Hide
    'End If
Case 1 'replace
    canvas(DD.Cform).datalist.list(Val(Ref)) = mesh11.STR2(Toolbox.finalsel)
    'Picture2(9).Visible = True ' pressing in
    canvas(DD.Cform).ReNew
    'upload default values
    Select Case Toolbox.finalsel
    Case Is < 10: op = "0"
    Case 10, 11: op = "R"
    Case 12, 13: op = "-CJ"
    Case 14, 15: op = "LJ"
    Case 16, 17: op = "V,0"
    Case 18 To 21: op = ".7,+0.025/I"
    Case 22: op = "50K,0+0.025/I,Hfe"
    End Select
    canvas(DD.Cform).values.list(Val(Ref)) = op
    Coldval = canvas(DD.Cform).values.list(Val(Ref))
    Ctyp = canvas(DD.Cform).datalist.list(Val(Ref))
    ReNew
Case 5 ' cancel
     'Line (240, 2760)-Step(5532, 1692), RGB(255, 255, 255), BF
     'T2(0).Text = ""
     'T2(1).Text = ""
     'T2(2).Text = ""
     canvas(DD.Cform).values.list(Val(Ref)) = Coldval
     pref2.Hide
Case 6 ' pressing Current
    Text1.Text = "<body bgcolor=#FFFFFF>"
    RichOutput
    With pref2
     If .Height = Backdrop.Height Then .Height = .Height * 2
    End With
    mesh11.MathDoc = True
    'Picture2(6).Visible = False
    'Picture2(20).Visible = True ' current stop
    With mesh11
        .LINKS.Clear
        .Resist.Clear
        .Volt.Clear
        .depres.Clear
        .depvolt.Clear
    End With
    linkup True
    'If Picture2(6).Visible Then GoTo leave
    If mesh11.LINKS.ListCount = 1 Then Text1.Text = "<body bgcolor=#EEEEEE>Current=Infintity": GoTo leave
    ' needs single node current detection
    volttot = mesh11.Volt.ListCount - 1
    restot = mesh11.Resist.ListCount - 1
    ReDim dep(volttot + restot + 1): ' Alpha = False
    ReDim dpp(volttot + restot + 1): ' dpc = 0 ' origion=1
    ReDim pis(volttot + restot + 1): ' pre iteration states
    For i = 0 To restot
    pis(i) = mesh11.Resist.list(i)
    If Not mesh11.numerical(mesh11.Resist.list(i)) Then alpha = True
    If mesh11.depres.list(i) <> "" Then
        dpc = dpc + 1: dep(dpc) = mesh11.depres.list(i): dpp(dpc) = i
        If Not mesh11.numerical(Left(dep(dpc), Len(dep(dpc)) - 3)) Then alpha = True
    End If
    Next
    For i = 0 To volttot
    pis(i + 1) = mesh11.Resist.list(i)
    If Not mesh11.numerical(mesh11.Volt.list(i)) Then alpha = True
    If mesh11.depvolt.list(i) <> "" Then
        dpc = dpc + 1: dep(dpc) = mesh11.depvolt.list(i): dpp(dpc) = i + restot + 1
        If Not mesh11.numerical(Left(dep(dpc), Len(dep(dpc)) - 3)) Then alpha = True
    End If
    Next
    If alpha Then ' alpha ascs
        Text1.Text = "<body bgcolor=#EEEEEE>"
        For i = dpc To 1 Step -1
          If mesh11.prcall = dep(i) Then
            mesh11.prcall = mesh11.rev(dep(i))
          Else: mesh11.prcall = dep(i)
          End If
          prcall = mesh11.Analyse()
          tline = Left(prcall, InStr(prcall, "/") - 1)
          bline = Mid(prcall, InStr(prcall, "/") + 1)
          If Left(Right(dep(i), 3), 1) = "/" Then swap = tline: tline = bline: bline = swap
          s(1) = Left(dep(i), Len(dep(i)) - 3)
          s(2) = tline
          tline = mesh11.expand(s(), Xmult)
          Text1.Text = Text1.Text + "[X" + LTrim(Str(i)) + "]=" + tline + "/" + bline + "<BR>"
          ascs(i) = tline + "/" + bline
          If Left(Right(dep(i), 3), 1) = "/" Then ascs(i) = "~" + ascs(i)
          If dpp(i) > restot Then
            mesh11.Volt.list(dpp(i) - restot - 1) = mesh11.Volt.list(dpp(i) - restot) + "+1[X" + LTrim(Str(i)) + "]"
          Else: mesh11.Resist.list(dpp(i)) = mesh11.Resist.list(dpp(i)) + "+1[X" + LTrim(Str(i)) + "]"
          End If
        Next
     End If
     If mesh11.prcall = Ccbran Then
        mesh11.prcall = mesh11.rev(Ccbran)
     Else: mesh11.prcall = Ccbran
     End If
     With mesh11
      prcall = .remove1(.Analyse)
     End With
     If alpha Then
        nl = Chr(13) + Chr(10)
        pofs = InStr(prcall, "/")
        tline = Left(prcall, pofs - 1)
        bline = Mid(prcall, pofs + 1)
        Text1.Text = Text1.Text + "Current=<center><u>" + tline + "</u><BR>"
        Text1.Text = Text1.Text + bline + "</center>"
        ascs(dpc + 1) = prcall: ASCSN = dpc + 1
         linkup False
         s(2) = Left(prcall, InStr(prcall, "/") - 1) _
             + "+" + Mid(prcall, InStr(prcall, "/") + 1)
         FxVer = mesh11.expand(s(), Xvar)
         DVBL = Mid(FxVer, 3)
         Select Case Left(FxVer, 1)
         Case "1": ' simple eq
            'Picture2(25).Visible = True
         Case "2": ' Complex eq
            s(1) = Left(prcall, InStr(prcall, "\") - 1)
            s(2) = Mid(prcall, InStr(prcall, "\") + 1)
            s(1) = mesh11.expand(s(), Xcplx)
            s(2) = "Both"
            Text1.Text = Text1.Text + "<BR>" + mesh11.expand(s(), Xphamp)
            'Picture2(32).Visible = True
            'Picture2(25).Visible = True
         Case "3": ' simultanious eq
            value1.Caption = "Upper Limit"
            value2.Caption = "Lower Limit"
            value3.Caption = ""
            Ctyp = "24"
            ReNew
            'Picture2(29).Visible = True
            'Picture2(27).Visible = True
         Case Else: ' unsolveable
         End Select
     Else
        tline = Val(Left(prcall, InStr(prcall, "/") - 1))
        bline = Val(Mid(prcall, InStr(prcall, "/") + 1))
        If bline = 0 Then bline = 0.0001
        Text1.Text = "Current=" + convert_SI(tline / bline) + " a<BR>"
        Value = canvas(DD.Cform).values.list(Val(Ref))
        Power = Val(Mid(Value, InStr(Value, ",") + 1)) * tline / bline
        If Power > 0 Then Text1.Text = Text1.Text + "Power dissapation=" + convert_SI(Power) + " watts<BR>"
     End If
leave:
    RichOutput
    'If Picture2(6).Visible Then Text1.Text = "<body bgcolor=#EEEEEE>"
    'If Picture2(25).Visible = False Then Picture2(6).Visible = True
    'Picture2(20).Visible = False
Case 7 ' print
    'If Picture2(29).Visible = False Then
     'printout
    'Else
     'Picture2(3).Visible = True
     'Picture2(7).Visible = False
     'DrawGraph 0, 0, Printer.Width, Printer.Height, 1
     'Printer.EndDoc
    'End If
Case 8 'ok
     'Line (240, 2760)-Step(5532, 1692), RGB(255, 255, 255), BF
     pref2.Hide
Case 10 'change from real->ideal
    'Picture2(10).Visible = False
    'Picture2(17).Visible = True
    'If Val(Ctyp) = 22 Then
    'T2(1) = "0"
    'Else
    'T2(0) = "0" ' Diode
    'End If
Case 17 'change from ideal->real
    'Picture2(17).Visible = False
    'Picture2(10).Visible = True
    If Val(Ctyp) = 22 Then
    'T2(1) = T2(1) + "+0.025/I"
    Else
    'T2(0) = T2(0) + "+0.025/I"
    End If
Case 19 ' ccode
    If Me.Height = Backdrop.Height Then Me.Height = Me.Height * 2
    Value = canvas(DD.Cform).values.list(Val(Ref))
    res = Mid(Value, InStr(Value, ",") + 1)
    Text1.Text = Text1.Text + "<BR>" + ccode(Val(res))
    RichOutput
Case 20 ' stop
    mesh11.LOOPS.Clear
    'Picture2(6).Visible = True
    'Picture2(20).Visible = False
    
Case 25 ' graph
     'mesh11.MathDoc = False
     'If Picture2(1).Visible Then
     ' ' intitialise graphing form
     ' Text1.Visible = False
     ' T2(0).Text = ""
     ' T2(1).Text = ""
     ' T2(2).Text = ""
     ' CUplim = 1
     ' Clowlim = -1
     ' For i = 0 To 31
     '  Select Case i
     '   Case 8, 9, 5, 0, 3: Picture2(i).Visible = True
     '   Case 28, 26, 9, 29, 30: Picture2(i).Visible = True
     '   Case Else: Picture2(i).Visible = False
     '  End Select
     ' Next ' automatically put on scope
     ' value1.Caption = "Upper Limit"
     ' value2.Caption = "Lower Limit"
     ' value3.Caption = ""
     ' Ctyp = "24"
     ' ReNew
     'End If ' draw graph immediately
     'Picture2(26).Visible = True
     'Picture2(25).Visible = False
     '' Line (0, 2760)-Step(5772, 1700), RGB(255, 255, 255), BF
     ' Integrate options:
     'DrawGraph 0, 0, pref2.Width / 2, pref2.Width / 2, 0
     'Picture2(25).Visible = True
     'Picture2(7).Visible = True
     'Picture2(3).Visible = False
Case 27 ' solve
    'If value1.Caption = "Upper Limit" Then
    ' If Picture2(6).Visible Then Ctyp = "26" Else Ctyp = "25"
    ' ReNew
    'Else
    ' If Picture2(6).Visible Then
    '  If T2(0).Text <> "" Then
    '   If mesh11.prcall = Ccbran Then
    '        mesh11.prcall = mesh11.rev(Ccbran)
    '   Else: mesh11.prcall = Ccbran
    '   End If
    '   prcall = mesh11.Analyse()
    '   For i = 1 To Len(FxVer)
    '    MFXS = Mid(FxVer, i, 1)
    '    If MFXS = "," Then
    '     If flag Then Exit For Else flag = True
    '    End If
    '    If flag And MFXS <> "," Then V1 = V1 + MFXS
    '   Next
    '   X1 = mesh11.expand(s(), Xcram + 10)
    '   Y1 = Mid(X1, InStr(X1, ",") + 1)
    '   X1 = Left(X1, InStr(X1, ",") - 1)
    '   s(1) = Mid(prcall, InStr(prcall, "/") + 1)
    '   s(2) = T2(0).Text
    '   s(2) = mesh11.expand(s(), Xmult)
    '   s(1) = Left(prcall, InStr(prcall, "/") - 1)
    '   nl = Chr(13) + Chr(10)
    '   s(1) = mesh11.expand(s(), Xcram)
    '   Text1.Text = Text1.Text + nl + "<BR>" + s(1)
    '   FxVer = Left(s(1), InStr(s(1), "=") - 1): DVBL = Mid(FxVer, i + 1)
    '   Text1.Text = Text1.Text + nl + "<BR>" + X1 + "=" + SI_convert(solve(0)): DVBL = V1
    '   FxVer = Mid(s(1), InStr(s(1), ",") + 1, InStr(Len(FxVer) + 2, s(1), "=") - 1)
    '   Text1.Text = Text1.Text + nl + "<BR>" + Y1 + "=" + SI_convert(solve(0))
    '  Else
    '   angle = Val(T2(2).Text) * 0.01745329251994
    '   s(1) = T2(1).Text
    '   s(2) = mesh11.STR2(Sin(angle))
    '   T2(0).Text = mesh11.expand(s(), Xmult)
    '   s(1) = T2(1).Text
    '   s(2) = mesh11.STR2(Cos(angle))
    '   s(1) = mesh11.expand(s(), Xmult)
    '   T2(0).Text = T2(0).Text + format(s(1)) + "J"
    '  End If
    ' Else
    '  If T2(1).Text = "" Then
    '   T2(1).Text = convert_SI(solve(Val(T2(0).Text)))
    '  ElseIf T2(0).Text = "" Then
    '    T2(0).Text = convert_SI(fx(Val(T2(1).Text)))
    '  End If
    ' End If
    'End If
Case 29 'scope
    ' Ctyp = "24"
    ' ReNew
Case 30
    'Picture2(26).Visible = False
    'Picture2(25).Visible = True
    'Picture2(30).Visible = False
    'Picture2(31).Visible = True
Case 31
    'Picture2(26).Visible = False
    'Picture2(25).Visible = True
    'Picture2(31).Visible = False
    'Picture2(30).Visible = True
Case 32
 'Picture2(32).Visible = False
 'Picture2(33).Visible = True
Case 33
 'Picture2(33).Visible = False
 'Picture2(32).Visible = True
End Select
End Sub
Private Sub Picture2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Picture2(Index).BorderStyle = 1 Then GoTo quickout
'For j = 0 To 22: Picture2(j).BorderStyle = 0: Next
Select Case Index
Case 0, 1, 5, 6, 7, 8, 10, 17, 19, 20, 25, 27, 29, 30, 31, 32, 33
'Picture2(Index).BorderStyle = 1
End Select
quickout:
End Sub
Public Function convert_SI(a) As String
For i = -12 To 9 Step 3
If 10 ^ i > Abs(a) Then i = i - 3: Exit For
Next i
a = a / 10 ^ i
a = Int(a * 100000) / 100000
Select Case i
Case -12: op = "f"
Case -9: op = "p"
Case -6: op = "u"
Case -3: op = "m"
Case 3: op = "K"
Case 6: op = "M"
Case 9: op = "G"
End Select
convert_SI = LTrim(Str(a)) + op
End Function
Public Function SI_convert(a) As Double
Dim expon As Double
SI_convert = -999
'remove trailing zeros -done
' must have sign - done
' all || -> | - done
' remove J's - done
If Len(a) = 0 Then GoTo out
If Left(a, 1) <> "+" And Left(a, 1) <> "-" Then a = "+" + a
If a = "+0" Then SI_convert = 0: GoTo out
Select Case Right(a, 1)
Case Is < "A", "p", "n", "u", "m", "K", "M"
Case Else: GoTo out
End Select
i = 1: Do While i < Len(a)
ca = Mid(a, i, 1)
If (la = "+" Or la = "-") And ca = "0" Then a = Left(a, i - 1) + Mid(a, i + 1): i = 1
If Mid(a, i, 2) = "--" Then a = Left(a, i) + Mid(a, i + 2): i = 1
If Mid(a, i, 2) = "+-" Then a = Left(a, i - 1) + Mid(a, i + 1): i = 1
If Mid(a, i, 2) = "-+" Then a = Left(a, i) + Mid(a, i + 2): i = 1
If Mid(a, i, 2) = "++" Then a = Left(a, i) + Mid(a, i + 2): i = 1
If la = "|" And ca = "|" Then a = Left(a, i - 1) + Mid(a, i + 1): i = 1
If ca = "J" Or ca = "j" Then a = Left(a, i - 1) + Mid(a, i + 1): i = 1
If (la = "+" Or la = "-") And ca >= "A" Then a = Left(a, i - 1) + "1" + Mid(a, i): i = 1
la = ca
i = i + 1: Loop
If Right(a, 1) = "|" Or Right(a, 1) < "0" Then GoTo out ' term can't end in |
ReDim p(Len(a) + 1) As Integer
ReDim pop(Len(a) \ 2) As Integer
Dim tot As Double
Do: cp = cp + 1
p1 = InStr(p(cp - 1) + 2, a, "+")
p2 = InStr(p(cp - 1) + 2, a, "-")
p3 = InStr(p(cp - 1) + 2, a, "|")
If p1 = 0 Then p1 = Len(a) + 1
If p2 = 0 Then p2 = Len(a) + 1
If p3 = 0 Then p3 = Len(a) + 1
If p1 < p2 And p1 < p3 Then p(cp) = p1
If p2 < p1 And p2 < p3 Then p(cp) = p2
If p3 < p1 And p3 < p1 Then p(cp) = p3: poc = poc + 1: pop(poc) = cp:
Loop Until p(cp) = 0
p(cp) = Len(a) + 1: p(0) = 1
ReDim el(cp)
ReDim num(cp) As Double
ReDim vr(cp)
For cx = 1 To cp
numb = Mid(a, p(cx - 1), p(cx) - p(cx - 1))
If Left(numb, 1) = "|" Then numb = Mid(numb, 2)
num(cx) = Val(numb)
vps = p(cx - 1) + Len(Str(num(cx)))
vr(cx) = Mid(a, vps + offset, p(cx) - vps)
Select Case vr(cx)
Case ""
Case "p"
expon = 1 / 1000000000
Case "u"
expon = 1 / 1000000
Case "m"
expon = 1 / 1000
Case "K", "k"
expon = 1000
Case "M"
expon = 1000000
Case "G"
expon = 1000000000
Case Else
GoTo out
End Select
If expon = 0 Then expon = 1
num(cx) = num(cx) * expon
Next
For i = 1 To poc
num(pop(i)) = num(pop(i) + 1) * num(pop(i)) / (num(pop(i) + 1) + num(pop(i)))
num(pop(i) + 1) = 0
Next: For i = 1 To cp
tot = tot + num(i)
Next: SI_convert = tot '
out:
End Function
Private Function ccode(num)
If num < 100 Then Exit Function
ReDim ccde(10)
ccde(0) = "Black"
ccde(1) = "Brown"
ccde(2) = "Red"
ccde(3) = "Orange"
ccde(4) = "Yellow"
ccde(5) = "Green"
ccde(6) = "Blue"
ccde(7) = "Violet"
ccde(8) = "Gray"
ccde(9) = "White"
a = LTrim(Str(num))
l = Len(a)
ReDim band(3)
band(1) = ccde(Val(Mid(a, 1, 1)))
band(2) = ccde(Val(Mid(a, 2, 1)))
band(3) = ccde(l - 2)
nl = Chr(13) + Chr(10)
ic = Chr(34)
op = Str(Int(num)) + "=" + nl
For i = 1 To 3
 op = op + "<a style=" + ic + "background-color:" + band(i)
 op = op + ";color=" + band(i) + ";" + ic + ">---</a>" + nl
Next: Cls
ccode = op
End Function


Private Sub T2_Change(Index As Integer)
Dim target As Object
Set target = canvas(DD.Cform)
ReDim enb(22), dsb(22), V(4)
enb(10) = 16: dsb(10) = 15
enb(11) = 16: dsb(11) = 15
enb(12) = 12: dsb(12) = 11
enb(13) = 12: dsb(13) = 11
enb(14) = 24: dsb(14) = 23
enb(15) = 24: dsb(15) = 23
enb(16) = 14: dsb(16) = 13
enb(17) = 14: dsb(17) = 13
enb(18) = 14: dsb(18) = 13
enb(19) = 14: dsb(19) = 13
enb(20) = 14: dsb(20) = 13
enb(21) = 14: dsb(21) = 13
enb(22) = 14: dsb(22) = 13
' check t2's for illegal entries
' 1.NULL
' 2.f=0 with Trc
' 3.Real Variable cap (ind ok)
' 4.multi-term Hoe or Hfe        [ SI(t2(1)) & ctyp="22" ]
' 5.Variable Frequency
' 6.bad dependancy format
If Val(Ctyp) < 24 Then
 'Picture2(4).Visible = False: Picture2(8).Visible = True
 'If (T2(0).Text = "") And value1.Caption <> "" Then Picture2(4).Visible = True: Picture2(8).Visible = False: GoTo out
 'If (T2(1).Text = "") And value2.Caption <> "" Then Picture2(4).Visible = True: Picture2(8).Visible = False: GoTo out
 'If (T2(2).Text = "") And value3.Caption <> "" Then Picture2(4).Visible = True: Picture2(8).Visible = False: GoTo out
 If Val(Ctyp) > 11 And Val(Ctyp) < 14 Then
        'If Val(freq) = 0 Or (SI_convert(T2(0).Text) = -999 And InStr(UCase(T2(0).Text), "J") = 0) Then
        ' Picture2(4).Visible = True: Picture2(8).Visible = False: GoTo out
        'End If
 End If
 'If Val(Ctyp) = 22 And (InStr(2, T2(0).Text, "+") Or InStr(2, T2(0).Text, "+")) Then Picture2(4).Visible = True: Picture2(8).Visible = False: GoTo out
 'If Val(Ctyp) = 22 And (InStr(2, T2(2).Text, "+") Or InStr(2, T2(2).Text, "+")) Then Picture2(4).Visible = True: Picture2(8).Visible = False: GoTo out
 'If Val(Ctyp) = 17 And SI_convert(T2(2).Text) = -999 Then Picture2(4).Visible = True: Picture2(8).Visible = False: GoTo out
 For z = 0 To 2 ' almost idiot proof
  'If Right(T2(z), 1) < "0" And T2(z) <> "" Then Picture2(4).Visible = True: Picture2(8).Visible = True: GoTo out
  'T2(z).Text = LTrim(T2(z).Text)
  'T2(z).Text = mesh11.rev(LTrim(mesh11.rev(T2(z).Text)))
 Next
 ' Icon changing loop
 For i = 1 To 3
  'V(i) = LTrim(Str(SI_convert(T2(i - 1).Text)))
  If value2.Caption = "" And i = 2 Then GoTo skipit
  If value3.Caption = "" And i = 3 Then GoTo skipit ' patch
  If V(i) = "-999" Then
     'V(i) = T2(i - 1).Text ' case sensitive!!
     If value1.Caption <> "" Then
      'Picture2(enb(Val(Ctyp))).Visible = True
      'Picture2(dsb(Val(Ctyp))).Visible = False
      If Val(Ctyp) = 10 Or Val(Ctyp) = 11 Then
         'Picture2(19).Visible = False
         'Picture2(18).Visible = True
      End If
      'Picture2(enb(Val(Ctyp))).Visible = True
      'Picture2(dsb(Val(Ctyp))).Visible = False
     End If
     If value2.Caption <> "" Then
      'Picture2(enb(Val(Ctyp))).Visible = True
      'Picture2(dsb(Val(Ctyp))).Visible = False
     End If
  Else ' entry is numeric
     If value1.Caption <> "" Then
      'Picture2(enb(Val(Ctyp))).Visible = False
      'Picture2(dsb(Val(Ctyp))).Visible = True
      If (Val(Ctyp) = 10 Or Val(Ctyp) = 11) And Val(V(i)) > 999 Then
        'Picture2(19).Visible = True
        'Picture2(18).Visible = False
       End If
     End If
     If value2.Caption <> "" Then
      'Picture2(enb(Val(Ctyp))).Visible = False
      'Picture2(dsb(Val(Ctyp))).Visible = True
     End If
   End If
skipit:
 Next
End If
'*************************************
indx = Val(Ref)
ReDim s(2) As String
Select Case Val(Ctyp)
Case 10, 11
    target.values.list(indx) = V(1)
Case 12 To 15
    'If InStr(UCase(T2(0).Text), "J") Then
    ' target.values.list(indx) = V(1)
    'Else
    ' s(1) = "6.28318530718"
    ' s(2) = freq
    ' s(1) = mesh11.expand(s(), Xmult)
    ' s(2) = V(1)
    ' s(1) = mesh11.expand(s(), Xmult) ' assuming s(1) is fixed
    ' If Val(Ctyp) = 12 Or Val(Ctyp) = 13 Then
    '  target.values.list(indx) = "-" + mesh11.STR2(1 / Val(s(1))) + "J"
    ' Else
    '  target.values.list(indx) = s(1) + "J"
    ' End If
    'End If
Case 16, 17
    target.values.list(indx) = V(2) + "," + V(1) 'voltage,resistance
    If Val(Ctyp) = 17 Then
        'freq = SI_convert(T2(2).Text)
        'v3 = LTrim(Str(SI_convert(T2(2).Text)))
        If Val(v3) > 0 Then
        'Picture2(22).Visible = True
        'Picture2(21).Visible = False
        Else
        'Picture2(22).Visible = False
        'Picture2(21).Visible = True
        End If
    End If
Case 18 To 21
    target.values.list(indx) = V(2) + "," + V(1)
Case 22
    target.values.list(indx) = V(1) + "," + V(2) + "," + V(3)
Case 25
    'If SI_convert(T2(0).Text) = -999 Xor SI_convert(T2(1).Text) = -999 Then
     'Picture2(27).Visible = True
     'Picture2(28).Visible = False
    'Else
    ' Picture2(28).Visible = True
    ' Picture2(27).Visible = False
    'End If
Case 26
    'test = mesh11.numerical(mesh11.strem(T2(0).Text, "J"))
    'If test Xor (SI_convert(T2(1).Text) = -999 And SI_convert(T2(2).Text) = -999) Then
    ' Picture2(27).Visible = True
    ' Picture2(28).Visible = False
    'Else
    ' Picture2(28).Visible = True
    ' Picture2(27).Visible = False
    'End If
Case 24
    ' uplim = SI_convert(T2(0).Text)
    ' lowlim = SI_convert(T2(1).Text)
    ' If uplim <> -999 And lowlim <> -999 And uplim > lowlim Then
    '  CUplim = Str(SI_convert(T2(0).Text))
    '  Clowlim = Str(SI_convert(T2(1).Text))
    '  Picture2(26).Visible = False
    '  Picture2(25).Visible = True
    '  Picture2(28).Visible = False
    '  Picture2(27).Visible = True
    ' Else
    '  Picture2(25).Visible = False
    '  Picture2(26).Visible = True
    '  Picture2(27).Visible = False
    '  Picture2(28).Visible = True
    ' End If
End Select
target.changed = True
MDI.Status.Panels(2).Text = "Changed"
target.Tile(Index).ToolTipText = target.values.list(Index)
out:
End Sub
Function Linked() As Boolean
Dim target As Object
Set target = canvas(DD.Cform)
Linked = True
ReDim C(20, 10) ' the piece in x,y
For Y = 0 To 9
For X = 0 To 19
C(X, Y) = target.datalist.list(X + (Y * 20))
If Val(C(X, Y)) = 22 Then
    C(X, Y) = "3": C(X - 1, Y) = "19": C(X, Y - 1) = "16"
End If
Next X, Y
C(20, 10) = "-1"
' set up IO udlr
ReDim udlr(50), dgb(10)
udlr(0) = "0110"
udlr(1) = "0111"
udlr(2) = "0101"
udlr(3) = "1110"
udlr(4) = "0011" 'H
udlr(5) = "1011"
udlr(6) = "1001"
udlr(7) = "1010"
udlr(8) = "1101"
udlr(9) = "1100" 'V
udlr(10) = "0011" 'H
udlr(11) = "1100" 'V
udlr(12) = "0011" 'H
udlr(13) = "1100" 'V
udlr(14) = "0011" 'H
udlr(15) = "1100" 'V
udlr(16) = "1100" 'V
udlr(17) = "1100" 'V
udlr(18) = "1100" 'V
udlr(19) = "0011" 'H
udlr(20) = "0011" 'H
udlr(21) = "1100" 'V
nopieces = 1

For Y = 0 To 9
For X = 0 To 19
'are all pieces connected
If C(X, Y) = "-1" Then GoTo skip_icheck
cpath = udlr(Val(C(X, Y)))
If Y > 0 Then
If C(X, Y - 1) <> "-1" Then
upath = udlr(Val(C(X, Y - 1)))
Else: upath = "0000"
End If
Else: upath = "0000"
End If
If Y < 9 Then
If C(X, Y + 1) <> "-1" Then
dpath = udlr(Val(C(X, Y + 1)))
Else: dpath = "0000"
End If
Else: dpath = "0000"
End If
If X > 0 Then
If C(X - 1, Y) <> "-1" Then
lpath = udlr(Val(C(X - 1, Y)))
Else: lpath = "0000"
End If
Else: lpath = "0000"
End If
If X < 19 Then
If C(X + 1, Y) <> "-1" Then
rpath = udlr(Val(C(X + 1, Y)))
Else: rpath = "0000"
End If
Else: rpath = "0000"
End If
upath = Mid(upath, 2, 1)
dpath = Mid(dpath, 1, 1)
lpath = Mid(lpath, 4, 1)
rpath = Mid(rpath, 3, 1) ' double checking
If Mid(cpath, 1, 1) <> upath Then Linked = False: Exit For
If Mid(cpath, 2, 1) <> dpath Then Linked = False: Exit For
If Mid(cpath, 3, 1) <> lpath Then Linked = False: Exit For
If Mid(cpath, 4, 1) <> rpath Then Linked = False: Exit For
If Val(C(X, Y)) = 16 Or Val(C(X, Y)) = 17 Then voltage = 1
skip_icheck:
Next X, Y
If voltage = 0 Then Linked = False
End Function
Sub linkup(DoLoops As Boolean)
Dim target As Object
Set target = canvas(DD.Cform)
For i = 0 To 50
 Foll(i) = 255
Next
ReDim n(20, 10) ' the node of c(x,y)
ReDim C(20, 10) ' the piece in x,y
ReDim V(200)
ReDim din(50)
ReDim Term(2) As String
With mesh11
        .LINKS.Clear
        .Resist.Clear
        .Volt.Clear
        .depres.Clear
        .depvolt.Clear
        .semic.Clear
 End With
For i = 1 To 50
 mesh11.depvolt.AddItem ""
 mesh11.depres.AddItem ""
Next i
For Y = 0 To 9
 For X = 0 To 19
 C(X, Y) = target.datalist.list(X + (Y * 20))
 V(X + (Y * 20)) = target.values.list(X + (Y * 20))
 If Val(C(X, Y)) = 22 Then
    C(X, Y) = "3": C(X - 1, Y) = "19": C(X, Y - 1) = "16"
    a = V(X + (Y * 20))
    Hoe = Left(a, InStr(a, ",") - 1): Hie = Mid(a, Len(Hoe) + 2)
    If SI_convert(Hoe) <> -999 Then Hoe = LTrim(Str(SI_convert(Hoe)))
    Hie = Left(Hie, InStr(Hie, ",") - 1)
    Hfe = Mid(a, Len(Hoe) + Len(Hie) + 2)
    'hoe1*hfe
    vres = Mid(V(X + (Y - 1) * 20), InStr(V(X + (Y - 1) * 20), ",") + 1)
    vres = format(vres)
    If vres = "0" Or vres = "+" Then vres = ""
    ' YFS=hoe*hfe*T,hoe+vres
    Term(1) = Hoe
    Term(2) = Hfe
    Yfs = mesh11.expand(Term(), Xmult) + "*T," + Hoe + vres
    V(X + (Y - 1) * 20) = "+0+" + Yfs
    hiepp = Mid(V(X - 1 + Y * 20), InStr(V(X - 1 + Y * 20), ",") + 1)
    hiepp = format(TermOut(hiepp))
    If hiepp = "0" Or hiepp = "+" Then hiepp = "" Else hiepp = hiepp + "+"
    V(X - 1 + Y * 20) = ".7," + hiepp + Hie
 End If
Next X, Y
C(20, 10) = "-1"
' set up IO udlr
' PathFinder Replaces from here--------------
ReDim udlr(50), dgb(10)
udlr(0) = "0110"
udlr(1) = "0111"
udlr(2) = "0101"
udlr(3) = "1110"
udlr(4) = "0011" 'H
udlr(5) = "1011"
udlr(6) = "1001"
udlr(7) = "1010"
udlr(8) = "1101"
udlr(9) = "1100" 'V
udlr(10) = "0011" 'H
udlr(11) = "1100" 'V
udlr(12) = "0011" 'H
udlr(13) = "1100" 'V
udlr(14) = "0011" 'H
udlr(15) = "1100" 'V
udlr(16) = "1100" 'V
udlr(17) = "1100" 'V
udlr(18) = "1100" 'V
udlr(19) = "0011" 'H
udlr(20) = "0011" 'H
udlr(21) = "1100" 'V
nopieces = 1
' no need to do linking check.
For Y = 0 To 10
 For X = 0 To 20
 If islink(C(X, Y)) And n(X, Y) = "" Then
    nodes = nodes + 1: ox = X: oy = Y
    n(X, Y) = Chr(64 + nodes)
    Do: GoSub next_link
    If X1 <> -1 Then n(X1, Y1) = Chr(64 + nodes)
    dgb(nodes) = dgb(nodes) + 1
    X = X1: Y = Y1
    Loop While X1 <> -1
    X = ox + 1: Y = oy
 End If
Next X, Y
' mop up extra pieces
pass = 1: Do
  is_links = False
  For X = 0 To 19
    For Y = 0 To 9
    'upgrade this to correct mislabeled tiles
     If islink(C(X, Y)) Then
      passup = True: GoSub join: passup = False
      passdn = True: GoSub join: passdn = False
      passlf = True: GoSub join: passlf = False
      passrt = True: GoSub join
     End If
  Next Y, X
Loop While is_links
' Pathfinder replaces from here ---------------------------
ReDim onod(50) As Integer
For Y = 1 To nodes
upperN(Y) = 0
LeftN(Y) = 0
Next
For Y = 0 To 9
  For X = 0 To 19
   If n(X, Y) <> "" Then onod(Asc(n(X, Y)) - 64) = onod(Asc(n(X, Y)) - 64) + 1
   If islink(C(X, Y)) Then
    i = Asc(n(X, Y)) - 64
    If X + 1 > LeftN(i) = 0 Then LeftN(i) = X + 1
    If Y > upperN(i) Then upperN(i) = Y
   End If
Next X, Y: offf = 0
For i = 1 To 50
 If onod(i) = 0 Then
  off = off + 1
 Else ' relabel every piece (- off nodes)
  For Y = 0 To 9
   For X = 0 To 19
   If n(X, Y) = Chr(i + 64) Then n(X, Y) = Chr(i + 64 - off)
  Next X, Y
 End If
Next i
nodes = 50 - off
ReDim ptr(50), ptv(nodes), ptd(50)
mesh11.Volt.Clear
For i = 1 To nodes
mesh11.Volt.AddItem ""
Next i

For X = 0 To 19
 For Y = 0 To 9
  If C(X, Y) <> "-1" And Not islink(C(X, Y)) Then
   'proxy node
   ln = "": un = ""
   lr = False: ur = False
   For i = 0 To 20
    oln = ln
    If i < X And islink(C(i, Y)) Then ln = n(i, Y)
   If oln <> ln Then lr = False
   If i < X And Val(C(i, Y)) > 9 Then lr = True
   If i > X And islink(C(i, Y)) Then rn = n(i, Y): Exit For
  Next
  For i = 0 To 10
   oun = un
   If i < Y And islink(C(X, i)) Then un = n(X, i)
   If oun <> un Then ur = False
   If i < Y And Val(C(X, i)) > 9 Then ur = True
   If i > Y And islink(C(X, i)) Then DN = n(X, i): Exit For
  Next
  If udlr(C(X, Y)) = "1100" Then
   link = un + DN: pxr = ur
  Else: link = ln + rn: pxr = lr
  End If
  ' check for link - note CPS has 0 origin
  For cps = 0 To mesh11.LINKS.ListCount - 1
   eqlink = mesh11.LINKS.list(cps)
   dblink = eqlink = link Or eqlink = mesh11.rev(link)
   If dblink Then Exit For
  Next ' proxy-ammended lines
  off = 0
  If Not dblink Then
   'serial
   mesh11.LINKS.AddItem link
   mesh11.Resist.AddItem ""
  ElseIf pxr = False Then
    nodes = nodes + 1 ' in parellel
    With mesh11
     .LINKS.AddItem Left(link, 1) + Chr(64 + nodes)
     .LINKS.AddItem Chr(64 + nodes) + Right(link, 1)
     .Resist.AddItem "+0"
     .Resist.AddItem ""
     .Volt.AddItem ""
    End With
    LeftN(nodes) = X + 1
    upperN(nodes) = Y
    olink = Left(link, 1)
    link = Chr(64 + nodes) + Right(link, 1)
    off = 2
    ' convert all left(link,1) to chr(64+nodes)
    For tx = 0 To 19
     For ty = 0 To 9
     If n(tx, ty) = olink Then n(tx, ty) = Chr(64 + nodes)
    Next ty, tx
  End If
  If X + Y * 20 = Val(Ref) Then
   Ccbran = link
  End If
  cps = cps + 1
  Value = V(X + Y * 20)
  res = Mid(Value, InStr(Value, ",") + 1)
  'ennumerate resistance
    la = Left(res, 1)
    If la <> "+" Or la <> "-" Then res = "+" + res: la = "+"
    For i = 2 To Len(res) + 1
    ca = Mid(res, i, 1)
    If (la = "+" Or la = "-") And ca >= "A" Then res = Left(res, i - 1) + "1" + Mid(res, i)
    la = ca:  Next
    depres = Mid(res, Len(TermOut(res)) + 1)
    If depres <> "" Then
     If Right(depres, 1) = "I" Then depres = Left(depres, Len(depres) - 1) + link
     If Right(depres, 1) = "T" Then
        ' find T link.
        n2 = n(X, Y + 1)
        For i = X - 1 To 0 Step -1
        If islink(C(i, Y + 1)) And n(i, Y + 1) <> n2 Then n1 = n(i, Y + 1): Exit For
        Next: depres = Left(depres, Len(depres) - 1) + n1 + n2
     End If
    End If
    For rpos = 0 To mesh11.LINKS.ListCount - 1
     If mesh11.LINKS.list(rpos) = link Then Exit For
    Next
    If mesh11.Resist.list(rpos) = "+0" Then mesh11.Resist.list(rpos) = ""
    mesh11.Resist.list(rpos) = mesh11.Resist.list(rpos) + TermOut(res)
    If mesh11.depres.list(rpos) = "" Then mesh11.depres.list(rpos) = depres
    Volt = Left(Value, InStr(Value, ","))
    If Volt <> "" Then
    'enumerate voltage
    la = Left(Volt, 1)
    If la <> "+" And la <> "-" Then Volt = "+" + Volt: la = "+"
    For i = 2 To Len(Volt)
     ca = Mid(Volt, i, 1)
     If (la = "+" Or la = "-") And ca >= "A" Then Volt = Left(Volt, i - 1) + "1" + Mid(Volt, i)
     la = ca: Next
    Volt = Left(Volt, Len(Volt) - 1) ' this should get rid of the ,
    depvolt = Mid(Volt, Len(TermOut(Volt)) + 1)
    If depvolt <> "" Then
     If Right(depvolt, 1) = "I" Then depvolt = Left(depvolt, Len(depvolt) - 1) + link
     If Right(depvolt, 1) = "T" Then
        ' find T link.
        n2 = n(X, Y + 1)
        For i = X - 1 To 0 Step -1
        If islink(C(i, Y + 1)) And n(i, Y + 1) <> n2 Then n1 = n(i, Y + 1): Exit For
        Next: depvolt = Left(depvolt, Len(depvolt) - 1) + n1 + n2
     End If
    End If
    If Left(Volt, 1) <> "+" And Left(Volt, 1) <> "-" Then Volt = "+" + Volt
    ' change un to left(link,1)
    vltp = Asc(Left(link, 1)) - 65
    vltp2 = Asc(Right(link, 1)) - 65
    If mesh11.Volt.list(vltp) = "GND" Then swap = vltp: vltp = vltp2: vltp2 = swap
    If mesh11.Volt.list(vltp) = "+0" Then mesh11.Volt.list(vltp) = ""
    If Val(C(X, Y)) < 18 Then
     If mesh11.Volt.list(vltp) <> "" And Foll(vltp) <> vltp2 Then Volt = mesh11.invert(Volt)
      If Foll(vltp) = 255 Then Foll(vltp) = vltp2
      mesh11.Volt.list(vltp2) = "GND"
    Else
     Foll(vltp) = 254
     If Val(C(X, Y)) = 18 Or Val(C(X, Y)) = 20 Then link = mesh11.rev(link)
     mesh11.semic.AddItem link
     If mesh11.Volt.list(vltp) <> "" Then Volt = mesh11.invert(Volt)
    End If
    mesh11.Volt.list(vltp) = mesh11.Volt.list(vltp) + TermOut(Volt)
    mesh11.depvolt.list(vltp) = depvolt
 End If
End If
Next Y, X
' fill in ?'s
For i = 0 To nodes
If mesh11.Volt.list(i) = "" Then mesh11.Volt.list(i) = "?"
Next
' multi circuit support -----------------
ReDim oknode(nodes)
ln = Left(Ccbran, 1)
rn = Right(Ccbran, 1)
oknode(Asc(ln) - 64) = 1
oknode(Asc(rn) - 64) = 1
For pass = 1 To nodes: mto = False
For i = mesh11.LINKS.ListCount - 1 To 0 Step -1
 ln = Left(mesh11.LINKS.list(i), 1)
 rn = Right(mesh11.LINKS.list(i), 1)
 If oknode(Asc(ln) - 64) Or oknode(Asc(rn) - 64) Then
  oknode(Asc(ln) - 64) = 1
  oknode(Asc(rn) - 64) = 1
 Else
  mto = True
 End If
Next i, pass: gnd = 0
' remove bad grounds
If mto = False Then
 For i = 0 To nodes - 1
  If mesh11.Volt.list(i) = "GND" Then
   If gnd Then mesh11.Volt.list(i) = "?" Else gnd = 1
  End If
 Next i
End If
For i = mesh11.LINKS.ListCount - 1 To 0 Step -1
 ln = Left(mesh11.LINKS.list(i), 1)
 rn = Right(mesh11.LINKS.list(i), 1)
 If oknode(Asc(ln) - 64) Then
 Else ' remove ln
 mesh11.Volt.list(Asc(ln) - 65) = "NULL"
 mesh11.depvolt.list(Asc(ln) - 65) = "NULL"
 mesh11.Resist.list(i) = "NULL"
 mesh11.depres.list(i) = "NULL"
 mesh11.LINKS.list(i) = "NULL"
 LeftN(Asc(ln) - 64) = 22
 upperN(Asc(ln) - 64) = 22
 Foll(Asc(ln) - 64) = 255
 End If
 If oknode(Asc(rn) - 64) Then
 Else ' remove rn
 mesh11.Volt.list(Asc(rn) - 65) = "NULL"
 mesh11.depvolt.list(Asc(rn) - 65) = "NULL"
 mesh11.Resist.list(i) = "NULL"
 mesh11.depres.list(i) = "NULL"
 mesh11.LINKS.list(i) = "NULL"
 LeftN(Asc(rn) - 64) = 22
 upperN(Asc(rn) - 64) = 22
 Foll(Asc(ln) - 64) = 255
 End If
Next
For i = mesh11.LINKS.ListCount - 1 To 0 Step -1
 If mesh11.LINKS.list(i) = "NULL" Then
  mesh11.LINKS.RemoveItem i
  mesh11.Resist.RemoveItem i
  mesh11.depres.RemoveItem i
 End If
Next

For i = nodes To 0 Step -1
 If mesh11.Volt.list(i) = "NULL" Then
  mesh11.Volt.RemoveItem i
  mesh11.depvolt.RemoveItem i
  For z = i + 1 To nodes
   LeftN(z) = LeftN(z + 1)
   upperN(z) = upperN(z + 1)
   Foll(z) = Foll(z + 1)
  Next
 End If
Next: off = 0
' relabel links
For i = 1 To nodes
If oknode(i) Then oknode(i) = off Else off = off + 1
Next
 ln = Left(Ccbran, 1)
 rn = Right(Ccbran, 1)
 ln = Chr(Asc(ln) - oknode(Asc(ln) - 64))
 rn = Chr(Asc(rn) - oknode(Asc(rn) - 64))
 Ccbran = ln + rn
For i = 0 To mesh11.LINKS.ListCount - 1
 ln = Left(mesh11.LINKS.list(i), 1)
 rn = Right(mesh11.LINKS.list(i), 1)
 ln = Chr(Asc(ln) - oknode(Asc(ln) - 64))
 rn = Chr(Asc(rn) - oknode(Asc(rn) - 64))
 mesh11.LINKS.list(i) = ln + rn
  If i < mesh11.semic.ListCount - 1 Then
  ln = Left(mesh11.semic.list(i), 1)
  rn = Right(mesh11.semic.list(i), 1)
  ln = Chr(Asc(ln) - oknode(Asc(ln) - 64))
  rn = Chr(Asc(rn) - oknode(Asc(rn) - 64))
  mesh11.semic.list(i) = ln + rn
 End If
Next i
nodes = mesh11.Volt.ListCount
' end of multi circuit support ---------
'update cbran
'loop detection
mesh11.NODE.Caption = LTrim(Str(nodes))
frmProgress.Progress.Visible = True
If DoLoops Then
 mesh11.LOOPS.Clear: cnt = 0: i = 0
 EI = enumb(nodes, nodes)
 Do
 i = i + 1
 If frmProgress.stopAnalysis Then GoTo end_linkup
 ps = permute(p, nodes): p = i
 frmProgress.Progress.Value = (i / EI) * 90
 vlps = valid(ps)
 If vlps > 0 Then
  mesh11.LOOPS.AddItem Left(ps, vlps)
  cnt = cnt + 1
 End If
 Loop Until i = EI
 mesh11.Number.Caption = Str(cnt)
 doubles
 frmProgress.Progress.Value = 0
 frmProgress.Progress.Visible = False
End If
GoTo end_linkup
'*******************************************************************************
next_link: '********************************************************************
X1 = -1 '***********************************************************************
dcp = udlr(Val(C(X, Y))) '******************************************************
' pass=0 :- adjacent link must be unlabeled to be accepted
' pass=1 :- ignore labelling
    'up
If Val(Mid(dcp, 1, 1)) And Y > 0 Then
    If C(X, Y - 1) <> "-1" Then
    conn = Val(Mid(udlr(Val(C(X, Y - 1))), 2, 1))
    If conn And (n(X, Y - 1) = "" Or passup) And islink(C(X, Y - 1)) Then X1 = X: Y1 = Y - 1
End If
End If
    'down
If Val(Mid(dcp, 2, 1)) And Y < 9 Then
    If C(X, Y + 1) <> "-1" Then
    conn = Val(Mid(udlr(Val(C(X, Y + 1))), 1, 1))
    If conn And (n(X, Y + 1) = "" Or passdn) And islink(C(X, Y + 1)) Then X1 = X: Y1 = Y + 1
End If
End If
    'right
If Val(Mid(dcp, 4, 1)) And X < 19 Then
    If C(X + 1, Y) <> "-1" Then
    conn = Val(Mid(udlr(Val(C(X + 1, Y))), 3, 1))
    If conn And (n(X + 1, Y) = "" Or passrt) And islink(C(X + 1, Y)) Then X1 = X + 1: Y1 = Y
End If
End If
    'left
If Val(Mid(dcp, 3, 1)) And X > 0 Then
    If C(X - 1, Y) <> "-1" Then
    conn = Val(Mid(udlr(Val(C(X - 1, Y))), 4, 1))
    If conn And (n(X - 1, Y) = "" Or passlf) And islink(C(X - 1, Y)) Then X1 = X - 1: Y1 = Y
End If '************************************************************************
End If '************************************************************************
Return '************************************************************************
'*******************************************************************************
join:
GoSub next_link
If X1 <> -1 Then
 Select Case True
  Case n(X1, Y1) > n(X, Y)
    n(X, Y) = n(X1, Y1): is_links = True
  Case n(X, Y) < n(X1, Y1)
    n(X1, Y1) = n(X, Y): is_links = True
 End Select
End If
Return
end_linkup:
End Sub
Private Function islink(a) As Boolean
If Val(a) >= 0 And Val(a) < 10 And a <> "" Then islink = True
End Function
Private Function TermOut(Term)
'This function will remove the depenadant part of a term
TermOut = Term
p1 = InStr(Term, "/")
p2 = InStr(Term, "*")
If p1 Or p2 Then Else GoTo out
For i = p1 + p2 To 0 Step -1
If Mid(Term, i, 1) = "+" Or Mid(Term, i, 1) = "-" Then Exit For
Next i:
If i = 0 Then TermOut = "" Else TermOut = Left(Term, i - 1)
out:
End Function
Private Function enumb(n, b) As Double
h = (b \ 2 + (b / 2 <> b \ 2) ^ 2)
If n = 0 Then
enumb = 0
Else
enumb = h * b ^ (n - 1) + enumb(n - 1, b)
End If
End Function
Private Function valid(a) As Integer
'branches = Val(BRANCH.Caption)
branches = mesh11.LINKS.ListCount
ReDim link(branches)
For i = 1 To branches
link(i) = mesh11.LINKS.list(i - 1)
Next
ReDim used(branches)
If Len(a) > 3 Then
For i = 1 To 3 ' This must remesh11 at 3!!!
p = Mid(a, i, 2): ex = exist(p)
If ex = 0 Then GoTo NOTValid
used(ex) = 1: Next
For i = 3 To branches
p = Mid(a, i, 1) + Left(a, 1)
ex = exist(p)
If ex And used(ex) = 0 Then valid = i: GoTo ISvalid
Next
GoTo NOTValid
ISvalid:
For z = 3 To i - 1
p = Left(a, 1) + Mid(a, z, 1)
If exist(p) Then valid = 0: GoTo NOTValid
Next:
For pri = 2 To i - 2
For sec = pri + 2 To branches
p = Mid(a, pri, 1) + Mid(a, sec, 1)
If exist(p) Then valid = 0: GoTo NOTValid
Next sec, pri
Else
a = a + Left(a, 1)
For i = 1 To Len(a) - 1
If exist(Mid(a, i, 2)) = 0 Then valid = 0: GoTo NOTValid
Next: valid = Len(a) - 1
End If
NOTValid:
End Function
Private Function exist(a) As Integer
'branches = Val(BRANCH.Caption)
branches = mesh11.LINKS.ListCount
ReDim link(branches)
For i = 1 To branches
link(i) = mesh11.LINKS.list(i - 1)
Next
exist = 0: For i = 1 To branches
If a = link(i) Or mesh11.rev(a) = link(i) Then exist = i
Next
End Function
Private Function permute(X, b) As String
If X Mod 50 = 0 Then dummy = DoEvents
For C = b - 1 To 0 Step -1
p = p + Chr(((X \ (b ^ C)) Mod b) + 65)
Next: permute = p
End Function
Private Sub doubles()
'cnt = Val(Number.Caption)
cnt = mesh11.LOOPS.ListCount
ReDim loopz(cnt) As String
For i = 1 To cnt
loopz(i) = mesh11.LOOPS.list(i - 1)
Next
For pri = 0 To cnt
For sec = pri + 1 To cnt
If contains(loopz(pri), loopz(sec)) = 1 Then loopz(pri) = ""
Next sec, pri

For i = 1 To cnt
If loopz(i) <> "" Then tot = tot + 1
Next
ReDim Temp(tot): tot = 0
For i = 1 To cnt
If loopz(i) <> "" Then tot = tot + 1: Temp(tot) = loopz(i)
Next:
mesh11.LOOPS.Clear
mesh11.Number.Caption = Str(tot)
For i = 1 To tot
mesh11.LOOPS.AddItem Temp(i)
Next
End Sub
Private Function contains(a, b) As Integer
' this is a simplified version of the contain function,
' it assumes that
' (a) there is no repetition of letters
' (b) there are no sub loops
If Len(a) <> Len(b) Then GoTo notcontained
For i = 1 To Len(a)
If InStr(a, Mid(b, i, 1)) = 0 Then GoTo notcontained
Next: contains = 1
notcontained:
End Function


Private Sub Text1_DblClick()
Text1.Text = ""
End Sub

Public Sub ReNew()
Dim s(2) As String
Coldval = canvas(DD.Cform).values.list(Val(Ref))
'T2(0).Text = "" ' this triggers the
'T2(1).Text = "" ' t2() change event which
'T2(2).Text = "" ' wipes entry on values.list
canvas(DD.Cform).values.list(Val(Ref)) = Coldval
If Val(Ctyp) < 24 Then
 'For i = 8 To 33
 'Picture2(i).Visible = False
 'Next
 'Picture2(1).Visible = True
 'Picture2(6).Visible = False
 'Picture2(8).Visible = True
 'Picture2(4).Visible = False
 ReDim picon(23) As Integer
 ReDim Label1(23)
 ReDim Label2(23)
 ReDim Label3(23)
 picon(10) = 1518: Label1(10) = "Resistance (Ohms)": Label2(10) = "": Label3(10) = "":
 picon(11) = 1518: Label1(11) = "Resistance (Ohms)": Label2(11) = "": Label3(11) = "":
 picon(12) = 1100: Label1(12) = "Capacitance ": Label2(12) = "": Label3(12) = "":
 picon(13) = 1100: Label1(13) = "Capacitance ": Label2(13) = "": Label3(13) = "":
 picon(14) = 2300: Label1(14) = "Inductance (Henrys)": Label2(14) = "": Label3(14) = "":
 picon(15) = 2300: Label1(15) = "Inductance (Henrys)": Label2(15) = "": Label3(15) = "":
 picon(16) = 1321: Label1(16) = "Internal Resistance": Label2(16) = "Voltage": Label3(16) = "":
 picon(17) = 1322: Label1(17) = "Internal Resistance": Label2(17) = "Voltage": Label3(17) = "Frequency":
 picon(18) = 1310: Label1(18) = "Forward Bias": Label2(18) = "Diode Voltage": Label3(18) = "":
 picon(19) = 1310: Label1(19) = "Forward Bias": Label2(19) = "Diode Voltage": Label3(19) = "":
 picon(20) = 1310: Label1(20) = "Forward Bias": Label2(20) = "Diode Voltage": Label3(20) = "":
 picon(21) = 1310: Label1(21) = "Forward Bias": Label2(21) = "Diode Voltage": Label3(21) = "":
 picon(22) = 1000: Label1(22) = "1/Hoe": Label2(22) = "Hie": Label3(22) = "Hfe"
 piss = picon(Val(Ctyp))
 piss1 = Int(piss / 100)
 piss2 = piss - (piss1 * 100)
 'Picture2(piss1).Visible = True
 'Picture2(piss2).Visible = True
 value1.Caption = Label1(Val(Ctyp))
 value2.Caption = Label2(Val(Ctyp))
 value3.Caption = Label3(Val(Ctyp))
'check if linked
 'If Linked And Val(Ctyp) <> 22 Then Picture2(2).Visible = False: Picture2(6).Visible = True
 'If Not Linked Or Val(Ctyp) = 22 Then Picture2(6).Visible = False: Picture2(2).Visible = True
End If
Select Case Val(Ctyp)
Case Is < 10: pref2.Hide
Case 10 To 15
'T2(0).Text = canvas(DD.Cform).values.list(Val(Ref))
Case 16 To 21
a = canvas(DD.Cform).values.list(Val(Ref))
'T2(1).Text = Left(a, InStr(a, ",") - 1) ' resistance
'T2(0).Text = Mid(a, InStr(a, ",") + 1) 'voltage
'If Val(Ctyp) = 17 Then T2(2).Text = freq
Case 22
a = canvas(DD.Cform).values.list(Val(Ref))
'T2(0).Text = Left(a, InStr(a, ",") - 1)
'T2(1).Text = Mid(a, Len(T2(0).Text) + 2)
'T2(1).Text = Left(T2(1).Text, InStr(T2(1).Text, ",") - 1)
'T2(2).Text = Mid(a, Len(T2(0).Text) + Len(T2(1).Text) + 3)
Case 24
     'Backdrop.Height = 2760
     value1.Caption = "Upper Limit"
     value2.Caption = "Lower Limit"
     value3.Caption = ""
'     T2(0).Text = LTrim(CUplim)
 '    T2(1).Text = LTrim(Clowlim)
Case 25
     value1.Caption = "Current ="
     value2.Caption = "  " + DVBL + " ="
     value3.Caption = ""
  '   Picture2(28).Visible = True
   '  Picture2(27).Visible = False
Case 26
     value1.Caption = "Phazor   ="
     value2.Caption = "Amplitude="
     value3.Caption = "Phase    ="
   '  Picture2(28).Visible = True
    ' Picture2(27).Visible = False
End Select
End Sub
Public Sub printout()
Dim cpn(50) As String
With Printer
 .FontSize = 12
 PX = .Width / 20
 PY = .Height / 20
For Y = 0 To 9
 For X = 0 To 19
 If Val(canvas(DD.Cform).datalist.list(X + Y * 20)) > -1 Then _
 .PaintPicture canvas(DD.Cform).Tile(X + Y * 20).Picture, _
   X * PX, Y * PY, PX, PY: ly = Y
  If Val(canvas(DD.Cform).datalist.list(X + Y * 20)) >= 10 Then
 cp = cp + 1
 cpn(cp) = canvas(DD.Cform).values.list(X + Y * 20)
 .CurrentX = X * PX
 .CurrentY = PY * (Y + 0.75)
 Printer.Print Str(cp)
 End If
Next X, Y
.CurrentX = 0
.CurrentY = (ly + 1) * PY
Printer.Print "Mesh 2.0 Analysis - Copyright 1999 F.Reid"
For i = 1 To cp
Printer.Print "Component #" + LTrim(Str(i)) + " - " + cpn(i)
Next
Printer.Print Text1.Text
.EndDoc
End With
End Sub
Private Function VarUtil(VL) As String
' home of the varutil arse-up!
' if sent NULL then return dependant variable (if exists)
' if sent a "number" then modify volt&resist list.
ReDim p(20), vr(100)
restot = mesh11.Resist.ListCount
For i = 0 To restot + mesh11.Volt.ListCount
  If i > restot Then
   a = mesh11.Volt.list(i - restot - 1)
   Else: a = mesh11.Resist.list(i)
  End If
 If mesh11.numerical(a) Or InStr(a, "[") Then GoTo DontReplace
 a = TermOut(a)
 For K = 2 To Len(a)
    lk = Mid(a, K - 1, 1)
    ck = Mid(a, K, 1)
    If (lk = "+" Or lk = "-") And ck >= "A" Then a = Left(a, K - 1) + "1" + Mid(a, K)
 Next
 cp = 0: op = 0: Do: cp = cp + 1: p(cp) = 0
  p1 = InStr(p(cp - 1) + 2, a, "+")
  p2 = InStr(p(cp - 1) + 2, a, "-")
  If p1 < p2 And p1 Then p(cp) = p1
  If p1 > p2 And p2 Then p(cp) = p2
 Loop Until p(cp) = 0
 p(cp) = Len(a) + 1: p(0) = 1
 ReDim el(cp)
 For cx = 1 To cp
  num = Val(Mid(a, p(cx - 1), p(cx) - p(cx - 1)))
  vps = p(cx - 1) + Len(Str(num))
  If p(cx) >= vps Then 'patch
   vrb = Mid(a, vps, p(cx) - vps):
  If vrb <> "" Then
       vr(t) = vrb: t = t + 1
       num = num * SI_convert(convert_SI(Val(VL)))
       End If
   End If
  op = op + num
 Next cx
If VL <> "" Then
 If i > restot Then
   mesh11.Volt.list(i - restot - 1) = SI_convert(convert_SI(op))
   Else: mesh11.Resist.list(i) = SI_convert(convert_SI(op))
 End If
End If
DontReplace:
Next i
For sec = 0 To t - 1
 If vr(0) <> vr(sec) Then duplicate = True
Next
If duplicate = False Then VarUtil = vr(0)
End Function
Private Function solve(target)
Dim s(2) As String
Upper = Val(CUplim)
Lower = Val(CUplim)
If Lower > Upper Then swap = Upper: Upper = Lower: Lower = swap
d = (Upper - Lower) / 10
X = Lower
frmProgress.Progress.Visible = True
For i = 1 To 100
 tmpx = X
 fofx = fx(tmpx)
 'If Picture2(31).Visible Then fofx = ((fx(X) + 0.001) - fofx) / 0.001
 If fofx > target Xor d < 0 Then d = d / -2: gotcha = True
 If fofx = target Then Exit For
 X = Fix(X * 100000000) / 100000000 + d ' emulate dbl precision
 If i Mod 5 = 0 Then
 frmProgress.Progress.Value = i
 DoEvents
 End If
Next i
frmProgress.Progress.Value = 0
frmProgress.Progress.Visible = False
'T2(1).Text = Str(X)
'If gotcha = False Then T2(1).Text = "Scope is too narrow"
solve = X
End Function
Public Function fx(X) As Double
' >>> Nested Xpression <<<
ReDim tf(ASCSN), s(2) As String
tempdvbl = DVBL
If DVBL = "" Then MsgBox "DVBL is not set"
For i = 1 To ASCSN
 tf(i) = ascs(i)
Next
For pri = 1 To ASCSN
 If Left(ascs(pri), 1) = "~" Then ascs(pri) = Mid(ascs(pri), 2)
 s(1) = mesh11.STR2(X)
 s(2) = ascs(pri)
 ascs(pri) = mesh11.expand(s(), Xqfx)
 If mesh11.numerical(ascs(pri)) And Val(ascs(pri)) < 0 And Left(tf(pri), 1) = "~" Then
  fx = 0
  GoTo Finalize
 End If
 For sec = 2 To pri
   DVBL = "[X" + LTrim(Str(sec - 1)) + "]"
   s(1) = ascs(sec - 1)
   s(2) = ascs(pri)
   ascs(pri) = mesh11.expand(s(), Xqfx)
Next sec, pri
s(1) = ascs(ASCSN)
Select Case True
 'Case Picture2(32).Visible: s(2) = "Phase": fx = Val(mesh11.expand(s(), Xphamp))
 'Case Picture2(33).Visible: s(2) = "Amp": fx = Val(mesh11.expand(s(), Xphamp))
 Case Else: fx = Val(s(1))
End Select
Finalize:
For i = 1 To ASCSN
 ascs(i) = tf(i)
Next
DVBL = tempdvbl
End Function
Public Function format(a)
format = a
If Left(a, 1) <> "+" And Left(a, 1) <> "-" Then format = "+" + a
End Function

