VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Mesh"
   ClientHeight    =   4125
   ClientLeft      =   3285
   ClientTop       =   3660
   ClientWidth     =   5850
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grid"
            Object.ToolTipText     =   "Grid"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   0
      Top             =   3852
      Width           =   5844
      _ExtentX        =   10319
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4683
            Text            =   "Press F1 for Help"
            TextSave        =   "Press F1 for Help"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Unchanged"
            TextSave        =   "Unchanged"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Mesh1.cct"
            TextSave        =   "Mesh1.cct"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Filename"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   240
      Top             =   3120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   720
      Top             =   3120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":28B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":29C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2E0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnutoolbox 
         Caption         =   "Tool&box"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLabels 
         Caption         =   "&Labels"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "&Grid"
      End
      Begin VB.Menu mnubar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Mesh..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuParam1 
         Caption         =   "(PARAM1)"
      End
      Begin VB.Menu mnuParam2 
         Caption         =   "(PARAM2)"
      End
      Begin VB.Menu mnuParam3 
         Caption         =   "(PARAM3)"
      End
      Begin VB.Menu mnuSepPopup 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuAnalyse 
         Caption         =   "Analyse"
      End
      Begin VB.Menu mnuSimulate 
         Caption         =   "Simulate"
      End
      Begin VB.Menu mnuGraph 
         Caption         =   "Graph"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "Solve"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public popupLeft, popupTop
Public HTML
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Sub MDIForm_Load()
 ' updated 24.4.00
 Me.Left = GetSetting(App.title, "Settings", "MainLeft", CELL_SPACING)
 Me.Top = GetSetting(App.title, "Settings", "MainTop", CELL_SPACING)
 Me.Width = GetSetting(App.title, "Settings", "MainWidth", 6500)
 Me.Height = GetSetting(App.title, "Settings", "MainHeight", 6500)
 Me.Visible = True
 DoEvents
 hToolbar = FindWindowEx(Toolbar1.hwnd, 0&, "ToolbarWindow32", vbNullString)
 Style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
 Style = Style Or TBSTYLE_FLAT
 Call SendMessageLong(hToolbar, TB_SETSTYLE, 0, Style)
 Toolbar1.Refresh
 LoadNew
 'UDP.Bind
 DoEvents
 Toolbox.Show
 Toolbox.Left = Me.Width - Toolbox.Width
 Toolbox.Top = Me.Height - Toolbox.Height
 App.HelpFile = App.Path & "\mesh2.hlp"
 Me.mnuPopup.Visible = False
End Sub
Private Sub MDIForm_Resize()
 If Me.WindowState = vbMinimized Then
  Toolbox.Hide
  DoEvents
 End If
End Sub

Private Sub mnuAnalyse_Click()
 

 
 
 ' Direct cut & paste out of pref2
 Dim s(2) As String
 HTML = "<body bgcolor=#FFFFFF>"
 mesh11.MathDoc = True
 ' show progress
 frmProgress.Show
 frmProgress.stopAnalysis = False
 frmProgress.Progress = frmProgress.Progress.Min
 DoEvents
 With mesh11
  .LINKS.Clear
  .Resist.Clear
  .Volt.Clear
  .depres.Clear
  .depvolt.Clear
 End With
 pref2.Ref = canvas(DD.Cform).SelectedIndex
 pref2.linkup True ' create meshlist
 
 ' no resistance in circuit.
 If mesh11.LINKS.ListCount = 1 Then
  HTML = HTML & "Current=Infintity"
  showHTML HTML
  Exit Sub
 End If
 
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
   dpc = dpc + 1
   dep(dpc) = mesh11.depvolt.list(i)
   dpp(dpc) = i + restot + 1
   If Not mesh11.numerical(Left(dep(dpc), Len(dep(dpc)) - 3)) Then alpha = True
  End If
 Next
 If alpha Then ' alpha ascs
  HTML = "<body bgcolor=#EEEEEE>"
  For i = dpc To 1 Step -1
   If mesh11.prcall = dep(i) Then
    mesh11.prcall = mesh11.rev(dep(i))
   Else
    mesh11.prcall = dep(i)
   End If
   prcall = mesh11.Analyse()
   tline = Left(prcall, InStr(prcall, "/") - 1)
   bline = Mid(prcall, InStr(prcall, "/") + 1)
   If Left(Right(dep(i), 3), 1) = "/" Then
    swap = tline
    tline = bline
    bline = swap
   End If
   s(1) = Left(dep(i), Len(dep(i)) - 3)
   s(2) = tline
   tline = mesh11.expand(s(), Xmult)
   HTML = HTML & "[X" + LTrim(Str(i)) + "]=" + tline + "/" + bline + "<BR>"
   ascs(i) = tline + "/" + bline
   If Left(Right(dep(i), 3), 1) = "/" Then ascs(i) = "~" + ascs(i)
   If dpp(i) > restot Then
    mesh11.Volt.list(dpp(i) - restot - 1) = mesh11.Volt.list(dpp(i) - restot) + "+1[X" + LTrim(Str(i)) + "]"
   Else
    mesh11.Resist.list(dpp(i)) = mesh11.Resist.list(dpp(i)) + "+1[X" + LTrim(Str(i)) + "]"
   End If
  Next
 End If
 If mesh11.prcall = pref2.Ccbran Then
  mesh11.prcall = mesh11.rev(pref2.Ccbran)
 Else
  mesh11.prcall = pref2.Ccbran
 End If
 With mesh11
  prcall = .remove1(.Analyse)
 End With
 If alpha Then
  nl = Chr(13) + Chr(10)
  pofs = InStr(prcall, "/")
  tline = Left(prcall, pofs - 1)
  bline = Mid(prcall, pofs + 1)
  If strCount(tline & bline, "-") > strCount(tline & bline, "+") Then
   tline = mesh11.invert(tline)
   bline = mesh11.invert(bline)
  End If
  HTML = HTML & "Current=<center><u>" + tline + "</u><BR>"
  HTML = HTML & bline + "</center>"
  ascs(dpc + 1) = prcall
  pref2.ASCSN = dpc + 1
  pref2.linkup False
  s(2) = Left(prcall, InStr(prcall, "/") - 1) _
         + "+" + Mid(prcall, InStr(prcall, "/") + 1)
  FxVer = mesh11.expand(s(), Xvar)
  pref2.DVBL = Mid(FxVer, 3)
  mnuGraph.Enabled = False
  Select Case Left(FxVer, 1)
   Case "1": ' simple eq
    mnuSolve.Enabled = True
    mnuGraph.Enabled = True
   Case "2": ' Complex eq
    mnuSolve.Enabled = False
    s(1) = Left(prcall, InStr(prcall, "\") - 1)
    s(2) = Mid(prcall, InStr(prcall, "\") + 1)
    s(1) = mesh11.expand(s(), Xcplx)
    s(2) = "Both"
    HTML = HTML & "<BR>" + mesh11.expand(s(), Xphamp)
    'Picture2(32).Visible = True
    'Picture2(25).Visible = True
   Case "3": ' simultanious eq
    mnuSolve.Enabled = False
    ' NOTE: This is solveable, just need mechanisim to enter
    '       range of values.
    'value1.Caption = "Upper Limit"
    'value2.Caption = "Lower Limit"
    'value3.Caption = ""
    'Ctyp = "24"
    'ReNew
    'Picture2(29).Visible = True
    'Picture2(27).Visible = True
   Case Else: ' unsolveable
    mnuSolve.Enabled = False
  End Select
 Else
  tline = Val(Left(prcall, InStr(prcall, "/") - 1))
  bline = Val(Mid(prcall, InStr(prcall, "/") + 1))
  If bline = 0 Then bline = 0.0001
  HTML = "Current=" + pref2.convert_SI(tline / bline) + " a<BR>"
  Value = canvas(DD.Cform).values.list(Val(Ref))
  Power = Val(Mid(Value, InStr(Value, ",") + 1)) * tline / bline
  If Power > 0 Then
   HTML = HTML & "Power dissapation=" + pref2.convert_SI(Power) + " watts<BR>"
  End If
 End If

 
 If Not frmProgress.stopAnalysis Then showHTML (HTML)
End Sub
Public Sub showHTML(HTML)
 ' should load IE or shell
 frmProgress.Hide
 applicationPath = App.Path
 If Right(applicationPath, 1) <> "\" Then
  applicationPath = applicationPath & "\"
 End If
 Open applicationPath & "mesh.html" For Output As #1
  Print #1, HTML;
 Close
 On Error GoTo AlternativeFlow:
 Set IE = CreateObject("InternetExplorer.application")
 IE.Visible = True
 IE.navigate2 App.Path & "\mesh.html"
 Exit Sub
AlternativeFlow:
 ' used if there is any problem with IE,
 ' not compatible with Windows NT
 Shell "start " & App.Path & "\mesh.html", vbNormalFocus
End Sub

Private Sub mnuDelete_Click()
 On Error GoTo ThrowException
 
  If Not canvas(DD.Cform).ActiveObject Is Nothing Then
   Dim components(200)
   Dim properties(200)
   For i = 0 To canvas(DD.Cform).datalist.ListCount - 1
    components(i) = canvas(DD.Cform).datalist.list(i)
    properties(i) = canvas(DD.Cform).values.list(i)
   Next
   canvas(DD.Cform).ActiveObject.Delete components, properties, canvas(DD.Cform).SelectedIndex
  End If
 
 With canvas(DD.Cform)
  .datalist.list(.SelectedIndex) = "-1"
  .Tile(.SelectedIndex).Picture = Toolbox.Sym(99).Picture
 End With
 Exit Sub
ThrowException:
 MsgBox Err.Description
End Sub


Private Sub mnuGraph_Click()
 frmSetRange.Show vbModal
 If frmSetRange.wasCancelled Then Exit Sub
 mesh11.MathDoc = False
 pref2.CUplim = pref2.SI_convert(frmSetRange.txtMax)
 pref2.Clowlim = pref2.SI_convert(frmSetRange.txtMin)
 pref2.DrawGraph 0, 0, pref2.Width / 2, pref2.Width / 2, 0
End Sub
Private Sub mnuGrid_Click()
 If DD.Cform > 0 Then
  For i = 0 To 199
   canvas(DD.Cform).Tile(i).BorderStyle = _
   canvas(DD.Cform).Tile(i).BorderStyle Xor 1
  Next
 End If
 Refresh
End Sub

Private Sub mnuLabels_Click()
 mnuLabels.Checked = Not mnuLabels.Checked
 canvas(DD.Cform).ReNew
End Sub

Private Sub mnuParam1_Click()
 ComponentIndex = canvas(DD.Cform).SelectedIndex
 frmSetParameter.txtValue.text = parseValues(ComponentIndex, 1)
 frmSetParameter.selectedParameter = 1
 frmSetParameter.Show
End Sub
Private Sub mnuParam2_Click()
 ComponentIndex = canvas(DD.Cform).SelectedIndex
 frmSetParameter.txtValue.text = parseValues(ComponentIndex, 2)
 frmSetParameter.selectedParameter = 2
 frmSetParameter.Show
End Sub
Private Sub mnuParam3_Click()
 ComponentIndex = canvas(DD.Cform).SelectedIndex
 frmSetParameter.txtValue.text = parseValues(ComponentIndex, 3)
 frmSetParameter.selectedParameter = 3
 frmSetParameter.Show
End Sub

Private Sub mnuSimulate_Click()
 Dim components(200)
 Dim properties(200)
 For i = 0 To canvas(DD.Cform).datalist.ListCount - 1
  components(i) = canvas(DD.Cform).datalist.list(i)
  properties(i) = canvas(DD.Cform).values.list(i)
 Next
 ActiveReading = canvas(DD.Cform).ActiveObject.Analyse(components, properties, canvas(DD.Cform).SelectedIndex)
 showHTML ActiveReading
End Sub

Private Sub mnutoolbox_Click()
 Toolbox.Visible = Not Toolbox.Visible
 mnutoolbox.Checked = Toolbox.Visible
End Sub

Private Sub UDP_DataArrival(ByVal BytesTotal As Long)
 Dim buffer As String
 UDP.GetData buffer
 ' parse buffer
 
 If InStr(buffer, ":") Then SessionID = Val(Left(buffer, InStr(buffer, ":") - 1))
 found = 0
 For i = 1 To DD.Forms
  If canvas(i).DesignID = SessionID Then found = i
 Next
 buffer = Mid(buffer, InStr(buffer, ":") + 1)
 args = Split(buffer, ",")
 If args(0) <> "NEW" And found = 0 Then found = DD.Cform
 Debug.Print SessionID & ">" & buffer
 Select Case UCase(args(0))
  Case "SETFOCUS"
   canvas(found).SetFocus
   UDP.SendData "+OK,Canvas " & SessionID & " Activated" & vbCrLf
  Case "NEW"
   Success = LoadNew
   If Not Success Then
    UDP.SendData "-ERR,Too many simultanious users" + vbCrLf
    Exit Sub
   End If
   canvas(DD.Cform).DesignID = SessionID
   UDP.SendData "+OK,User " & SessionID & " Sucessfully logged in" & vbCrLf
  Case "CLOSE"
   canvas(found).SetFocus
   canvas(found).changed = False
   mnuFileClose_Click
   DoEvents
   UDP.SendData "+OK,User " & SessionID & " Sucessfully logged out" & vbCrLf
  Case "PLACE"
   ModXY = Val(args(2)) + Val(args(3)) * 20
   canvas(found).datalist.list(ModXY) = args(1)
   For i = 4 To UBound(args) - 1
    PieceValue = PieceValue + args(i) + ","
   Next
   PieceValue = PieceValue + args(UBound(args))
   canvas(found).values.list(ModXY) = PieceValue
   canvas(found).ReNew
   UDP.SendData "+OK,Component Successfully placed" + vbCrLf
  Case "RETR"
   ModXY = Val(args(1)) + Val(args(2)) * 20
   PieceNumber = canvas(found).datalist.list(ModXY)
   Value = canvas(found).values.list(ModXY)
   UDP.SendData "+OK," & PieceNumber & "," & Value & vbCrLf
  Case "ANALYSE"
   ModXY = Val(args(1)) + Val(args(2)) * 20
   Response = pref2.AnalyseRequest(ModXY) ' to be written
   UDP.SendData "+OK," & Response ' response is mesh21.htm
 End Select
' Extra code added 24.4.00 for Inet support
' Requires Winsock: name:UDP,localport 1050, remoteport 1055
End Sub

Private Function LoadNew()
' updated 24.4.00
If DD.Forms = 3 Then LoadNew = False: Exit Function
DD.Forms = DD.Forms + 1
ReDim Preserve canvas(DD.Forms)
Set canvas(DD.Forms) = New Design
canvas(DD.Forms).cvs = DD.Forms
canvas(DD.Forms).init
DD.Cform = DD.Forms
For i = 0 To 199
 canvas(DD.Forms).Tile(i).BorderStyle = 0
Next
inc = 1: Do
 title = "Mesh" + LTrim(Str(inc)) + ".cct"
 inc = inc + 1
Loop While FormExist(title)
canvas(DD.Forms).Caption = title
LoadNew = True
End Function
Public Function FormExist(ByVal title) As Boolean
title = UCase(title)
For i = 1 To DD.Forms
 If UCase(canvas(i).Caption) = title Then
  FormExist = True
  Exit For
 End If
Next
End Function


Private Sub MDIForm_Terminate()
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If Me.WindowState <> vbMinimized Then
   SaveSetting App.title, "Settings", "MainLeft", Me.Left
   SaveSetting App.title, "Settings", "MainTop", Me.Top
   SaveSetting App.title, "Settings", "MainWidth", Me.Width
   SaveSetting App.title, "Settings", "MainHeight", Me.Height
End If
End
End Sub
Public Sub Refresh()
If DD.Forms > 0 Then
 With canvas(DD.Cform)
  Select Case .changed
   Case False: Status.Panels(2).text = "Unchanged"
   Case True: Status.Panels(2).text = "Changed"
  End Select
  Status.Panels(3).text = .Caption
  Toolbar1.Buttons(8).Value = .Tile(0).BorderStyle

 End With
 Toolbar1.Buttons(3).Enabled = True
 Toolbar1.Buttons(5).Enabled = True
 Toolbar1.Buttons(8).Enabled = True
 mnuFileSave.Enabled = True
 mnuFileSaveAs.Enabled = True
 mnuFilePrint.Enabled = True
 mnuFileClose.Enabled = True
 If DD.Forms = 3 Then
  Toolbar1.Buttons(1).Enabled = False
  Toolbar1.Buttons(2).Enabled = False
  mnuFileOpen.Enabled = False
  mnuFileNew.Enabled = False
 Else
  Toolbar1.Buttons(1).Enabled = True
  Toolbar1.Buttons(2).Enabled = True
  mnuFileOpen.Enabled = True
  mnuFileNew.Enabled = True
 End If
Else
 Toolbar1.Buttons(3).Enabled = False
 Toolbar1.Buttons(5).Enabled = False
 Toolbar1.Buttons(8).Enabled = False
 mnuFileSave.Enabled = False
 mnuFileSaveAs.Enabled = False
 mnuFilePrint.Enabled = False
 mnuFileClose.Enabled = False
 Status.Panels(2).text = ""
 Status.Panels(3).text = ""
End If
End Sub
Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub
Private Sub mnuViewOptions_Click()
Options.Show vbModal, Me
End Sub
Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        Status.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        Status.Visible = True
        mnuViewStatusBar.Checked = True
    End If
End Sub
Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        Toolbar1.Visible = False
        mnuViewToolbar.Checked = False
    Else
        Toolbar1.Visible = True
        mnuViewToolbar.Checked = True
    End If
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    App.HelpFile = App.Path & "\mesh2.hlp"
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub
Private Sub mnuHelpSearch_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub
Private Sub mnuWindowArrangeIcons_Click()
Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowCascade_Click()
Me.Arrange vbCascade
End Sub
Private Sub mnuWindowNewWindow_Click()
LoadNew
End Sub
Private Sub mnuWindowTileHorizontal_Click()
Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowTileVertical_Click()
Me.Arrange vbTileVertical
End Sub
Private Sub mnuFileOpen_Click()
 On Error GoTo ThrowException
Dim sFile As String
With dlgCommonDialog
 .Filter = "Circuits (*.cct)|*.cct"
 .ShowOpen
 If Len(.Filename) = 0 Then Exit Sub
 sFile = .Filename
End With
Open sFile For Input As #1
Line Input #1, inputData
Select Case inputData
 Case "~~~~MESH 2.0 Data~~~~"
  mnuFileNew_Click
  canvas(DD.Cform).Caption = sFile
  canvas(DD.Cform).changed = False
  For i = 0 To 398
   Line Input #1, df
   If i < 200 Then
    If Val(df) = 23 Then df = "22"
    canvas(DD.Cform).datalist.list(i) = df
   Else
    canvas(DD.Cform).values.list(i - 200) = df
   End If
  Next i
  canvas(DD.Cform).ReNew
  Close #1
 Case "~~~~MESH 2.1 Data~~~~"
  mnuFileNew_Click
  canvas(DD.Cform).Caption = sFile
  canvas(DD.Cform).changed = False
  Line Input #1, df
  Set canvas(DD.Cform).ActiveObject = CreateObject(df)
  cctData = canvas(DD.Cform).ActiveObject.getCircuit
  cctDataRows = Split(cctData, vbCrLf)
  For i = 1 To 398
   df = cctDataRows(i)
   If i < 200 Then
    If Val(df) = 23 Then df = "22"
    canvas(DD.Cform).datalist.list(i) = df
   Else
    canvas(DD.Cform).values.list(i - 200) = df
   End If
  Next i
  canvas(DD.Cform).ReNew
  AboutBox = canvas(DD.Cform).ActiveObject.about
  If InStr(AboutBox, "<HTML>") > 1 Then
   showHTML AboutBox
  End If
  Close #1
 Case Else
  mtmp = MsgBox("Incompatable file type", 64, "Mesh2.1-Information")
  Exit Sub
End Select
Exit Sub
ThrowException:
 Close
 MsgBox "File Open failed due to " & Err.Description
End Sub
Public Sub mnuFileClose_Click()
If canvas(DD.Cform).changed = True Then
 If MsgBox("Circuit has been changed, Save Now?" _
 , vbInformation + vbOKCancel, "Mesh 2.1 Information") _
 = vbOK Then mnuFileSave_Click
End If
canvas(DD.Cform).Caption = "closing..."
Unload canvas(DD.Cform) ' added 27.05.00
For i = 1 To DD.Forms
 If i > DD.Cform And DD.Cform > 1 Then
    Set canvas(i - 1) = canvas(i)
    canvas(i - 1).cvs = i - 1
 End If
Next
DD.Forms = DD.Forms - 1
DD.Cform = DD.Forms
Refresh
End Sub
Private Sub mnuFileSave_Click()
SaveForm canvas(DD.Cform).Caption
End Sub
Private Sub mnuFileSaveAs_Click()
With dlgCommonDialog
 .Filter = "Circuits     (*.cct)|*.cct"
 .ShowSave
 If Len(.Filename) = 0 Then Exit Sub
SaveForm .Filename
End With
End Sub
Public Sub SaveForm(Filename)
Open Filename For Output As #1
Print #1, "~~~~MESH 2.0 Data~~~~"
For i = 0 To 398
 If i < 200 Then
  Print #1, canvas(DD.Cform).datalist.list(i)
 Else
  Print #1, canvas(DD.Cform).values.list(i - 200)
 End If
Next i
Close
canvas(DD.Cform).changed = False
Status.Panels(2).text = "Unchanged"
End Sub
Private Sub mnuFilePageSetup_Click()
dlgCommonDialog.ShowPrinter
End Sub
Private Sub mnuFilePrintPreview_Click()
Options.Show
End Sub
Private Sub mnuFilePrint_Click()
pref2.printout
End Sub
Private Sub mnuFileExit_Click()
Unload Me
End Sub
Private Sub mnuFileNew_Click()
LoadNew
Refresh
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1: LoadNew
Case 2: mnuFileOpen_Click
Case 3: mnuFileSave_Click
Case 5: mnuFilePrint_Click
Case 7:
        If Toolbox.Visible Then
            Toolbox.Visible = False
            Else
            Toolbox.Visible = True
        End If
Case 8:
       mnuGrid_Click
'Case 9: MsgBox "AutoWire Function not completed yet"
End Select
End Sub
Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'If MDI.Toolbar1.Buttons(7).Value = 0 Then Toolbox.Show
End Sub
Public Sub updatePopupMenu(cType)
 ' copy from pref2.renew
 Me.mnuParam1.Visible = True
 Me.mnuSepPopup.Visible = True
 Me.mnuParam2.Visible = False
 Me.mnuParam3.Visible = False
 Select Case cType
  Case 10, 11:
   Me.mnuParam1.Caption = "Resistance (Ohms)"
  Case 12, 13:
   Me.mnuParam1.Caption = "Capacitance"
  Case 14, 15:
   Me.mnuParam1.Caption = "Inductance"
  Case 16:
   Me.mnuParam1.Caption = "Internal Resistance"
   Me.mnuParam2.Caption = "Voltage"
   Me.mnuParam2.Visible = True
  Case 17:
   Me.mnuParam1.Caption = "Internal Resistance"
   Me.mnuParam2.Caption = "Voltage"
   Me.mnuParam3.Caption = "Frequency"
   Me.mnuParam2.Visible = True
   Me.mnuParam3.Visible = True
  Case 18, 19, 20, 21:
   Me.mnuParam1.Caption = "Forward Bias"
   Me.mnuParam2.Caption = "Diode Voltage"
   Me.mnuParam2.Visible = True
  Case 22:
   Me.mnuParam1.Caption = "1/Hoe"
   Me.mnuParam2.Caption = "Hie"
   Me.mnuParam3.Caption = "Hfe"
   Me.mnuParam2.Visible = True
   Me.mnuParam3.Visible = True
  Case Else
   Me.mnuParam1.Visible = False
   Me.mnuSepPopup.Visible = False
 End Select
 If pref2.Linked And cType >= 10 Then
  Me.mnuAnalyse.Enabled = True
  If Not .ActiveObject Is Nothing Then
   mnuSimulate.Enabled = True
  Else
   mnuSimulate.Enabled = False
  End If
 Else
  Me.mnuAnalyse.Enabled = False
 End If
End Sub
Public Function parseValues(ComponentIndex, Parameter)
 ComponentType = canvas(DD.Cform).datalist.list(ComponentIndex)
 componentValues = canvas(DD.Cform).values.list(ComponentIndex)
  
 Select Case Val(ComponentType)
  Case 10 To 15
   a = componentValues
  Case 16 To 21
   a = componentValues
   b = Left(a, InStr(a, ",") - 1) ' resistance
   a = Mid(a, InStr(a, ",") + 1)  ' voltage
   If ComponentType = 17 Then C = pref2.freq
  Case 22
   a = componentValues
   a = Left(a, InStr(a, ",") - 1)
   b = Mid(componentValues, Len(a) + 2)
   b = Left(b, InStr(b, ",") - 1)
   C = Mid(componentValues, Len(a) + Len(b) + 3)
 End Select
 Select Case Parameter
  Case 1
   parseValues = a
  Case 2
   parseValues = b
  Case 3
   parseValues = C
 End Select
End Function


