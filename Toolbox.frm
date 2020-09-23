VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Toolbox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1440
   ClientLeft      =   7650
   ClientTop       =   6810
   ClientWidth     =   5145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Toolbox.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Toolbox.frx":27A2
   ScaleHeight     =   1440
   ScaleWidth      =   5145
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   900
      ButtonWidth     =   826
      ButtonHeight    =   804
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Wires"
            Object.ToolTipText     =   "Wires"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Resistors"
            Object.ToolTipText     =   "Resistors"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "capacitors"
            Object.ToolTipText     =   "capacitors & inductors"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Power"
            Object.ToolTipText     =   "power sources"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "diode"
            Object.ToolTipText     =   "semiconductors"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "back"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "forward"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1000
         Left            =   0
         Picture         =   "Toolbox.frx":2BE4
         ScaleHeight     =   1005
         ScaleWidth      =   495
         TabIndex        =   38
         Top             =   1200
         Width           =   492
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3600
         Top             =   1080
      End
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   9
      Left            =   3000
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   37
      Top             =   1560
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   8
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   36
      Top             =   1560
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   7
      Left            =   1575
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   35
      Top             =   1560
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   6
      Left            =   2520
      MouseIcon       =   "Toolbox.frx":3026
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   34
      Top             =   612
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   5
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   33
      Top             =   1560
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   4
      Left            =   2280
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   32
      Top             =   1560
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   3
      Left            =   1896
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   31
      Top             =   612
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   2
      Left            =   1296
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   30
      Top             =   612
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   1
      Left            =   708
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   29
      Top             =   612
      Width           =   564
   End
   Begin VB.PictureBox box 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   708
      Index           =   0
      Left            =   120
      MouseIcon       =   "Toolbox.frx":3178
      ScaleHeight     =   675
      ScaleWidth      =   540
      TabIndex        =   28
      Top             =   588
      Width           =   564
   End
   Begin VB.Frame Frame1 
      Caption         =   "Invisible Data:"
      Height          =   1692
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   8172
      Begin VB.PictureBox Sym 
         Height          =   720
         Index           =   99
         Left            =   1200
         ScaleHeight     =   660
         ScaleWidth      =   1200
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.PictureBox Sym 
         Height          =   930
         Index           =   0
         Left            =   168
         Picture         =   "Toolbox.frx":35BA
         ScaleHeight     =   870
         ScaleWidth      =   705
         TabIndex        =   24
         Top             =   384
         Width           =   768
      End
      Begin VB.PictureBox Sym 
         Height          =   492
         Index           =   1
         Left            =   600
         Picture         =   "Toolbox.frx":3804
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   23
         Top             =   1320
         Width           =   732
      End
      Begin VB.PictureBox Sym 
         Height          =   732
         Index           =   2
         Left            =   4080
         Picture         =   "Toolbox.frx":3A4E
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   22
         Top             =   1200
         Width           =   852
      End
      Begin VB.PictureBox Sym 
         Height          =   732
         Index           =   4
         Left            =   2520
         Picture         =   "Toolbox.frx":3C98
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   21
         Top             =   1320
         Width           =   852
      End
      Begin VB.PictureBox Sym 
         Height          =   612
         Index           =   7
         Left            =   1704
         Picture         =   "Toolbox.frx":3EE2
         ScaleHeight     =   555
         ScaleWidth      =   675
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.PictureBox Sym 
         Height          =   492
         Index           =   10
         Left            =   1080
         Picture         =   "Toolbox.frx":412C
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   19
         Top             =   2160
         Width           =   732
      End
      Begin VB.PictureBox Sym 
         Height          =   612
         Index           =   12
         Left            =   1080
         Picture         =   "Toolbox.frx":4376
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   18
         Top             =   2520
         Width           =   612
      End
      Begin VB.PictureBox Sym 
         Height          =   492
         Index           =   13
         Left            =   1800
         Picture         =   "Toolbox.frx":45C0
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.PictureBox Sym 
         Height          =   492
         Index           =   14
         Left            =   2520
         Picture         =   "Toolbox.frx":480A
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.PictureBox Sym 
         Height          =   492
         Index           =   15
         Left            =   3240
         Picture         =   "Toolbox.frx":4A54
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   15
         Top             =   2520
         Width           =   612
      End
      Begin VB.PictureBox Sym 
         Height          =   492
         Index           =   16
         Left            =   1080
         Picture         =   "Toolbox.frx":4C9E
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   14
         Top             =   3000
         Width           =   612
      End
      Begin VB.PictureBox Sym 
         Height          =   492
         Index           =   17
         Left            =   1800
         Picture         =   "Toolbox.frx":4EE8
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   13
         Top             =   3000
         Width           =   612
      End
      Begin VB.PictureBox Sym 
         Height          =   852
         Index           =   18
         Left            =   4320
         Picture         =   "Toolbox.frx":5132
         ScaleHeight     =   795
         ScaleWidth      =   795
         TabIndex        =   12
         Top             =   2520
         Width           =   852
      End
      Begin VB.PictureBox Sym 
         Height          =   612
         Index           =   21
         Left            =   3252
         Picture         =   "Toolbox.frx":537C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label selected 
         Caption         =   "1"
         Height          =   252
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   852
      End
      Begin VB.Label final_sel 
         Caption         =   "17"
         Height          =   252
         Left            =   480
         TabIndex        =   26
         Top             =   3960
         Width           =   372
      End
   End
   Begin VB.PictureBox Sym 
      Height          =   492
      Index           =   8
      Left            =   4560
      Picture         =   "Toolbox.frx":55C6
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Sym 
      Height          =   612
      Index           =   9
      Left            =   3600
      Picture         =   "Toolbox.frx":5810
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Sym 
      Height          =   612
      Index           =   5
      Left            =   4200
      Picture         =   "Toolbox.frx":5A5A
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.PictureBox Sym 
      Height          =   852
      Index           =   6
      Left            =   4560
      Picture         =   "Toolbox.frx":5CA4
      ScaleHeight     =   795
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Sym 
      Height          =   732
      Index           =   3
      Left            =   3000
      Picture         =   "Toolbox.frx":5EEE
      ScaleHeight     =   675
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Sym 
      Height          =   492
      Index           =   11
      Left            =   4200
      Picture         =   "Toolbox.frx":6138
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Sym 
      Height          =   612
      Index           =   19
      Left            =   2124
      Picture         =   "Toolbox.frx":6382
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   5256
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Sym 
      Height          =   612
      Index           =   20
      Left            =   2844
      Picture         =   "Toolbox.frx":65CC
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   5256
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Sym 
      Height          =   612
      Index           =   22
      Left            =   4524
      Picture         =   "Toolbox.frx":6816
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   5256
      Visible         =   0   'False
      Width           =   612
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   3840
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Toolbox.frx":6A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Toolbox.frx":6D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Toolbox.frx":7094
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Toolbox.frx":8466
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Toolbox.frx":98B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Toolbox.frx":AD0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Toolbox.frx":C25C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Toolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dx, dy, scroll, finalsel, obpx, obpy, Csel, otw, oth
Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
For i = 0 To 9
 box(i).Width = CELL_SPACING
 box(i).Height = CELL_SPACING
 box(i).BorderStyle = 0
Next
Csel = 1
fill_holders
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MouseIcon = box(0).MouseIcon
Me.MousePointer = 99
End Sub
Private Sub Form_Resize()
wt1 = oth - tbToolBar.Height > CELL_SPACING
wt2 = Me.Height - tbToolBar.Height > CELL_SPACING
If wt1 Xor wt2 Then fill_holders
yh = CELL_SPACING + tbToolBar.Height
If Me.Height < yh Or Me.Width < CELL_SPACING * 6 Then
 Me.Height = yh
 Me.Width = CELL_SPACING * 6
End If
oth = Me.Height
otw = Me.Width
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case Is < 6
  If Csel <> Button.Index Then
   For i = 1 To 5
    tbToolBar.Buttons(i).Value = 0
   Next
   Csel = Button.Index
   tbToolBar.Buttons(Csel).Value = 1
   fill_holders
  End If
Case Else
  scroller (Button.Index)
End Select
End Sub
Public Sub scroller(dir)
wiretest = Csel = 1 _
 And Me.Height - tbToolBar.Height < CELL_SPACING
If wiretest = False Or scroll <> -999 Then
 If dir = 10 And box(9).Left + box(9).Width > Me.Width Then dir = -1
 If dir = 9 And box(0).Left < 0 Then dir = 1
 If dir < 2 Then
  scroll = -999
  For i = 0 To 9
   box(i).Left = box(i).Left + 150 * dir
   DoEvents
  Next
  scroll = 0
 End If
End If
End Sub
Private Sub tbToolBar_MouseDown(Bu4tton As Integer, Shift As Integer, X As Single, Y As Single)
dx = X: dy = Y
End Sub
Private Sub tbToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
If Button Then
If Abs(X - dx) > CELL_SPACING Or Abs(Y - dy) > CELL_SPACING Then X = dx: Y = dy
Toolbox.Left = Toolbox.Left + X - dx
Toolbox.Top = Toolbox.Top + Y - dy
End If
End Sub
Public Sub fill_holders()
wiretest = Csel = 1 _
 And Me.Height - tbToolBar.Height > CELL_SPACING + 400
Dim count(5) As Byte
count(1) = 10
count(2) = 2
count(3) = 4
count(4) = 2
count(5) = 5
For i = 1 To Csel - 1
 range = range + count(i)
Next
For i = 0 To 9
 PX = CELL_SPACING * i * 1.2
 PY = tbToolBar.Height
 If wiretest And i > 4 Then
  PX = CELL_SPACING * (i - 5) * 1.2
  PY = CELL_SPACING + 600
 End If
 box(i).Left = PX
 box(i).Top = PY
 If i < count(Csel) Then
  box(i).Picture = Sym(i + range).Picture
  box(i).Visible = True
 Else
  box(i).Visible = False
 End If
Next
End Sub
Private Sub box_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 dx = X: dy = Y
 obpx = box(Index).Left
 obpy = box(Index).Top
 DD.Left = Me.Left + obpx
 DD.Top = Me.Top + obpy
 DD.Show
 DD.Width = CELL_SPACING
 DD.Height = CELL_SPACING
 DD.Mobile.Picture = box(Index).Picture
 For finalsel = 0 To 22
 If box(Index).Picture = Sym(finalsel).Picture Then Exit For
 Next
End If
End Sub
Private Sub box_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim target As Object
box(Index).MouseIcon = box(0).MouseIcon
box(Index).MousePointer = 99
If Button = 1 Then
  box(Index).Left = box(Index).Left + X - dx
 box(Index).Top = box(Index).Top + Y - dy
 DD.Left = DD.Left + X - dx
 DD.Top = DD.Top + Y - dy
End If
End Sub
Private Sub box_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 DD.Hide
 X = X + DD.Left - MDI.Left
 Y = Y + DD.Top - MDI.Top - 1400
 If DD.Cform <> 0 Then
  canvas(DD.Cform).drop X, Y, Shift
 End If
 box(Index).Left = obpx
 box(Index).Top = obpy
 dx = 0: dy = 0
End If
End Sub

Private Sub Timer1_Timer()
If tbToolBar.Buttons(9).Value Then scroller 9
If tbToolBar.Buttons(10).Value Then scroller 10
End Sub
