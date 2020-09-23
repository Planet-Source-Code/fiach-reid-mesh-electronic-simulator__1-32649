VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5055
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Tag             =   "Apply"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3828.617
      ScaleMode       =   0  'User
      ScaleWidth      =   5733.383
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2029
         Left            =   506
         TabIndex        =   8
         Tag             =   "Sample 4"
         Top             =   504
         Width           =   2039
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3828.617
      ScaleMode       =   0  'User
      ScaleWidth      =   5733.383
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.CheckBox bmathd 
         Caption         =   "Math Documentation"
         Height          =   372
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3372
      End
      Begin VB.CheckBox bcns 
         Caption         =   "Complex Number support"
         Height          =   372
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3252
      End
      Begin VB.CheckBox bccl 
         Caption         =   "Cross cancelling"
         Height          =   372
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Value           =   1  'Checked
         Width           =   3612
      End
      Begin VB.CheckBox buisce 
         Caption         =   "Iterative Semiconductor Analysis"
         Height          =   372
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   1  'Checked
         Width           =   3492
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   192
      ScaleHeight     =   3828.617
      ScaleMode       =   0  'User
      ScaleWidth      =   5733.383
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   5685
      Begin VB.ComboBox labeling 
         Height          =   288
         ItemData        =   "Options.frx":0000
         Left            =   3708
         List            =   "Options.frx":0010
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   3240
         Width           =   1812
      End
      Begin ComctlLib.Slider Slider 
         Height          =   492
         Left            =   3600
         TabIndex        =   13
         Top             =   2400
         Width           =   1932
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   10
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.OptionButton dgraph 
         Caption         =   "Line"
         Height          =   252
         Index           =   1
         Left            =   3720
         TabIndex        =   12
         Top             =   840
         Width           =   1692
      End
      Begin VB.Frame Frame1 
         Caption         =   "Graph Styles"
         Height          =   1092
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1932
         Begin VB.OptionButton dgraph 
            Caption         =   "Plot"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1692
         End
      End
      Begin VB.Frame fraSample1 
         Caption         =   "Print preview"
         Height          =   3348
         Left            =   208
         TabIndex        =   4
         Tag             =   "Sample 1"
         Top             =   120
         Width           =   3120
         Begin VB.PictureBox arena 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1332
            Left            =   120
            ScaleHeight     =   1305
            ScaleWidth      =   2865
            TabIndex        =   21
            Top             =   240
            Width           =   2892
         End
         Begin VB.PictureBox page 
            BackColor       =   &H00FFFFFF&
            Height          =   3012
            Left            =   120
            ScaleHeight     =   2955
            ScaleWidth      =   2835
            TabIndex        =   9
            Top             =   240
            Width           =   2892
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Labeling:"
         Height          =   252
         Left            =   3720
         TabIndex        =   15
         Top             =   3000
         Width           =   1332
      End
      Begin VB.Label Perc 
         Caption         =   "Graphic Area : 50%"
         Height          =   252
         Left            =   3840
         TabIndex        =   14
         Top             =   2160
         Width           =   1572
      End
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Printer Setup"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Maths engine"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dx, dy, uisce, ccl, cns, mathd, cstyle, gstyle, Dstyle, Atop, Aleft, Asize, Grid, AW
Private Sub arena_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dx = X: dy = Y
End Sub
Private Sub arena_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then
 ax = arena.Left + arena.Width + X - dx
 ay = arena.Top + arena.Height + Y - dy
 PX = page.Left + page.Width
 PY = page.Top + page.Height
 If X > dx And ax > PX Then dx = X
 If X < dx And arena.Left + X - dx < page.Left Then dx = X
 If Y > dy And ay > PY Then dy = Y
 If Y < dy And arena.Top + Y - dy < page.Top Then dy = Y
 arena.Left = arena.Left + X - dx
 arena.Top = arena.Top + Y - dy
End If
End Sub

Private Sub cmdApply_Click()
If dgraph(0).Value Then gstyle = 1 Else gstyle = 2
Atop = (page.Left - arena.Left) * 100 / page.Width
Aleft = (page.Top - arena.Top) * 100 / page.Height
Asize = Slider.Value
uisce = buisce.Value
ccl = bccl.Value
cns = bcns.Value
mathd = bmathd.Value
Grid = BGrid.Value
MDI.mnuGrid.Checked = Grid
AW = BAW.Value
Dstyle = BDDM.ListIndex + 1
cstyle = labeling.ListIndex + 1
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub dgraph_Click(Index As Integer)
If Index = 0 Then dgraph(0).Value = False _
 Else dgraph(1).Value = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Form_Load()
labeling.ListIndex = 0

End Sub

Private Sub Slider_Scroll()
arena.Left = page.Left
arena.Top = page.Top
arena.Width = page.Width * Slider.Value / 100
arena.Height = page.Height * Slider.Value / 200
Perc.Caption = "Graphic Area :" + Str(Slider.Value) + "%"
End Sub
Private Sub tbsOptions_Click()
   Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
End Sub
