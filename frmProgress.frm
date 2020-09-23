VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin ComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblProgress 
      Caption         =   "Analysing..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public stopAnalysis
Private Sub cmdCancel_Click()
 stopAnalysis = True
 Hide
End Sub

