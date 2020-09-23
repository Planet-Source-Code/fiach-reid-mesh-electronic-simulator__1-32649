VERSION 5.00
Begin VB.Form frmSetRange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Range"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Max:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblValue 
      Caption         =   "Min:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmSetRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wasCancelled
Private Sub cmdCancel_Click()
 wasCancelled = True
 Hide
End Sub
Private Sub cmdOK_Click()
 wasCancelled = False
 Hide
End Sub
