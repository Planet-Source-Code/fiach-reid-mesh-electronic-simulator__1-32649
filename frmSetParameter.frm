VERSION 5.00
Begin VB.Form frmSetParameter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Parameter"
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
   Begin VB.CommandButton Command1 
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
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblValue 
      Caption         =   "Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSetParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public selectedParameter

Private Sub cmdOK_Click()
  On Error GoTo ThrowException:
  ComponentIndex = canvas(DD.Cform).SelectedIndex
  componentValues = canvas(DD.Cform).values.list(ComponentIndex)
  ComponentType = canvas(DD.Cform).datalist.list(ComponentIndex)
    
  If Not canvas(DD.Cform).ActiveObject Is Nothing Then
   Dim components(200)
   Dim properties(200)
   For i = 0 To canvas(DD.Cform).datalist.ListCount - 1
    components(i) = canvas(DD.Cform).datalist.list(i)
    properties(i) = canvas(DD.Cform).values.list(i)
   Next
   canvas(DD.Cform).ActiveObject.Update components, properties, ComponentIndex, selectedParameter, txtValue.Text
  End If
    
  ReDim s(2) As String
  Select Case ComponentType
   Case 10, 11
    canvas(DD.Cform).values.list(ComponentIndex) = txtValue.Text
   Case 12 To 15
    If InStr(UCase(txtValue.Text), "J") > 0 Then
     canvas(DD.Cform).values.list(ComponentIndex) = txtValue.Text
    Else
     s(1) = "6.28318530718" ' 2pi
     s(2) = pref2.freq
     s(1) = mesh11.expand(s(), Xmult)
     s(2) = txtValue
     s(1) = mesh11.expand(s(), Xmult) ' assuming s(1) is fixed
     If ComponentType = 12 Or ComponentType = 13 Then
      canvas(DD.Cform).values.list(ComponentIndex) = _
       "-" + mesh11.STR2(1 / Val(s(1))) + "J"
     Else
      canvas(DD.Cform).values.list(ComponentIndex) = s(1) + "J"
     End If
    End If
Case 16, 17
    If selectedParameter = 1 Then
     voltage = MDI.parseValues(ComponentIndex, 2)
     resistance = txtValue.Text
    Else
     resistance = MDI.parseValues(ComponentIndex, 1)
     voltage = txtValue.Text
    End If
    canvas(DD.Cform).values.list(ComponentIndex) = voltage + "," + resistance
    If ComponentType = 17 And selectedParameter = 3 Then
     pref2.freq = pref2.SI_convert(txtValue.Text)
    End If
Case 18 To 21
    If selectedParameter = 1 Then
     resistance = MDI.parseValues(ComponentIndex, 2)
     voltage = txtValue.Text
    Else
     voltage = MDI.parseValues(ComponentIndex, 1)
     resistance = txtValue.Text
    End If
    canvas(DD.Cform).values.list(ComponentIndex) = _
     voltage + "," + resistance
Case 22
    Hoe = MDI.parseValues(ComponentIndex, 1)
    Hfe = MDI.parseValues(ComponentIndex, 2)
    Hie = MDI.parseValues(ComponentIndex, 3)
    Select Case selectedParameter
     Case 1
      Hoe = txtValue.Text
     Case 2
      Hfe = txtValue.Text
     Case 3
      Hie = txtValue.Text
    End Select
    canvas(DD.Cform).values.list(ComponentIndex) = _
     Hoe & "," & Hfe & "," & Hie
 End Select
 Hide
 canvas(DD.Cform).ReNew
 Exit Sub
ThrowException:
 MsgBox Err.Description
 Hide
End Sub

Private Sub Command1_Click()
 Hide
End Sub
Private Sub Form_Activate()
 Me.Left = MDI.popupLeft
 Me.Top = MDI.popupTop
 Select Case selectedParameter
  Case 1
   Me.Caption = MDI.mnuParam1.Caption
  Case 2
   Me.Caption = MDI.mnuParam2.Caption
  Case 3
   Me.Caption = MDI.mnuParam3.Caption
 End Select
   
End Sub
Private Sub txtValue_Change()
 
 cmdOK.Enabled = True
 ComponentIndex = canvas(DD.Cform).SelectedIndex
 
 ComponentValue = canvas(DD.Cform).values.list(ComponentIndex)
 ComponentType = canvas(DD.Cform).datalist.list(ComponentIndex)

 ' null
 If txtValue.Text = "" Then cmdOK.Enabled = False
 
 ' non-complex Variable for TRC
 If ComponentType > 11 And ComponentType < 14 Then
        If Val(pref2.freq) = 0 Then cmdOK.Enabled = False
        If (pref2.SI_convert(txtValue.Text) = -999 And InStr(UCase(txtValue.Text), "J") = 0) Then
         cmdOK.Enabled = False
        End If
 End If
 
 ' multi-term BJT
 If ComponentType = 22 And (InStr(2, txtValue.Text, "+") Or InStr(2, txtValue.Text, "+")) Then
  cmdOK.Enabled = False
 End If
 
 ' Variable Frequency
  If ComponentType = 17 _
     And pref2.SI_convert(txtValue.Text) = -999 _
     And selectedParameter = 3 Then
  cmdOK.Enabled = False
 End If
 If cmdOK.Enabled And pref2.SI_convert(txtValue.Text) <> -999 Then
  txtValue.Text = pref2.SI_convert(txtValue.Text)
 End If
End Sub
