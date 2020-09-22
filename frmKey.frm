VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKey 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Key Check Demo Form"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   HelpContextID   =   340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   1995
      TabIndex        =   7
      Top             =   855
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Hide typing"
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   2220
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2115
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   2115
      Width           =   1170
   End
   Begin VB.TextBox txtConfirm 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   1590
      Width           =   4320
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   435
      Width           =   4320
   End
   Begin VB.Label lblQuality 
      Alignment       =   1  'Right Justify
      Caption         =   "Key Quality"
      Height          =   225
      Left            =   210
      TabIndex        =   8
      Top             =   855
      Width           =   1695
   End
   Begin VB.Label lblConfirm 
      Caption         =   "Confirm the Key"
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1380
      Width           =   4320
   End
   Begin VB.Label lblCode 
      Caption         =   "Enter the Key"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   225
      Width           =   3900
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' >>>>>>>>>>>>>>> Key Quality Demo Form <<<<<<<<<<<<<<<<<

Private Sub Form_Activate()
'set form for new input
Me.cmdOK.Enabled = False
Me.txtCode.Text = ""
Me.txtConfirm.Text = ""
Me.txtCode.PasswordChar = "*"
Me.txtConfirm.PasswordChar = "*"
Me.Check1.Value = 1
Me.ProgressBar1.Value = 0
Me.txtCode.SetFocus
End Sub

Private Sub txtCode_Change()

' >>> THIS IS THE ACTUAL QUALITY CHECKING CODE <<<

' update progressbar's quality value
Me.ProgressBar1.Value = KeyQuality(Me.txtCode.Text)

' refuse poor keys
If KeyQuality(Me.txtCode.Text) > 0 Then
    Me.cmdOK.Enabled = True
    Else
    Me.cmdOK.Enabled = False
    End If
End Sub

Private Sub cmdOK_Click()
' validation of the key
' check if key and confirmation are identical
If Me.txtCode.Text <> Me.txtConfirm.Text Or Me.txtCode.Text = "" Then
    MsgBox "key and confirmation do not match.", vbCritical
    Me.txtCode.Text = ""
    Me.txtConfirm.Text = ""
    Me.txtCode.SetFocus
    Exit Sub
    End If
' Check key quality, in this case more than 20 percent is asked
If KeyQuality(Me.txtCode.Text) < 20 Then
    MsgBox "The key does not meet the requirements!", vbCritical
    Else
    '>>>> write here the code to apply your key <<<<
    MsgBox "Your key: " & Me.txtCode.Text & vbCrLf & vbCrLf & "Key Quality: " & Str(KeyQuality(Me.txtCode.Text)) & " percent", vbInformation
    End If
'me.Hide
'reset form
Me.cmdOK.Enabled = False
Me.txtCode.Text = ""
Me.txtConfirm.Text = ""
Me.txtCode.PasswordChar = "*"
Me.txtConfirm.PasswordChar = "*"
Me.Check1.Value = 1
Me.ProgressBar1.Value = 0
Me.txtCode.SetFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
'goto confirm box on ENTER
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.txtCode <> "" Then Me.txtConfirm.SetFocus
    End If
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
'goto OK on ENTER
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.txtConfirm <> "" And Me.cmdOK.Enabled = True Then cmdOK_Click
    End If
End Sub

Private Sub Check1_Click()
'show or hide password chars
If Me.Check1.Value = 1 Then
    Me.txtCode.PasswordChar = "*"
    Me.txtConfirm.PasswordChar = "*"
    Else
    Me.txtCode.PasswordChar = ""
    Me.txtConfirm.PasswordChar = ""
    End If
End Sub

Private Sub cmdCancel_Click()
'get out of here
Me.txtCode.Text = ""
Me.txtConfirm.Text = ""
Unload Me
End Sub

