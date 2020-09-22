VERSION 5.00
Begin VB.Form YouDie 
   BackColor       =   &H000000FF&
   Caption         =   "You got killed!"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   Icon            =   "YouDie.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1080
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   240
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H000000FF&
      Caption         =   "&Continue"
      Height          =   495
      Left            =   5880
      MaskColor       =   &H8000000A&
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblDeath 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Death HAS COME!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   6375
   End
End
Attribute VB_Name = "YouDie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public colortick As Integer

Private Sub cmdContinue_Click()
Unload YouDie
End Sub

Private Sub Form_Load()
colortick = 5
End Sub

Private Sub Timer1_Timer()
If lblDeath.Visible = True Then
  lblDeath.Visible = False
Else
  lblDeath.Visible = True
End If
'If cmdContinue.Visible = True Then
'  cmdContinue.Visible = False
'Else
'  cmdContinue.Visible = True
'End If




End Sub

Private Sub Timer2_Timer()
colortick = colortick - 1
If colortick = 0 Then colortick = 5
Select Case colortick
  Case 1
    YouDie.BackColor = vbRed
  Case 2
    YouDie.BackColor = vbBlue
  Case 3
    YouDie.BackColor = vbGreen
  Case 4
    YouDie.BackColor = vbOrange
  Case 5
    YouDie.BackColor = vbYellow
End Select


End Sub
