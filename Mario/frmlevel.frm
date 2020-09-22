VERSION 5.00
Begin VB.Form frmlevel 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "U BEAT THE LEVEL!"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   -375
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "CLICK HERE TO EXIT.  WAIT TO CONTINUE."
      Height          =   1575
      Left            =   4080
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   360
   End
   Begin VB.Label lblWIN 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "YOU BEAT THE LEVEL!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   11295
   End
End
Attribute VB_Name = "frmlevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
SoundName = App.Path & "\Sound\woohoo.wav"
Answer = sndPlaySound(SoundName, SND_ASYNC)
'End
level = level + 1
BGSet
lblWIN.Width = frmlevel.Width - 20
lblWIN.Caption = "YOU BEAT THE LEVEL WITH " & ScoreNum & " Points!"
'Select Case ScoreNum
'  Case Is < 20
'    lblWIN.Caption = lblWIN.Caption & "You just made it!"
'  Case Is < 300
'    lblWIN.Caption = lblWIN.Caption & "Below average score."
'  Case Is < 600
'    lblWIN.Caption = lblWIN.Caption & "Not bad!"
'  Case Is < 1000
'    lblWIN.Caption = lblWIN.Caption & "THAT IS TALENT!"
'  Case Is < 1500
'    lblWIN.Caption = lblWIN.Caption & "You are a master!"
'  Case Is < 2000
'    lblWIN.Caption = lblWIN.Caption & "AWESOME!"
'  Case Else
'    lblWIN.Caption = lblWIN.Caption & "Hacker"
'End Select

End Sub

Private Sub Timer1_Timer()

frmlevel.Hide
Mainform.Show
If level = 3 Then
  End
End If



'End
End Sub

