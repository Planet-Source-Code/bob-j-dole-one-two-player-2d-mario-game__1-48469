VERSION 5.00
Begin VB.Form frmHighScore 
   Caption         =   "High Scores"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   Icon            =   "frmHighScore.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'End
End Sub

Private Sub Form_Load()
Dim strFileName As String
Dim strText As String
Dim FileHandle%
Dim strBuffer


strFileName = App.Path & "\HighScore.txt"
FileHandle% = FreeFile
Open strFileName For Input As #FileHandle%

MousePointer = vbHourglass
Do While Not EOF(FileHandle%)
  Line Input #FileHandle%, strBuffer
  strText = strText & strBuffer & vbCrLf
Loop

MousePointer = vbDefault
Close #FileHandle%
Text1.Text = strText
  
End Sub
