VERSION 5.00
Begin VB.Form frmSpeed 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CHOOSE YOUR SPEED!"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4680
   Icon            =   "frmSpeed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CONTINUE"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Text            =   "ENTER SPEED..."
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   $"frmSpeed.frx":08CA
      Height          =   1695
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "WHAT SPEED DO YOU LIKE?"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()

Mainform.Timer1.Interval = Text1.Text
Unload Me
End Sub


Private Sub Text1_Click()
Text1.Text = ""
End Sub

