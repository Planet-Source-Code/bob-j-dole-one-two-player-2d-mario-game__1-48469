VERSION 5.00
Begin VB.Form frmCharacter 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose the number of players"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4515
   Icon            =   "frmCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.OptionButton opttwo 
      BackColor       =   &H80000007&
      Caption         =   "Two Player"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optone 
      BackColor       =   &H80000007&
      Caption         =   "Single Player"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   3120
      Picture         =   "frmCharacter.frx":08CA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   2280
      Picture         =   "frmCharacter.frx":18E8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   600
      Picture         =   "frmCharacter.frx":2892
      Stretch         =   -1  'True
      Top             =   360
      Width           =   810
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdExit_Click()
End
End Sub

Private Sub CmdOk_Click()
If optone.Value = True Then
  MultiPlayer = False
ElseIf opttwo.Value = True Then
  MultiPlayer = True
End If

If MultiPlayer = False Then
  Debug.Print "there will be only one player"
Else
  Debug.Print "this will be a two player game"
End If
Unload Me
End Sub


