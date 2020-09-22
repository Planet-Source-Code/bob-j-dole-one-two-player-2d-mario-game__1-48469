VERSION 5.00
Begin VB.Form frmControls 
   Caption         =   "CHANGE YOUR CONTROLS!"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "frmControls.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Luigi"
      Height          =   5415
      Left            =   4440
      TabIndex        =   11
      Top             =   240
      Width           =   3735
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   1680
         TabIndex        =   21
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtljump 
         Height          =   615
         Left            =   1680
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtlleft 
         Height          =   615
         Left            =   1680
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtlright 
         Height          =   615
         Left            =   1680
         TabIndex        =   17
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   1680
         TabIndex        =   16
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Fireball"
         Height          =   495
         Left            =   480
         TabIndex        =   20
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Down"
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Right"
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Left"
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Jump"
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   360
         Picture         =   "frmControls.frx":08CA
         Stretch         =   -1  'True
         Top             =   600
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mario"
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.TextBox txtmleft 
         Height          =   615
         Left            =   1800
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtmright 
         Height          =   615
         Left            =   1800
         TabIndex        =   9
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtmdown 
         Height          =   615
         Left            =   1800
         TabIndex        =   8
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtmfire 
         Height          =   615
         Left            =   1800
         TabIndex        =   7
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtmjump 
         Height          =   615
         Left            =   1800
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Down"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fireball"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Right"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Left"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Jump"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   1680
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   360
         Picture         =   "frmControls.frx":18E8
         Stretch         =   -1  'True
         Top             =   600
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

txtmleft.Text = KeyReturn(mariokey.left)
txtmright.Text = KeyReturn(mariokey.right)
txtmjump.Text = KeyReturn(mariokey.Up)
txtmdown.Text = KeyReturn(mariokey.Down)
txtmfire.Text = KeyReturn(mariokey.fireball)

txtlleft.Text = KeyReturn(luigikey.left)
txtlright.Text = KeyReturn(luigikey.right)
txtljump.Text = KeyReturn(luigikey.Up)
End Sub





Public Function KeyReturn(Keycode As Long)
'Similar to above, as a new Keyboard button is being remapped,
    'it will be displayed to the user the selected button
    Select Case Keycode
    Case 8:
        KeyReturn = "Backspace"
    Case 13:
        KeyReturn = "Enter"
    Case 16:
        KeyReturn = "Shift"
    Case 17:
        KeyReturn = "Ctrl"
    Case 18:
        KeyReturn = "Alt"
    Case 19:
        KeyReturn = "Pause"
    Case 20:
        KeyReturn = "Caps Lock"
    Case 32:
        KeyReturn = "Spacebar"
    Case 33:
        KeyReturn = "Page Up"
    Case 34:
        KeyReturn = "Page Down"
    Case 35:
        KeyReturn = "End"
    Case 36:
        KeyReturn = "Home"
    Case 37:
        KeyReturn = "Left"
    Case 38:
        KeyReturn = "Up"
    Case 39:
        KeyReturn = "Right"
    Case 40:
        KeyReturn = "Down"
    Case 45:
        KeyReturn = "Insert"
    Case 46:
        KeyReturn = "Delete"
    Case 48:
        KeyReturn = "0"
    Case 49:
        KeyReturn = "1"
    Case 50:
        KeyReturn = "2"
    Case 51:
        KeyReturn = "3"
    Case 52:
        KeyReturn = "4"
    Case 53:
        KeyReturn = "5"
    Case 54:
        KeyReturn = "6"
    Case 55:
        KeyReturn = "7"
    Case 56:
        KeyReturn = "8"
    Case 57:
        KeyReturn = "9"
    Case 65:
        KeyReturn = "A"
    Case 66:
        KeyReturn = "B"
    Case 67:
        KeyReturn = "C"
    Case 68:
        KeyReturn = "D"
    Case 69:
        KeyReturn = "E"
    Case 70:
        KeyReturn = "F"
    Case 71:
        KeyReturn = "G"
    Case 72:
        KeyReturn = "H"
    Case 73:
        KeyReturn = "I"
    Case 74:
        KeyReturn = "J"
    Case 75:
        KeyReturn = "K"
    Case 76:
        KeyReturn = "L"
    Case 77:
        KeyReturn = "M"
    Case 78:
        KeyReturn = "N"
    Case 79:
        KeyReturn = "O"
    Case 80:
        KeyReturn = "P"
    Case 81:
        KeyReturn = "Q"
    Case 82:
        KeyReturn = "R"
    Case 83:
        KeyReturn = "S"
    Case 84:
        KeyReturn = "T"
    Case 85:
        KeyReturn = "U"
    Case 86:
        KeyReturn = "V"
    Case 87:
        KeyReturn = "W"
    Case 88:
        KeyReturn = "X"
    Case 89:
        KeyReturn = "Y"
    Case 90:
        KeyReturn = "Z"
    Case 96:
        KeyReturn = "Numpad 0"
    Case 97:
        KeyReturn = "Numpad 1"
    Case 98:
        KeyReturn = "Numpad 2"
    Case 99:
        KeyReturn = "Numpad 3"
    Case 100:
        KeyReturn = "Numpad 4"
    Case 101:
        KeyReturn = "Numpad 5"
    Case 102:
        KeyReturn = "Numpad 6"
    Case 103:
        KeyReturn = "Numpad 7"
    Case 104:
        KeyReturn = "Numpad 8"
    Case 105:
        KeyReturn = "Numpad 9"
    Case 106:
        KeyReturn = "Numpad *"
    Case 107:
        KeyReturn = "Numpad +"
    Case 109:
        KeyReturn = "Numpad -"
    Case 110:
        KeyReturn = "Numpad ."
    Case 111:
        KeyReturn = "Numpad /"
    Case 112:
        KeyReturn = "F1"
    Case 113:
        KeyReturn = "F2"
    Case 114:
        KeyReturn = "F3"
    Case 115:
        KeyReturn = "F4"
    Case 116:
        KeyReturn = "F5"
    Case 117:
        KeyReturn = "F6"
    Case 118:
        KeyReturn = "F7"
    Case 119:
        KeyReturn = "F8"
    Case 120:
        KeyReturn = "F9"
    Case 121:
        KeyReturn = "F10"
    Case 144:
        KeyReturn = "Num Lock"
    Case 145:
        KeyReturn = "Scroll Lock"
    Case 186:
        KeyReturn = ";"
    Case 187:
        KeyReturn = "-"
    Case 188:
        KeyReturn = ","
    Case 189:
        KeyReturn = "="
    Case 190:
        KeyReturn = "."
    Case 191:
        KeyReturn = "/"
    Case 192:
        KeyReturn = "`"
    Case 219:
        KeyReturn = "["
    Case 220:
        KeyReturn = "\"
    Case 221:
        KeyReturn = "]"
    Case 222:
        KeyReturn = "'"
    Case Else:
        KeyReturn = "Undefined"
    End Select
End Function
