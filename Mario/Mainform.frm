VERSION 5.00
Begin VB.Form Mainform 
   BackColor       =   &H80000007&
   Caption         =   "Welcome to Mario"
   ClientHeight    =   6615
   ClientLeft      =   780
   ClientTop       =   1005
   ClientWidth     =   9600
   Icon            =   "Mainform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9600
   Begin VB.TextBox txtspdy 
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtspdx 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtposy 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtposx 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "fps"
      Top             =   6720
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   840
      Top             =   480
   End
   Begin VB.TextBox txtfps 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   6720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   480
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Option 
      Caption         =   "&Option"
      Begin VB.Menu Controls 
         Caption         =   "&Controls"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu ReadMe 
         Caption         =   "&Read Me"
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 Option Explicit

Private Sub Controls_Click()
PauseGame
frmControls.Show
End Sub

Private Sub Exit_Click()
End
End Sub

'changes speed values and check for possible cheats
Private Sub Form_KeyDown(Keycode As Integer, Shift As Integer)

Checkcheats Keycode
If mario.State <> Dying Then
  
    
  If Keycode = mariokey.right Then
    mario.speed.x = 12
    platformmov = False
  ElseIf Keycode = mariokey.left Then
    mario.speed.x = -12
    platformmov = False
  ElseIf Keycode = mariokey.Up And (Jetpack = True Or mario.onfloor = True) Then
    
    If Jetpack = True Then
        SoundName = App.Path & "\Sound\jet.wav"
        Answer = sndPlaySound(SoundName, SND_ASYNC)
    Else
      SoundName = App.Path & "\Sound\jump.wav"
      Answer = sndPlaySound(SoundName, SND_ASYNC)
    End If
    platformmov = False
    If mario.onfloor = True Then
      mario.speed.y = -mario.Startspeed.y
      mario.onfloor = False
    Else
      mario.speed.y = mario.speed.y - 2
    End If
    'mario.onfloor = False
    'mario.speed.y = -mario.Startspeed.y
  ElseIf Keycode = mariokey.Down Then
    platformmov = False
  ElseIf Keycode = mariokey.fireball Then
    Dim index As Integer
    For index = 0 To firemax
      If fired(index) = False Then
        TempPos(index) = mario.position.y
        fired(index) = True
        ScoreNum = ScoreNum - 50
        If mario.speed.x >= 0 Then
          fireball(index).speed.x = fireball(index).Startspeed.x
        ElseIf mario.speed.x < 0 Then
          fireball(index).speed.x = -fireball(index).Startspeed.x
        End If
        Exit For
      End If
      
    Next index
   
   
  End If
End If

If MultiPlayer = True Then
  If luigi.State <> Dying Then
  
      
    If Keycode = luigikey.right Then
      luigi.speed.x = 12
      platformmov = False
    ElseIf Keycode = luigikey.left Then
      luigi.speed.x = -12
      platformmov = False
    ElseIf Keycode = luigikey.Up And (Jetpack = True Or luigi.onfloor = True) Then
      If Jetpack = True Then
          SoundName = App.Path & "\Sound\jet.wav"
          Answer = sndPlaySound(SoundName, SND_ASYNC)
      Else
        SoundName = App.Path & "\Sound\jump.wav"
        Answer = sndPlaySound(SoundName, SND_ASYNC)
      End If
      platformmov = False
      If luigi.onfloor = True Then
        luigi.speed.y = -luigi.Startspeed.y
        luigi.onfloor = False
      Else
        luigi.speed.y = luigi.speed.y - 2
      End If
      'luigi.onfloor = False
      'mario.speed.y = -mario.Startspeed.y
    ElseIf Keycode = vbKeyDown Then
      platformmov = False
    'ElseIf Keycode = vbKeySpace Then
    '  Dim index As Integer
    '  For index = 0 To firemax
    '    If fired(index) = False Then
    '      TempPos(index) = mario.position.y
    '      fired(index) = True
    '      ScoreNum = ScoreNum - 50
    '      If mario.speed.x >= 0 Then
    '        fireball(index).speed.x = fireball(index).Startspeed.x
    '      ElseIf mario.speed.x < 0 Then
    '        fireball(index).speed.x = -fireball(index).Startspeed.x
    '      End If
    '      Exit For
    '    End If
        
    '  Next index
     
     
    End If
  End If
End If
End Sub
'stops mario
Private Sub Form_KeyUp(Keycode As Integer, Shift As Integer)
If Keycode = vbKeyRight Or Keycode = vbKeyLeft Then
    mario.speed.x = 0
End If
If Keycode = vbKeyD Or Keycode = vbKeyA Then
    luigi.speed.x = 0
End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
PicFree
End
End Sub

Private Sub New_Click()
Main
End Sub

'calls movement and drawings, draws the buffer
Private Sub Timer1_Timer()

SoundName = "Play " & App.Path & "\Sound\background.mid"
Error = mciSendString(SoundName, AnswerString, Len(AnswerString) - 1, 0&)
Answer = mciGetErrorString(Error, ErrorString, 255)

Dim firecenter(0 To firemax) As Point
If MultiPlayer = True Then
  Dim lcenter As Point
End If
Dim mCenter As Point
Dim bgCenter As Point
Dim collided As Boolean
Dim firecollided(0 To firemax) As Boolean
Dim index As Integer
Dim index2 As Integer
Dim B2Center As Point
collided = False
If BumpedPow = True Then
  BumpedPow = False
  BitBlt OffscreenDc, Backdrop.position.x, Backdrop.position.y, 800, 600, Backdrop.hsourcedc, 0, -5, vbSrcCopy
Else
  Backdraw
End If
'------------pow stuff----------------------------------------------
PowDraw
Pow.celltop = PowAnimate


'--------------------------------------------------------------------


'collisions
'checks if "C" is smaller than both radius combined
For index = 0 To bgmax
  Dim findex As Integer
  For findex = 0 To firemax
    If (Badguy(index).State = Normal Or Badguy(index).State = Flipped) And fired(findex) = True Then
      firecenter(findex).x = fireball(findex).position.x + fireball(findex).cellwidth / 2
      firecenter(findex).y = fireball(findex).position.y + fireball(findex).cellheight / 2
      bgCenter.x = Badguy(index).position.x + Badguy(index).cellwidth / 2
      bgCenter.y = Badguy(index).position.y + Badguy(index).cellheight / 2
      firecollided(findex) = collide(firecenter(findex), fireball(findex).radius, bgCenter, Badguy(index).radius)
      If firecollided(findex) = True Then
        Badguy(index).speed.y = -20
        Badguy(index).State = Dying
        firecollided(findex) = False
      End If
    End If
  Next findex
    
    
    
    
    
    '----------------------Checks to see if you kill a badguy-------------------
    If Badguy(index).State <> Dead And Badguy(index).State <> Waiting Then
        If mario.State = Normal Then
            mCenter.x = mario.position.x + mario.cellwidth / 2
            mCenter.y = mario.position.y + mario.cellheight / 2
            bgCenter.x = Badguy(index).position.x + Badguy(index).cellwidth / 2
            bgCenter.y = Badguy(index).position.y + Badguy(index).cellheight / 2
            collided = collide(mCenter, mario.radius, bgCenter, Badguy(index).radius)
            If collided = True Then
                Bgcollided Badguy(index), mario
                collided = False
            End If
        End If
        If MultiPlayer = True Then
          If luigi.State = Normal Then
              lcenter.x = luigi.position.x + luigi.cellwidth / 2
              lcenter.y = luigi.position.y + luigi.cellheight / 2
              collided = collide(lcenter, luigi.radius, bgCenter, Badguy(index).radius)
              If collided = True Then
                  Bgcollided Badguy(index), luigi
                  collided = False
              End If
          End If
        End If
        '--------------------------------------------------------
        
        'this checks if two badguys collide, if they do, make them go opposite directions
        For index2 = index + 1 To bgmax
            B2Center.x = Badguy(index2).position.x + Badguy(index2).cellwidth / 2
            B2Center.y = Badguy(index2).position.y + Badguy(index2).cellheight / 2
            collided = collide(bgCenter, Badguy(index).radius, B2Center, Badguy(index2).radius)
            If collided = True Then
              'Badguy(index).speed.x = -Badguy(index).speed.x
              'Badguy(index2).speed.x = -Badguy(index2).speed.x
              'collided = False
              If Badguy(index).position.x < Badguy(index2).position.x Then
                  Badguy(index).speed.x = -Badguy(index).speed.x
                  Badguy(index2).speed.x = Badguy(index2).speed.x
              Else
                  Badguy(index).speed.x = Badguy(index).speed.x
                  Badguy(index2).speed.x = -Badguy(index2).speed.x
              End If
          End If
        Next index2
        If Badguy(index).State = Flipped Then
          If Badguy(index).fliptime = 1 Then
            Badguy(index).State = Normal
            Dim r As Integer
            r = Int(Rnd * 2)
            If r = 0 Then
              Badguy(index).speed.x = 6
            Else
              Badguy(index).speed.x = -6
            End If
            Badguy(index).speed.y = -20
            Badguy(index).fliptime = 0
          ElseIf Badguy(index).fliptime > 1 Then
            Badguy(index).fliptime = Badguy(index).fliptime - 1
          End If
        End If
        BGMove Badguy(index)
        Badguy(index).celltop = BGAnimate(Badguy(index))
        BGdraw Badguy(index)
    End If
Next index
Bgcreate
DoEvents
MarioMove
mario.celltop = MarioAnimate
Mariodraw

If MultiPlayer = True Then
  LuigiMove
  luigi.celltop = LuigiAnimate
  LuigiDraw
End If

'use scores
ScoreGet

If mlive > -5 Then mlive = mlive - 1
If llive > -5 Then llive = llive - 1

'if he falls down far, the he dies, reset him
If mario.position.y > 480 Then
  'Death.position.x = mario.position.x - 90
  'Death.position.y = 200
  'deathtimer = 30
  'YouDie.Show vbModal
  mario.position.y = 0
  mario.speed.x = 0
  mario.speed.y = -20
  mario.position.x = 310
  mlive = 30
  mario.State = Normal
  platformmov = True
End If
If luigi.position.y > 480 Then
  'YouDie.Show vbModal
  luigi.position.y = 0
  luigi.speed.x = 0
  luigi.speed.y = -20
  luigi.position.x = 310
  llive = 30
  luigi.State = Normal
  platformmov = True
End If
'If deathtimer > 0 Then
'  DeathDraw
'  deathtimer = deathtimer - 1
'End If

'checks to see howmany badguys have died
Dim x As Integer
For x = 0 To bgmax
    If Badguy(x).State <> Dead Then
        alldead = False
        Exit For
    End If
    alldead = True
Next x
'if all dead, you beat the game
If alldead = True Then
    frmlevel.Show
End If

If platformmov = False Then
  platform.position.y = -10
  Floors(8).left = platform.position.x
  Floors(8).right = platform.position.x + platform.cellwidth
  Floors(8).top = platform.position.y
  Floors(8).bottom = platform.position.y + platform.cellheight
  Floors(8).ceiling = platform.position.y + platform.cellheight
End If
PlatfromMove
platformDraw
DoEvents
If mario.State = Normal Then
  ScoreNum = ScoreNum - 1
  If ScoreNum <= 0 Then
    MsgBox "You ran out of life (points)"
    End
  End If
End If
'--------------------------Fireball stuff-------------------------------
For index = 0 To firemax
  If fired(index) = True Then
    FireDraw (index)
  End If
  FireMove (index)
  fireball(index).celltop = FireballAnimate(index)
Next index
DoEvents
'------------------------Coins----------------------------------------
For index = 0 To coinmax
  If coin(index).State <> Dead And coin(index).State <> Waiting Then
    coin(index).celltop = coinanimate(coin(index))
    
    Coinmove coin(index)
    CoinDraw coin(index)
    Dim coincenter As Point
    Dim mariocenter As Point
    Dim luigicenter As Point
    Dim coincollided As Boolean
    
    coincenter.x = 0
    coincenter.y = 0
    mariocenter.x = 0
    mariocenter.y = 0
    coincollided = False
    
    coincenter.x = coin(index).position.x + coin(index).cellwidth / 2
    coincenter.y = coin(index).position.y + coin(index).cellheight / 2
    mariocenter.x = mario.position.x + mario.cellwidth / 2
    mariocenter.y = mario.position.y + mario.cellheight / 2
    luigicenter.x = luigi.position.x + luigi.cellwidth / 2
    luigicenter.y = luigi.position.y + luigi.cellheight / 2
    coincollided = collide(coincenter, coin(index).radius, mariocenter, mario.radius)
    If coincollided = True And coin(index).State = Normal And mario.State = Normal Then
      Debug.Print "Hit coin" & index
      SoundName = App.Path & "\Sound\money.wav"
      Answer = sndPlaySound(SoundName, SND_ASYNC)
      coin(index).State = Dying
      coin(index).speed.x = mario.speed.x
      coin(index).speed.y = -20
      ScoreNum = ScoreNum + 35
      coincollided = False
    End If
    coincollided = collide(coincenter, coin(index).radius, luigicenter, luigi.radius)
    If coincollided = True And coin(index).State = Normal And luigi.State = Normal Then
      SoundName = App.Path & "\Sound\money.wav"
      Answer = sndPlaySound(SoundName, SND_ASYNC)
      Debug.Print "Hit coin" & index
      coin(index).State = Dying
      coin(index).speed.x = luigi.speed.x
      coin(index).speed.y = -20
      ScoreNum = ScoreNum + 35
      coincollided = False
    End If
  End If
Next index
cncreate
'--------------------------------------------------------------------

fps = fps + 1
txtposx.Text = mario.position.x
txtposy.Text = mario.position.y
txtspdx.Text = mario.speed.x
txtspdy.Text = mario.speed.y
BitBlt Mainform.hdc, 0, 0, 640, 480, OffscreenDc, 0, 0, vbSrcCopy
'Label1.Refresh
If doubleplay = True Then
  SoundName = App.Path & "\Sound\doublekill.wav"
  Answer = sndPlaySound(SoundName, SND_ASYNC)
  doubleplay = False
End If


End Sub
'fps
Private Sub Timer2_Timer()
txtfps.Text = CStr(fps)
BumpedPow = False
fps = 0
End Sub


