Attribute VB_Name = "Central"
Option Explicit

Type Control
  Up As Long
  Down As Long
  left As Long
  right As Long
  fireball As Long
End Type
Public mariokey As Control
Public luigikey As Control
Public doubleplay As Boolean
Public doublekill As Integer
Public MultiPlayer As Boolean
Public Const Full = 20
Public Const Used = 19
Public Const Drop = 18
Public Const emptie = 17
Public BumpedPow As Boolean
Public level As Integer
Public Jetpack As Boolean
Dim Cheatcode As String
Public alldead As Boolean


'function that starts when you open the program

Public Sub Main()
frmCharacter.Show vbModal
frmSpeed.Show vbModal
Randomize (time)
If MultiPlayer = True Then
  LuigiSet
End If
MarioSet
PowSet
BGSet
Cheatset
Backset
Screenset
ScoreSet
FireSet
coinSet
ControlSet
DeathLoad
Load Mainform
Mainform.Show
DoEvents
Mainform.Timer1.Enabled = True
SoundName = App.Path & "\Sound\start.wav"
Answer = sndPlaySound(SoundName, SND_ASYNC)
level = 1


End Sub
'sets the cheats for use
Private Sub Cheatset()
Cheatcode = ""
Jetpack = False



End Sub
'adds the letter to a variable, if completed turn on or off the cheat
Public Sub Checkcheats(Keycode As Integer)
If Keycode >= vbKeyA And Keycode <= vbKeyZ Then
      Cheatcode = Cheatcode & Chr(Keycode)
      If Cheatcode = "JET" Then
        If Jetpack = False Then
            Jetpack = True
            MsgBox "you turned the jetpack on", vbOKOnly + vbInformation
        Else
            Jetpack = False
            MsgBox "you turned the jetpack off", vbOKOnly + vbInformation
        End If
      Cheatcode = ""
      End If

End If
If Keycode = vbKeyEscape Then
    Cheatcode = ""
End If


End Sub

Public Sub PowSet()
'pow1 = 0
'pow2 = 30
'pow3 = 54


  Pow.position.x = 296.5
  Pow.position.y = 290

  Dim HTempDC As Long
  Dim HOldbmp As Long
  HTempDC = GetDC(0)
  'normal
  Set Pow.picsource = New StdPicture
  Set Pow.picsource = LoadPicture(App.Path & "\Pics\pow.bmp")
  Pow.hsourcedc = CreateCompatibleDC(HTempDC)
  HOldbmp = SelectObject(Pow.hsourcedc, Pow.picsource.Handle)

  'variables for animation
  Dim bmtemp As Bitmap
  Pow.celltop = 0
  Pow.cellcount = 3
  GetObject Pow.picsource.Handle, Len(bmtemp), bmtemp
  Pow.cellwidth = bmtemp.BmWidth
  Pow.cellheight = 31
  Pow.State = Full
BumpedPow = False
If Pow.State = Full Then
  Floors(9).bottom = Pow.position.y + 31
  Floors(9).ceiling = Pow.position.y + 31
ElseIf Pow.State = Used Then
  Floors(9).bottom = Pow.position.y + 24
  Floors(9).ceiling = Pow.position.y + 24
ElseIf Pow.State = Drop Then
  Floors(9).bottom = Pow.position.y + 17
  Floors(9).ceiling = Pow.position.y + 17
End If
  
  
  ReleaseDC 0, HTempDC
End Sub


Public Sub PowDraw()
Dim height As Long
If Pow.State = Full Then
  height = 31
ElseIf Pow.State = Used Then
  height = 24
ElseIf Pow.State = Drop Then
  height = 17
End If


BitBlt OffscreenDc, Pow.position.x, Pow.position.y, Pow.cellwidth, height, Pow.hsourcedc, 0, Pow.celltop, vbSrcCopy

End Sub

Public Function PowAnimate()
Dim powTop As Long

If Pow.State = Full Then
  powTop = 0
ElseIf Pow.State = Used Then
  powTop = 31
ElseIf Pow.State = Drop Then
  powTop = 54
ElseIf Pow.State = emptie Then
  
Else
  MsgBox "Error in pow classification"
End If

        

  
      
PowAnimate = powTop

End Function

Public Sub BumpAll()
If Pow.State <> emptie Then
    SoundName = App.Path & "\Sound\pow.wav"
    Answer = sndPlaySound(SoundName, SND_ASYNC)
    Select Case Pow.State
      Case Full
        Pow.State = Used
        Floors(9).bottom = Pow.position.y + 24
        Floors(9).ceiling = Pow.position.y + 24
      Case Used
        Pow.State = Drop
        Floors(9).bottom = Pow.position.y + 17
        Floors(9).ceiling = Pow.position.y + 17
      Case Drop
        Pow.State = emptie
    End Select
    Dim index As Integer
    For index = 0 To bgmax
      If Badguy(index).State = Normal Then
        Badguy(index).State = Flipped
        Badguy(index).speed.x = 0
        Badguy(index).speed.y = -20
      ElseIf Badguy(index).State = Flipped Then
        Badguy(index).State = Normal
        Dim r As Integer
        r = Int(Rnd * 2)
        If r = 0 Then
          Badguy(index).speed.x = 6
        Else
          Badguy(index).speed.x = -6
        End If
        Badguy(index).speed.y = -20
      End If
    Next index
    For index = 0 To coinmax
      If coin(index).State = Normal Then
        coin(index).State = Dying
        coin(index).speed.x = 0
        coin(index).speed.y = -35
      End If
    Next index
    Else

End If
  
End Sub

Public Sub PicFree()
Dim HOldbmp As Long
Dim index As Integer

Set Mainform.Picture = Nothing

Set mario.picsource = Nothing
SelectObject mario.hsourcedc, HOldbmp
DeleteDC mario.hsourcedc

Set mario.picmask = Nothing
SelectObject mario.hmaskdc, HOldbmp
DeleteDC mario.hmaskdc

Set luigi.picsource = Nothing
SelectObject luigi.hsourcedc, HOldbmp
DeleteDC luigi.hsourcedc

Set luigi.picmask = Nothing
SelectObject luigi.hmaskdc, HOldbmp
DeleteDC luigi.hmaskdc

For index = 0 To bgmax
  Set Badguy(index).picsource = Nothing
  SelectObject Badguy(index).hsourcedc, HOldbmp
  DeleteDC Badguy(index).hsourcedc
  
  Set Badguy(index).picmask = Nothing
  SelectObject Badguy(index).hmaskdc, HOldbmp
  DeleteDC Badguy(index).hmaskdc
Next index

Set Backdrop.picsource = Nothing
SelectObject Backdrop.hsourcedc, HOldbmp
DeleteDC Backdrop.hsourcedc



Offscreenbmp = 0
SelectObject OffscreenDc, HOldbmp
DeleteDC OffscreenDc

For index = 0 To firemax
    Set fireball(index).picsource = Nothing
    SelectObject fireball(index).hsourcedc, HOldbmp
    DeleteDC fireball(index).hsourcedc
    
    Set fireball(index).picmask = Nothing
    SelectObject fireball(index).hmaskdc, HOldbmp
    DeleteDC fireball(index).hmaskdc

Next index
For index = 0 To coinmax
    Set coin(index).picsource = Nothing
    SelectObject coin(index).hsourcedc, HOldbmp
    DeleteDC coin(index).hsourcedc
    
    Set coin(index).picmask = Nothing
    SelectObject coin(index).hmaskdc, HOldbmp
    DeleteDC coin(index).hmaskdc
Next index

Set Pow.picsource = Nothing
SelectObject Pow.hsourcedc, HOldbmp
DeleteDC Pow.hsourcedc

Set platform.picsource = Nothing
SelectObject platform.hsourcedc, HOldbmp
DeleteDC platform.hsourcedc

SoundName = "Stop " & App.Path & "\Sound\background.mid"
Error = mciSendString(SoundName, AnswerString, Len(AnswerString) - 1, 0&)
Answer = mciGetErrorString(Error, ErrorString, 255)


End Sub

Public Sub ControlSet()
mariokey.Up = vbKeyUp
mariokey.left = vbKeyLeft
mariokey.right = vbKeyRight
mariokey.Down = vbKeyDown
mariokey.fireball = vbKeyControl

luigikey.Up = vbKeyW
luigikey.left = vbKeyA
luigikey.right = vbKeyD
End Sub

Public Sub PauseGame()
Mainform.Timer1.Enabled = False
Mainform.Timer1.Enabled = False

End Sub

Public Sub Resumegame()
Mainform.Timer1.Enabled = True
Mainform.Timer1.Enabled = True
End Sub

Public Sub DeathLoad()

Dim HTempDC As Long
Dim HOldbmp As Long
HTempDC = GetDC(0)
'normal
Set Death.picsource = New StdPicture
Set Death.picsource = LoadPicture(App.Path & "\Pics\Death.bmp")
Death.hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Death.hsourcedc, Death.picsource.Handle)
'mask
Set Death.picmask = New StdPicture
Set Death.picmask = LoadPicture(App.Path & "\Pics\Deathmsk.bmp")
Death.hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Death.hmaskdc, Death.picmask.Handle)
'variables for animation
Dim bmtemp As Bitmap
Death.celltop = 0
Death.cellcount = 1
GetObject Death.picsource.Handle, Len(bmtemp), bmtemp
Death.cellwidth = bmtemp.BmWidth
Death.cellheight = bmtemp.BmHeight / Death.cellcount
If Death.cellwidth < Death.cellheight Then
    Death.radius = Death.cellwidth / 2
Else
    Death.radius = Death.cellheight / 2
End If



ReleaseDC 0, HTempDC



End Sub

Public Sub DeathDraw()
BitBlt OffscreenDc, Death.position.x, Death.position.y, Death.cellwidth, Death.cellheight, Death.hmaskdc, 0, Death.celltop, vbSrcAnd
BitBlt OffscreenDc, Death.position.x, Death.position.y, Death.cellwidth, Death.cellheight, Death.hsourcedc, 0, Death.celltop, vbSrcPaint

End Sub
