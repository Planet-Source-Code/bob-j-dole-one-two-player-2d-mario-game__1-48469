Attribute VB_Name = "background"
Option Explicit

Public Backdrop As Background
Public OffscreenDc As Long
Public Offscreenbmp As Long

Type Background
 position As Point
 picsource As StdPicture
 hsourcedc As Long
 leftside As Integer
 rightside As Integer
End Type

Type Floor
    top As Integer
    left As Integer
    right As Integer
    bottom As Integer
    ceiling As Integer
End Type

Type BackPic
 position As Point
 picsource As StdPicture
 hsourcedc As Long
 picmask As StdPicture
 hmaskdc As Long
 celltop As Integer
 cellheight As Integer
 cellwidth As Integer
 cellcount As Integer
End Type

Type scores
  deaths As Integer
  bumps As Integer
  kills As Integer
  time As Integer
  fireball As Integer
  total As Integer
End Type
Public marioscore As scores
Public luigiscore As scores
Public Score As BackPic
Public Floors(0 To 9) As Floor
Public ScoreNum As Long
Public BumpPoints As Integer
Public Points As Integer
Public ScorePos As Point


Public Sub Backset()

FloorSet
Backdrop.position.x = 0
Backdrop.position.y = 0
Dim HTempDC As Long
Dim HOldbmp As Long
HTempDC = GetDC(0)
Set Backdrop.picsource = New StdPicture
Set Backdrop.picsource = LoadPicture(App.Path & "\Pics\bckgrnd.bmp")
Backdrop.hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Backdrop.hsourcedc, Backdrop.picsource.Handle)
ReleaseDC 0, HTempDC
Backdrop.leftside = 0
Backdrop.rightside = 640



End Sub

Public Sub Backdraw()

BitBlt OffscreenDc, Backdrop.position.x, Backdrop.position.y, 800, 600, Backdrop.hsourcedc, 0, 0, vbSrcCopy

End Sub

Public Sub Screenset()

Dim HTempDC As Long
Dim HOldbmp As Long
HTempDC = GetDC(0)
OffscreenDc = CreateCompatibleDC(HTempDC)
Offscreenbmp = CreateCompatibleBitmap(HTempDC, 640, 480)
HOldbmp = SelectObject(OffscreenDc, Offscreenbmp)
ReleaseDC 0, HTempDC


End Sub


Private Sub FloorSet()
'initilizes floors
Floors(0).left = -50
Floors(0).right = 700
Floors(0).top = 415
Floors(0).bottom = 440
Floors(0).ceiling = Floors(0).top - 5

Floors(1).left = -50
Floors(1).right = 235
Floors(1).top = 320
Floors(1).bottom = 334
Floors(1).ceiling = Floors(1).top - 5

Floors(2).left = 406
Floors(2).right = 700
Floors(2).top = 320
Floors(2).bottom = 334
Floors(2).ceiling = Floors(2).top - 5

Floors(3).left = -50
Floors(3).right = 99
Floors(3).top = 218
Floors(3).bottom = 232
Floors(3).ceiling = Floors(3).top - 5

Floors(4).left = 548
Floors(4).right = 700
Floors(4).top = 218
Floors(4).bottom = 232
Floors(4).ceiling = Floors(4).top - 5

Floors(5).left = 166
Floors(5).right = 460
Floors(5).top = 204
Floors(5).bottom = 218
Floors(5).ceiling = Floors(5).top - 5

Floors(6).left = -50
Floors(6).right = 266
Floors(6).top = 96
Floors(6).bottom = 110
Floors(6).ceiling = Floors(6).top - 5

Floors(7).left = 380
Floors(7).right = 700
Floors(7).top = 96
Floors(7).bottom = 110
Floors(7).ceiling = Floors(7).top - 5

platform.State = Normal
platform.position.x = 310
platform.position.y = 0

Dim HTempDC As Long
Dim HOldbmp As Long
HTempDC = GetDC(0)
'normal
Set platform.picsource = New StdPicture
Set platform.picsource = LoadPicture(App.Path & "\Pics\platform.bmp")
platform.hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(platform.hsourcedc, platform.picsource.Handle)

'variables for animation
Dim bmtemp As Bitmap
platform.celltop = 0
platform.cellcount = 1
GetObject platform.picsource.Handle, Len(bmtemp), bmtemp
platform.cellwidth = bmtemp.BmWidth
platform.cellheight = bmtemp.BmHeight / platform.cellcount

Floors(8).left = platform.position.x
Floors(8).right = platform.position.x + platform.cellwidth
Floors(8).top = platform.position.y
Floors(8).bottom = platform.position.y + platform.cellheight
Floors(8).ceiling = platform.position.y + platform.cellheight


ReleaseDC 0, HTempDC

Floors(9).left = Pow.position.x
Floors(9).right = Pow.position.x + Pow.cellwidth
Floors(9).top = Pow.position.y
'Floors(9).bottom = Pow.position.y + Pow.cellheight
'Floors(9).ceiling = Pow.position.y + Pow.cellheight




End Sub

Public Sub ScoreSet()

ScorePos.x = 200
ScorePos.y = 417
Score.position = ScorePos

Dim HTempDC As Long
Dim HOldbmp As Long
HTempDC = GetDC(0)
'normal
Set Score.picsource = New StdPicture
Set Score.picsource = LoadPicture(App.Path & "\Pics\numbers.bmp")
Score.hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Score.hsourcedc, Score.picsource.Handle)
'mask
Set Score.picmask = New StdPicture
Set Score.picmask = LoadPicture(App.Path & "\Pics\msknum.bmp")
Score.hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Score.hmaskdc, Score.picmask.Handle)
'variables for animation
Dim bmtemp As Bitmap
Score.celltop = 0
Score.cellcount = 10
GetObject Score.picsource.Handle, Len(bmtemp), bmtemp
Score.cellwidth = bmtemp.BmWidth
Score.cellheight = bmtemp.BmHeight / Score.cellcount

ReleaseDC 0, HTempDC
ScoreNum = 2000
'killing
Points = 150
'bumping
BumpPoints = 60
End Sub

Public Sub ScoreDraw()
Dim scoretop As Long
Dim scoreleft As Long
Dim scoreacross As Long
Dim scoredown As Long
scoretop = Score.position.y
scoreleft = Score.position.x
scoreacross = Score.cellwidth
scoredown = Score.cellheight

BitBlt OffscreenDc, scoreleft, scoretop, scoreacross, scoredown, Score.hmaskdc, 0, Score.celltop, vbSrcAnd
BitBlt OffscreenDc, scoreleft, scoretop, scoreacross, scoredown, Score.hsourcedc, 0, Score.celltop, vbSrcPaint
End Sub

Public Sub ScoreGet()
Dim Digit As Integer
Dim tscore As Long


If ScoreNum = 0 Then
Score.celltop = 0
ScoreDraw
Else
tscore = ScoreNum

Do While tscore > 0
Digit = tscore Mod 10

tscore = Int(tscore / 10)

Score.celltop = Digit * Score.cellheight
Score.position.x = Score.position.x - Score.cellwidth
ScoreDraw
Loop
End If
Score.position = ScorePos
End Sub


Public Sub platformDraw()
If platformmov = True Then
  BitBlt OffscreenDc, platform.position.x, platform.position.y, platform.cellwidth, platform.cellheight, platform.hsourcedc, 0, 0, vbSrcCopy
End If
End Sub


