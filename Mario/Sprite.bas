Attribute VB_Name = "Sprite"
Option Explicit
Dim chk As Integer
'horizontal and vertical positions
Type Point
    x As Integer
    y As Integer
End Type
    

'Anything that moves in the game
Type Sprite
 'placement
 position As Point
 'velocity
 speed As Point
 Startspeed As Point
 onfloor As Boolean
 picsource As StdPicture
 hsourcedc As Long
 picmask As StdPicture
 hmaskdc As Long
 celltop As Integer
 cellheight As Integer
 cellwidth As Integer
 cellcount As Integer
 radius As Long
 State As Integer
 fliptime As Integer
 Name As String
End Type

Public Death As Sprite
Public Pow As Sprite
Public Const firemax As Integer = 5
Public TempPos(0 To firemax) As Integer
Public fireball(0 To firemax) As Sprite

Public fired(0 To firemax) As Boolean

Public platformmov As Boolean
Public luigi As Sprite
Public mario As Sprite
Public platform As Sprite
Public Const bgmax = 10
Public Badguy(0 To bgmax) As Sprite
Const Gravity = 4
Const Floortop = 415
Public Const Normal = 0 'moving
Public Const Flipped = 1 'flipped
Public Const Dying = 2 'touched, flying in air
Public Const Dead = 3 ' dead
Public Const Waiting = 4 'turtle is about to go out
Public Const Flipstart = 180

Public fps As Long
Public bgNumber As Integer
Public bgticker As Integer
Public cnNumber As Integer
Public cnticker As Integer
Public mlive As Integer
Public llive As Integer
Public Const coinmax = 8
Public coin(0 To coinmax) As Sprite

'Adds velocity to position, change mario's position
Public Sub MarioMove()

GetSpeed mario
If mario.position.x > Backdrop.rightside Then
    mario.position.x = Backdrop.leftside
Else
    mario.position.x = mario.position.x + mario.speed.x
End If

If mario.position.x < Backdrop.leftside Then
    mario.position.x = Backdrop.rightside
Else
    mario.position.x = mario.position.x + mario.speed.x
End If

mario.position.y = mario.position.y + mario.speed.y

End Sub
'sets mario's pictures and speeds/positions
Public Sub MarioSet()
mario.Name = "mario"
mario.State = Normal
mario.position.x = 200
mario.position.y = 300
mario.speed.x = 0
mario.speed.y = 0
mario.Startspeed.x = 20
mario.Startspeed.y = 35
mario.onfloor = True
Dim HTempDC As Long
Dim HOldbmp As Long
HTempDC = GetDC(0)
'normal
Set mario.picsource = New StdPicture
Set mario.picsource = LoadPicture(App.Path & "\Pics\marios.bmp")
mario.hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(mario.hsourcedc, mario.picsource.Handle)
'mask
Set mario.picmask = New StdPicture
Set mario.picmask = LoadPicture(App.Path & "\Pics\mariomsk.bmp")
mario.hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(mario.hmaskdc, mario.picmask.Handle)
'variables for animation
Dim bmtemp As Bitmap
mario.celltop = 0
mario.cellcount = 14
GetObject mario.picsource.Handle, Len(bmtemp), bmtemp
mario.cellwidth = bmtemp.BmWidth
mario.cellheight = bmtemp.BmHeight / mario.cellcount
If mario.cellwidth < mario.cellheight Then
    mario.radius = mario.cellwidth / 2
Else
    mario.radius = mario.cellheight / 2
End If



ReleaseDC 0, HTempDC




End Sub
'draws mario to buffer
Public Sub Mariodraw()

BitBlt OffscreenDc, mario.position.x, mario.position.y, 38, 34, mario.hmaskdc, 0, mario.celltop, vbSrcAnd
BitBlt OffscreenDc, mario.position.x, mario.position.y, 38, 34, mario.hsourcedc, 0, mario.celltop, vbSrcPaint


End Sub
'animates mario according to speed and position (jumping?)
Public Function MarioAnimate()
If mario.State = Dying Then
  MarioAnimate = 13 * mario.cellheight
  Exit Function
End If
  
Dim MarioTop As Long
MarioTop = mario.celltop + mario.cellheight
If mario.speed.x > 0 Then
    If mario.onfloor = False Then
        MarioTop = 272
    Else
        
        If MarioTop >= 102 Then
            MarioTop = 0
        End If
    End If
ElseIf mario.speed.x < 0 Then
    If mario.onfloor = False Then
        MarioTop = 306
    Else
        If MarioTop >= 204 Or MarioTop < 102 Then
            MarioTop = 102
        End If
    End If
ElseIf Int(Rnd * 30) = 0 Then
    MarioTop = 238
Else
    MarioTop = 204
End If
MarioAnimate = MarioTop
End Function
'Detects collisions with the floor using (guy) so it can use it for many sprites
Private Sub GetSpeed(guy As Sprite)

Dim guyTop As Long
Dim stepdown As Integer
Dim stepup As Integer
Dim GuyBottom As Long
Dim guyLeft As Long
Dim guyRight As Long
Dim index As Integer

guy.speed.y = guy.speed.y + Gravity
GuyBottom = guy.position.y + guy.cellheight
guyLeft = guy.position.x
guyRight = guy.position.x + guy.cellwidth
guyTop = guy.position.y

stepdown = GuyBottom + guy.speed.y
stepup = guyTop + guy.speed.y
guy.onfloor = False

If guy.State <> Dying Then
    For index = 0 To 9
        If platformmov = False And index = 8 Then
            index = 9
        End If
        If Pow.State = emptie And index = 9 Then
          Exit For
        End If
        'if in the floor area
        If guyRight > Floors(index).left And guyLeft < Floors(index).right Then
            'If in the air and the speed is fast enough to go though (make the speed just enought to hit the top
            'Or if has enough jump to go high enough to pass the floor, set the speed just enough to hit the bottom
            If (guyTop >= Floors(index).ceiling And stepup < Floors(index).ceiling) Or (GuyBottom <= Floors(index).top And stepdown > Floors(index).top) Then
            
                If guy.speed.y > 0 Then
                    'landing
                    guy.speed.y = Floors(index).top - GuyBottom
                    guy.onfloor = True
                Else
                    'jumping
                    guy.speed.y = Floors(index).ceiling - guyTop
                    If index = 9 And BumpedPow = False Then
                      BumpedPow = True
                      Debug.Print "Bump"
                      
                      BumpAll
                    End If
                End If
            End If
        End If

    Next index
End If



 

End Sub
'sets the BG position, pictures etc.
Public Sub BGSet()

    Dim HOldbmp As Long
    Dim bmtemp As Bitmap
    Dim HTempDC As Long
    bgNumber = 0
    bgticker = 50
    Badguy(0).position.x = 0
    Badguy(0).position.y = 20
    Badguy(0).Startspeed.x = 10
    Badguy(0).Startspeed.y = 35
    Badguy(0).speed.x = 6
    Badguy(0).speed.y = 35
    Badguy(0).onfloor = True
    HTempDC = GetDC(0)
    'color Badguy(0) squares
    Set Badguy(0).picsource = New StdPicture
    Set Badguy(0).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
    Badguy(0).hsourcedc = CreateCompatibleDC(HTempDC)
    HOldbmp = SelectObject(Badguy(0).hsourcedc, Badguy(0).picsource.Handle)
        'The blacky shady squares
    Set Badguy(0).picmask = New StdPicture
    Set Badguy(0).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
    Badguy(0).hmaskdc = CreateCompatibleDC(HTempDC)
    HOldbmp = SelectObject(Badguy(0).hmaskdc, Badguy(0).picmask.Handle)
    Badguy(0).celltop = 0
    Badguy(0).cellcount = 24
    GetObject Badguy(0).picsource.Handle, Len(bmtemp), bmtemp
    Badguy(0).cellheight = bmtemp.BmHeight / Badguy(0).cellcount
    Badguy(0).cellwidth = bmtemp.BmWidth
    Badguy(0).radius = (Badguy(0).cellheight + Badguy(0).cellwidth) / 4
    Badguy(0).State = Waiting
    



    Badguy(1).position.x = 630
    Badguy(1).position.y = 20
    Badguy(1).Startspeed.x = 10
    Badguy(1).Startspeed.y = 35
    Badguy(1).speed.x = -6
    Badguy(1).speed.y = 35
    Badguy(1).onfloor = True
    HTempDC = GetDC(1)
    'color Badguy(1) squares
    Set Badguy(1).picsource = New StdPicture
    Set Badguy(1).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
    Badguy(1).hsourcedc = CreateCompatibleDC(HTempDC)
    HOldbmp = SelectObject(Badguy(1).hsourcedc, Badguy(1).picsource.Handle)
    'The blacky shady squares
    Set Badguy(1).picmask = New StdPicture
    Set Badguy(1).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
    Badguy(1).hmaskdc = CreateCompatibleDC(HTempDC)
    HOldbmp = SelectObject(Badguy(1).hmaskdc, Badguy(1).picmask.Handle)
    Badguy(1).celltop = 0
    Badguy(1).cellcount = 24
    GetObject Badguy(1).picsource.Handle, Len(bmtemp), bmtemp
    Badguy(1).cellheight = bmtemp.BmHeight / Badguy(1).cellcount
    Badguy(1).cellwidth = bmtemp.BmWidth
    Badguy(1).radius = (Badguy(1).cellheight + Badguy(1).cellwidth) / 4
    Badguy(1).State = Waiting
    
  '''''''''''''''''Initialize badguy(2)'s Position and Speed
Badguy(2).position.x = 0
Badguy(2).position.y = 20
Badguy(2).Startspeed.x = 10
Badguy(2).Startspeed.y = 35
Badguy(2).speed.x = 6
Badguy(2).speed.y = 2
Badguy(2).onfloor = True

''''''''''''''''''Load badguy(2)'s Picture
Set Badguy(2).picsource = New StdPicture
Set Badguy(2).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(2).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(2).hsourcedc, Badguy(2).picsource.Handle)

''''''''''''''''' Load badguy(2)'s Mask
Set Badguy(2).picmask = New StdPicture
Set Badguy(2).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(2).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(2).hmaskdc, Badguy(2).picmask.Handle)

'''''''''''''''' Initialize badguy(2)'s cell information
Badguy(2).celltop = 0
Badguy(2).cellcount = 24
GetObject Badguy(2).picsource.Handle, Len(bmtemp), bmtemp
Badguy(2).cellwidth = bmtemp.BmWidth
Badguy(2).cellheight = bmtemp.BmHeight / Badguy(2).cellcount

If Badguy(2).cellwidth < Badguy(2).cellheight Then
    Badguy(2).radius = Badguy(2).cellwidth / 2
Else
    Badguy(2).radius = Badguy(2).cellheight / 2
End If
Badguy(2).State = Waiting


'''''''''''''''''Initialize badguy(3)'s Position and Speed
Badguy(3).position.x = 630
Badguy(3).position.y = 20
Badguy(3).Startspeed.x = 10
Badguy(3).Startspeed.y = 35
Badguy(3).speed.x = -6
Badguy(3).speed.y = 2
Badguy(3).onfloor = True

''''''''''''''''''Load badguy(3)'s Picture
Set Badguy(3).picsource = New StdPicture
Set Badguy(3).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(3).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(3).hsourcedc, Badguy(3).picsource.Handle)

''''''''''''''''' Load badguy(3)'s Mask
Set Badguy(3).picmask = New StdPicture
Set Badguy(3).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(3).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(3).hmaskdc, Badguy(3).picmask.Handle)

'''''''''''''''' Initialize badguy(3)'s cell information
Badguy(3).celltop = 0
Badguy(3).cellcount = 24
GetObject Badguy(3).picsource.Handle, Len(bmtemp), bmtemp
Badguy(3).cellwidth = bmtemp.BmWidth
Badguy(3).cellheight = bmtemp.BmHeight / Badguy(3).cellcount

If Badguy(3).cellwidth < Badguy(3).cellheight Then
    Badguy(3).radius = Badguy(3).cellwidth / 2
Else
    Badguy(3).radius = Badguy(3).cellheight / 2
End If
Badguy(3).State = Waiting


'''''''''''''''''Initialize badguy(4)'s Position and Speed
Badguy(4).position.x = 0
Badguy(4).position.y = 20
Badguy(4).Startspeed.x = 10
Badguy(4).Startspeed.y = 35
Badguy(4).speed.x = 6
Badguy(4).speed.y = 2
Badguy(4).onfloor = True

''''''''''''''''''Load badguy(4)'s Picture
Set Badguy(4).picsource = New StdPicture
Set Badguy(4).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(4).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(4).hsourcedc, Badguy(4).picsource.Handle)

''''''''''''''''' Load badguy(4)'s Mask
Set Badguy(4).picmask = New StdPicture
Set Badguy(4).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(4).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(4).hmaskdc, Badguy(4).picmask.Handle)

'''''''''''''''' Initialize badguy(4)'s cell information
Badguy(4).celltop = 0
Badguy(4).cellcount = 24
GetObject Badguy(4).picsource.Handle, Len(bmtemp), bmtemp
Badguy(4).cellwidth = bmtemp.BmWidth
Badguy(4).cellheight = bmtemp.BmHeight / Badguy(4).cellcount

If Badguy(4).cellwidth < Badguy(4).cellheight Then
    Badguy(4).radius = Badguy(4).cellwidth / 2
Else
    Badguy(4).radius = Badguy(4).cellheight / 2
End If
Badguy(4).State = Waiting


'''''''''''''''''Initialize badguy(5)'s Position and Speed
Badguy(5).position.x = 0
Badguy(5).position.y = 20
Badguy(5).Startspeed.x = 10
Badguy(5).Startspeed.y = 35
Badguy(5).speed.x = 6
Badguy(5).speed.y = 2
Badguy(5).onfloor = True

''''''''''''''''''Load badguy(5)'s Picture
Set Badguy(5).picsource = New StdPicture
Set Badguy(5).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(5).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(5).hsourcedc, Badguy(5).picsource.Handle)

''''''''''''''''' Load badguy(5)'s Mask
Set Badguy(5).picmask = New StdPicture
Set Badguy(5).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(5).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(5).hmaskdc, Badguy(5).picmask.Handle)

'''''''''''''''' Initialize badguy(5)'s cell information
Badguy(5).celltop = 0
Badguy(5).cellcount = 24
GetObject Badguy(5).picsource.Handle, Len(bmtemp), bmtemp
Badguy(5).cellwidth = bmtemp.BmWidth
Badguy(5).cellheight = bmtemp.BmHeight / Badguy(5).cellcount

If Badguy(5).cellwidth < Badguy(5).cellheight Then
    Badguy(5).radius = Badguy(5).cellwidth / 2
Else
    Badguy(5).radius = Badguy(5).cellheight / 2
End If
Badguy(5).State = Waiting


'''''''''''''''''Initialize badguy(6)'s Position and Speed
Badguy(6).position.x = 630
Badguy(6).position.y = 20
Badguy(6).Startspeed.x = 10
Badguy(6).Startspeed.y = 35
Badguy(6).speed.x = -6
Badguy(6).speed.y = 2
Badguy(6).onfloor = True

''''''''''''''''''Load badguy(6)'s Picture
Set Badguy(6).picsource = New StdPicture
Set Badguy(6).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(6).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(6).hsourcedc, Badguy(6).picsource.Handle)

''''''''''''''''' Load badguy(6)'s Mask
Set Badguy(6).picmask = New StdPicture
Set Badguy(6).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(6).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(6).hmaskdc, Badguy(6).picmask.Handle)

'''''''''''''''' Initialize badguy(6)'s cell information
Badguy(6).celltop = 0
Badguy(6).cellcount = 24
GetObject Badguy(6).picsource.Handle, Len(bmtemp), bmtemp
Badguy(6).cellwidth = bmtemp.BmWidth
Badguy(6).cellheight = bmtemp.BmHeight / Badguy(6).cellcount

If Badguy(6).cellwidth < Badguy(6).cellheight Then
    Badguy(6).radius = Badguy(6).cellwidth / 2
Else
    Badguy(6).radius = Badguy(6).cellheight / 2
End If
Badguy(6).State = Waiting


'''''''''''''''''Initialize badguy(7)'s Position and Speed
Badguy(7).position.x = 630
Badguy(7).position.y = 20
Badguy(7).Startspeed.x = 10
Badguy(7).Startspeed.y = 35
Badguy(7).speed.x = -6
Badguy(7).speed.y = 2
Badguy(7).onfloor = True

''''''''''''''''''Load badguy(7)'s Picture
Set Badguy(7).picsource = New StdPicture
Set Badguy(7).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(7).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(7).hsourcedc, Badguy(7).picsource.Handle)

''''''''''''''''' Load badguy(7)'s Mask
Set Badguy(7).picmask = New StdPicture
Set Badguy(7).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(7).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(7).hmaskdc, Badguy(7).picmask.Handle)

'''''''''''''''' Initialize badguy(7)'s cell information
Badguy(7).celltop = 0
Badguy(7).cellcount = 24
GetObject Badguy(7).picsource.Handle, Len(bmtemp), bmtemp
Badguy(7).cellwidth = bmtemp.BmWidth
Badguy(7).cellheight = bmtemp.BmHeight / Badguy(7).cellcount

If Badguy(7).cellwidth < Badguy(7).cellheight Then
    Badguy(7).radius = Badguy(7).cellwidth / 2
Else
    Badguy(7).radius = Badguy(7).cellheight / 2
End If
Badguy(7).State = Waiting


'''''''''''''''''Initialize badguy(8)'s Position and Speed
Badguy(8).position.x = 0
Badguy(8).position.y = 20
Badguy(8).Startspeed.x = 10
Badguy(8).Startspeed.y = 35
Badguy(8).speed.x = 6
Badguy(8).speed.y = 2
Badguy(8).onfloor = True

''''''''''''''''''Load badguy(8)'s Picture
Set Badguy(8).picsource = New StdPicture
Set Badguy(8).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(8).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(8).hsourcedc, Badguy(8).picsource.Handle)

''''''''''''''''' Load badguy(8)'s Mask
Set Badguy(8).picmask = New StdPicture
Set Badguy(8).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(8).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(8).hmaskdc, Badguy(8).picmask.Handle)

'''''''''''''''' Initialize badguy(8)'s cell information
Badguy(8).celltop = 0
Badguy(8).cellcount = 24
GetObject Badguy(8).picsource.Handle, Len(bmtemp), bmtemp
Badguy(8).cellwidth = bmtemp.BmWidth
Badguy(8).cellheight = bmtemp.BmHeight / Badguy(8).cellcount

If Badguy(8).cellwidth < Badguy(8).cellheight Then
    Badguy(8).radius = Badguy(8).cellwidth / 2
Else
    Badguy(8).radius = Badguy(8).cellheight / 2
End If
Badguy(8).State = Waiting


'''''''''''''''''Initialize badguy(9)'s Position and Speed
Badguy(9).position.x = 630
Badguy(9).position.y = 20
Badguy(9).Startspeed.x = 10
Badguy(9).Startspeed.y = 35
Badguy(9).speed.x = -6
Badguy(9).speed.y = 2
Badguy(9).onfloor = True

''''''''''''''''''Load badguy(9)'s Picture
Set Badguy(9).picsource = New StdPicture
Set Badguy(9).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(9).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(9).hsourcedc, Badguy(9).picsource.Handle)

''''''''''''''''' Load badguy(9)'s Mask
Set Badguy(9).picmask = New StdPicture
Set Badguy(9).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(9).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(9).hmaskdc, Badguy(9).picmask.Handle)

'''''''''''''''' Initialize badguy(9)'s cell information
Badguy(9).celltop = 0
Badguy(9).cellcount = 24
GetObject Badguy(9).picsource.Handle, Len(bmtemp), bmtemp
Badguy(9).cellwidth = bmtemp.BmWidth
Badguy(9).cellheight = bmtemp.BmHeight / Badguy(9).cellcount

If Badguy(9).cellwidth < Badguy(9).cellheight Then
    Badguy(9).radius = Badguy(9).cellwidth / 2
Else
    Badguy(9).radius = Badguy(9).cellheight / 2
End If
Badguy(9).State = Waiting

'''''''''''''''''Initialize badguy(10)'s Position and Speed
Badguy(10).position.x = 630
Badguy(10).position.y = 20
Badguy(10).Startspeed.x = 10
Badguy(10).Startspeed.y = 35
Badguy(10).speed.x = -6
Badguy(10).speed.y = 2
Badguy(10).onfloor = True

''''''''''''''''''Load badguy(10)'s Picture
Set Badguy(10).picsource = New StdPicture
Set Badguy(10).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
Badguy(10).hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(10).hsourcedc, Badguy(10).picsource.Handle)

''''''''''''''''' Load badguy(10)'s Mask
Set Badguy(10).picmask = New StdPicture
Set Badguy(10).picmask = LoadPicture(App.Path + "\Pics\turtsmsk.bmp")
Badguy(10).hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(Badguy(10).hmaskdc, Badguy(10).picmask.Handle)

'''''''''''''''' Initialize badguy(10)'s cell information
Badguy(10).celltop = 0
Badguy(10).cellcount = 24
GetObject Badguy(10).picsource.Handle, Len(bmtemp), bmtemp
Badguy(10).cellwidth = bmtemp.BmWidth
Badguy(10).cellheight = bmtemp.BmHeight / Badguy(10).cellcount

If Badguy(10).cellwidth < Badguy(10).cellheight Then
    Badguy(10).radius = Badguy(10).cellwidth / 2
Else
    Badguy(10).radius = Badguy(10).cellheight / 2
End If
Badguy(10).State = Waiting
ReleaseDC 0, HTempDC

If level = 2 Then
  Dim index As Integer
  For index = 0 To bgmax
    ''''''''''''''''''Load badguy(10)'s Picture
    Set Badguy(index).picsource = New StdPicture
    Set Badguy(index).picsource = LoadPicture(App.Path + "\Pics\crabsspr.bmp")
    Badguy(index).hsourcedc = CreateCompatibleDC(HTempDC)
    HOldbmp = SelectObject(Badguy(index).hsourcedc, Badguy(index).picsource.Handle)
    
    ''''''''''''''''' Load badguy(index)'s Mask
    Set Badguy(index).picmask = New StdPicture
    Set Badguy(index).picmask = LoadPicture(App.Path + "\Pics\crabsmsk.bmp")
    Badguy(index).hmaskdc = CreateCompatibleDC(HTempDC)
    HOldbmp = SelectObject(Badguy(index).hmaskdc, Badguy(index).picmask.Handle)
  Next index
ElseIf level = 3 Then
  End
End If
End Sub


'draws the bad guy to the buffer
Public Sub BGdraw(bg As Sprite)
Dim BGTop As Long
Dim BGLeft As Long
Dim BGAcross As Long
Dim BGDown As Long

BGTop = bg.position.y
BGLeft = bg.position.x
BGAcross = bg.cellwidth
BGDown = bg.cellheight

    
BitBlt OffscreenDc, bg.position.x, bg.position.y, 38, 34, bg.hmaskdc, 0, bg.celltop, vbSrcAnd
BitBlt OffscreenDc, bg.position.x, bg.position.y, 38, 34, bg.hsourcedc, 0, bg.celltop, vbSrcPaint

End Sub

Public Sub BGMove(bg As Sprite)
'moves the badguy
'if he goes of the screen at the bottom the reappear at top going the opposite direction
Dim BgBottom As Integer
BgBottom = bg.position.y + bg.cellheight

GetSpeed bg
If bg.position.x > Backdrop.rightside Then
    If BgBottom = Floors(0).top Then
        bg.position.y = 10
        bg.speed.x = -bg.speed.x
    Else
        bg.position.x = Backdrop.leftside
    End If
Else
    bg.position.x = bg.position.x + bg.speed.x
End If

If bg.position.x < Backdrop.leftside Then
    If BgBottom = Floors(0).top Then
        bg.position.y = 10
        bg.speed.x = -bg.speed.x
    Else
        bg.position.x = Backdrop.rightside
    End If
Else
    bg.position.x = bg.position.x + bg.speed.x
End If

bg.position.y = bg.position.y + bg.speed.y

If bg.position.y > 480 Then
    bg.State = Dead
End If

End Sub
'animates the badguy
Public Function BGAnimate(bg As Sprite)
Dim BGTop As Long
BGTop = bg.celltop + bg.cellheight
Select Case bg.State
    Case Normal
        If bg.speed.x > 0 Then
            If BGTop > 36 Then
                BGTop = 0
            End If
        Else
            If bg.speed.x < 0 Then
                If BGTop > 108 Or BGTop < 36 Then
                    BGTop = 72
                End If
            End If
   
        End If
    Case Flipped
        If bg.speed.y <> 0 Then
            If BGTop > 324 Or BGTop < 288 Then
                BGTop = 288
            End If
        Else
            If BGTop > 180 Then
                BGTop = 144
            End If
        End If
    Case Dying
        If BGTop > 324 Or BGTop < 288 Then
            BGTop = 288
        End If
        
        
End Select
BGAnimate = BGTop

End Function
'checks for collision using the center and radius (Pythrogum theroem)
Public Function collide(mCenter As Point, mRadius As Long, bgCenter As Point, bgRadius As Long)
Dim A As Long
Dim B As Long
Dim C As Long

A = mCenter.x - bgCenter.x
B = mCenter.y - bgCenter.y
C = Sqr(A * A + B * B)

If C < mRadius + bgRadius Then
    collide = True
Else
    collide = False
End If



End Function
'allows guy to bump the badguy, and many other combonations
Public Sub Bgcollided(bg As Sprite, guy As Sprite)
Dim bump As Boolean
Dim index As Integer
Dim guyTop As Integer
Dim BgBottom As Integer
Dim guyLeft As Integer
Dim guyRight As Integer
Const floormax = 8

'initilization
guyRight = guy.position.x + guy.cellwidth
guyLeft = guy.position.x
guyTop = guy.position.y
BgBottom = bg.position.y + bg.cellheight
bump = False

index = 0


Do While index <= floormax
    If guyTop >= Floors(index).ceiling And BgBottom <= Floors(index).top Then
        If guyRight > Floors(index).left And guyLeft < Floors(index).right Then
            bump = True
        End If
    End If
    index = index + 1
Loop

If bump = True And bg.State = Normal Then
    SoundName = App.Path & "\Sound\bgBump.wav"
    Answer = sndPlaySound(SoundName, SND_ASYNC)
    bg.State = Flipped
    bg.speed.x = 0
    bg.speed.y = -bg.Startspeed.y
    ScoreNum = ScoreNum + BumpPoints
    bg.fliptime = Flipstart
ElseIf bump = True And bg.State = Flipped Then
    bg.State = Normal
    Dim r As Integer
    r = Int(Rnd * 2)
    If r = 0 Then
      bg.speed.x = 6
    Else
      bg.speed.x = -6
    End If
    bg.speed.y = -20
    'bg.speed.x = bg.Startspeed.x
Else
    Dim live As Integer
    If guy.Name = "luigi" Then
      live = llive
    ElseIf guy.Name = "mario" Then
      live = mlive
    End If
    If bg.State = Normal And live < 0 And guy.State <> Dying Then
        'YouDie.Show vbModal
        'guy.position.y = 0
        'guy.speed.x = 0
        'guy.speed.y = 0
        'guy.position.x = 310
        'mlive = 30
        guy.speed.x = 0
        If guy.speed.y >= 0 Then
          guy.speed.y = -25
        Else
          guy.speed.y = guy.speed.y * 1.5
        End If
        guy.State = Dying
        ScoreNum = ScoreNum - 200
        
        chk = Int(Rnd * 3)
        If chk = 0 Then
          SoundName = App.Path & "\Sound\fallscream.wav"
        ElseIf chk = 1 Then
          SoundName = App.Path & "\Sound\ahh.wav"
        ElseIf chk = 2 Then
          SoundName = App.Path & "\Sound\makill.wav"
        End If
        Answer = sndPlaySound(SoundName, SND_ASYNC)
    Else
        If bg.State = Flipped Then
            If doublekill < 0 Then
              doublekill = 10
            Else
              doubleplay = True
            End If
            bg.State = Dying
            bg.speed.y = -bg.Startspeed.y
            bg.speed.x = guy.speed.x
            ScoreNum = ScoreNum + Points
        End If
    End If
    

End If

End Sub

'Ceates badguys until the limit
Public Sub Bgcreate()
bgticker = bgticker - 1
If bgticker = 0 And bgNumber <= bgmax Then
    Badguy(bgNumber).State = Normal
    bgticker = 30
    bgNumber = bgNumber + 1
End If

End Sub
'moves the platform, when dead
Public Sub PlatfromMove()
If platformmov = True And platform.position.y < 50 Then
  platform.position.y = platform.position.y + 2
  Floors(8).left = platform.position.x
  Floors(8).right = platform.position.x + platform.cellwidth
  Floors(8).top = platform.position.y
  Floors(8).bottom = platform.position.y + platform.cellheight
  Floors(8).ceiling = platform.position.y + platform.cellheight
End If
End Sub
'Sets up the fireballs
Public Sub FireSet()
Dim index As Integer

For index = 0 To firemax
  fireball(index).position.x = mario.position.x
  fireball(index).position.y = mario.position.y + mario.cellheight
  fireball(index).speed.x = 0
  fireball(index).speed.y = 0
  fireball(index).Startspeed.x = 15
  fireball(index).Startspeed.y = 55

  Dim HTempDC As Long
  Dim HOldbmp As Long
  HTempDC = GetDC(0)
  'normal
  Set fireball(index).picsource = New StdPicture
  Set fireball(index).picsource = LoadPicture(App.Path & "\Pics\firespr.bmp")
  fireball(index).hsourcedc = CreateCompatibleDC(HTempDC)
  HOldbmp = SelectObject(fireball(index).hsourcedc, fireball(index).picsource.Handle)
  'mask
  Set fireball(index).picmask = New StdPicture
  Set fireball(index).picmask = LoadPicture(App.Path & "\Pics\firemsk.bmp")
  fireball(index).hmaskdc = CreateCompatibleDC(HTempDC)
  HOldbmp = SelectObject(fireball(index).hmaskdc, fireball(index).picmask.Handle)
  'variables for animation
  Dim bmtemp As Bitmap
  fireball(index).celltop = 0
  fireball(index).cellcount = 4
  GetObject fireball(index).picsource.Handle, Len(bmtemp), bmtemp
  fireball(index).cellwidth = bmtemp.BmWidth
  fireball(index).cellheight = bmtemp.BmHeight / fireball(index).cellcount
  If fireball(index).cellwidth < fireball(index).cellheight Then
    fireball(index).radius = fireball(index).cellwidth / 2
  Else
    fireball(index).radius = fireball(index).cellheight / 2
  End If

  ReleaseDC 0, HTempDC

Next index
End Sub
'Draws the fireball that needs to be
Public Sub FireDraw(number As Integer)

Dim fireballacross As Long
Dim fireballdown As Long
fireballacross = fireball(number).cellwidth
fireballdown = fireball(number).cellheight

BitBlt OffscreenDc, fireball(number).position.x, fireball(number).position.y, fireballacross, fireballdown, fireball(number).hmaskdc, 0, fireball(number).celltop, vbSrcAnd
BitBlt OffscreenDc, fireball(number).position.x, fireball(number).position.y, fireballacross, fireballdown, fireball(number).hsourcedc, 0, fireball(number).celltop, vbSrcPaint

End Sub


'animates the paticular fireball
Public Function FireballAnimate(number As Integer)
Dim fireballTop As Long
fireballTop = fireball(number).celltop + fireball(number).cellheight

        
If fireballTop > 72 Then
    fireballTop = 0
End If
  
      
FireballAnimate = fireballTop

End Function
'move the fireball, follows mario or bounces
Public Sub FireMove(number As Integer)
Dim FireInc As Integer
If fireball(number).position.x > Backdrop.rightside Then
  FireReset (number)
ElseIf fireball(number).position.x < Backdrop.leftside Then
  FireReset (number)
ElseIf fireball(number).position.y < 0 Then
  FireReset (number)
Else

  If fired(number) = True Then
    If fireball(number).position.y < (TempPos(number) + 3) Then
      FireInc = 5
    ElseIf fireball(number).position.y > (TempPos(number) + 15) Then
      FireInc = -5
    End If
    fireball(number).speed.y = fireball(number).speed.y + FireInc
    fireball(number).position.y = fireball(number).position.y + fireball(number).speed.y
    fireball(number).position.x = fireball(number).position.x + fireball(number).speed.x
  Else
    fireball(number).position.x = mario.position.x
    fireball(number).position.y = mario.position.y
  End If
End If
End Sub
'if the fireball goes of the screen reset!
Public Sub FireReset(number As Integer)

fired(number) = False
fireball(number).position.x = mario.position.x
fireball(number).position.y = mario.position.y
fireball(number).speed.x = 0
fireball(number).speed.y = 0

End Sub



Public Sub coinSet()
cnticker = 35
Dim index As Integer
For index = 0 To coinmax
  coin(index).State = Waiting
  If index Mod 2 = 1 Then
    coin(index).position.x = 630
    coin(index).position.y = 20
    coin(index).speed.x = -5
    coin(index).speed.y = 0
  Else
    coin(index).position.x = 0
    coin(index).position.y = 20
    coin(index).speed.x = 5
    coin(index).speed.y = 0
  End If
  coin(index).Startspeed.x = 20
  coin(index).Startspeed.y = 35
  coin(index).onfloor = True
  Dim HTempDC As Long
  Dim HOldbmp As Long
  HTempDC = GetDC(0)
  'normal
  Set coin(index).picsource = New StdPicture
  Set coin(index).picsource = LoadPicture(App.Path & "\Pics\coinspr.bmp")
  coin(index).hsourcedc = CreateCompatibleDC(HTempDC)
  HOldbmp = SelectObject(coin(index).hsourcedc, coin(index).picsource.Handle)
  'mask
  Set coin(index).picmask = New StdPicture
  Set coin(index).picmask = LoadPicture(App.Path & "\Pics\coinmsk.bmp")
  coin(index).hmaskdc = CreateCompatibleDC(HTempDC)
  HOldbmp = SelectObject(coin(index).hmaskdc, coin(index).picmask.Handle)
  'variables for animation
  Dim bmtemp As Bitmap
  coin(index).celltop = 0
  coin(index).cellcount = 3
  GetObject coin(index).picsource.Handle, Len(bmtemp), bmtemp
  coin(index).cellwidth = bmtemp.BmWidth
  coin(index).cellheight = bmtemp.BmHeight / coin(index).cellcount
  If coin(index).cellwidth < coin(index).cellheight Then
      coin(index).radius = coin(index).cellwidth / 2
  Else
      coin(index).radius = coin(index).cellheight / 2
  End If
  
  
  
  ReleaseDC 0, HTempDC
Next index
End Sub

Public Sub CoinDraw(cn As Sprite)
Dim cnacross As Integer
Dim cndown As Integer
cnacross = cn.cellwidth
cndown = cn.cellheight

BitBlt OffscreenDc, cn.position.x, cn.position.y, cnacross, cndown, cn.hmaskdc, 0, cn.celltop, vbSrcAnd
BitBlt OffscreenDc, cn.position.x, cn.position.y, cnacross, cndown, cn.hsourcedc, 0, cn.celltop, vbSrcPaint

End Sub

Public Sub Coinmove(cn As Sprite)
Dim cnBottom As Integer
cnBottom = cn.position.y + cn.cellheight

GetSpeed cn
If cn.position.x > Backdrop.rightside Then
    If cnBottom = Floors(0).top Then
        cn.position.y = 10
        cn.speed.x = -cn.speed.x
    Else
        cn.position.x = Backdrop.leftside
    End If
Else
    cn.position.x = cn.position.x + cn.speed.x
End If

If cn.position.x < Backdrop.leftside Then
    If cnBottom = Floors(0).top Then
        cn.position.y = 10
        cn.speed.x = -cn.speed.x
    Else
        cn.position.x = Backdrop.rightside
    End If
Else
    cn.position.x = cn.position.x + cn.speed.x
End If

cn.position.y = cn.position.y + cn.speed.y

If cn.position.y > 480 Then
    cn.State = Dead
End If

End Sub

Public Function coinanimate(cn As Sprite)
Dim cnTop As Long
If cn.State = Dying Then
  coinanimate = (cn.cellheight * 2)
  Exit Function
End If
cnTop = cn.celltop + cn.cellheight
If cnTop > cn.cellheight Then
  cnTop = 0
End If
coinanimate = cnTop
End Function

Public Sub cncreate()
cnticker = cnticker - 1
If cnticker = 0 And cnNumber <= coinmax Then
    coin(cnNumber).State = Normal
    cnticker = 53
    cnNumber = cnNumber + 1
End If
End Sub

Public Sub LuigiSet()
luigi.Name = "luigi"
luigi.State = Normal
luigi.position.x = 415
luigi.position.y = 300
luigi.speed.x = 0
luigi.speed.y = 0
luigi.Startspeed.x = 20
luigi.Startspeed.y = 35
luigi.onfloor = True
Dim HTempDC As Long
Dim HOldbmp As Long
HTempDC = GetDC(0)
'normal
Set luigi.picsource = New StdPicture
Set luigi.picsource = LoadPicture(App.Path & "\Pics\luigis.bmp")
luigi.hsourcedc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(luigi.hsourcedc, luigi.picsource.Handle)
'mask
Set luigi.picmask = New StdPicture
Set luigi.picmask = LoadPicture(App.Path & "\Pics\luigimsk.bmp")
luigi.hmaskdc = CreateCompatibleDC(HTempDC)
HOldbmp = SelectObject(luigi.hmaskdc, luigi.picmask.Handle)
'variables for animation
Dim bmtemp As Bitmap
luigi.celltop = 0
luigi.cellcount = 14
GetObject luigi.picsource.Handle, Len(bmtemp), bmtemp
luigi.cellwidth = bmtemp.BmWidth
luigi.cellheight = bmtemp.BmHeight / luigi.cellcount
If luigi.cellwidth < luigi.cellheight Then
    luigi.radius = luigi.cellwidth / 2
Else
    luigi.radius = luigi.cellheight / 2
End If



ReleaseDC 0, HTempDC





End Sub


Public Function LuigiAnimate()
If luigi.State = Dying Then
  LuigiAnimate = 13 * luigi.cellheight
  Exit Function
End If
  
Dim luigiTop As Long
luigiTop = luigi.celltop + luigi.cellheight
If luigi.speed.x > 0 Then
    If luigi.onfloor = False Then
        luigiTop = 272
    Else
        
        If luigiTop >= 102 Then
            luigiTop = 0
        End If
    End If
ElseIf luigi.speed.x < 0 Then
    If luigi.onfloor = False Then
        luigiTop = 306
    Else
        If luigiTop >= 204 Or luigiTop < 102 Then
            luigiTop = 102
        End If
    End If
ElseIf Int(Rnd * 30) = 0 Then
    luigiTop = 238
Else
    luigiTop = 204
End If
LuigiAnimate = luigiTop
End Function


Public Sub LuigiDraw()

BitBlt OffscreenDc, luigi.position.x, luigi.position.y, 38, 34, luigi.hmaskdc, 0, luigi.celltop, vbSrcAnd
BitBlt OffscreenDc, luigi.position.x, luigi.position.y, 38, 34, luigi.hsourcedc, 0, luigi.celltop, vbSrcPaint

End Sub


Public Sub LuigiMove()

GetSpeed luigi
If luigi.position.x > Backdrop.rightside Then
    luigi.position.x = Backdrop.leftside
Else
    luigi.position.x = luigi.position.x + luigi.speed.x
End If

If luigi.position.x < Backdrop.leftside Then
    luigi.position.x = Backdrop.rightside
Else
    luigi.position.x = luigi.position.x + luigi.speed.x
End If

luigi.position.y = luigi.position.y + luigi.speed.y

End Sub

