Attribute VB_Name = "modGame"
'The BitBlt function allows for fast and smooth drawing to the form
'and to picture boxes, but isn't great for animation

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal animX As Long, ByVal animY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'origionally I had a bunch of cool sound effects, but it REALLY slowed the game down.
'maybe I'll get 'em working in a later version.
'allows the playing of wav files
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'for the sound function
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound


'for the bitblt function
Public Const SRCCOPY = &HCC0020   'Copies the source over the destination
Public Const SRCINVERT = &H660046 'Copies and inverts the source over the destination
Public Const SRCAND = &H8800C6    'Adds the source to the destination

Public Walkable(0 To 164) As Integer
Public Texture(0 To 164) As Integer
Public TileLeft(0 To 164) As Integer
Public TileTOP(0 To 164) As Integer

Public Const fLEFT As Integer = 0    'left animation
Public Const fUP As Integer = 100    'up animation
Public Const fRIGHT As Integer = 200 'right animation
Public Const fDOWN As Integer = 300  'down animation

Public MapX As Integer
Public MapY As Integer
Public tLEFT(0 To 254) As Integer
Public tTOP(0 To 254) As Integer
Public Object_Data(0 To 254) As String
Public tENEMY(0 To 254) As Integer
Public tENEMY_LEFT(0 To 254) As Integer
Public tENEMY_TOP(0 To 254) As Integer
Public tENEMY_DIRECTION(0 To 254)
Public tENEMY_frameX(0 To 254) As Integer
Public tENEMY_frameY(0 To 254) As Integer
Public Walk(0 To 254) As Integer
Public FrameX As Integer
Public FrameY As Integer
Public PlayerX As Integer
Public PlayerY As Integer
Public wSPEED As Integer
Public eSPEED As Integer
Public Direction
Public dHIT
Public Health As Integer
Public Enemies As Integer
Public Magic_Direction(0 To 30) As String
Public Magic_Type(0 To 30) As String
Public Magic_Speed(0 To 30) As Integer
Public Magic_Left(0 To 30) As Integer
Public Magic_Top(0 To 30) As Integer
Sub wait(howlong)
' USAGE: wait #ofseconds; example wait 3 will wait 3 seconds
temptime = Timer
Do
DoEvents
Loop While Timer < temptime + howlong
End Sub

Public Sub NewMap()
Randomize
X = 0
Y = 0
Open App.Path & "\x" & MapX & "y" & MapY & ".map" For Input As #1
For land = 0 To 224
Input #1, t, w, obj, obj_tag, obj_dat
tLEFT(land) = X
tTOP(land) = Y
Walk(land) = w
Object_Data(land) = obj_dat
a = BitBlt(frmMain.picRefresh.hdc, X, Y, 40, 40, frmTiles.tile(t).hdc, 0, 0, SRCCOPY)
If obj = 1 Then
a = BitBlt(frmMain.picRefresh.hdc, X, Y, 40, 40, frmTiles.object(obj_tag).hdc, 40, 0, SRCAND)
a = BitBlt(frmMain.picRefresh.hdc, X, Y, 40, 40, frmTiles.object(obj_tag).hdc, 0, 0, SRCINVERT)
End If
Enemies = 0
If Walk(land) = 1 Then Enemies = Int(100 * Rnd)
If Enemies = 1 Or Enemies = 50 Then
tENEMY(land) = 1
tENEMY_LEFT(land) = tLEFT(land)
tENEMY_TOP(land) = tTOP(land)
direct = Int(4 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(land) = "left"
If direct = 1 Then tENEMY_DIRECTION(land) = "up"
If direct = 2 Then tENEMY_DIRECTION(land) = "right"
If direct = 3 Then tENEMY_DIRECTION(land) = "down"
Else
tENEMY(land) = 0
tENEMY_LEFT(land) = -1
tENEMY_TOP(land) = -1
tENEMY_DIRECTION(land) = ""
End If
X = X + 40
If X >= 40 * 15 Then
X = 0
Y = Y + 40
End If
Next land
Close #1
frmMain.picRefresh.Refresh

a = BitBlt(frmMain.picMain.hdc, 0, 0, frmMain.picMain.Width, frmMain.picMain.Height, frmMain.picRefresh.hdc, 0, 0, SRCCOPY)
For t = 0 To 254
If tENEMY(t) = 1 Then a = BitBlt(frmMain.picMain.hdc, tENEMY_LEFT(t), tENEMY_TOP(t), 50, 50, frmTiles.picSoldier.hdc, 50, 0, SRCAND)
If tENEMY(t) = 1 Then a = BitBlt(frmMain.picMain.hdc, tENEMY_LEFT(t), tENEMY_TOP(t), 50, 50, frmTiles.picSoldier.hdc, 0, 0, SRCINVERT)
Next t
frmMain.picMain.Refresh
End Sub
Public Sub Damien_Hit_Left()
dHIT = 1
FrameX = fRIGHT
FrameY = 0
For hit = 1 To 25
PlayerX = PlayerX - 3
For t = 0 To 254
If PlayerX + 11 - wSPEED >= tLEFT(t) And PlayerX + 11 - wSPEED <= tLEFT(t) + 40 And PlayerY + 38 >= tTOP(t) And PlayerY + 38 <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerX = PlayerX + 3
If PlayerX + 11 - wSPEED >= tLEFT(t) And PlayerX + 11 - wSPEED <= tLEFT(t) + 40 And PlayerY + 50 >= tTOP(t) And PlayerY + 50 <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerX = PlayerX + 3
Next t
a = BitBlt(frmMain.picMain.hdc, 0, 0, frmMain.picMain.Width, frmMain.picMain.Height, frmMain.picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX, FrameY, SRCINVERT)
frmMain.picMain.Refresh
Next hit
Health = Health - 10
frmMain.picHealth.Cls
frmMain.picHealth.Line (0, 0)-(Health, frmMain.picHealth.ScaleHeight), QBColor(9), BF
frmMain.lblHealth.Caption = "Health: " & Health
dHIT = 0
End Sub

Public Sub Damien_Hit_Right()
dHIT = 1
FrameX = fLEFT
FrameY = 0
For hit = 1 To 25
PlayerX = PlayerX + 3
For t = 0 To 254
If PlayerX + 38 + wSPEED >= tLEFT(t) And PlayerX + 38 + wSPEED <= tLEFT(t) + 40 And PlayerY + 35 >= tTOP(t) And PlayerY + 35 <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerX = PlayerX - 3
If PlayerX + 38 + wSPEED >= tLEFT(t) And PlayerX + 38 + wSPEED <= tLEFT(t) + 40 And PlayerY + 50 >= tTOP(t) And PlayerY + 50 <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerX = PlayerX - 3
Next t
a = BitBlt(frmMain.picMain.hdc, 0, 0, frmMain.picMain.Width, frmMain.picMain.Height, frmMain.picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX, FrameY, SRCINVERT)
frmMain.picMain.Refresh
Next hit
Health = Health - 10
frmMain.picHealth.Cls
frmMain.picHealth.Line (0, 0)-(Health, frmMain.picHealth.ScaleHeight), QBColor(9), BF
frmMain.lblHealth.Caption = "Health: " & Health
dHIT = 0
End Sub

Public Sub Damien_Hit_Up()
dHIT = 1
FrameX = fUP
FrameY = 0
For hit = 1 To 25
PlayerY = PlayerY - 3
For t = 0 To 254
If PlayerX + 11 >= tLEFT(t) And PlayerX + 11 <= tLEFT(t) + 40 And PlayerY + 35 - wSPEED >= tTOP(t) And PlayerY + 35 - wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerY = PlayerY + 3
If PlayerX + 38 >= tLEFT(t) And PlayerX + 38 <= tLEFT(t) + 40 And PlayerY + 35 - wSPEED >= tTOP(t) And PlayerY + 35 - wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerY = PlayerY + 3
Next t
a = BitBlt(frmMain.picMain.hdc, 0, 0, frmMain.picMain.Width, frmMain.picMain.Height, frmMain.picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX, FrameY, SRCINVERT)
frmMain.picMain.Refresh
Next hit
Health = Health - 10
frmMain.picHealth.Cls
frmMain.picHealth.Line (0, 0)-(Health, frmMain.picHealth.ScaleHeight), QBColor(9), BF
frmMain.lblHealth.Caption = "Health: " & Health
dHIT = 0
End Sub

Public Sub Damien_Hit_Down()
dHIT = 1
FrameX = fDOWN
FrameY = 0
For hit = 1 To 25
PlayerY = PlayerY + 3
For t = 0 To 254
If PlayerX + 11 >= tLEFT(t) And PlayerX + 11 <= tLEFT(t) + 40 And PlayerY + 50 + wSPEED >= tTOP(t) And PlayerY + 50 + wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerY = PlayerY - 3
If PlayerX + 38 >= tLEFT(t) And PlayerX + 38 <= tLEFT(t) + 40 And PlayerY + 50 + wSPEED >= tTOP(t) And PlayerY + 50 + wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then PlayerY = PlayerY - 3
Next t
a = BitBlt(frmMain.picMain.hdc, 0, 0, frmMain.picMain.Width, frmMain.picMain.Height, frmMain.picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamienHit.hdc, FrameX, FrameY, SRCINVERT)
frmMain.picMain.Refresh
Next hit
Health = Health - 10
frmMain.picHealth.Cls
frmMain.picHealth.Line (0, 0)-(Health, frmMain.picHealth.ScaleHeight), QBColor(9), BF
frmMain.lblHealth.Caption = "Health: " & Health
dHIT = 0
End Sub

Public Sub Cast_Magic_Up(cast As String)
For m = 0 To 30
If Magic_Direction(m) = "" Then
Magic_Direction(m) = "up"
Magic_Speed(m) = 10
Magic_Type(m) = cast
Magic_Left(m) = PlayerX
Magic_Top(m) = PlayerY
Exit For
End If
Next m
End Sub

Public Sub Cast_Magic_Down(cast As String)
For m = 0 To 30
If Magic_Direction(m) = "" Then
Magic_Direction(m) = "down"
Magic_Speed(m) = 10
Magic_Type(m) = cast
Magic_Left(m) = PlayerX
Magic_Top(m) = PlayerY
Exit For
End If
Next m
End Sub

Public Sub Cast_Magic_Left(cast As String)
For m = 0 To 30
If Magic_Direction(m) = "" Then
Magic_Direction(m) = "left"
Magic_Speed(m) = 10
Magic_Type(m) = cast
Magic_Left(m) = PlayerX
Magic_Top(m) = PlayerY
Exit For
End If
Next m
End Sub

Public Sub Cast_Magic_Right(cast As String)
For m = 0 To 30
If Magic_Direction(m) = "" Then
Magic_Direction(m) = "right"
Magic_Speed(m) = 10
Magic_Type(m) = cast
Magic_Left(m) = PlayerX
Magic_Top(m) = PlayerY
Exit For
End If
Next m
End Sub
