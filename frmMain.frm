VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   2175
   ClientTop       =   1815
   ClientWidth     =   5820
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5820
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer tmrMagic 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   360
   End
   Begin VB.PictureBox picHealth 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   360
      ScaleHeight     =   10
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer tmrEnemy 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   360
   End
   Begin VB.Timer tmrMusic 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   360
   End
   Begin VB.CheckBox chkMusic 
      BackColor       =   &H00000000&
      Caption         =   "Music"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin MCI.MMControl music 
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1085
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   1695
      Left            =   3240
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.PictureBox picRefresh 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   3240
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblHealth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Health: 100"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DM_DISPLAYFLAGS = &H200000
Const DM_DISPLAYFREQUENCY = &H400000

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const BITSPIXEL = 12

' /* Flags for ChangeDisplaySettings */
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H2
Const CDS_FULLSCREEN = &H4
Const CDS_GLOBAL = &H8
Const CDS_SET_PRIMARY = &H10
Const CDS_RESET = &H40000000
Const CDS_SETRECT = &H20000000
Const CDS_NORESET = &H10000000

' /* Return values for ChangeDisplaySettings */
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const DISP_CHANGE_FAILED = -1
Const DISP_CHANGE_BADMODE = -2
Const DISP_CHANGE_NOTUPDATED = -3
Const DISP_CHANGE_BADFLAGS = -4
Const DISP_CHANGE_BADPARAM = -5

Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Dim D() As DEVMODE, lNumModes As Long
Private res(0 To 50)
Dim oRES
Public Sub reschange()
Dim Y As Long
Dim Flags As Long, X As Long
For checkres = 0 To 50
If res(checkres) = "800x600x16" Then X = checkres
Next checkres
D(X).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    Flags = CDS_UPDATEREGISTRY
    Y = ChangeDisplaySettings(D(X), Flags)
    Select Case Y
        Case DISP_CHANGE_RESTART
            Y = MsgBox("This change will not take effect until you reboot the system.  Reboot now?", vbYesNo)
            If Y = vbYes Then
                Flags = 0
                Y = ExitWindowsEx(EWX_REBOOT, Flags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        Case Else
            MsgBox "Error changing resolution! Returned: " & Y
    End Select
End Sub

Private Sub cmdExit_Click()
Unload frmMain
End Sub

Private Sub cmdStart_Click()
Health = 100
picMain.Enabled = True
PlayerX = picMain.ScaleWidth / 2
PlayerY = picMain.ScaleHeight / 2
Call NewMap
If chkMusic.Value = 1 Then
music.Command = "Stop"
music.Command = "Close"
music.FileName = App.Path & "\Battle-unreleased.mid"
music.Command = "Open"
music.Command = "Play"
tmrMusic.Enabled = True
End If
tmrEnemy.Enabled = True
tmrMagic.Enabled = True
cmdStart.Visible = False
chkMusic.Visible = False
lblHealth.Visible = True
picHealth.Visible = True
picMain.SetFocus
End Sub
Private Sub Form_DblClick()
Unload frmMain
End Sub

Private Sub Form_Load()
wSPEED = 7
eSPEED = 7

Dim l As Long, lMaxModes As Long
    Dim lBits As Long, lWidth As Long, lHeight As Long
    lBits = GetDeviceCaps(hdc, BITSPIXEL)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    lMaxModes = 8
    ReDim D(0 To lMaxModes) As DEVMODE
    lNumModes = 0
    l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Do While l
        res(r) = D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel
        r = r + 1
        If lBits = D(lNumModes).dmBitsPerPel And lWidth = D(lNumModes).dmPelsWidth And lHeight = D(lNumModes).dmPelsHeight Then oRES = D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve D(0 To lMaxModes) As DEVMODE
        End If
        l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Loop
    lNumModes = lNumModes - 1

Call reschange
End Sub
Private Sub Form_Resize()
picMain.Move 3000, 0
picMain.Width = frmMain.ScaleWidth - 3000
picMain.Height = picMain.Width
'picMain.Scale (0, 0)-(150, 150)

picRefresh.Move 3000, 0
picRefresh.Width = frmMain.ScaleWidth - 3000
picRefresh.Height = picRefresh.Width
'picRefresh.Scale (0, 0)-(150, 150)

lblHealth.Move 500, 250
picHealth.Move 500, lblHealth.Top + lblHealth.Height
picHealth.Line (0, 0)-(picHealth.ScaleWidth, picHealth.ScaleHeight), QBColor(9), BF
cmdExit.Move cmdStart.Left, frmMain.ScaleHeight - (cmdExit.Height + 150)
End Sub

Private Sub Form_Unload(Cancel As Integer)
music.Command = "Stop"
music.Command = "Close"
Call resreset
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
Randomize
If dHIT = 1 Then Exit Sub

If KeyCode = vbKeySpace And FrameX = fLEFT Then Call Cast_Magic_Left("fireball")
If KeyCode = vbKeySpace And FrameX = fRIGHT Then Call Cast_Magic_Right("fireball")
If KeyCode = vbKeySpace And FrameX = fUP Then Call Cast_Magic_Up("fireball")
If KeyCode = vbKeySpace And FrameX = fDOWN Then Call Cast_Magic_Down("fireball")

For t = 0 To 254
If KeyCode = vbKeyE And PlayerX + 25 >= tLEFT(t) And PlayerX + 25 <= tLEFT(t) + 40 And PlayerY + 25 >= tTOP(t) And PlayerY + 25 <= tTOP(t) + 40 And Object_Data(t) <> "NA" Then
If Object_Data(t) <> "House" Then MsgBox Object_Data(t)
End If
Next t

If KeyCode = 37 Then
For t = 0 To 254
If PlayerX + 11 - wSPEED >= tLEFT(t) And PlayerX + 11 - wSPEED <= tLEFT(t) + 40 And PlayerY + 38 >= tTOP(t) And PlayerY + 38 <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
If PlayerX + 11 - wSPEED >= tLEFT(t) And PlayerX + 11 - wSPEED <= tLEFT(t) + 40 And PlayerY + 50 >= tTOP(t) And PlayerY + 50 <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
'If playerx >= tLEFT(t) And playerx <= tLEFT(t) + 40 And playery >= tTOP(t) And playery <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx >= tLEFT(t) And playerx <= tLEFT(t) + 40 And playery + 50 >= tTOP(t) And playery + 50 <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx >= tLEFT(t) And playerx <= tLEFT(t) + 40 And playery + 25 >= tTOP(t) And playery + 25 <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
Next t
Direction = dleft
FrameX = fLEFT
FrameY = FrameY + 50
PlayerX = PlayerX - wSPEED
If FrameY >= 50 * 8 Then FrameY = 0
a = BitBlt(picMain.hdc, 0, 0, picMain.Width, picMain.Height, picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX, FrameY, SRCINVERT)
If PlayerX < 0 Then
PlayerX = picMain.ScaleWidth
MapX = MapX - 1
Call NewMap
End If
End If

If KeyCode = 38 Then
For t = 0 To 254
If PlayerX + 11 >= tLEFT(t) And PlayerX + 11 <= tLEFT(t) + 40 And PlayerY + 35 - wSPEED >= tTOP(t) And PlayerY + 35 - wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
If PlayerX + 38 >= tLEFT(t) And PlayerX + 38 <= tLEFT(t) + 40 And PlayerY + 35 - wSPEED >= tTOP(t) And PlayerY + 35 - wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
'If playerx >= tLEFT(t) And playerx <= tLEFT(t) + 40 And playery >= tTOP(t) And playery <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx + 50 >= tLEFT(t) And playerx + 50 <= tLEFT(t) + 40 And playery >= tTOP(t) And playery <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx + 25 >= tLEFT(t) And playerx + 25 <= tLEFT(t) + 40 And playery >= tTOP(t) And playery <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
Next t
Direction = dUP
FrameX = fUP
FrameY = FrameY + 50
PlayerY = PlayerY - wSPEED
If FrameY >= 50 * 8 Then FrameY = 0
a = BitBlt(picMain.hdc, 0, 0, picMain.Width, picMain.Height, picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX, FrameY, SRCINVERT)
If PlayerY < 0 Then
PlayerY = picMain.ScaleHeight
MapY = MapY + 1
Call NewMap
End If
End If

If KeyCode = 39 Then
For t = 0 To 254
If PlayerX + 38 + wSPEED >= tLEFT(t) And PlayerX + 38 + wSPEED <= tLEFT(t) + 40 And PlayerY + 35 >= tTOP(t) And PlayerY + 35 <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
If PlayerX + 38 + wSPEED >= tLEFT(t) And PlayerX + 38 + wSPEED <= tLEFT(t) + 40 And PlayerY + 50 >= tTOP(t) And PlayerY + 50 <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
'If playerx + 50 >= tLEFT(t) And playerx + 50 <= tLEFT(t) + 40 And playery >= tTOP(t) And playery <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx + 50 >= tLEFT(t) And playerx + 50 <= tLEFT(t) + 40 And playery + 50 >= tTOP(t) And playery + 50 <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx + 50 >= tLEFT(t) And playerx + 50 <= tLEFT(t) + 40 And playery + 25 >= tTOP(t) And playery + 25 <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
Next t
Direction = dRIGHT
FrameX = fRIGHT
FrameY = FrameY + 50
PlayerX = PlayerX + wSPEED
If FrameY >= 50 * 8 Then FrameY = 0
a = BitBlt(picMain.hdc, 0, 0, picMain.Width, picMain.Height, picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX, FrameY, SRCINVERT)
If PlayerX > picMain.ScaleWidth Then
PlayerX = 0
MapX = MapX + 1
Call NewMap
End If
End If

If KeyCode = 40 Then
For t = 0 To 254
If PlayerX + 11 >= tLEFT(t) And PlayerX + 11 <= tLEFT(t) + 40 And PlayerY + 50 + wSPEED >= tTOP(t) And PlayerY + 50 + wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
If PlayerX + 38 >= tLEFT(t) And PlayerX + 38 <= tLEFT(t) + 40 And PlayerY + 50 + wSPEED >= tTOP(t) And PlayerY + 50 + wSPEED <= tTOP(t) + 40 And Walk(t) = 0 Then Exit Sub
'If playerx >= tLEFT(t) And playerx <= tLEFT(t) + 40 And playery + 50 >= tTOP(t) And playery + 50 <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx + 50 >= tLEFT(t) And playerx + 50 <= tLEFT(t) + 40 And playery + 50 >= tTOP(t) And playery + 50 <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
'If playerx + 25 >= tLEFT(t) And playerx + 25 <= tLEFT(t) + 40 And playery + 50 >= tTOP(t) And playery + 50 <= tTOP(t) + 40 And tENEMY(t) = 1 Then
'Call battle
'Exit Sub
'End If
Next t
Direction = dDOWN
FrameX = fDOWN
FrameY = FrameY + 50
PlayerY = PlayerY + wSPEED
If FrameY >= 50 * 8 Then FrameY = 0
a = BitBlt(picMain.hdc, 0, 0, picMain.Width, picMain.Height, picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX, FrameY, SRCINVERT)
If PlayerY > picMain.ScaleHeight Then
PlayerY = 0
MapY = MapY - 1
Call NewMap
End If
End If

For t = 0 To 254
If tENEMY(t) = 1 Then
a = BitBlt(picMain.hdc, tENEMY_LEFT(t), tENEMY_TOP(t), 50, 50, frmTiles.picSoldier.hdc, tENEMY_frameX(t) + 50, tENEMY_frameY(t), SRCAND)
a = BitBlt(picMain.hdc, tENEMY_LEFT(t), tENEMY_TOP(t), 50, 50, frmTiles.picSoldier.hdc, tENEMY_frameX(t), tENEMY_frameY(t), SRCINVERT)
End If
Next t

For m = 0 To 30
If Magic_Direction(m) <> "" Then
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 40, 0, SRCAND)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 0, 0, SRCINVERT)
End If
Next m

picMain.Refresh
End Sub
Public Sub resreset()
Dim Y As Long
Dim Flags As Long, X As Long
For checkres = 0 To 50
If res(checkres) = oRES Then X = checkres
Next checkres
D(X).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    Flags = CDS_UPDATEREGISTRY
    Y = ChangeDisplaySettings(D(X), Flags)
    Select Case Y
        Case DISP_CHANGE_RESTART
            Y = MsgBox("This change will not take effect until you reboot the system.  Reboot now?", vbYesNo)
            If Y = vbYes Then
                Flags = 0
                Y = ExitWindowsEx(EWX_REBOOT, Flags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        Case Else
            MsgBox "Error changing resolution! Returned: " & Y
    End Select
Unload frmMain
Unload frmTiles
End
End Sub

Private Sub tmrEnemy_Timer()
Randomize
If dHIT = 1 Then Exit Sub
a = BitBlt(frmMain.picMain.hdc, 0, 0, frmMain.picRefresh.ScaleWidth, frmMain.picRefresh.ScaleHeight, frmMain.picRefresh.hdc, 0, 0, SRCCOPY)
For r = 0 To 254
If tENEMY(r) = 1 Then
If tENEMY_DIRECTION(r) = "left" Then
For l = 0 To 254
If tENEMY_LEFT(r) + 11 - wSPEED >= tLEFT(l) And tENEMY_LEFT(r) + 11 - wSPEED <= tLEFT(l) + 40 And tENEMY_TOP(r) + 38 >= tTOP(l) And tENEMY_TOP(r) + 38 <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "up"
If direct = 1 Then tENEMY_DIRECTION(r) = "right"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
End If
If tENEMY_LEFT(r) + 11 - wSPEED >= tLEFT(l) And tENEMY_LEFT(r) + 11 - wSPEED <= tLEFT(l) + 40 And tENEMY_TOP(r) + 50 >= tTOP(l) And tENEMY_TOP(r) + 50 <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "up"
If direct = 1 Then tENEMY_DIRECTION(r) = "right"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
End If
Next l
If tENEMY_LEFT(r) + 11 >= PlayerX + 11 And tENEMY_LEFT(r) + 11 <= PlayerX + 38 And tENEMY_TOP(r) + 35 <= PlayerY + 50 And tENEMY_TOP(r) + 35 >= PlayerY + 35 Then Call Damien_Hit_Left
If tENEMY_LEFT(r) + 11 >= PlayerX + 11 And tENEMY_LEFT(r) + 11 <= PlayerX + 38 And tENEMY_TOP(r) + 50 <= PlayerY + 50 And tENEMY_TOP(r) + 50 >= PlayerY + 35 Then Call Damien_Hit_Left
End If

If tENEMY_DIRECTION(r) = "up" Then
For l = 0 To 254
If tENEMY_LEFT(r) + 11 >= tLEFT(l) And tENEMY_LEFT(r) + 11 <= tLEFT(l) + 40 And tENEMY_TOP(r) + 35 - wSPEED >= tTOP(l) And tENEMY_TOP(r) + 35 - wSPEED <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "right"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
End If
If tENEMY_LEFT(r) + 38 >= tLEFT(l) And tENEMY_LEFT(r) + 38 <= tLEFT(l) + 40 And tENEMY_TOP(r) + 35 - wSPEED >= tTOP(l) And tENEMY_TOP(r) + 35 - wSPEED <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "right"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
End If
Next l
If tENEMY_LEFT(r) + 11 >= PlayerX + 11 And tENEMY_LEFT(r) + 11 <= PlayerX + 38 And tENEMY_TOP(r) + 35 <= PlayerY + 50 And tENEMY_TOP(r) + 35 >= PlayerY + 35 Then Call Damien_Hit_Up
If tENEMY_LEFT(r) + 38 >= PlayerX + 11 And tENEMY_LEFT(r) + 38 <= PlayerX + 38 And tENEMY_TOP(r) + 35 <= PlayerY + 50 And tENEMY_TOP(r) + 35 >= PlayerY + 35 Then Call Damien_Hit_Up
End If

If tENEMY_DIRECTION(r) = "right" Then
For l = 0 To 254
If tENEMY_LEFT(r) + 38 + wSPEED >= tLEFT(l) And tENEMY_LEFT(r) + 38 + wSPEED <= tLEFT(l) + 40 And tENEMY_TOP(r) + 35 >= tTOP(l) And tENEMY_TOP(r) + 35 <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "up"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
End If
If tENEMY_LEFT(r) + 38 + wSPEED >= tLEFT(l) And tENEMY_LEFT(r) + 38 + wSPEED <= tLEFT(l) + 40 And tENEMY_TOP(r) + 50 >= tTOP(l) And tENEMY_TOP(r) + 50 <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "up"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
End If
Next l
If tENEMY_LEFT(r) + 38 >= PlayerX + 11 And tENEMY_LEFT(r) + 38 <= PlayerX + 38 And tENEMY_TOP(r) + 35 <= PlayerY + 50 And tENEMY_TOP(r) + 35 >= PlayerY + 35 Then Call Damien_Hit_Right
If tENEMY_LEFT(r) + 38 >= PlayerX + 11 And tENEMY_LEFT(r) + 38 <= PlayerX + 38 And tENEMY_TOP(r) + 50 <= PlayerY + 50 And tENEMY_TOP(r) + 50 >= PlayerY + 35 Then Call Damien_Hit_Right
End If

If tENEMY_DIRECTION(r) = "down" Then
For l = 0 To 254
If tENEMY_LEFT(r) + 11 >= tLEFT(l) And tENEMY_LEFT(r) + 11 <= tLEFT(l) + 40 And tENEMY_TOP(r) + 50 + wSPEED >= tTOP(l) And tENEMY_TOP(r) + 50 + wSPEED <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "up"
If direct = 2 Then tENEMY_DIRECTION(r) = "right"
End If
If tENEMY_LEFT(r) + 38 >= tLEFT(l) And tENEMY_LEFT(r) + 38 <= tLEFT(l) + 40 And tENEMY_TOP(r) + 50 + wSPEED >= tTOP(l) And tENEMY_TOP(r) + 50 + wSPEED <= tTOP(l) + 40 And Walk(l) = 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "up"
If direct = 2 Then tENEMY_DIRECTION(r) = "right"
End If
Next l
If tENEMY_LEFT(r) + 38 >= PlayerX + 11 And tENEMY_LEFT(r) + 38 <= PlayerX + 38 And tENEMY_TOP(r) + 50 <= PlayerY + 50 And tENEMY_TOP(r) + 50 >= PlayerY + 35 Then Call Damien_Hit_Down
If tENEMY_LEFT(r) + 11 >= PlayerX + 11 And tENEMY_LEFT(r) + 11 <= PlayerX + 38 And tENEMY_TOP(r) + 50 <= PlayerY + 50 And tENEMY_TOP(r) + 50 >= PlayerY + 35 Then Call Damien_Hit_Down
End If

If tENEMY_LEFT(r) < 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "up"
If direct = 1 Then tENEMY_DIRECTION(r) = "right"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
tENEMY_LEFT(r) = 1
End If
If tENEMY_LEFT(r) + 50 > frmMain.picRefresh.ScaleWidth Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "up"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
tENEMY_LEFT(r) = frmMain.picRefresh.ScaleWidth - 51
End If
If tENEMY_TOP(r) < 0 Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "right"
If direct = 2 Then tENEMY_DIRECTION(r) = "down"
tENEMY_TOP(r) = 1
End If
If tENEMY_TOP(r) + 50 > frmMain.picRefresh.ScaleHeight Then
direct = Int(3 * Rnd)
If direct = 0 Then tENEMY_DIRECTION(r) = "left"
If direct = 1 Then tENEMY_DIRECTION(r) = "up"
If direct = 2 Then tENEMY_DIRECTION(r) = "right"
tENEMY_TOP(r) = frmMain.picRefresh.ScaleHeight - 51
End If

If tENEMY_DIRECTION(r) = "left" Then
tENEMY_frameX(r) = fLEFT
tENEMY_frameY(r) = tENEMY_frameY(r) + 50
If tENEMY_frameY(r) >= 400 Then tENEMY_frameY(r) = 0
tENEMY_LEFT(r) = tENEMY_LEFT(r) - eSPEED
End If

If tENEMY_DIRECTION(r) = "up" Then
tENEMY_frameX(r) = fLEFT
tENEMY_frameY(r) = tENEMY_frameY(r) + 50
If tENEMY_frameY(r) >= 400 Then tENEMY_frameY(r) = 0
tENEMY_TOP(r) = tENEMY_TOP(r) - eSPEED
End If

If tENEMY_DIRECTION(r) = "right" Then
tENEMY_frameX(r) = fLEFT
tENEMY_frameY(r) = tENEMY_frameY(r) + 50
If tENEMY_frameY(r) >= 400 Then tENEMY_frameY(r) = 0
tENEMY_LEFT(r) = tENEMY_LEFT(r) + eSPEED
End If

If tENEMY_DIRECTION(r) = "down" Then
tENEMY_frameX(r) = fLEFT
tENEMY_frameY(r) = tENEMY_frameY(r) + 50
If tENEMY_frameY(r) >= 400 Then tENEMY_frameY(r) = 0
tENEMY_TOP(r) = tENEMY_TOP(r) + eSPEED
End If

a = BitBlt(frmMain.picMain.hdc, tENEMY_LEFT(r), tENEMY_TOP(r), 50, 50, frmTiles.picSoldier.hdc, tENEMY_frameX(r) + 50, tENEMY_frameY(r), SRCAND)
a = BitBlt(frmMain.picMain.hdc, tENEMY_LEFT(r), tENEMY_TOP(r), 50, 50, frmTiles.picSoldier.hdc, tENEMY_frameX(r), tENEMY_frameY(r), SRCINVERT)
End If

Next r

a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX, FrameY, SRCINVERT)

For m = 0 To 30
If Magic_Direction(m) <> "" Then
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 40, 0, SRCAND)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 0, 0, SRCINVERT)
End If
Next m

frmMain.picMain.Refresh
End Sub

Private Sub tmrMagic_Timer()
a = BitBlt(frmMain.picMain.hdc, 0, 0, frmMain.picRefresh.ScaleWidth, frmMain.picRefresh.ScaleHeight, frmMain.picRefresh.hdc, 0, 0, SRCCOPY)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX + 50, FrameY, SRCAND)
a = BitBlt(frmMain.picMain.hdc, PlayerX, PlayerY, 50, 50, frmTiles.picDamien.hdc, FrameX, FrameY, SRCINVERT)
For t = 0 To 254
If tENEMY(t) = 1 Then
a = BitBlt(frmMain.picMain.hdc, tENEMY_LEFT(t), tENEMY_TOP(t), 50, 50, frmTiles.picSoldier.hdc, tENEMY_frameX(t) + 50, tENEMY_frameY(t), SRCAND)
a = BitBlt(frmMain.picMain.hdc, tENEMY_LEFT(t), tENEMY_TOP(t), 50, 50, frmTiles.picSoldier.hdc, tENEMY_frameX(t), tENEMY_frameY(t), SRCINVERT)
End If
Next t

For m = 0 To 30

If Magic_Type(m) = "fireball" Then

If Magic_Direction(m) = "up" Then
Magic_Top(m) = Magic_Top(m) - Magic_Speed(m)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 40, 0, SRCAND)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 0, 0, SRCINVERT)
If Magic_Top(m) + 50 < 0 Then Magic_Direction(m) = ""
For t = 0 To 254
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
Next t
End If

If Magic_Direction(m) = "down" Then
Magic_Top(m) = Magic_Top(m) + 10
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 40, 0, SRCAND)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 0, 0, SRCINVERT)
If Magic_Top(m) > picMain.ScaleHeight Then Magic_Direction(m) = ""
For t = 0 To 254
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
Next t
End If

If Magic_Direction(m) = "left" Then
Magic_Left(m) = Magic_Left(m) - Magic_Speed(m)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 40, 0, SRCAND)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 0, 0, SRCINVERT)
If Magic_Left(m) + 50 < 0 Then Magic_Direction(m) = ""
For t = 0 To 254
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
Next t
End If

If Magic_Direction(m) = "right" Then
Magic_Left(m) = Magic_Left(m) + Magic_Speed(m)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 40, 0, SRCAND)
a = BitBlt(frmMain.picMain.hdc, Magic_Left(m), Magic_Top(m), 40, 40, frmTiles.picFireball.hdc, 0, 0, SRCINVERT)
If Magic_Left(m) > picMain.ScaleWidth Then Magic_Direction(m) = ""
For t = 0 To 254
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) >= tENEMY_TOP(t) And Magic_Top(m) <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) >= tENEMY_LEFT(t) And Magic_Left(m) <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
If tENEMY(t) = 1 And Magic_Left(m) + 40 >= tENEMY_LEFT(t) And Magic_Left(m) + 40 <= tENEMY_LEFT(t) + 50 And Magic_Top(m) + 40 >= tENEMY_TOP(t) And Magic_Top(m) + 40 <= tENEMY_TOP(t) + 50 Then
tENEMY(t) = 0
Magic_Direction(m) = ""
End If
Next t
End If

End If

Next m

frmMain.picMain.Refresh

End Sub
Private Sub tmrMusic_Timer()
If chkMusic.Value = 1 Then
If music.Position = music.Length Then
music.Command = "Stop"
music.Command = "Close"
music.FileName = App.Path & "\Battle-unreleased.mid"
music.Command = "Open"
music.Command = "Play"
End If
End If
End Sub


