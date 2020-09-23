VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   3750
   ClientTop       =   3030
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSplash 
      Interval        =   100
      Left            =   3840
      Top             =   2520
   End
   Begin VB.PictureBox picSplash 
      AutoRedraw      =   -1  'True
      Height          =   4155
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6120
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Splash As Integer

Private Sub Form_DblClick()
Unload frmSplash
End Sub

Private Sub Form_Load()
frmSplash.Move (Screen.Width / 2) - (frmSplash.ScaleWidth / 2), (Screen.Height / 2) - (frmSplash.ScaleHeight / 2)
End Sub


Private Sub Timer1_Timer()

End Sub


Private Sub Form_Unload(Cancel As Integer)
Load frmMain
frmMain.Visible = True
End Sub

Private Sub tmrSplash_Timer()
Randomize
Splash = Splash + 1
For dither = 0 To 1000
X = Int(picSplash.ScaleWidth * Rnd)
Y = Int(picSplash.ScaleHeight * Rnd)
a = BitBlt(frmSplash.hdc, X, Y, 1, 1, picSplash.hdc, X, Y, SRCCOPY)
Next dither
frmSplash.Refresh
If Splash = 50 Then
a = BitBlt(frmSplash.hdc, 0, 0, 400, 250, picSplash.hdc, 0, 0, SRCCOPY)
frmSplash.Refresh
tmrSplash.Enabled = False
wait (5)
Unload frmSplash
End If
End Sub


