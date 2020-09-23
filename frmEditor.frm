VERSION 5.00
Begin VB.Form frmEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8595
   ClientLeft      =   1965
   ClientTop       =   1755
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.Frame frmObjects 
      Caption         =   "Objects"
      Height          =   5775
      Left            =   0
      TabIndex        =   77
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   10
         Left            =   720
         Picture         =   "frmEditor.frx":0000
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   88
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   9
         Left            =   120
         Picture         =   "frmEditor.frx":1304
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   87
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   8
         Left            =   1320
         Picture         =   "frmEditor.frx":2608
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   86
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   7
         Left            =   720
         Picture         =   "frmEditor.frx":390C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   85
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   6
         Left            =   120
         Picture         =   "frmEditor.frx":4C10
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   84
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   5
         Left            =   1320
         Picture         =   "frmEditor.frx":5F14
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   83
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   4
         Left            =   720
         Picture         =   "frmEditor.frx":7218
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   82
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   3
         Left            =   120
         Picture         =   "frmEditor.frx":851C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   81
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   2
         Left            =   1320
         Picture         =   "frmEditor.frx":9820
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   80
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmEditor.frx":AB24
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   79
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox object 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmEditor.frx":BE28
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   78
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frmSnow 
      Caption         =   "Dirt/Road"
      Height          =   5775
      Left            =   0
      TabIndex        =   55
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   50
         Left            =   1320
         Picture         =   "frmEditor.frx":D12C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   74
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   49
         Left            =   1320
         Picture         =   "frmEditor.frx":E42E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   56
         Top             =   5040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   48
         Left            =   720
         Picture         =   "frmEditor.frx":F730
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   57
         Top             =   5040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   47
         Left            =   120
         Picture         =   "frmEditor.frx":10A32
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   58
         Top             =   5040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   46
         Left            =   1320
         Picture         =   "frmEditor.frx":11D34
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   59
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   45
         Left            =   720
         Picture         =   "frmEditor.frx":13036
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   60
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   44
         Left            =   120
         Picture         =   "frmEditor.frx":14338
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   61
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   43
         Left            =   720
         Picture         =   "frmEditor.frx":1563A
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   62
         Top             =   3840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   42
         Left            =   120
         Picture         =   "frmEditor.frx":1693C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   64
         Top             =   3840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   41
         Left            =   720
         Picture         =   "frmEditor.frx":17C3E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   63
         Top             =   3240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   40
         Left            =   120
         Picture         =   "frmEditor.frx":18F40
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   65
         Top             =   3240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   39
         Left            =   720
         Picture         =   "frmEditor.frx":1A242
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   66
         Top             =   2640
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   38
         Left            =   120
         Picture         =   "frmEditor.frx":1B544
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   67
         Top             =   2640
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   37
         Left            =   720
         Picture         =   "frmEditor.frx":1C846
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   68
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   36
         Left            =   120
         Picture         =   "frmEditor.frx":1DB48
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   69
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   35
         Left            =   720
         Picture         =   "frmEditor.frx":1EE4A
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   70
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   34
         Left            =   120
         Picture         =   "frmEditor.frx":2014C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   71
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   33
         Left            =   720
         Picture         =   "frmEditor.frx":2144E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   72
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   32
         Left            =   120
         Picture         =   "frmEditor.frx":22750
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   73
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   31
         Left            =   720
         Picture         =   "frmEditor.frx":23A52
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   75
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   30
         Left            =   120
         Picture         =   "frmEditor.frx":24D54
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.FileListBox mademaps 
      Height          =   1065
      Left            =   960
      Pattern         =   "*.map"
      TabIndex        =   54
      Top             =   0
      Width           =   1575
   End
   Begin VB.CheckBox chkRandom 
      Caption         =   "Random"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   53
      Top             =   8280
      Width           =   975
   End
   Begin VB.Frame frmWater 
      Caption         =   "Water/Shore"
      Height          =   3375
      Left            =   0
      TabIndex        =   39
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   63
         Left            =   720
         Picture         =   "frmEditor.frx":26056
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   52
         Top             =   2640
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   62
         Left            =   120
         Picture         =   "frmEditor.frx":27358
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   51
         Top             =   2640
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   61
         Left            =   720
         Picture         =   "frmEditor.frx":2865A
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   50
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   60
         Left            =   120
         Picture         =   "frmEditor.frx":2995C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   49
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   59
         Left            =   1320
         Picture         =   "frmEditor.frx":2AC5E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   48
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   58
         Left            =   720
         Picture         =   "frmEditor.frx":2BF60
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   47
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   57
         Left            =   120
         Picture         =   "frmEditor.frx":2D262
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   46
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   56
         Left            =   1320
         Picture         =   "frmEditor.frx":2E564
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   45
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   55
         Left            =   720
         Picture         =   "frmEditor.frx":2F866
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   44
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   54
         Left            =   120
         Picture         =   "frmEditor.frx":30B68
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   43
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   53
         Left            =   1320
         Picture         =   "frmEditor.frx":31E6A
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   52
         Left            =   720
         Picture         =   "frmEditor.frx":3316C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   51
         Left            =   120
         Picture         =   "frmEditor.frx":3446E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Random"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   7560
      Width           =   1215
   End
   Begin VB.ComboBox tiletype 
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Text            =   "Tile Type"
      ToolTipText     =   "Select tile type here"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.PictureBox picBlank 
      Height          =   615
      Left            =   960
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Map"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Map"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame frmGrass 
      Caption         =   "Grass"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   8
         Left            =   1320
         Picture         =   "frmEditor.frx":35770
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   38
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   7
         Left            =   720
         Picture         =   "frmEditor.frx":36A72
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   37
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   6
         Left            =   120
         Picture         =   "frmEditor.frx":37D74
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   36
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   5
         Left            =   1320
         Picture         =   "frmEditor.frx":39076
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   35
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   4
         Left            =   720
         Picture         =   "frmEditor.frx":3A378
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   3
         Left            =   120
         Picture         =   "frmEditor.frx":3B67A
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   2
         Left            =   1320
         Picture         =   "frmEditor.frx":3C97C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmEditor.frx":3DC7E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmEditor.frx":3EF80
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected"
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   855
      Begin VB.PictureBox selected 
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frmDirt 
      Caption         =   "Dirt/Road"
      Height          =   5775
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   28
         Left            =   1320
         Picture         =   "frmEditor.frx":40282
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   29
         Top             =   5040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   27
         Left            =   720
         Picture         =   "frmEditor.frx":41584
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   31
         Top             =   5040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   26
         Left            =   120
         Picture         =   "frmEditor.frx":42886
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   30
         Top             =   5040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   25
         Left            =   1320
         Picture         =   "frmEditor.frx":43B88
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   33
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   24
         Left            =   720
         Picture         =   "frmEditor.frx":44E8A
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   28
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   23
         Left            =   120
         Picture         =   "frmEditor.frx":4618C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   32
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   22
         Left            =   720
         Picture         =   "frmEditor.frx":4748E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   26
         Top             =   3840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   20
         Left            =   720
         Picture         =   "frmEditor.frx":48790
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   25
         Top             =   3240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   21
         Left            =   120
         Picture         =   "frmEditor.frx":49A92
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   24
         Top             =   3840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   19
         Left            =   120
         Picture         =   "frmEditor.frx":4AD94
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   23
         Top             =   3240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   18
         Left            =   720
         Picture         =   "frmEditor.frx":4C096
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   22
         Top             =   2640
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   17
         Left            =   120
         Picture         =   "frmEditor.frx":4D398
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   20
         Top             =   2640
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   16
         Left            =   720
         Picture         =   "frmEditor.frx":4E69A
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   21
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   15
         Left            =   120
         Picture         =   "frmEditor.frx":4F99C
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   14
         Left            =   720
         Picture         =   "frmEditor.frx":50C9E
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   13
         Left            =   120
         Picture         =   "frmEditor.frx":51FA0
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   14
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   12
         Left            =   720
         Picture         =   "frmEditor.frx":532A2
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   11
         Left            =   120
         Picture         =   "frmEditor.frx":545A4
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   29
         Left            =   1320
         Picture         =   "frmEditor.frx":558A6
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   10
         Left            =   720
         Picture         =   "frmEditor.frx":56BA8
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox tile 
         Height          =   615
         Index           =   9
         Left            =   120
         Picture         =   "frmEditor.frx":57EAA
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Image imgObject 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   0
      Left            =   3000
      Tag             =   "NA"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   0
      Left            =   3000
      Tag             =   "NA"
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj_data(0 To 224)
Private Sub cmdLoad_Click()
mapnumber = InputBox("Enter the map number in X and Y coorinate form. E.G. 'x1y0' would be the map directly left of the first map.")
If mapnumber = "" Then Exit Sub

For t = 0 To 224
imgTile(t).Picture = picBlank.Picture
imgObject(t).Picture = picBlank.Picture
Next t

Open App.Path & "\" & mapnumber & ".map" For Input As #1

For t = 0 To 224
Input #1, texture, walk, obj, obj_tag, obj_dat
If IsNumeric(texture) = True Then
imgTile(t).Picture = tile(texture).Picture
imgTile(t).Tag = texture
End If
If obj = 1 Then
imgObject(t).Picture = object(obj_tag).Picture
obj_data(t) = obj_dat
End If
Next t
Close #1


End Sub

Private Sub cmdRandom_Click()
Randomize
If tiletype.Text = "Grass" Then
If MsgBox("CAUTION: This will erase all work you have done so far on this map! Press OK to continue.", vbYesNo) = vbYes Then
For t = 0 To 224
t2 = Int(9 * Rnd)
imgTile(t).Picture = tile(t2).Picture
imgTile(t).Tag = tile(t2).Tag
Next t
End If
End If

If tiletype.Text = "Water" Then
If MsgBox("CAUTION: This will erase all work you have done so far on this map! Press OK to continue.", vbYesNo) = vbYes Then
For t = 0 To 224
imgTile(t).Picture = tile(55).Picture
imgTile(t).Tag = tile(55).Tag
Next t
End If
End If

If tiletype.Text = "Objects" Then
If MsgBox("CAUTION: This will erase all objects on this map! Press OK to continue.", vbYesNo) = vbYes Then
For t = 0 To 224
imgObject(t).Picture = picBlank.Picture
Next t
End If
End If
End Sub
Private Sub cmdSave_Click()
mapnumber = InputBox("Enter the map number in X and Y coorinate form. E.G. 'x1y0' would be the map directly right of the first map.")
If mapnumber = "" Then
MsgBox "Map NOT Saved!"
Exit Sub
End If
Open App.Path & "\" & mapnumber & ".map" For Output As #1
For t = 0 To 224
block = imgTile(t).Tag
If imgTile(t).Tag >= 51 Then
walk = 0
Else
walk = 1
End If

If IsNumeric(imgObject(t).Tag) Then
If imgObject(t).Tag >= 11 Then
walk = 1
obj = 1
obj_tag = imgObject(t).Tag
Else
walk = 0
obj = 1
obj_tag = imgObject(t).Tag
End If
Else
obj = 0
obj_tag = "NA"
End If

If obj_data(t) = "" Then obj_data(t) = "NA"

Print #1, block & "," & walk & "," & obj & "," & obj_tag & "," & obj_data(t)
Next t
Close #1
mademaps.Refresh
MsgBox "Map Saved Successfully!"
End Sub

Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Load()
tiletype.AddItem "Grass"
tiletype.AddItem "Dirt"
tiletype.AddItem "Snow"
tiletype.AddItem "Water"
tiletype.AddItem "Objects"
For t = 0 To 63
tile(t).Tag = t
Next t
For t = 0 To 10
object(t).Tag = t
Next t
mademaps.Path = App.Path
End Sub


Private Sub Form_Resize()
imgObject(0).Move frmEditor.ScaleWidth - imgObject(0).Width * 15, 0
imgTile(0).Move frmEditor.ScaleWidth - imgTile(0).Width * 15, 0
For t = 1 To 15 * 15 - 1
Load imgObject(t)
Load imgTile(t)
imgTile(t).Visible = True
If t Mod 15 <> 0 Then
imgObject(t).Move imgObject(t - 1).Left + imgObject(t - 1).Width, imgObject(t - 1).Top
imgTile(t).Move imgTile(t - 1).Left + imgTile(t - 1).Width, imgTile(t - 1).Top
Else
imgObject(t).Move imgObject(0).Left, imgObject(t - 1).Top + imgObject(t - 1).Height
imgTile(t).Move imgTile(0).Left, imgTile(t - 1).Top + imgTile(t - 1).Height
End If
Next t
For t = 0 To 15
Line (imgTile(t).Left, 0)-(imgTile(t).Left, frmEditor.ScaleHeight), QBColor(8)
Next t
For t = 0 To 225 Step 10
Line (imgTile(0).Left, imgTile(t).Top)-(frmEditor.ScaleWidth, imgTile(t).Top), QBColor(8)
Next t
End Sub
Private Sub imgObject_Click(Index As Integer)
imgObject(Index).Picture = selected.Picture
imgObject(Index).Tag = selected.Tag
If selected.Tag <= 8 Then obj_data(Index) = "House"
If selected.Tag = 9 Then obj_data(Index) = InputBox("Enter the message for this sign to display.")
End Sub

Private Sub imgTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize
imgTile(Index).Picture = selected.Picture
imgTile(Index).Tag = selected.Tag
If tiletype = "Grass" And chkRandom.Value = 1 Then
t = Int(9 * Rnd)
selected.Picture = tile(t).Picture
selected.Tag = tile(t).Tag
End If
End Sub


Private Sub mademaps_DblClick()
For t = 0 To 224
imgTile(t).Picture = picBlank.Picture
imgObject(t).Picture = picBlank.Picture
Next t

Open App.Path & "\" & mademaps.filename For Input As #1
For t = 0 To 224
Input #1, texture, walk, obj, obj_tag, obj_dat
If IsNumeric(texture) = True Then
imgTile(t).Picture = tile(texture).Picture
imgTile(t).Tag = texture
End If
If obj = 1 Then
imgObject(t).Picture = object(obj_tag).Picture
obj_data(t) = obj_dat
End If
Next t
Close #1
End Sub
Private Sub object_Click(Index As Integer)
selected.Picture = object(Index).Picture
selected.Tag = object(Index).Tag
End Sub

Private Sub tile_Click(Index As Integer)
selected.Picture = tile(Index).Picture
selected.Tag = tile(Index).Tag
End Sub


Private Sub tiletype_Click()
If tiletype.Text = "Grass" Then
For t = 0 To 224
imgObject(t).Visible = False
Next t
frmGrass.Visible = True
frmDirt.Visible = False
frmSnow.Visible = False
frmWater.Visible = False
frmObjects.Visible = False
cmdRandom.Enabled = True
chkRandom.Enabled = True
cmdRandom.Caption = "Random"
End If
If tiletype.Text = "Dirt" Then
For t = 0 To 224
imgObject(t).Visible = False
Next t
frmGrass.Visible = False
frmDirt.Visible = True
frmSnow.Visible = False
frmWater.Visible = False
frmObjects.Visible = False
cmdRandom.Enabled = False
chkRandom.Enabled = False
End If
If tiletype.Text = "Snow" Then
For t = 0 To 224
imgObject(t).Visible = False
Next t
frmGrass.Visible = False
frmDirt.Visible = False
frmSnow.Visible = True
frmWater.Visible = False
frmObjects.Visible = False
cmdRandom.Enabled = False
chkRandom.Enabled = False
End If
If tiletype.Text = "Water" Then
For t = 0 To 224
imgObject(t).Visible = False
Next t
frmGrass.Visible = False
frmDirt.Visible = False
frmSnow.Visible = False
frmWater.Visible = True
frmObjects.Visible = False
cmdRandom.Enabled = True
chkRandom.Enabled = False
cmdRandom.Caption = "Fill"
End If
If tiletype.Text = "Objects" Then
For t = 0 To 224
imgObject(t).Visible = True
Next t
frmGrass.Visible = False
frmDirt.Visible = False
frmSnow.Visible = False
frmWater.Visible = False
frmObjects.Visible = True
cmdRandom.Enabled = True
chkRandom.Enabled = False
cmdRandom.Caption = "Erase Objects"
End If
End Sub
