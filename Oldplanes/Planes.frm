VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   0  'None
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   816
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   360
      TabIndex        =   45
      Top             =   240
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Return to Main Menu"
         Height          =   195
         Left            =   960
         TabIndex        =   49
         ToolTipText     =   "Return to main menu"
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cancel"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         ToolTipText     =   "Close this window and return to game"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quit Game"
         Height          =   195
         Left            =   2880
         TabIndex        =   47
         ToolTipText     =   "Quit this game"
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Are you sure you want to quit this game?"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   2865
      End
   End
   Begin VB.PictureBox PlayerPlane 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   3360
      Picture         =   "Planes.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   43
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   4320
      Picture         =   "Planes.frx":3042
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   5280
      Picture         =   "Planes.frx":6084
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   41
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   6240
      Picture         =   "Planes.frx":90C6
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   40
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   7200
      Picture         =   "Planes.frx":C108
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   39
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneChute 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   8160
      Picture         =   "Planes.frx":F14A
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   38
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneChute 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   8160
      Picture         =   "Planes.frx":1218C
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   32
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Misson over"
      ForeColor       =   &H8000000D&
      Height          =   3975
      Left            =   2520
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox Button 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2040
         Picture         =   "Planes.frx":151CE
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   44
         ToolTipText     =   "Return to menu"
         Top             =   3480
         Width           =   2610
      End
      Begin VB.PictureBox BlankPic 
         Height          =   375
         Left            =   2040
         Picture         =   "Planes.frx":1853C
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   37
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Medal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   585
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Medal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox RIP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   990
         Index           =   0
         Left            =   4440
         Picture         =   "Planes.frx":1A366
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.PictureBox PlanePic 
         Height          =   375
         Left            =   1440
         Picture         =   "Planes.frx":1D3A8
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   31
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox ChutePic 
         Height          =   375
         Left            =   840
         Picture         =   "Planes.frx":1F1D2
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   30
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox RipPic 
         Height          =   375
         Left            =   240
         Picture         =   "Planes.frx":20FFC
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   29
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox RIP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   990
         Index           =   1
         Left            =   1200
         Picture         =   "Planes.frx":22E26
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Timer Over 
         Enabled         =   0   'False
         Left            =   3240
         Top             =   2880
      End
      Begin VB.PictureBox Medal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   585
         TabIndex        =   25
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Gold 
         AutoSize        =   -1  'True
         Height          =   1020
         Left            =   4200
         Picture         =   "Planes.frx":25E68
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   24
         Top             =   -720
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox Bronze 
         AutoSize        =   -1  'True
         Height          =   1020
         Left            =   3000
         Picture         =   "Planes.frx":28EAA
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   23
         Top             =   -720
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox Silver 
         AutoSize        =   -1  'True
         Height          =   1020
         Left            =   2160
         Picture         =   "Planes.frx":2BEEC
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   22
         Top             =   -720
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox Veteran 
         AutoSize        =   -1  'True
         Height          =   1020
         Left            =   1080
         Picture         =   "Planes.frx":2EF2E
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   21
         Top             =   -720
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox Medal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   585
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Victory 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   2520
         TabIndex        =   50
         Top             =   360
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "PLAYER 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5400
         TabIndex        =   36
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   3360
         TabIndex        =   28
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   1200
         TabIndex        =   27
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   3360
         TabIndex        =   20
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   1200
         TabIndex        =   19
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "PLAYER 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.PictureBox Enemy 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   3
      Left            =   10920
      Picture         =   "Planes.frx":31F70
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   15
      Top             =   7320
      Width           =   960
   End
   Begin VB.PictureBox Player 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   480
      Picture         =   "Planes.frx":34FB2
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   14
      Top             =   2640
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   7320
      Picture         =   "Planes.frx":37FF4
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   13
      Top             =   7920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   6360
      Picture         =   "Planes.frx":3B036
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   12
      Top             =   7920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   5400
      Picture         =   "Planes.frx":3E078
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4440
      Picture         =   "Planes.frx":410BA
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlane 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   3480
      Picture         =   "Planes.frx":440FC
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   9
      Top             =   7920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   7200
      Picture         =   "Planes.frx":4713E
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   5880
      Top             =   120
   End
   Begin VB.PictureBox PlayerPlaneDamage3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   6240
      Picture         =   "Planes.frx":4A180
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   5280
      Picture         =   "Planes.frx":4D1C2
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   4320
      Picture         =   "Planes.frx":50204
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlane 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   3360
      Picture         =   "Planes.frx":53246
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer MainLoop 
      Interval        =   1
      Left            =   5400
      Top             =   120
   End
   Begin VB.PictureBox Player 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   480
      Picture         =   "Planes.frx":56288
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   3600
      Width           =   960
   End
   Begin VB.PictureBox Enemy 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   2
      Left            =   11160
      Picture         =   "Planes.frx":592CA
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   5280
      Width           =   960
   End
   Begin VB.PictureBox Enemy 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   10440
      Picture         =   "Planes.frx":5C30C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   2640
      Width           =   960
   End
   Begin VB.PictureBox Enemy 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   11160
      Picture         =   "Planes.frx":5F34E
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   360
      Width           =   960
   End
   Begin VB.Line Spline 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   688
      X2              =   712
      Y1              =   528
      Y2              =   528
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   0
      X2              =   144
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Spline 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   696
      X2              =   720
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Line Spline 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   656
      X2              =   680
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   16
      X2              =   160
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Spline 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   704
      X2              =   728
      Y1              =   56
      Y2              =   56
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Sub AI()


Dim I As Integer




For I = 0 To 3
If Spline(I).X2 < 0 Then EnemyBulletFlying(I) = False 'Destroy bullets that go off the screen


If EnemyBulletFlying(I) = False And EnemyDead(I) = False Then 'Enemy planes automated for firing bullets
Spline(I).X1 = Enemy(I).Left
Spline(I).X2 = Enemy(I).Left + 40
Spline(I).Y1 = Enemy(I).Top + Enemy(I).Height / 2
Spline(I).Y2 = Enemy(I).Top + Enemy(I).Height / 2


EnemyBulletFlying(I) = True
End If



Next I
End Sub

Public Sub CheckIfEnemyDown()

End Sub

Sub Collisions()
Dim I As Integer
Dim I2 As Integer

For I = 0 To 3
For I2 = 0 To 1

If Spline(I).Y1 > Player(I2).Top And Spline(I).Y1 < Player(I2).Top + Player(I2).Height _
And Spline(I).X1 > Player(I2).Left And Spline(I).X1 < Player(I2).Left + Player(I2).Width And EnemyBulletFlying(I) = True And PlayerDamage(I2) < 3 Then
If PlayerDamage(I2) < 3 And Not PlayerEjected(I2) = True Then PlayerDamage(I2) = PlayerDamage(I2) + EnemyBulletDamage(I)
EnemyBulletFlying(I) = False

End If

Next I2
Next I


For I = 0 To 3
For I2 = 0 To 1

If Line1(I2).Y1 > Enemy(I).Top And Line1(I2).Y1 < Enemy(I).Top + Enemy(I).Height _
And Line1(I2).X1 > Enemy(I).Left And Line1(I2).X1 < Enemy(I).Left + Enemy(I).Width And BulletFlying(I2) = True And EnemyDamage(I) < 3 Then
If EnemyDamage(I) < 3 Then EnemyDamage(I) = EnemyDamage(I) + PlayerBulletDamage(I2)
BulletFlying(I2) = False

If EnemyDamage(I) >= 3 Then PlayerKills(I2) = PlayerKills(I2) + 1 'If we downed an enemy plane add it to our kill list!

End If

Next I2
Next I

End Sub

Sub EndMission()

'THe game will end if:

'IN SINGLE PLAYER
'-if the player is killed
'-if all enemy planes are destroyed
'-if we ejected

'IN MULTIPLAYER
'-if all players are killed
'-if all enemy planes are destroyes
'-if both players have ejected

'IF MULTIPLAYER,WE LOST THE BATTLE,ALL OUR PLANES ARE SHOT DOWN
If PlayerDamage(0) = 4 And PlayerDamage(1) = 4 And MultiPlayer = True Then
Frame1.Visible = True
Over_Timer
Victory.Caption = "Mission Failed"
LevelVictory = False
End If
'We were killed in action,return to main menu

'IF SINGLEPLAYER,AND WE ARE KILLED
If PlayerDamage(0) = 4 And MultiPlayer = False Then
Frame1.Visible = True
Over_Timer
Victory.Caption = "Mission Failed"
LevelVictory = False
End If
'We were killed in action,return to main menu


'WE ARE VICTORIUS
If EnemyDead(0) = True And EnemyDead(1) = True And EnemyDead(2) = True And EnemyDead(3) = True Then
Frame1.Visible = True
Over_Timer
Victory.Caption = "Level completed"
LevelVictory = True
End If
'Lets go to next level


'IF WE EJECTED IN SINGLE PLAYER
If PlayerEjected(0) = True And MultiPlayer = False Then
Frame1.Visible = True
Over_Timer
Victory.Caption = "Mission Failed"
LevelVictory = False
End If
'We ejected and are still alive so we have another chance of playing a mission

'IF WE EJECTED IN MULTIPLAYER
If PlayerEjected(0) = True And PlayerEjected(1) = True And MultiPlayer = True Then
Frame1.Visible = True
Over_Timer
Victory.Caption = "Mission Failed"
LevelVictory = False
End If
'We both ejected and are still alive so we have another chance of playing a mission


End Sub

Sub EnemyPlanes()






For I = 0 To 3
If EnemyDamage(I) = 0 Then Enemy(I).Picture = EnemyPlane.Picture
If EnemyDamage(I) >= 1 And EnemyDamage(I) < 2 Then Enemy(I).Picture = EnemyPlaneDamage1.Picture
If EnemyDamage(I) >= 2 And EnemyDamage(I) < 3 Then Enemy(I).Picture = EnemyPlaneDamage2.Picture
If EnemyDamage(I) >= 3 And EnemyDamage(I) < 4 Then Enemy(I).Picture = EnemyPlaneDamage3.Picture
If EnemyDamage(I) >= 4 Then Enemy(I).Picture = EnemyPlaneDamage4.Picture

Next I
'ASSIGN PICTURES*******************************************


For I = 0 To 3
If Not EnemyDead(I) = True And Not EnemyDamage(I) > 2 Then Enemy(I).Left = Enemy(I).Left - EnemyPlaneSpeed(I) 'ENemy planes moving to the left side of the screen

'ENemy plane is CRASHING'
If Not EnemyDead(I) = True And EnemyDamage(I) > 2 Then Enemy(I).Left = Enemy(I).Left - 0.2

If Enemy(I).Left + Enemy(I).ScaleWidth < 0 Then Enemy(I).Left = MainForm.ScaleWidth 'if enemy planes move too far away on the left redraw them on the far right

If Spline(I).X1 < 0 Or EnemyBulletFlying(I) = False Then EnemyBulletFlying(I) = False: Spline(I).Visible = False 'if enemy bullet is off the screen destroy it

If EnemyBulletFlying(I) = True Then 'Move enemy bullets to the left side
Spline(I).X1 = Spline(I).X1 - EnemyBulletSpeed(I)
Spline(I).X2 = Spline(I).X2 - EnemyBulletSpeed(I)
Spline(I).Visible = True
End If

Next I
End Sub

Sub PlayerPlanes()
For I = 0 To 1 'Movement of player planes
If PlayerDirection(I) = GoUp And PlayerDamage(I) < 3 Then Player(I).Top = Player(I).Top - PlayerPlaneSpeed(I)
If PlayerDirection(I) = GoDown And PlayerDamage(I) < 3 Then Player(I).Top = Player(I).Top + PlayerPlaneSpeed(I)
If PlayerDirection(I) = GoRight And PlayerDamage(I) < 3 Then Player(I).Left = Player(I).Left + PlayerPlaneSpeed(I)
If PlayerDirection(I) = GoLeft And PlayerDamage(I) < 3 Then Player(I).Left = Player(I).Left - PlayerPlaneSpeed(I)
Next I


'ASSIGN PICTURES*******************************************
For I = 0 To 1
If PlayerDamage(I) = 3 Then Player(I).Left = Player(I).Left + 0.2

If PlayerDamage(I) = 0 Then Player(I).Picture = PlayerPlane(I).Picture
If PlayerDamage(I) >= 1 And PlayerDamage(I) < 2 Then Player(I).Picture = PlayerPlaneDamage1(I).Picture
If PlayerDamage(I) >= 2 And PlayerDamage(I) < 3 Then Player(I).Picture = PlayerPlaneDamage2(I).Picture
If PlayerDamage(I) >= 3 And PlayerDamage(I) < 4 Then Player(I).Picture = PlayerPlaneDamage3(I).Picture
If PlayerDamage(I) >= 4 And Not PlayerEjected(I) = True Then Player(I).Picture = PlayerPlaneDamage4(I).Picture
Next I


'Parachuting players from planes
For I = 0 To 1

  If PlayerParachuting(I) = True And Not PlayerDamage(I) = 4 Then
    PlayerTimeTillChute(I) = PlayerTimeTillChute(I) + 1
  End If

  If PlayerTimeTillChute(I) >= 50 Then PlayerEjected(I) = True: PlayerDamage(I) = 4: Player(I).Picture = PlayerPlaneChute(I).Picture

Next I


For I = 0 To 1

    If Line1(I).X1 >= MainForm.ScaleWidth Or BulletFlying(I) = False Then BulletFlying(I) = False: Line1(I).Visible = False 'destroy our bullet if it goes off the screen

    If BulletFlying(I) = True Then 'Add movement to player planes bullets
    Line1(I).X1 = Line1(I).X1 + PlayerBulletSpeed(I)
    Line1(I).X2 = Line1(I).X2 + PlayerBulletSpeed(I)
    Line1(I).Visible = True
    End If

Next I

End Sub


Private Sub Button_Click()



MainForm.Visible = False
Frame1.Visible = False

Playing = False



'Give Medals to PLAYER ONE
If Medal(2).Visible = True Then PlayerSpecialMedals(0) = PlayerSpecialMedals(0) + 1
If Medal(0).Picture = Gold.Picture Then PlayerGoldMedals(0) = PlayerGoldMedals(0) + 1
If Medal(0).Picture = Silver.Picture Then PlayerSilverMedals(0) = PlayerSilverMedals(0) + 1
If Medal(0).Picture = Bronze.Picture Then PlayerBronzeMedals(0) = PlayerBronzeMedals(0) + 1

'GIVE MEDALS TO PLAYER TWO
If Medal(3).Visible = True Then PlayerSpecialMedals(1) = PlayerSpecialMedals(1) + 1
If Medal(1).Picture = Gold.Picture Then PlayerGoldMedals(1) = PlayerGoldMedals(1) + 1
If Medal(1).Picture = Silver.Picture Then PlayerSilverMedals(1) = PlayerSilverMedals(1) + 1
If Medal(1).Picture = Bronze.Picture Then PlayerBronzeMedals(1) = PlayerBronzeMedals(1) + 1



'After closing we got to reset the awards!
'Lets give medals to our player 1
Medal(0).Picture = BlankPic.Picture: Medal(0).Visible = False: Label9.Caption = ""
Medal(2).Picture = BlankPic.Picture: Medal(2).Visible = False: Label11.Caption = ""

RIP(0).Visible = False
RIP(0).Visible = False: RIP(0).Picture = BlankPic.Picture
RIP(0).Visible = False: RIP(0).Picture = BlankPic.Picture

'Lets give medals to our player 2
Medal(1).Picture = BlankPic.Picture: Medal(1).Visible = False: Label8.Caption = ""
Medal(3).Picture = BlankPic.Picture: Medal(3).Visible = False: Label12.Caption = ""

RIP(1).Visible = False
RIP(1).Visible = False: RIP(1).Picture = BlankPic.Picture
RIP(1).Visible = False: RIP(1).Picture = BlankPic.Picture

'if we passed level 5 then this must mean we completed the game!!!!
'Or we are DEAD!!!
If LevelVictory = True And GameLevel = 5 Or LevelVictory = False And PlayerEjected(0) = False Then
GameLevel = 1
LevelVictory = False
MainForm.Visible = False
GameOver.Visible = True


If MultiPlayer = False Then
GameOver.Label3.Visible = False
GameOver.Label4.Visible = False
GameOver.Label5.Visible = False
GameOver.Label6.Visible = False
GameOver.Label1.Visible = False

GameOver.Gold.Visible = False
GameOver.Silver.Visible = False
GameOver.Bronze.Visible = False
GameOver.Veteran.Visible = False
End If
If MultiPlayer = True Then
GameOver.Label3.Visible = True
GameOver.Label4.Visible = True
GameOver.Label5.Visible = True
GameOver.Label6.Visible = True
GameOver.Label1.Visible = True

GameOver.Gold.Visible = True
GameOver.Silver.Visible = True
GameOver.Bronze.Visible = True
GameOver.Veteran.Visible = True
End If


If LevelVictory = True Then GameOver.Info.Caption = "Well, CONGRATULATIONS, you actually finished this game, you passed level 5!!!! Be sure to check out when this game is finished for some more fun."
If LevelVictory = False Then GameOver.Info.Caption = "You were Killed in action, maybe you'l live next time you play."

GameOver.Label2.Caption = PlayerName(0)
GameOver.Label1.Caption = PlayerName(1)

GameOver.Label7.Caption = "Gold medals awarded: " & PlayerGoldMedals(0)
GameOver.Label8.Caption = "Silver medals awarded: " & PlayerSilverMedals(0)
GameOver.Label9.Caption = "Bronze medals awarded: " & PlayerBronzeMedals(0)
GameOver.Label10.Caption = "Special medals awarded: " & PlayerSpecialMedals(0)

GameOver.Label3.Caption = "Gold medals awarded: " & PlayerGoldMedals(1)
GameOver.Label4.Caption = "Silver medals awarded: " & PlayerSilverMedals(1)
GameOver.Label5.Caption = "Bronze medals awarded: " & PlayerBronzeMedals(1)
GameOver.Label6.Caption = "Special medals awarded: " & PlayerSpecialMedals(1)
End If


'If we ejected give another chance
If LevelVictory = False And PlayerEjected(0) = True Then
GameLevel = 1
PlaneMenu.Visible = True
LevelVictory = False
End If

'If we passed the mission advance one level
If LevelVictory = True Then
PlaneMenu.Visible = True
GameLevel = GameLevel + 1
LevelVictory = False
End If
PlaneMenu.Level.Caption = "Game level: " & GameLevel





End Sub

Private Sub Button_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MainForm.Button.Picture = Graphics.Button(1).Picture
End If


End Sub

Private Sub Button_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MainForm.Button.Picture = Graphics.Button(0).Picture
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Not PlayerDamage(0) >= 4 Then
If KeyCode = vbKeyLeft Then PlayerDirection(0) = GoLeft
If KeyCode = vbKeyRight Then PlayerDirection(0) = GoRight
If KeyCode = vbKeyUp Then PlayerDirection(0) = GoUp
If KeyCode = vbKeyDown Then PlayerDirection(0) = GoDown
End If

If KeyCode = vbKeyControl And BulletFlying(0) = False And Not PlayerDead(0) = True Then
Line1(0).X1 = Player(0).Left + Player(0).Width
Line1(0).X2 = (Player(0).Left + Player(0).Width) + 40
Line1(0).Y1 = Player(0).Top + Player(0).Height / 2
Line1(0).Y2 = Player(0).Top + Player(0).Height / 2
BulletFlying(0) = True
End If

If KeyCode = vbKeyShift Then PlayerParachuting(0) = True


'COntrols for player 2,only used in multiplayer

If MultiPlayer = True Then

    If Not PlayerDamage(1) >= 4 Then
    If KeyCode = vbKeyA Then PlayerDirection(1) = GoLeft
    If KeyCode = vbKeyD Then PlayerDirection(1) = GoRight
    If KeyCode = vbKeyW Then PlayerDirection(1) = GoUp
    If KeyCode = vbKeyS Then PlayerDirection(1) = GoDown
    End If

    If KeyCode = vbKey1 And BulletFlying(1) = False And Not PlayerDead(1) = True Then
    Line1(1).X1 = Player(1).Left + Player(1).Width
    Line1(1).X2 = (Player(1).Left + Player(1).Width) + 40
    Line1(1).Y1 = Player(1).Top + Player(1).Height / 2
    Line1(1).Y2 = Player(1).Top + Player(1).Height / 2
    BulletFlying(1) = True
    End If

    If KeyCode = vbKey2 Then PlayerParachuting(1) = True

End If




If KeyCode = vbKeyF1 Then DebugForm.Visible = True
If KeyCode = vbKeyF2 Then DebugForm.Visible = False


If KeyCode = vbKeyEscape Then Frame2.Visible = True: Playing = False

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then PlayerDirection(0) = 0
If KeyCode = vbKeyRight Then PlayerDirection(0) = 0
If KeyCode = vbKeyUp Then PlayerDirection(0) = 0
If KeyCode = vbKeyDown Then PlayerDirection(0) = 0




If KeyCode = vbKeyA Then PlayerDirection(1) = 0
If KeyCode = vbKeyD Then PlayerDirection(1) = 0
If KeyCode = vbKeyW Then PlayerDirection(1) = 0
If KeyCode = vbKeyS Then PlayerDirection(1) = 0

End Sub


Private Sub Label2_Click()
Frame2.Visible = False

MainForm.Visible = False
Frame1.Visible = False

Playing = False


Unload MainForm
Unload DebugForm
Unload MainMenu
Unload PlaneMenu
Unload HelpForm
Unload Graphics
Unload GameOver

End

End Sub

Private Sub Label3_Click()
Frame2.Visible = False
Playing = True
End Sub


Private Sub Label4_Click()
Frame2.Visible = False

MainMenu.Visible = True
MainForm.Visible = False
Frame1.Visible = False

Playing = False


End Sub

Private Sub MainLoop_Timer()







If Playing = True Then

  AI
  Collisions
  EnemyPlanes
  PlayerPlanes
  EndMission
  
End If

End Sub

Private Sub Over_Timer()
Label6.Caption = PlayerName(0)

If MultiPlayer = True Then
  Label7.Caption = PlayerName(1)
Else
  Label7.Caption = ""
End If


'Lets give medals to our player 1
If PlayerKills(0) = 2 Then Medal(0).Picture = Bronze.Picture: Medal(0).Visible = True: Label9.Caption = "Bronze Medal given for average combat performance"
If PlayerKills(0) = 3 Then Medal(0).Picture = Silver.Picture: Medal(0).Visible = True: Label9.Caption = "Silver Medal given for above average combat performance"
If PlayerKills(0) = 4 Then Medal(0).Picture = Gold.Picture: Medal(0).Visible = True: Label9.Caption = "Gold Medal given for excelling at combat,and above average combat capabilities."
If PlayerDamage(0) = 0 Then Medal(2).Picture = Veteran.Picture: Medal(2).Visible = True: Label11.Caption = "Special Medal given for none taken hits to the aircraft."

If PlayerDead(0) = True Then RIP(0).Visible = True: RIP(0).Picture = RipPic.Picture
If PlayerDead(0) = False And Not PlayerEjected(0) = True Then RIP(0).Visible = True: RIP(0).Picture = PlanePic.Picture
If PlayerEjected(0) = True Then RIP(0).Visible = True: RIP(0).Picture = ChutePic.Picture



If MultiPlayer = True Then

    'Lets give medals to our player 2
    If PlayerKills(1) = 2 Then Medal(1).Picture = Bronze.Picture: Medal(1).Visible = True: Label8.Caption = "Bronze Medal given for average combat performance"
    If PlayerKills(1) = 3 Then Medal(1).Picture = Silver.Picture: Medal(1).Visible = True: Label8.Caption = "Silver Medal given for above average combat performance"
    If PlayerKills(1) = 4 Then Medal(1).Picture = Gold.Picture: Medal(1).Visible = True: Label8.Caption = "Gold Medal given for excelling at combat,and above average combat capabilities."
    If PlayerDamage(1) = 0 Then Medal(3).Picture = Veteran.Picture: Medal(3).Visible = True: Label12.Caption = "Special Medal given for none taken hits to the aircraft."

    If PlayerDead(1) = True Then RIP(1).Visible = True:: RIP(1).Picture = RipPic.Picture
    If PlayerDead(1) = False And Not PlayerEjected(1) = True Then RIP(1).Visible = True: RIP(1).Picture = PlanePic.Picture
    If PlayerEjected(1) = True Then RIP(1).Visible = True: RIP(1).Picture = ChutePic.Picture

End If




End Sub

Private Sub Timer2_Timer()
'CRASHINGS OF PLANES


'PLAYER PLANES
For I = 0 To 1
If PlayerDamage(I) >= 2 And PlayerDamage(I) < 3 Then PlayerTimeTillCrash(I) = PlayerTimeTillCrash(I) + 1
If PlayerDamage(I) >= 3 Then PlayerTimeTillCrash(I) = PlayerTimeTillCrash(I) + 10

If PlayerTimeTillCrash(I) > 50 Then PlayerDamage(I) = 4: PlayerDead(I) = True: PlayerTimeTillCrash(I) = 0
Next I



'ENEMY PLANES
For I = 0 To 3
If EnemyDamage(I) >= 2 And EnemyDamage(I) < 3 Then EnemyTimeTillCrash(I) = EnemyTimeTillCrash(I) + 1
If EnemyDamage(I) >= 3 Then EnemyTimeTillCrash(I) = EnemyTimeTillCrash(I) + 10

If EnemyTimeTillCrash(I) > 50 Then EnemyDamage(I) = 4: EnemyDead(I) = True: EnemyTimeTillCrash(I) = 0
Next I



End Sub


