VERSION 5.00
Begin VB.Form HelpForm 
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Blue 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   9480
      Picture         =   "HelpForm.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   23
      Top             =   1200
      Width           =   960
   End
   Begin VB.PictureBox Green 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   8640
      Picture         =   "HelpForm.frx":3042
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   22
      Top             =   1200
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   720
      Picture         =   "HelpForm.frx":6084
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   19
      Top             =   1680
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1680
      Picture         =   "HelpForm.frx":90C6
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   18
      Top             =   1680
      Width           =   960
   End
   Begin VB.PictureBox Player 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   8160
      Picture         =   "HelpForm.frx":C108
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   9
      Top             =   3360
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneChute 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   6360
      Picture         =   "HelpForm.frx":F14A
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   8
      Top             =   3960
      Width           =   960
   End
   Begin VB.PictureBox PlayerPlaneDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4800
      Picture         =   "HelpForm.frx":1218C
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   7
      Top             =   4080
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< Main Menu"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Return to Main Menu"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Enemy 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   3360
      Picture         =   "HelpForm.frx":151CE
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   3600
      Width           =   960
   End
   Begin VB.PictureBox Player 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   7800
      Picture         =   "HelpForm.frx":18210
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "http://visualbasicfreegames.blogspot.com/"
      Height          =   195
      Left            =   7200
      TabIndex        =   25
      Top             =   5400
      Width           =   3045
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Find more about this game at:"
      Height          =   195
      Left            =   4920
      TabIndex        =   24
      Top             =   5400
      Width           =   2085
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8040
      TabIndex        =   21
      Top             =   4440
      Width           =   165
   End
   Begin VB.Label Label13 
      Caption         =   "Following these instruction will keep you alive."
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   $"HelpForm.frx":1B252
      Height          =   1365
      Left            =   840
      TabIndex        =   17
      Top             =   120
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line22 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   1200
      X2              =   1560
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Line Line21 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   1560
      X2              =   1320
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line20 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   1080
      Y2              =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "6."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   15
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5160
      TabIndex        =   14
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   12
      Top             =   0
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   11
      Top             =   0
      Width           =   270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   1
      X1              =   8520
      X2              =   10200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Defend yourself by moving the plane and attacking the enemy"
      Height          =   615
      Left            =   8520
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line19 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9240
      Y1              =   3120
      Y2              =   3600
   End
   Begin VB.Line Line18 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9120
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9360
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line16 
      X1              =   6240
      X2              =   6120
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line15 
      X1              =   6240
      X2              =   6120
      Y1              =   4440
      Y2              =   4320
   End
   Begin VB.Line Line14 
      X1              =   5640
      X2              =   6240
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   $"HelpForm.frx":1B2F2
      Height          =   1560
      Left            =   5520
      TabIndex        =   6
      Top             =   2520
      Width           =   1785
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   6120
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   6360
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   3480
      Y2              =   3240
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3480
      X2              =   3840
      Y1              =   3000
      Y2              =   3480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "These are the enemy planes,try to avoid their fire"
      Height          =   675
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   1860
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   8280
      X2              =   8400
      Y1              =   960
      Y2              =   840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   8280
      X2              =   8160
      Y1              =   960
      Y2              =   840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   8280
      X2              =   8280
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "This are the planes you are flying,they look exactly the same like those of the enemy, except they are red,green or blue coloured."
      Height          =   975
      Left            =   8400
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5400
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5400
      X2              =   5400
      Y1              =   1200
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5760
      X2              =   5400
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   $"HelpForm.frx":1B39E
      Height          =   1950
      Left            =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   1665
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderStyle     =   4  'Dash-Dot
      Index           =   0
      X1              =   4440
      X2              =   6120
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
MainMenu.Visible = True
HelpForm.Visible = False
End Sub


