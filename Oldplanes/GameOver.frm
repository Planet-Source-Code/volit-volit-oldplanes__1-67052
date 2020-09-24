VERSION 5.00
Begin VB.Form GameOver 
   BorderStyle     =   0  'None
   Caption         =   "Game Over"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Main Menu"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   8040
      Picture         =   "GameOver.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   10
      Top             =   5400
      Width           =   960
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   8040
      Picture         =   "GameOver.frx":3042
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   9
      Top             =   3240
      Width           =   960
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   8040
      Picture         =   "GameOver.frx":6084
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   8
      Top             =   4320
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   8040
      Picture         =   "GameOver.frx":90C6
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   960
   End
   Begin VB.PictureBox Veteran 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   240
      Picture         =   "GameOver.frx":C108
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   5520
      Width           =   960
   End
   Begin VB.PictureBox Silver 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   240
      Picture         =   "GameOver.frx":F14A
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   3360
      Width           =   960
   End
   Begin VB.PictureBox Bronze 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   240
      Picture         =   "GameOver.frx":1218C
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   2
      Top             =   4440
      Width           =   960
   End
   Begin VB.PictureBox Gold 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   240
      Picture         =   "GameOver.frx":151CE
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   960
   End
   Begin VB.Shape Shape1 
      Height          =   5055
      Left            =   120
      Top             =   1560
      Width           =   8895
   End
   Begin VB.Line Line4 
      X1              =   64
      X2              =   528
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   5400
      TabIndex        =   18
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   5400
      TabIndex        =   17
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   5400
      TabIndex        =   16
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   5400
      TabIndex        =   15
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1320
      TabIndex        =   13
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1320
      TabIndex        =   12
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "PLAYER 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "PLAYER 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Info 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "GameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MainMenu.Visible = True
GameOver.Visible = False


End Sub

