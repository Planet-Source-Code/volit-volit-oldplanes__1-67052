VERSION 5.00
Begin VB.Form Graphics 
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BackButton 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   840
      Picture         =   "Graphics.frx":0000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   28
      Top             =   6240
      Width           =   1110
   End
   Begin VB.PictureBox BackButton 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   840
      Picture         =   "Graphics.frx":1622
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   27
      Top             =   5880
      Width           =   1110
   End
   Begin VB.PictureBox StartButton 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   1920
      Picture         =   "Graphics.frx":2C44
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   26
      Top             =   6240
      Width           =   1110
   End
   Begin VB.PictureBox StartButton 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   1920
      Picture         =   "Graphics.frx":4266
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   25
      Top             =   5880
      Width           =   1110
   End
   Begin VB.PictureBox Button 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   600
      Picture         =   "Graphics.frx":5888
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   174
      TabIndex        =   24
      Top             =   5520
      Width           =   2610
   End
   Begin VB.PictureBox Button 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   600
      Picture         =   "Graphics.frx":8BF6
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   174
      TabIndex        =   23
      Top             =   5160
      Width           =   2610
   End
   Begin VB.PictureBox EnemyPlane 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   480
      Picture         =   "Graphics.frx":BF64
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1440
      Picture         =   "Graphics.frx":EFA6
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2400
      Picture         =   "Graphics.frx":11FE8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   3360
      Picture         =   "Graphics.frx":1502A
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox EnemyPlaneDamage4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4320
      Picture         =   "Graphics.frx":1806C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Blue 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   480
      Picture         =   "Graphics.frx":1B0AE
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox BlueDamage1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1440
      Picture         =   "Graphics.frx":1E0F0
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox BlueDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2400
      Picture         =   "Graphics.frx":21132
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox BlueDamage3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   3360
      Picture         =   "Graphics.frx":24174
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox BlueDamage4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4320
      Picture         =   "Graphics.frx":271B6
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox BlueChute 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   5280
      Picture         =   "Graphics.frx":2A1F8
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Green 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   480
      Picture         =   "Graphics.frx":2D23A
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox GreenDamage1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1440
      Picture         =   "Graphics.frx":3027C
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox GreenDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2400
      Picture         =   "Graphics.frx":332BE
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox GreenDamage3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   3360
      Picture         =   "Graphics.frx":36300
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox GreenDamage4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4320
      Picture         =   "Graphics.frx":39342
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox GreenChute 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   5280
      Picture         =   "Graphics.frx":3C384
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Red 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   480
      Picture         =   "Graphics.frx":3F3C6
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox RedDamage1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1440
      Picture         =   "Graphics.frx":42408
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox RedDamage2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2400
      Picture         =   "Graphics.frx":4544A
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox RedDamage3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   3360
      Picture         =   "Graphics.frx":4848C
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox RedDamage4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4320
      Picture         =   "Graphics.frx":4B4CE
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox RedChute 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   5280
      Picture         =   "Graphics.frx":4E510
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
