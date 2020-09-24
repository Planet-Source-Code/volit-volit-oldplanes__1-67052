VERSION 5.00
Begin VB.Form MainMenu 
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10365
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5400
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Player Controls"
      Height          =   2655
      Left            =   8760
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   6120
         TabIndex        =   23
         ToolTipText     =   "Close this window"
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Eject: Shift "
         Height          =   195
         Left            =   4680
         TabIndex        =   22
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fire: Ctrl"
         Height          =   195
         Left            =   4680
         TabIndex        =   21
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Right: Arrow right"
         Height          =   195
         Left            =   4680
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Left: Arrow left "
         Height          =   195
         Left            =   4680
         TabIndex        =   19
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Down: Arrow down "
         Height          =   195
         Left            =   4680
         TabIndex        =   18
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Up: Arrow up"
         Height          =   195
         Left            =   4680
         TabIndex        =   17
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Eject: 2 "
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fire: 1 "
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Right: D"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Left: A"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Down: S "
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Up: W"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Player 2"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Player 1"
         Height          =   195
         Left            =   4680
         TabIndex        =   9
         Top             =   480
         Width           =   570
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Single Player Campaign"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      ToolTipText     =   "Start 1 human player game"
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Help"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "See help instructions about the game"
      Top             =   4080
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hall of Fame"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "See player high scores *Not implemented yet*"
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Quit this game"
      Top             =   4680
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Controls"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "See player plane controls"
      Top             =   2880
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Multi Player Skirmish"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Start 2 human players game"
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Single Player Skirmish"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "Start 1 human player game"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   7200
      Picture         =   "MainMenu.frx":030A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "OldPlanes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4200
      TabIndex        =   24
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   0
      Picture         =   "MainMenu.frx":8B5F4
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "v 0.5"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   5160
      Width           =   360
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


A = InputBox("Enter your name   Controls: Arrows,Ctrl,Shift")
PlayerName(0) = A

If A = "" Then

 MsgBox "You must enter a name!"

 A = InputBox("Enter your name   Controls: Arrows,Ctrl,Shift")
 PlayerName(0) = A

End If

If A = "" Then
 MsgBox "Since you havent entered your name you will be given one."
 PlayerName(0) = "Unknown Pilot"
End If


StartSinglePlayerGame
PlaneMenu.LabelName.Caption = PlayerName(0)


MainMenu.Visible = False
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = True
Command2.Font.Bold = False
Command3.Font.Bold = False
Command4.Font.Bold = False
Command5.Font.Bold = False
Command6.Font.Bold = False
Command7.Font.Bold = False

End Sub


Private Sub Command2_Click()


A = InputBox("Player 1 enter your name   Controls: Arrows,Ctrl,Shift")
PlayerName(0) = A

If A = "" Then MsgBox "You must enter a name!": Exit Sub

B = InputBox("Player 2 enter your name   Controls: A,S,W,D,1,2")
PlayerName(1) = B

If B = "" Then MsgBox "You must enter a name!": Exit Sub

StartMultiPlayerGame
PlaneMenu.LabelName.Caption = PlayerName(0)

MainMenu.Visible = False
End Sub


Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
Command2.Font.Bold = True
Command3.Font.Bold = False
Command4.Font.Bold = False
Command5.Font.Bold = False
Command6.Font.Bold = False
Command7.Font.Bold = False
End Sub


Private Sub Command3_Click()
Frame1.Visible = True

Frame1.Left = 2040
Frame1.Top = 1080


End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
Command2.Font.Bold = False
Command3.Font.Bold = True
Command4.Font.Bold = False
Command5.Font.Bold = False
Command6.Font.Bold = False
Command7.Font.Bold = False
End Sub


Private Sub Command4_Click()

Unload MainForm
Unload DebugForm
Unload MainMenu
Unload PlaneMenu
Unload HelpForm
Unload Graphics
Unload GameOver

End

End Sub


Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
Command2.Font.Bold = False
Command3.Font.Bold = False
Command4.Font.Bold = True
Command5.Font.Bold = False
Command6.Font.Bold = False
Command7.Font.Bold = False
End Sub


Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
Command2.Font.Bold = False
Command3.Font.Bold = False
Command4.Font.Bold = False
Command5.Font.Bold = True
Command6.Font.Bold = False
Command7.Font.Bold = False
End Sub


Private Sub Command6_Click()
HelpForm.Visible = True
End Sub


Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
Command2.Font.Bold = False
Command3.Font.Bold = False
Command4.Font.Bold = False
Command5.Font.Bold = False
Command6.Font.Bold = True
Command7.Font.Bold = False
End Sub


Private Sub Command7_Click()

A = InputBox("Enter your name   Controls: Arrows,Ctrl,Shift")
PlayerName(0) = A

If A = "" Then

 MsgBox "You must enter a name!"

 A = InputBox("Enter your name   Controls: Arrows,Ctrl,Shift")
 PlayerName(0) = A

End If

If A = "" Then
 MsgBox "Since you havent entered your name you will be given one."
 PlayerName(0) = "Unknown Pilot"
End If


StartSinglePlayerCampaignGame
PlaneMenu.LabelName.Caption = PlayerName(0)


MainMenu.Visible = False
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font.Bold = False
Command2.Font.Bold = False
Command3.Font.Bold = False
Command4.Font.Bold = False
Command5.Font.Bold = False
Command6.Font.Bold = False
Command7.Font.Bold = True
End Sub


Private Sub Label15_Click()
Frame1.Visible = False
End Sub


