VERSION 5.00
Begin VB.Form PlaneMenu 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   LinkTopic       =   "Form2"
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   764
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 9: "
      Height          =   375
      Index           =   8
      Left            =   7560
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 8: "
      Height          =   375
      Index           =   7
      Left            =   7560
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 7: "
      Height          =   375
      Index           =   6
      Left            =   7560
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 6: "
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 5: "
      Height          =   375
      Index           =   4
      Left            =   7560
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 4: "
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 3: "
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 2: "
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Btn1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "Planes2.frx":0000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   7
      ToolTipText     =   "Go back to main menu"
      Top             =   8400
      Width           =   1110
      Begin VB.Line Line1 
         X1              =   0
         X2              =   0
         Y1              =   24
         Y2              =   0
      End
   End
   Begin VB.PictureBox Btn2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   10320
      Picture         =   "Planes2.frx":1622
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   6
      ToolTipText     =   "Start the game"
      Top             =   8400
      Width           =   1110
   End
   Begin VB.CommandButton MsnBtn 
      Caption         =   "Mission 1: "
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Player 2 Airplane"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Choose a plane for player 2"
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Player 1 Airplane"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Choose a plane for player 1"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   11040
      Top             =   2640
   End
   Begin VB.Label Level 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   10320
      TabIndex        =   19
      Top             =   8040
      Width           =   480
   End
   Begin VB.Label Player1Name 
      AutoSize        =   -1  'True
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   10
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Player2Name 
      AutoSize        =   -1  'True
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   525
      TabIndex        =   8
      Top             =   3000
      Width           =   45
   End
   Begin VB.Image Image3 
      Height          =   810
      Left            =   480
      Picture         =   "Planes2.frx":2C44
      Stretch         =   -1  'True
      ToolTipText     =   "Select Bomber plane"
      Top             =   6120
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   825
      Left            =   2160
      Top             =   5280
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   810
      Left            =   480
      Picture         =   "Planes2.frx":B8A6
      Stretch         =   -1  'True
      ToolTipText     =   "Select Fighter plane"
      Top             =   4920
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   480
      Picture         =   "Planes2.frx":14508
      Stretch         =   -1  'True
      ToolTipText     =   "Select Scout plane"
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label MsnInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   7560
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label LabelName 
      AutoSize        =   -1  'True
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select your plane to fly with."
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Plane information box"
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "PlaneMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn1_Click()
PlaneMenu.Visible = False
MainMenu.Visible = True

End Sub

Private Sub Btn1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PlaneMenu.Btn1.Picture = Graphics.BackButton(1).Picture
End Sub


Private Sub Btn1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PlaneMenu.Btn1.Picture = Graphics.BackButton(0).Picture

End Sub


Private Sub Btn2_Click()

If Campaign = True Then
MsgBox "Sorry, the campaign mode is not finished yet."
Exit Sub
End If


If PlayerPlane(0) = "" Or PlayerPlane(1) = "" Then
 'if we are in 2 player mode!!
 If Command3.Visible = True Then
   MsgBox "All pilots must select their planes!"
   Exit Sub
 End If
End If

'if we are in single player mode
If PlayerPlane(0) = "" And Command3.Visible = False Then
  MsgBox "You haven't selected your plane!"
  Exit Sub
End If



MainForm.Visible = True
PlaneMenu.Visible = False
MainForm.Frame1.Visible = False

Playing = True



'Update player plane positions
MainForm.Player(0).Left = 0
MainForm.Player(1).Left = 0
'Y position
MainForm.Player(0).Top = MainForm.Top / 6
MainForm.Player(1).Top = MainForm.Top / 3

'Reset player Configuration
For I = 0 To 1
PlayerDamage(I) = 0
PlayerDead(I) = False
PlayerKills(I) = 0
PlayerEjected(I) = False
PlayerParachuting(I) = False
PlayerTimeTillChute(I) = 0
PlayerTimeTillCrash(I) = 0
BulletFlying(I) = False
Next I

'Update enemy plane position
MainForm.Enemy(0).Left = 744
MainForm.Enemy(1).Left = 700
MainForm.Enemy(2).Left = 680
MainForm.Enemy(3).Left = 780
'Y position
MainForm.Enemy(0).Top = 50
MainForm.Enemy(1).Top = 200
MainForm.Enemy(2).Top = 400
MainForm.Enemy(3).Top = 500

For I = 0 To 3
EnemyDamage(I) = 0
EnemyDead(I) = False
EnemyBulletFlying(I) = False
EnemyStartX(I) = 0
EnemyStartY(I) = 0
EnemyDead(I) = 0
EnemyTimeTillCrash(I) = 0
EnemyEjected(I) = False
Next I



If GameLevel = 1 Then
For I = 0 To 3
   EnemyBulletDamage(I) = 0.5
   EnemyBulletSpeed(I) = 4
   EnemyPlaneSpeed(I) = 1
Next I
End If
   
If GameLevel = 2 Then
For I = 0 To 3
   EnemyBulletDamage(I) = 1
   EnemyBulletSpeed(I) = 6
   EnemyPlaneSpeed(I) = 1
Next I
End If
   
If GameLevel = 3 Then
For I = 0 To 3
   EnemyBulletDamage(I) = 1
   EnemyBulletSpeed(I) = 8
   EnemyPlaneSpeed(I) = 1.5
Next I
End If
   
If GameLevel = 4 Then
For I = 0 To 3
   EnemyBulletDamage(I) = 1.5
   EnemyBulletSpeed(I) = 10
   EnemyPlaneSpeed(I) = 2
Next I
End If

If GameLevel = 5 Then
For I = 0 To 3
   EnemyBulletDamage(I) = 2
   EnemyBulletSpeed(I) = 15
   EnemyPlaneSpeed(I) = 2.5
Next I
End If


End Sub



Private Sub Btn2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

PlaneMenu.Btn2.Picture = Graphics.StartButton(1).Picture



End Sub

Private Sub Btn2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PlaneMenu.Btn2.Picture = Graphics.StartButton(0).Picture
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
LabelName.Caption = PlayerName(0)



End Sub

Private Sub Command3_Click()
LabelName.Caption = PlayerName(1)



End Sub


Private Sub Command5_Click()
GameLevel = 1
MsnInfo.Caption = "Mission 1: Kill all enemy planes,and return ALIVE."

AllPlayersMustSurvive = True
AllPlanesMustSurvive = False
NonePlanesDamaged = False
DestroyOpposition = True

End Sub


Private Sub Image1_Click()
Shape1.Visible = True
Shape1.Left = Image1.Left
Shape1.Top = Image1.Top


If LabelName = PlayerName(0) Then

   PlayerPlane(0) = "Scout plane"
   Player1Name.Caption = PlayerName(0) & " - flying Scout plane"
   PlayerBulletDamage(0) = 0.5
   PlayerBulletSpeed(0) = 5
   PlayerPlaneSpeed(0) = 2.5
   PlayerPlaneType(0) = Green
   
    MainForm.PlayerPlane(0).Picture = Graphics.Green.Picture
    MainForm.PlayerPlaneDamage1(0).Picture = Graphics.GreenDamage1
    MainForm.PlayerPlaneDamage2(0).Picture = Graphics.GreenDamage2
    MainForm.PlayerPlaneDamage3(0).Picture = Graphics.GreenDamage3
    MainForm.PlayerPlaneDamage4(0).Picture = Graphics.GreenDamage4
    MainForm.PlayerPlaneChute(0).Picture = Graphics.GreenChute

   
End If

If LabelName = PlayerName(1) Then

   PlayerPlane(1) = "Scout plane"
   Player2Name.Caption = PlayerName(1) & " - flying Scout plane"
   PlayerBulletDamage(1) = 0.5
   PlayerBulletSpeed(1) = 5
   PlayerPlaneSpeed(1) = 2.5
   PlayerPlaneType(1) = Green
   
    MainForm.PlayerPlane(1).Picture = Graphics.Green.Picture
    MainForm.PlayerPlaneDamage1(1).Picture = Graphics.GreenDamage1
    MainForm.PlayerPlaneDamage2(1).Picture = Graphics.GreenDamage2
    MainForm.PlayerPlaneDamage3(1).Picture = Graphics.GreenDamage3
    MainForm.PlayerPlaneDamage4(1).Picture = Graphics.GreenDamage4
    MainForm.PlayerPlaneChute(1).Picture = Graphics.GreenChute

End If






Label1.Caption = "Green plane: High plane speed, Low bullet damage, High bullet speed."
Label2.Caption = "Scout plane"
End Sub


Private Sub Image2_Click()
Shape1.Visible = True
Shape1.Left = Image2.Left
Shape1.Top = Image2.Top

If LabelName = PlayerName(0) Then

   PlayerPlane(0) = "Fighter plane"
   Player1Name.Caption = PlayerName(0) & " - flying Fighter plane"
   PlayerBulletDamage(0) = 1
   PlayerBulletSpeed(0) = 3
   PlayerPlaneSpeed(0) = 2
   PlayerPlaneType(0) = Red
   
    MainForm.PlayerPlane(0).Picture = Graphics.Red.Picture
    MainForm.PlayerPlaneDamage1(0).Picture = Graphics.RedDamage1
    MainForm.PlayerPlaneDamage2(0).Picture = Graphics.RedDamage2
    MainForm.PlayerPlaneDamage3(0).Picture = Graphics.RedDamage3
    MainForm.PlayerPlaneDamage4(0).Picture = Graphics.RedDamage4
    MainForm.PlayerPlaneChute(0).Picture = Graphics.RedChute

   
End If

If LabelName = PlayerName(1) Then

   PlayerPlane(1) = "Fighter plane"
   Player2Name.Caption = PlayerName(1) & " - flying Fighter plane"
   PlayerBulletDamage(1) = 1
   PlayerBulletSpeed(1) = 3
   PlayerPlaneSpeed(1) = 2
   PlayerPlaneType(1) = Red
   
    MainForm.PlayerPlane(1).Picture = Graphics.Red.Picture
    MainForm.PlayerPlaneDamage1(1).Picture = Graphics.RedDamage1
    MainForm.PlayerPlaneDamage2(1).Picture = Graphics.RedDamage2
    MainForm.PlayerPlaneDamage3(1).Picture = Graphics.RedDamage3
    MainForm.PlayerPlaneDamage4(1).Picture = Graphics.RedDamage4
    MainForm.PlayerPlaneChute(1).Picture = Graphics.RedChute

End If




Label1.Caption = "Red plane: Medium plane speed, Medium bullet damage, Medium bullet speed."
Label2.Caption = "Fighter plane"
End Sub


Private Sub Image3_Click()
Shape1.Visible = True
Shape1.Left = Image3.Left
Shape1.Top = Image3.Top


If LabelName = PlayerName(0) Then

   PlayerPlane(0) = "Bomber plane"
   Player1Name.Caption = PlayerName(0) & " - flying Bomber plane"
   PlayerBulletDamage(0) = 1.5
   PlayerBulletSpeed(0) = 2
   PlayerPlaneSpeed(0) = 1.5
   PlayerPlaneType(0) = Blue
   
    MainForm.PlayerPlane(0).Picture = Graphics.Blue.Picture
    MainForm.PlayerPlaneDamage1(0).Picture = Graphics.BlueDamage1
    MainForm.PlayerPlaneDamage2(0).Picture = Graphics.BlueDamage2
    MainForm.PlayerPlaneDamage3(0).Picture = Graphics.BlueDamage3
    MainForm.PlayerPlaneDamage4(0).Picture = Graphics.BlueDamage4
    MainForm.PlayerPlaneChute(0).Picture = Graphics.BlueChute

   
End If

If LabelName = PlayerName(1) Then

   PlayerPlane(1) = "Bomber plane"
   Player2Name.Caption = PlayerName(1) & " - flying Bomber plane"
   PlayerBulletDamage(1) = 1.5
   PlayerBulletSpeed(1) = 2
   PlayerPlaneSpeed(1) = 1.5
   PlayerPlaneType(1) = Blue
   
    MainForm.PlayerPlane(1).Picture = Graphics.Blue.Picture
    MainForm.PlayerPlaneDamage1(1).Picture = Graphics.BlueDamage1
    MainForm.PlayerPlaneDamage2(1).Picture = Graphics.BlueDamage2
    MainForm.PlayerPlaneDamage3(1).Picture = Graphics.BlueDamage3
    MainForm.PlayerPlaneDamage4(1).Picture = Graphics.BlueDamage4
    MainForm.PlayerPlaneChute(1).Picture = Graphics.BlueChute

End If







Label1.Caption = "Blue plane: Low plane speed, High bullet damage, Low bullet speed."
Label2.Caption = "Bomber plane"

End Sub

Private Sub Timer1_Timer()

If LabelName = PlayerName(0) And Not PlaneName.Caption = PlayerPlane(0) Then PlaneName.Caption = PlayerPlane(0)
If LabelName = PlayerName(0) And PlayerPlane(0) = "" And Not PlaneName.Caption = "No plane selected" Then PlaneName.Caption = "No plane selected"

If LabelName = PlayerName(1) And Not PlaneName.Caption = PlayerPlane(1) Then PlaneName.Caption = PlayerPlane(1)
If LabelName = PlayerName(1) And PlayerPlane(1) = "" And Not PlaneName.Caption = "No plane selected" Then PlaneName.Caption = "No plane selected"

 




'If Combo1.Text = "" Then
'Label1.Caption = "Choose the airplanes to play with!"
'Picture1.Picture = Plane1.Picture
'End If


'If Combo1.Text = "Red Hawk" Then
'Label1.Caption = "This plane has better defensive capabilities then the Red Sparrow but it lacks its speed. Advantages: Defense"
'Picture1.Picture = Plane1.Picture
'End If

'If Combo1.Text = "Red Sparrow" Then
'Label1.Caption = "This plane has higher speed capabilities then the Red Hawk but it lacks its defensive capabilities. Advantages: Speed"
'Picture1.Picture = Plane2.Picture
'End If

End Sub


