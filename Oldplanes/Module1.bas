Attribute VB_Name = "MainModule"

'***********************************************************
'**************COPYRIGHT INFORMATION************************
'***********************************************************

'You DO NOT have the permission to use the source code
'of this game,or any graphics, for using it in your own
'projects,or modifying this game in any way.



Option Explicit


Public A, B, C, D As String
Public I, I2 As Single
Public GameLevel As Single
Public LevelVictory As Boolean

Public Playing As Boolean
Public Campaign As Boolean

Public MultiPlayer As Boolean
Public PlayerName(0 To 1) As String
Public PlayerPlane(0 To 1) As String

Public PlayerPlaneType(0 To 1) As Byte
'Mission parameters
Public AllPlayersMustSurvive As Boolean
Public AllPlanesMustSurvive As Boolean
Public NonePlanesDamaged As Boolean
Public DestroyOpposition As Boolean
'End Mission parameters

Public PlayerDirection(0 To 2) As Byte


Public BulletFlying(0 To 2) As Boolean

Public PlayerGoldMedals(0 To 1) As Integer
Public PlayerSilverMedals(0 To 1) As Integer
Public PlayerBronzeMedals(0 To 1) As Integer
Public PlayerSpecialMedals(0 To 1) As Integer



Public PlayerDamage(2) As Single
Public PlayerStartX(0 To 2) As Single
Public PlayerStartY(0 To 2) As Single
Public PlayerDead(0 To 2) As Boolean
Public PlayerKills(0 To 2) As Single
Public PlayerEjected(0 To 2) As Boolean

                                                    Public PlayerBulletDamage(0 To 2) As Single
                                                    Public PlayerBulletSpeed(0 To 2) As Single
                                                    Public PlayerPlaneSpeed(0 To 2) As Single


Public PlayerParachuting(0 To 2) As Boolean
Public PlayerTimeTillChute(0 To 2) As Single
Public PlayerTimeTillCrash(0 To 2) As Single



Public EnemyBulletFlying(0 To 3) As Boolean





Public EnemyDamage(3) As Single
Public EnemyStartX(0 To 3) As Single
Public EnemyStartY(0 To 3) As Single
Public EnemyDead(0 To 3) As Single
Public EnemyTimeTillCrash(0 To 3) As Single
Public EnemyEjected(0 To 3) As Boolean

                                                    Public EnemyBulletDamage(0 To 3) As Single
                                                    Public EnemyBulletSpeed(0 To 3) As Single
                                                    Public EnemyPlaneSpeed(0 To 3) As Single

Public Const GoUp = 1
Public Const GoDown = 2
Public Const GoLeft = 3
Public Const GoRight = 4

Const Red = 50
Const Green = 51
Const Blue = 52


Public Const EnemyPlane1Speed = 1.5
Public Const EnemyPlane2Speed = 1.5
Public Const Plane1Speed = 0
Public Const Plane2Speed = 0




Public Sub GoToNextLevel()

End Sub


Public Sub StartSinglePlayerGame()
GameLevel = 1
Campaign = False

PlaneMenu.Level.Caption = "Game level: " & GameLevel

PlaneMenu.Visible = True

PlaneMenu.Command2.Visible = False
PlaneMenu.Command3.Visible = False

PlaneMenu.Player1Name.Caption = PlayerName(0)
PlaneMenu.Player2Name.Caption = ""


MainForm.Player(1).Visible = False
MultiPlayer = False

For I = 0 To 8
PlaneMenu.MsnBtn(I).Visible = False
Next I
PlaneMenu.MsnInfo.Visible = False

End Sub


Public Sub StartSinglePlayerCampaignGame()
GameLevel = 1
Campaign = True

PlaneMenu.Visible = True

PlaneMenu.Command2.Visible = False
PlaneMenu.Command3.Visible = False

PlaneMenu.Player1Name.Caption = PlayerName(0)
PlaneMenu.Player2Name.Caption = ""


MainForm.Player(1).Visible = False
MultiPlayer = False

For I = 0 To 8
PlaneMenu.MsnBtn(I).Visible = True
Next I
PlaneMenu.MsnInfo.Visible = True

End Sub


Public Sub StartMultiPlayerGame()
'Star at level1
GameLevel = 1
'We are not in campaign mode
Campaign = False

PlaneMenu.Level.Caption = "Game level: " & GameLevel

PlaneMenu.Visible = True

PlaneMenu.Command2.Visible = True
PlaneMenu.Command3.Visible = True

PlaneMenu.Command2.Caption = PlayerName(0) & "'s Plane"
PlaneMenu.Command3.Caption = PlayerName(1) & "'s Plane"

PlaneMenu.Player1Name.Caption = PlayerName(0)
PlaneMenu.Player2Name.Caption = PlayerName(1)

MainForm.Player(1).Visible = True
MultiPlayer = True

For I = 0 To 8
PlaneMenu.MsnBtn(I).Visible = False
Next I
PlaneMenu.MsnInfo.Visible = False

End Sub

