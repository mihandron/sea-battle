VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim PlayerWin As Integer
Dim EnemyWin As Integer

Sub init()
    PlayerWin = 0
    EnemyWin = 0
End Sub

Sub WinCheck(isItPlayer As Boolean)
    If isItPlayer = True Then
        PlayerWin = PlayerWin + 1
        If PlayerWin = 10 Then
            playerWinMessage
        End If
    Else
        EnemyWin = EnemyWin + 1
        If EnemyWin = 10 Then
            enemyWinMessage
        Else
            EnemyPlaying.EnemyStep
        End If
    End If
End Sub

Sub playerWinMessage()
    GameStage = 7
    Field.Information (18)
    enemyField.Interior.Color = RGB(127, 185, 165)
    ClearUp.Clearing
End Sub

Sub enemyWinMessage()
    GameStage = 7
    Field.Information (19)
    playerField.Interior.Color = RGB(127, 185, 165)
    ClearUp.Clearing
End Sub
