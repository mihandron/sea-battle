Attribute VB_Name = "Game"
Option Explicit

Public playerField As Range
Public enemyField As Range
Public Field As initializingField
Public ButtonCriterion As Boolean
Public Choose As Choosing
Public GameStage As Integer
Public ShipCounter As Integer
Public ChooseEnemy As ChoosingEnemy
Public PlayerPlaying As PlayerPlay
Public EnemyPlaying As EnemyPlay
Public TimerP As New TimerPause
Public ClearUp As New ClearUp
Public WinChecking As New WinChecking

Sub theGame()
    ShipCounter = 0
    
    Set Field = New initializingField
    Field.Squarings
    
    Set Choose = New Choosing
    Choose.startingProcedure
    
    Set ChooseEnemy = New ChoosingEnemy

    
    'Debug.Print Selection.Address
    
End Sub

Sub theButton()
    If GameStage = 1 Then
        Choose.Battleship
    ElseIf GameStage = 2 Then
        Choose.Destroyer
    ElseIf GameStage = 3 Then
        Choose.Cruiser
    ElseIf GameStage = 4 Then
        Choose.Submarine
    ElseIf GameStage = 5 Then
        ChooseEnemy.EnemyStartProcedure
        Set PlayerPlaying = New PlayerPlay
        Set EnemyPlaying = New EnemyPlay
        EnemyPlaying.initital
        GameStage = 6
        ClearUp.Clearing
        Field.Information 8
    ElseIf GameStage = 6 Then
        PlayerPlaying.PlayerStep
    End If
End Sub

Sub EnemyButton()
    ThisWorkbook.Worksheets("helpdesk").Range("A1:W12").Value = 0
    enemyField.Interior.Color = RGB(174, 220, 206)
    ChooseEnemy.EnemyStartProcedure
End Sub
