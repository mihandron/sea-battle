VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlayerPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub PlayerStep()
    Dim SelectedCells As Range
    Dim inteS As Range
    
    Set SelectedCells = ThisWorkbook.Worksheets("Main_field").Range(Selection(1, 1).Address)
    Set inteS = Application.Intersect(SelectedCells, enemyField)
    If inteS Is Nothing Then
        Wronger
    ElseIf ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value = 0 Or _
            ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value = 1 Then
        Missing inteS, False
        GameStage = 7
        EnemyPlaying.EnemyStep
    ElseIf ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value < 0 Then
        AlreadyShot
    Else
        DamageToEnemy inteS
    End If
End Sub

Sub DamageToEnemy(inteS As Range)
    Dim cell As Range
    Dim criterion As Boolean
    criterion = False
    ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value = -ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value
    For Each cell In ThisWorkbook.Worksheets("helpdesk").Range(enemyField.Address)
        If cell.Value = -ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value Then
            criterion = True
        End If
    Next
    If criterion = False Then
        Field.Information (6)
        For Each cell In ThisWorkbook.Worksheets("helpdesk").Range(enemyField.Address)
            If cell.Value = ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value Then
                Missing ThisWorkbook.Worksheets("Main_field").Range(cell.Address), True, ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value
            End If
        Next
        For Each cell In ThisWorkbook.Worksheets("helpdesk").Range(enemyField.Address)
            If cell.Value = ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value Then
                ThisWorkbook.Worksheets("Main_field").Range(cell.Address).Interior.Color = vbRed
            End If
        Next
        MakeMinusOne
        ClearUp.Clearing
        TimerP.TimerToPause (1)
        Field.Information (8)
        WinChecking.WinCheck True
    Else
        Field.Information (7)
        inteS.Interior.Color = vbYellow
        TimerP.TimerToPause (1)
        Field.Information (8)
    End If
End Sub

Sub Missing(inteS As Range, destroyed As Boolean, Optional Value As Integer)
    inteS.Interior.Color = RGB(127, 185, 165)
    If destroyed = False Then
        ThisWorkbook.Worksheets("helpdesk").Range(inteS.Address).Value = -1
        Misser inteS
    Else
        With inteS
            .Offset(, 1).Interior.Color = RGB(127, 185, 165)
            .Offset(, -1).Interior.Color = RGB(127, 185, 165)
            .Offset(1).Interior.Color = RGB(127, 185, 165)
            .Offset(-1).Interior.Color = RGB(127, 185, 165)
            .Offset(-1, -1).Interior.Color = RGB(127, 185, 165)
            .Offset(1, -1).Interior.Color = RGB(127, 185, 165)
            .Offset(-1, 1).Interior.Color = RGB(127, 185, 165)
            .Offset(1, 1).Interior.Color = RGB(127, 185, 165)
        End With
    End If
End Sub

Sub MakeMinusOne()
    Dim cell As Range
    For Each cell In ThisWorkbook.Worksheets("Main_field").Range(enemyField.Address)
        If cell.Interior.Color = RGB(127, 185, 165) Then
            ThisWorkbook.Worksheets("helpdesk").Range(cell.Address).Value = -1
        End If
    Next
End Sub

Sub Wronger()
    Field.Information (9)
    TimerP.TimerToPause (1)
    Field.Information (8)
End Sub

Sub Misser(inteS As Range)
    Field.Information (10)
    TimerP.TimerToPause (1)
    'Field.Information (8)
End Sub

Sub AlreadyShot()
    Field.Information (11)
    TimerP.TimerToPause (2)
    Field.Information (8)
End Sub
