VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnemyPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim shipIsDamaged As Boolean
Dim enemySearching As Range
Dim memory1 As Range
Dim memory2 As Range
Dim memory3 As Range

Sub initital()
    shipIsDamaged = False
    Set enemySearching = ThisWorkbook.Worksheets("helpdesk2").Range("B2:K11")
End Sub

Sub EnemyStep()
    Dim counter As Integer
    Dim cell As Range
    Dim target As Integer
    Dim targetCell As Range
    Dim criterion As Boolean
    
    Field.Information (17)
    TimerP.TimerToPause (2)
    
    If shipIsDamaged = False Then  ' the case when in the last step enemy DID NOT damaged players ship
        counter = 0
        For Each cell In enemySearching  ' counter to know how much cells that enemy should shot
            If cell.Value = 0 Then
                counter = counter + 1
            End If
        Next
        target = Int((counter - 1 + 1) * Rnd + 1)  ' random target
        counter = 0
        criterion = True
        For Each cell In enemySearching  ' find target cell
            If cell.Value = 0 Then
                counter = counter + 1
            End If
            If counter = target And criterion = True Then
                Set targetCell = cell
                criterion = False
            End If
        Next
        targetShot targetCell
        
    Else  ' the case when in one of the last step enemy DAMAGED players ship
        chooseTheSide
    End If
End Sub


Sub chooseTheSide()
    Dim counter As Integer
    Dim target As Integer
        Dim side1 As Range
        Dim side2 As Range
        Dim side3 As Range
        Dim side4 As Range
            Dim unitedMemory As Range
    Set side1 = Nothing
    Set side2 = Nothing
    Set side3 = Nothing
    Set side4 = Nothing
        Set unitedMemory = Nothing
    counter = 0
    
    If memory2 Is Nothing Then  ' the case when only just one cell hitted
        If memory1.Offset(-1, 0).Value = 0 Then  ' give objects to sides fot choosing random extinct borders and shot places
            counter = counter + 1
            Set side1 = memory1.Offset(-1, 0)
        End If
        If memory1.Offset(1, 0).Value = 0 Then
            counter = counter + 1
            If side1 Is Nothing Then
                Set side1 = memory1.Offset(1, 0)
            Else
                Set side2 = memory1.Offset(1, 0)
            End If
        End If
        If memory1.Offset(0, -1).Value = 0 Then
            counter = counter + 1
            If side1 Is Nothing Then
                Set side1 = memory1.Offset(0, -1)
            ElseIf side2 Is Nothing Then
                Set side2 = memory1.Offset(0, -1)
            Else
                Set side3 = memory1.Offset(0, -1)
            End If
        End If
        If memory1.Offset(0, 1).Value = 0 Then
            counter = counter + 1
            If side1 Is Nothing Then
                Set side1 = memory1.Offset(0, 1)
            ElseIf side2 Is Nothing Then
                Set side2 = memory1.Offset(0, 1)
            ElseIf side3 Is Nothing Then
                Set side3 = memory1.Offset(0, 1)
            Else
                Set side4 = memory1.Offset(0, 1)
            End If
        End If
        target = Int((counter - 1 + 1) * Rnd + 1)
        If target = 1 Then
            targetShot side1
        ElseIf target = 2 Then
            targetShot side2
        ElseIf target = 3 Then
            targetShot side3
        ElseIf target = 4 Then
            targetShot side4
        End If
        
    ElseIf memory3 Is Nothing Then  ' the case when two close parts were damaged
        If memory1.Offset(-1, 0).Value = memory2.Value Then  ' m1 is under m2
            If memory2.Offset(-1, 0).Value = 0 Then
                counter = counter + 1
                Set side1 = memory2.Offset(-1, 0)
            End If
            If memory1.Offset(1, 0).Value = 0 Then
                counter = counter + 1
                If side1 Is Nothing Then
                    Set side1 = memory1.Offset(1, 0)
                Else
                    Set side2 = memory1.Offset(1, 0)
                End If
            End If
        ElseIf memory1.Offset(1, 0).Value = memory2.Value Then  ' m2 is under m1
            If memory2.Offset(1, 0).Value = 0 Then
                counter = counter + 1
                Set side1 = memory2.Offset(1, 0)
            End If
            If memory1.Offset(-1, 0).Value = 0 Then
                counter = counter + 1
                If side1 Is Nothing Then
                    Set side1 = memory1.Offset(-1, 0)
                Else
                    Set side2 = memory1.Offset(-1, 0)
                End If
            End If
        ElseIf memory1.Offset(0, -1).Value = memory2.Value Then  ' m2 is to m1 left
            If memory2.Offset(0, -1).Value = 0 Then
                counter = counter + 1
                Set side1 = memory2.Offset(0, -1)
            End If
            If memory1.Offset(0, 1).Value = 0 Then
                counter = counter + 1
                If side1 Is Nothing Then
                    Set side1 = memory1.Offset(0, 1)
                Else
                    Set side2 = memory1.Offset(0, 1)
                End If
            End If
        ElseIf memory1.Offset(0, 1).Value = memory2.Value Then  ' m1 is to m2 left
            If memory2.Offset(0, 1).Value = 0 Then
                counter = counter + 1
                Set side1 = memory2.Offset(0, 1)
            End If
            If memory1.Offset(0, -1).Value = 0 Then
                counter = counter + 1
                If side1 Is Nothing Then
                    Set side1 = memory1.Offset(0, -1)
                Else
                    Set side2 = memory1.Offset(0, -1)
                End If
            End If
        End If
        target = Int((counter - 1 + 1) * Rnd + 1)
        If target = 1 Then
            targetShot side1
        ElseIf target = 2 Then
            targetShot side2
        End If
    Else  ' the case when three close parts were damaged
        Set unitedMemory = Union(memory1, memory2)
        If unitedMemory.Offset(-1, 0).Cells(1, 1).Value = memory3.Value _
                    Or unitedMemory.Offset(2, 0).Cells(1, 1).Value = memory3.Value Then ' vertical
            Set unitedMemory = Union(memory1, memory2, memory3)
            If unitedMemory.Offset(-1, 0).Cells(1, 1).Value = 0 Then
                counter = counter + 1
                Set side1 = unitedMemory.Offset(-1, 0).Cells(1, 1)
            End If
            If unitedMemory.Offset(3, 0).Cells(1, 1).Value = 0 Then
                counter = counter + 1
                If side1 Is Nothing Then
                    Set side1 = unitedMemory.Offset(3, 0).Cells(1, 1)
                Else
                    Set side2 = unitedMemory.Offset(3, 0).Cells(1, 1)
                End If
            End If

        ElseIf unitedMemory.Offset(0, -1).Cells(1, 1).Value = memory3.Value _
                    Or unitedMemory.Offset(0, 2).Cells(1, 1).Value = memory3.Value Then  ' gorizontal
            Set unitedMemory = Union(memory1, memory2, memory3)
            If unitedMemory.Offset(0, -1).Cells(1, 1).Value = 0 Then
                counter = counter + 1
                Set side1 = unitedMemory.Offset(0, -1).Cells(1, 1)
            End If
            If unitedMemory.Offset(0, 3).Cells(1, 1).Value = 0 Then
                counter = counter + 1
                If side1 Is Nothing Then
                    Set side1 = unitedMemory.Offset(0, 3).Cells(1, 1)
                Else
                    Set side2 = unitedMemory.Offset(0, 3).Cells(1, 1)
                End If
            End If
        End If
        
        target = Int((counter - 1 + 1) * Rnd + 1)
        If target = 1 Then
            targetShot side1
        ElseIf target = 2 Then
            targetShot side2
        End If
    End If
End Sub


Sub targetShot(targetCell As Range)
    If ThisWorkbook.Worksheets("helpdesk").Range(targetCell.Address) > 1 Then
        targetCell.Value = ThisWorkbook.Worksheets("helpdesk").Range(targetCell.Address).Value
        damageToPlayer targetCell
    Else
        targetCell.Value = 1
        enemyMissed targetCell
    End If
End Sub


Sub damageToPlayer(targetCell As Range)
    If targetCell.Value >= 8 Then
        enemyDestroyedShip targetCell, targetCell.Value
    ElseIf targetCell.Value >= 5 Then
        If memory1 Is Nothing Then
            enemyDamagedShip targetCell
        Else
            enemyDestroyedShip Union(targetCell, memory1), targetCell.Value
        End If
    ElseIf targetCell.Value >= 3 Then
        If memory1 Is Nothing Or memory2 Is Nothing Then
            enemyDamagedShip targetCell
        Else
            enemyDestroyedShip Union(targetCell, memory1, memory2), targetCell.Value
        End If
    ElseIf targetCell.Value = 2 Then
        If memory1 Is Nothing Or memory2 Is Nothing Or memory3 Is Nothing Then
            enemyDamagedShip targetCell
        Else
            enemyDestroyedShip Union(targetCell, memory1, memory2, memory3), targetCell.Value
        End If
    End If
End Sub


Sub enemyMissed(targetCell As Range)
    Field.Information (14)
    ThisWorkbook.Worksheets("Main_field").Range(targetCell.Address).Interior.Color = RGB(127, 185, 165)
    TimerP.TimerToPause (1)
    Field.Information (8)
    GameStage = 6
End Sub


Sub enemyDamagedShip(targetCell As Range)
    If memory1 Is Nothing Then
        Set memory1 = targetCell
    ElseIf memory2 Is Nothing Then
        Set memory2 = targetCell
    ElseIf memory3 Is Nothing Then
        Set memory3 = targetCell
    End If
    
    Field.Information (15)
    ThisWorkbook.Worksheets("Main_field").Range(targetCell.Address).Interior.Color = vbYellow
    TimerP.TimerToPause (1)

    shipIsDamaged = True
    EnemyStep
End Sub


Sub enemyDestroyedShip(ship As Range, num As Integer)
    Field.Information (16)
        With ThisWorkbook.Worksheets("Main_field").Range(ship.Address)
            .Offset(, 1).Interior.Color = RGB(127, 185, 165)
            .Offset(, -1).Interior.Color = RGB(127, 185, 165)
            .Offset(1).Interior.Color = RGB(127, 185, 165)
            .Offset(-1).Interior.Color = RGB(127, 185, 165)
            .Offset(-1, -1).Interior.Color = RGB(127, 185, 165)
            .Offset(1, -1).Interior.Color = RGB(127, 185, 165)
            .Offset(-1, 1).Interior.Color = RGB(127, 185, 165)
            .Offset(1, 1).Interior.Color = RGB(127, 185, 165)
        End With
        ClearUp.Clearing
        With ship
            .Offset(, 1).Value = 1
            .Offset(, -1).Value = 1
            .Offset(1).Value = 1
            .Offset(-1).Value = 1
            .Offset(-1, -1).Value = 1
            .Offset(1, -1).Value = 1
            .Offset(-1, 1).Value = 1
            .Offset(1, 1).Value = 1
        End With
        ThisWorkbook.Worksheets("Main_field").Range(ship.Address).Interior.Color = vbRed
        ship.Value = num
    TimerP.TimerToPause (1)
    
    Set memory1 = Nothing
    Set memory2 = Nothing
    Set memory3 = Nothing
     
    shipIsDamaged = False
    WinChecking.WinCheck False
End Sub
