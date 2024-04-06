VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChoosingEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim numberShip As Integer

'2 is for 4-cell battleship ¹1
'3 is for 3-cell battleship ¹1
'4 is for 3-cell battleship ¹2
'5 is for 2-cell battleship ¹1
'6 is for 2-cell battleship ¹2
'7 is for 2-cell battleship ¹3
'8 is for 1-cell battleship ¹1
'9 is for 1-cell battleship ¹2
'10 is for 1-cell battleship ¹3
'11 is for 1-cell battleship ¹4

Sub EnemyStartProcedure()
    numberShip = 1
    EnemyShip (4)
    EnemyShip (3)
    EnemyShip (3)
    EnemyShip (2)
    EnemyShip (2)
    EnemyShip (2)
    EnemyShip (1)
    EnemyShip (1)
    EnemyShip (1)
    EnemyShip (1)
    
End Sub



Sub EnemyShip(length As Integer)
    Dim Path As Integer
    Dim RndRow As Integer
    Dim RndColumn As Integer
    Dim SelfCrit As Boolean
    Dim i As Range
    Do While True
        Path = Int((2 * Rnd) + 1)
        SelfCrit = True
        If Path = 1 Then 'to bottom
            RndRow = Int((7 * Rnd) + 1)
            RndColumn = Int((10 * Rnd) + 1)
            For Each i In ThisWorkbook.Worksheets("helpdesk").Range(ThisWorkbook.Worksheets("helpdesk").Cells((1 + RndRow), (12 + RndColumn)), ThisWorkbook.Worksheets("helpdesk").Cells((1 + length - 1 + RndRow), (12 + RndColumn)))
                Debug.Print i.Value
                If i.Value <> 0 Then
                    SelfCrit = False
                End If
            Next
            If SelfCrit = True Then
                EnemyBackground (1 + RndRow), (12 + RndColumn), (1 + length - 1 + RndRow), (12 + RndColumn), length
                Exit Do
            End If
        ElseIf Path = 2 Then 'to right
            RndRow = Int((10 * Rnd) + 1)
            RndColumn = Int((7 * Rnd) + 1)
            For Each i In ThisWorkbook.Worksheets("helpdesk").Range(ThisWorkbook.Worksheets("helpdesk").Cells((1 + RndRow), (12 + RndColumn)), ThisWorkbook.Worksheets("helpdesk").Cells((1 + RndRow), (12 + length - 1 + RndColumn)))
                Debug.Print i.Value
                If i.Value <> 0 Then
                    SelfCrit = False
                End If
            Next
            If SelfCrit = True Then
                EnemyBackground (1 + RndRow), (12 + RndColumn), (1 + RndRow), (12 + length - 1 + RndColumn), length
                Exit Do
            End If
        End If
    Loop
End Sub

Sub EnemyBackground(y0 As Integer, x0 As Integer, y1 As Integer, x1 As Integer, length As Integer)
    With ThisWorkbook.Worksheets("helpdesk").Range(ThisWorkbook.Worksheets("Main_field").Cells(y0, x0).Address, ThisWorkbook.Worksheets("Main_field").Cells(y1, x1).Address)
        .Offset(, 1).Value = 1
        .Offset(, -1).Value = 1
        .Offset(1).Value = 1
        .Offset(-1).Value = 1
        .Offset(-1, -1).Value = 1
        .Offset(1, -1).Value = 1
        .Offset(-1, 1).Value = 1
        .Offset(1, 1).Value = 1
    End With
    numberShip = numberShip + 1
    ThisWorkbook.Worksheets("helpdesk").Range(ThisWorkbook.Worksheets("Main_field").Cells(y0, x0).Address, ThisWorkbook.Worksheets("Main_field").Cells(y1, x1).Address).Value = numberShip
    'ThisWorkbook.Worksheets("Main_field").Range(ThisWorkbook.Worksheets("Main_field").Cells(y0, x0).Address, ThisWorkbook.Worksheets("Main_field").Cells(y1, x1).Address).Interior.Color = vbBlack
End Sub
