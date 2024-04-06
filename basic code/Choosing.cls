VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Choosing"
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
'-n is for damaged sections of ships

Sub startingProcedure()
    Field.Information (1)
    GameStage = 1
End Sub

Sub Battleship()
    Dim BtShCr As Boolean
    Dim cellCounter As Integer
    Dim i As Range
    Dim isect As Range
    Dim SelectedCells As Range
    
    numberShip = 1
    
    Field.Information (1)
    cellCounter = 0
    BtShCr = True
    Set SelectedCells = Range(Selection.Address)
    For Each i In SelectedCells
        Set isect = Application.Intersect(i, playerField)
        If isect Is Nothing Then
            BtShCr = False
        Else
            cellCounter = cellCounter + 1
        End If
    Next
    Debug.Print cellCounter & " + " & BtShCr _
        & " + " & GameStage
    If cellCounter = 4 And BtShCr = True Then
        SelectedCells.Interior.Color = 568569
        GameStage = 2
        Field.Information (2)
        Background
    Else
        If BtShCr = False Then
            Wronger 13
        Else
            Wronger 12
        End If
    End If
    
End Sub

Sub Destroyer()
    Dim BtShCr As Boolean
    Dim BgCrit As Boolean
    Dim cellCounter As Integer
    Dim i As Range
    Dim isect As Range
    Dim SelectedCells As Range
    
    Field.Information (2)
    cellCounter = 0
    BtShCr = True
    BgCrit = True
    Set SelectedCells = Range(Selection.Address)
    For Each i In SelectedCells
        Set isect = Application.Intersect(i, playerField)
        If isect Is Nothing Then
            BtShCr = False
        Else
            If ThisWorkbook.Worksheets("helpdesk").Range(i.Address).Value = 0 Then
                cellCounter = cellCounter + 1
            Else
                BgCrit = False
            End If
        End If
    Next
    If cellCounter = 3 And BtShCr = True And BgCrit = True Then
        SelectedCells.Interior.Color = 568569
        ShipCounter = ShipCounter + 1
        Background
    Else
        If BtShCr = False Or BgCrit = False Then
            Wronger 13
        Else
            Wronger 12
        End If
    End If
    Debug.Print cellCounter & " + " & BtShCr _
    & " + " & GameStage; " + " & ShipCounter
    If ShipCounter >= 2 Then
        GameStage = 3
        Field.Information (3)
    End If
End Sub

Sub Cruiser()
    Dim BtShCr As Boolean
    Dim BgCrit As Boolean
    Dim cellCounter As Integer
    Dim i As Range
    Dim isect As Range
    Dim SelectedCells As Range
    
    Field.Information (3)
    cellCounter = 0
    BtShCr = True
    BgCrit = True
    Set SelectedCells = Range(Selection.Address)
    For Each i In SelectedCells
        Set isect = Application.Intersect(i, playerField)
        If isect Is Nothing Then
            BtShCr = False
        Else
            If ThisWorkbook.Worksheets("helpdesk").Range(i.Address).Value = 0 Then
                cellCounter = cellCounter + 1
            Else
                BgCrit = False
            End If
        End If
    Next
    Debug.Print cellCounter & " + " & BtShCr _
        & " + " & GameStage
    If cellCounter = 2 And BtShCr = True And BgCrit = True Then
        SelectedCells.Interior.Color = 568569
        ShipCounter = ShipCounter + 1
        Background
    Else
        If BtShCr = False Or BgCrit = False Then
            Wronger 13
        Else
            Wronger 12
        End If
    End If
    If ShipCounter = 5 Then
        GameStage = 4
        Field.Information (4)
    End If
End Sub

Sub Submarine()
    Dim BtShCr As Boolean
    Dim BgCrit As Boolean
    Dim cellCounter As Integer
    Dim i As Range
    Dim isect As Range
    Dim SelectedCells As Range
    
    Field.Information (4)
    cellCounter = 0
    BtShCr = True
    BgCrit = True
    Set SelectedCells = Range(Selection.Address)
    For Each i In SelectedCells
        Set isect = Application.Intersect(i, playerField)
        If isect Is Nothing Then
            BtShCr = False
        Else
            If ThisWorkbook.Worksheets("helpdesk").Range(i.Address).Value = 0 Then
                cellCounter = cellCounter + 1
            Else
                BgCrit = False
            End If
        End If
    Next
    Debug.Print cellCounter & " + " & BtShCr _
        & " + " & GameStage
    If cellCounter = 1 And BtShCr = True And BgCrit = True Then
        SelectedCells.Interior.Color = 568569
        ShipCounter = ShipCounter + 1
        Background
    Else
        If BtShCr = False Or BgCrit = False Then
            Wronger 13
        Else
            Wronger 12
        End If
    End If
    If ShipCounter = 9 Then
        GameStage = 5
        Field.Information (5)
    End If
End Sub

Sub Background()
    Dim BackgroundCells As Range
    Set BackgroundCells = _
        ThisWorkbook.Worksheets("helpdesk").Range(Selection.Address)
    With BackgroundCells
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
    BackgroundCells.Value = numberShip
End Sub

Sub Wronger(num As Integer)
    Field.Information (num)
    TimerP.TimerToPause (1)
    Field.Information (8)
End Sub
