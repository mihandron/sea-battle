VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "initializingField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub Squarings()
    Set playerField = ThisWorkbook.Worksheets("Main_field").Range("B2:K11")
    Set enemyField = ThisWorkbook.Worksheets("Main_field").Range("M2:V11")
    playerField.Interior.Color = RGB(174, 220, 206)
    enemyField.Interior.Color = RGB(174, 220, 206)
    ThisWorkbook.Worksheets("Main_Field").Range("Y5:AM5").Interior.Color = vbWhite
    ThisWorkbook.Worksheets("helpdesk").Range("A1:W12").Value = 0
    ThisWorkbook.Worksheets("helpdesk2").Range("B2:K11").Value = 0
    WinChecking.init
    With playerField.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With playerField.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With playerField.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With playerField.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With enemyField.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With enemyField.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With enemyField.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With enemyField.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With ThisWorkbook.Worksheets("Main_Field").Range("Y2")
        .Value = "Tips: "
        .Font.Name = "Times New Roman"
        .Font.Size = 16
    End With
    With ThisWorkbook.Worksheets("Main_Field").Range("Y3")
        .Font.Name = "Times New Roman"
        .Font.Size = 16
    End With
End Sub

Sub Information(message)
    If message = 1 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Choose cells for a battleship and press the button"
        ThisWorkbook.Worksheets("Main_Field").Range("Y5:AB5").Interior.Color = 568569
    ElseIf message = 2 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y5:AM5").Interior.Color = vbWhite
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Choose cells for two destroyers and press the button"
        ThisWorkbook.Worksheets("Main_Field").Range("Y5:AA5,AC5:AE5").Interior.Color = 568569
    ElseIf message = 3 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y5:AM5").Interior.Color = vbWhite
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Choose cells for three cruisers and press the button"
        ThisWorkbook.Worksheets("Main_Field").Range("Y5:Z5,AB5:AC5,AE5:AF5").Interior.Color = 568569
    ElseIf message = 4 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y5:AM5").Interior.Color = vbWhite
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Choose cells for four submarines and press the button"
        ThisWorkbook.Worksheets("Main_Field").Range("Y5,AA5,AC5,AE5").Interior.Color = 568569
    ElseIf message = 5 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y5:AM5").Interior.Color = vbWhite
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Tap Next to autoplace enemy ships"
    ElseIf message = 6 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "The ship is destroyed"
    ElseIf message = 7 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "You damaged the ship"
    ElseIf message = 8 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Choose cell to attack enemy"
    ElseIf message = 9 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Wrong cell!!!"
    ElseIf message = 10 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "You missed!!!"
    ElseIf message = 11 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "You've already shot there!!!"
    ElseIf message = 12 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Wrong amount of cells for the ship!!!"
    ElseIf message = 13 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Wrong place for the ship!!!"
    ElseIf message = 14 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Enemy missed!!!"
    ElseIf message = 15 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Enemy damaged your ship!!!"
    ElseIf message = 16 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Enemy destroyed your ship!!!"
    ElseIf message = 17 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "Enemy is thinking..."
    ElseIf message = 18 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "CONGRATULATIONS!!! YOU WON!!!"
    ElseIf message = 19 Then
        ThisWorkbook.Worksheets("Main_Field").Range("Y3").Value = _
            "YOU LOST..."
    End If
End Sub
