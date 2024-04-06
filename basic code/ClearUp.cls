VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClearUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub Clearing()
    Dim areaToDelete As Range
    Set areaToDelete = Union(ThisWorkbook.Worksheets("Main_field").Range("A1:W1"), _
                            ThisWorkbook.Worksheets("Main_field").Range("A12:W12"), _
                            ThisWorkbook.Worksheets("Main_field").Range("A2:A11"), _
                            ThisWorkbook.Worksheets("Main_field").Range("W2:W11"), _
                            ThisWorkbook.Worksheets("Main_field").Range("L2:L11"))
    ThisWorkbook.Worksheets("Main_field").Range(areaToDelete.Address).Interior.Color = vbWhite
    ThisWorkbook.Worksheets("helpdesk").Range(areaToDelete.Address).Clear
End Sub
