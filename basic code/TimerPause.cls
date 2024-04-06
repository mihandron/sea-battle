VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerPause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Sub TimerToPause(pause As Integer)
    Dim Start As Single
    Start = Timer
    Do While Timer < Start + pause
        DoEvents
    Loop
End Sub
