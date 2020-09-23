Attribute VB_Name = "mTimer"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long


Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long

Public lTimer As Long

'// if you break without killing the timer, the IDE will crash.. trust me - haha..
'// I am looking for a better timer.. maybe tomorrow.
Private Sub Clock_Tick(ByVal lInt As Long)

On Error Resume Next

    lTimer = SetTimer(0, 0, lInt, AddressOf Main_Loop)

On Error GoTo 0

End Sub

Public Sub Kill_Timer()
    KillTimer 0, lTimer
End Sub

Public Sub Start_Timer(ByVal lInt As Long)
    Clock_Tick lInt
End Sub

Public Sub Main_Loop()

On Error GoTo Handler

    DoEvents

'Exit Sub
Handler:
Kill_Timer

End Sub
