Option Compare Database
Option Explicit

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, ByVal nIDEvent As Long) As Long


Public Event OnTimer()

Private TimerID As Long

'Start timer
Public Sub Startit(IntervalMs As Long)
    TimerID = SetTimer(Application.hWndAccessApp, ObjPtr(Me), IntervalMs, AddressOf Timers.TimerProc)
End Sub

'Stop timer
Public Sub Stopit()
If TimerID <> -1 Then
    KillTimer Application.hWndAccessApp, TimerID
    Debug.Print "(" & Now() & ") Stop Timer Process, TimerID: " & TimerID
    TimerID = 0
End If

End Sub

'Trigger Public event
Public Sub RaiseTimerEvent()
    RaiseEvent OnTimer
End Sub
