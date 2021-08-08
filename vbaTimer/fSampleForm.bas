Option Compare Database
Option Explicit

Private Const CurrentFormName = "fSampleForm"

Private WithEvents TimerControler As clsTimer
Const TimerControler_ScanInterval = 5000           ' 5 sek

Private Sub TimerControler_OnTimer()
  'Some Event on time
End Sub

Private Sub Form_Load()
  Call TimeControlerLoad(Me.Form)
  Set TimerControler = New clsTimer
  TimerControler.Startit TimerControler_ScanInterval
End Sub

Private Sub Form_Close()
  Call TimerControler.Stopit
End Sub
