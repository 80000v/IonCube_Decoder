Attribute VB_Name = "basAPITimer"
Option Explicit

' Module: basAPITimer
' API Timer Functions - from Litwin, Getz, Gilbert

Declare Function wu_GetTime Lib "winmm.dll" Alias _
    "timeGetTime" () As Long
Private mlStartTime As Long

Public Sub ap_StartTimer()
    mlStartTime = wu_GetTime()
End Sub

Public Function ap_EndTimer() As Long
    ap_EndTimer = wu_GetTime() - mlStartTime
End Function

