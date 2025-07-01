
Module Timers
    Friend Delegate Sub TimerProcDelegate()
End Module


Class MultimediaTimer

    ' ===============================================================================================
    '  MULTIMEDIA TIMER - Resolution=1ms - MinInterval=3ms 
    ' ===============================================================================================
    Private Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Int32) As Int32
    Private Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Int32, _
                                                           ByVal uResolution As Int32, _
                                                           ByVal lpFunction As Timer_Delegate, _
                                                           ByVal dwUser As Int32, _
                                                           ByVal uFlags As Int32) As Int32
    Private m_hMMTimer As Int32
    Private m_TimerElapsedDelegate As Timer_Delegate
    Private m_TimerProcDelegate As TimerProcDelegate

    Private Delegate Sub Timer_Delegate(ByVal uID As Int32, ByVal uMsg As Int32, _
                                        ByVal dwUser As Int32, ByVal dw1 As Int32, ByVal dw2 As Int32)
    Private Sub Timer_Elapsed(ByVal uID As Int32, ByVal uMsg As Int32, _
                              ByVal dwUser As Int32, ByVal dw1 As Int32, ByVal dw2 As Int32)
        m_TimerProcDelegate()
    End Sub
    Friend Sub Timer_Start(ByVal millisec As Int32, ByRef TimerProc As TimerProcDelegate)
        If TimerProc Is Nothing Then
            MsgBox("MMTimer_Start - You must provide a TimerProc.")
            Return
        End If
        Timer_Stop()
        If m_hMMTimer = 0 Then
            Const CALLBACK_PERIODIC As Int32 = 1
            m_TimerProcDelegate = TimerProc
            m_TimerElapsedDelegate = AddressOf Timer_Elapsed
            m_hMMTimer = timeSetEvent(millisec, 0, m_TimerElapsedDelegate, 0, CALLBACK_PERIODIC)
        End If
    End Sub
    Friend Sub Timer_Stop()
        If m_hMMTimer <> 0 Then
            timeKillEvent(m_hMMTimer)
            m_hMMTimer = 0
        End If
    End Sub

End Class
