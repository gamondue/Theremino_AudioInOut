Imports System.Runtime.InteropServices

Namespace RecPlayer
    Module RecPlayer

        Private wp1 As WavePlayer = New WavePlayer
        Private m_Player As AudioOut
        Private m_Recorder As AudioIn
        Private m_Fifo As FIFO = New FIFO(440000)

        Private Sub Filler(ByRef data() As Int16)
            m_Fifo.GetArray(data)
        End Sub

        Private Sub DataArrived(ByRef data() As Int16)
            m_Fifo.AddArray(data)
        End Sub

        Friend Sub RecPlayStart(Optional ByVal OutputDeviceIndex As Int32 = -1, Optional ByVal InDevice As Int32 = -1)
            RecPlayStop()
            m_Player = New AudioOut()
            m_Recorder = New AudioIn()
            m_Recorder.InputOn(InDevice, 44100, 1, 1024, 8, New AudioIn.BufferDoneEventHandler(AddressOf DataArrived))
            Threading.Thread.Sleep(500)
            m_Player.OutputOn(OutputDeviceIndex, 44100, 1, 1024, 8, New AudioOut.BufferFillEventHandler(AddressOf Filler))
            ' ------------------------------ and add 500 mS of wave
            wp1.PlayStart(OutputDeviceIndex, 1000, -10, WavePlayer.WaveTypes.Sine)
            Threading.Thread.Sleep(500)
            wp1.PlayStop()
        End Sub

        Friend Sub RecPlayStop()
            If m_Player IsNot Nothing Then
                m_Player.OutputOff()
            End If
            If m_Recorder IsNot Nothing Then
                m_Recorder.InputOff()
            End If
            m_Fifo.Flush()
        End Sub

    End Module
End Namespace
