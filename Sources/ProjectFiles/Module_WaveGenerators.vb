Namespace WaveGenerators
    Module WaveGenerators

        Private wp1 As WavePlayer = New WavePlayer
        Private wp2 As WavePlayer = New WavePlayer
        Private wp3 As WavePlayer = New WavePlayer

        Friend Sub StartStop(ByVal OutputDeviceIndex As Int32, _
                             ByVal Gen1 As Boolean, _
                             ByVal Gen2 As Boolean, _
                             ByVal Gen3 As Boolean)
            If Gen1 Then
                wp1.PlayStart(OutputDeviceIndex)
            Else
                wp1.PlayStop()
            End If
            If Gen2 Then
                wp2.PlayStart(OutputDeviceIndex)
            Else
                wp2.PlayStop()
            End If
            If Gen3 Then
                wp3.PlayStart(OutputDeviceIndex)
            Else
                wp3.PlayStop()
            End If
        End Sub

        Friend Sub SetFreq(ByVal f1 As Single, ByVal f2 As Single, ByVal f3 As Single)
            wp1.SetFrequency(f1)
            wp2.SetFrequency(f2)
            wp3.SetFrequency(f3)
        End Sub

        Friend Sub SetOutLevel(ByVal db1 As Single, ByVal db2 As Single, ByVal db3 As Single)
            wp1.SetOutLevel(db1)
            wp2.SetOutLevel(db2)
            wp3.SetOutLevel(db3)
        End Sub

        Friend Sub SetWavetypes(ByVal w1 As Int32, ByVal w2 As Int32, ByVal w3 As Int32)
            wp1.SetWaveType(CType(w1, WavePlayer.WaveTypes))
            wp2.SetWaveType(CType(w2, WavePlayer.WaveTypes))
            wp3.SetWaveType(CType(w3, WavePlayer.WaveTypes))
        End Sub

        Friend Sub CloseAll()
            ' ----------------------------------------------------------------- reduce smoothly the volume
            wp1.PlayPrepareForStop()
            wp2.PlayPrepareForStop()
            wp3.PlayPrepareForStop()
            ' ----------------------------------------------------------------- wait to avoid closing noise
            Dim waittime As Int32 = 0
            waittime = Math.Max(waittime, wp1.GetTotalDelayMillisec)
            waittime = Math.Max(waittime, wp2.GetTotalDelayMillisec)
            waittime = Math.Max(waittime, wp3.GetTotalDelayMillisec)
            Threading.Thread.Sleep(waittime + 10)
            ' ----------------------------------------------------------------- and close
            wp1.PlayStopImmediate()
            wp2.PlayStopImmediate()
            wp3.PlayStopImmediate()
        End Sub

    End Module
End Namespace
