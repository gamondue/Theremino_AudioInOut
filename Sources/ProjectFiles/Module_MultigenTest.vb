Namespace MultigenTest
    Module MultigenTest

        Private wp() As WavePlayer

        Friend Sub Initialize(ByVal OutputDeviceIndex As Int32)
            ReDim wp(9)
            For i As Int32 = 0 To wp.Length - 1
                wp(i) = New WavePlayer
                wp(i).PlayStart(OutputDeviceIndex, 100 + i * 60, -30, WavePlayer.WaveTypes.Sine)
            Next
        End Sub

        Friend Sub CloseAll()
            If wp Is Nothing Then Return
            For i As Int32 = 0 To wp.Length - 1
                wp(i).PlayPrepareForStop()
            Next
            Threading.Thread.Sleep(wp(0).GetTotalDelayMillisec + 10)
            For i As Int32 = 0 To wp.Length - 1
                wp(i).PlayStopImmediate()
            Next
            ReDim wp(-1)
        End Sub

    End Module
End Namespace
