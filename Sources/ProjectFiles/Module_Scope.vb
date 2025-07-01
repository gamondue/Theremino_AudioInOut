Namespace Scope
    Module Scope

        Private m_SampleFreq As Int32 = 176400 ' 44100 or 176400 or 192000 or 198450
        Private m_Channels As Int32 = 1
        Private m_NumBuffers As Int32 = 4
        Private m_LenBuffers As Int32 = 4096

        Private FifoLen As Int32 = m_SampleFreq * 2
        Private m_fifo As FIFO = New FIFO(FifoLen)
        Private m_Recorder As AudioIn

        Private m_mvDiv As Single = 500
        Private m_msDiv As Single = 100

        Private textFont As Font = New Font("Tahoma", 8)
        Private BackColor As Color = Color.FromArgb(30, 90, 120)
        Private scalePen As Pen = New Pen(Color.FromArgb(45, 58, 70))
        Private tracePen As Pen = New Pen(Color.FromArgb(240, 255, 255))

        Private Sub DataArrived(ByRef data() As Int16)
            m_fifo.AddArray(data)
        End Sub

        Friend Sub Initialize(ByVal pbox As PictureBox, _
                              Optional ByVal InputDeviceIndex As Int32 = -1)
            CloseAll()
            m_Recorder = New AudioIn()
            m_Recorder.InputOn(InputDeviceIndex, _
                               m_SampleFreq, _
                               m_Channels, _
                               m_LenBuffers, _
                               m_NumBuffers, _
                               New AudioIn.BufferDoneEventHandler(AddressOf DataArrived))
        End Sub

        Friend Sub CloseAll()
            If m_Recorder IsNot Nothing Then
                m_Recorder.InputOff()
            End If
        End Sub

        Friend Sub SetVoltage(ByVal mvDiv As Single)
            m_mvDiv = mvDiv
        End Sub

        Friend Sub SetTime(ByVal msDiv As Single)
            m_msDiv = msDiv
        End Sub

        Friend Sub Update(ByVal pbox As PictureBox, _
                          ByVal neg As Boolean, _
                          ByVal trig As Boolean)

            'Dim sw1 As Diagnostics.Stopwatch = New Diagnostics.Stopwatch : sw1.Start()
            ' ----------------------------------------------------------------- 
            InitPictureboxImage(pbox)
            If pbox.Image Is Nothing Then Return
            ' ----------------------------------------------------------------- 
            Dim i As Int32
            Dim x As Single
            Dim y As Single
            Dim oldx As Int32
            Dim oldy As Single
            Dim hpix As Int32 = pbox.Image.Height
            Dim wpix As Int32 = pbox.Image.Width
            Dim h2 As Single = hpix / 2.0F
            Dim kmv As Single = hpix * 1000.0F / 8.0F / m_mvDiv
            Dim kw As Single = m_SampleFreq / 100.0F * m_msDiv / wpix
            Dim preTrigSamples As Int32
            Dim FifoReadIndex As Int32
            ' ----------------------------------------------------------------- 
            Dim g As Graphics = Graphics.FromImage(pbox.Image)
            g.Clear(BackColor)
            ' ----------------------------------------------------------------- draw scale
            For i = 0 To 9
                x = wpix * i \ 10
                g.DrawLine(scalePen, x, 0, x, hpix)
            Next
            For i = 0 To 7
                y = hpix * i \ 8
                g.DrawLine(scalePen, 0, y, wpix, y)
            Next

       

            ' ----------------------------------------------------------------- prepare the FIFO read position
            Dim min As Single = Single.MaxValue
            Dim max As Single = Single.MinValue
            Dim v As Single
            preTrigSamples = Math.Min(2 * CInt(kw * wpix), FifoLen \ 2)
            FifoReadIndex = CInt(m_fifo.GetWriteIndex - preTrigSamples)
            Dim trigpos As Int32 = FifoReadIndex
            ' ----------------------------------------------------------------- adaptive trigger
            Static TrigMin As Single
            Static TrigMax As Single
            min = Single.MaxValue
            max = Single.MinValue
            Dim TrigState As Int32 = 0
            For ix As Int32 = 0 To preTrigSamples \ 2
                v = m_fifo.GetElement(FifoReadIndex + ix) / 32768.0F
                If neg Then v = -v
                If TrigState = 0 AndAlso v > TrigMax Then TrigState = 1
                If TrigState = 1 AndAlso v < TrigMin Then TrigState = 2
                If TrigState = 2 AndAlso v > TrigMax / 10 Then
                    TrigState = -1
                    trigpos = ix - CInt(wpix * kw * 0.11F)
                End If
                If v > max Then max = v
                If v < min Then min = v
            Next
            If trig Then
                Dim margin As Single = (max - min) * 0.1F
                TrigMax = max - margin
                TrigMin = min + margin
                If TrigState = -1 Then
                    FifoReadIndex += trigpos
                Else
                    FifoReadIndex = CInt(m_fifo.GetWriteIndex - kw * wpix - 1)
                End If
            Else
                FifoReadIndex = CInt(m_fifo.GetWriteIndex - kw * wpix - 1)
            End If


            ' TODO - NEW 
            '' ----------------------------------------------------------------- prepare min max and read position
            'Dim min As Single = Single.MaxValue
            'Dim max As Single = Single.MinValue
            'Dim v As Single
            'preTrigSamples = Math.Min(2 * CInt(kw * wpix), FifoLen \ 2)
            'FifoReadIndex = CInt(m_fifo.GetWriteIndex - preTrigSamples)
            'Dim trigpos As Int32 = FifoReadIndex
            '' ----------------------------------------------------------------- adaptive trigger
            'Static TrigMin As Single
            'Static TrigMax As Single
            'min = Single.MaxValue
            'max = Single.MinValue
            'Dim TrigState As Int32 = 0
            'For ix As Int32 = 0 To preTrigSamples \ 2
            '    v = m_fifo.GetElement(FifoReadIndex + ix) / 32768.0F
            '    If neg Then v = -v
            '    If TrigState = 0 AndAlso v > TrigMax Then TrigState = 1
            '    If TrigState = 1 AndAlso v < TrigMin Then TrigState = 2
            '    If TrigState = 2 AndAlso v > TrigMax / 10 Then
            '        TrigState = -1
            '        trigpos = ix - CInt(wpix * kw * 0.11F)
            '    End If
            '    If v > max Then max = v
            '    If v < min Then min = v
            'Next
            'If trig Then
            '    Dim margin As Single = (max - min) * 0.1F
            '    TrigMax = max - margin
            '    TrigMin = min + margin
            '    If TrigState = -1 Then
            '        FifoReadIndex += trigpos
            '    Else
            '        FifoReadIndex = CInt(m_fifo.GetWriteIndex - kw * wpix - 1)
            '    End If
            'Else
            '    FifoReadIndex = CInt(m_fifo.GetWriteIndex - kw * wpix - 1)
            'End If



            ' ----------------------------------------------------------------- wave
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            oldx = -1
            oldy = 0
            For ix As Int32 = 0 To wpix - 1
                v = m_fifo.GetElement(FifoReadIndex + CInt(ix * kw)) / 32768.0F
                If neg Then v = -v
                y = h2 - kmv * v
                If oldx >= 0 Then
                    g.DrawLine(tracePen, oldx, oldy, ix, y)
                End If
                oldx = ix
                oldy = y
            Next
            ' ----------------------------------------------------------------- text
            If m_mvDiv >= 1000 Then
                g.DrawString((m_mvDiv / 1000).ToString(GCI) & " V", textFont, Brushes.Azure, 5, hpix - 20)
            ElseIf m_mvDiv >= 1 Then
                g.DrawString(m_mvDiv.ToString(GCI) & " mV", textFont, Brushes.Azure, 5, hpix - 20)
            Else
                g.DrawString((m_mvDiv * 1000).ToString(GCI) & " uV", textFont, Brushes.Azure, 5, hpix - 20)
            End If
            If m_msDiv >= 1 Then
                g.DrawString(m_msDiv.ToString() & " mS", textFont, Brushes.Azure, wpix - 46, hpix - 20)
            Else
                g.DrawString((m_msDiv * 1000).ToString() & " uS", textFont, Brushes.Azure, wpix - 45, hpix - 20)
            End If
            Dim vpp As Single = max - min
            If vpp > 1.9 Then
                g.FillRectangle(Brushes.Yellow, 82, 5, 70, 13)
                g.DrawString("SATURATION", textFont, Brushes.Black, 82, 5)
            End If
            If vpp > 0.1 Then
                g.DrawString(vpp.ToString("0.00") & " Vpp", textFont, Brushes.Azure, 5, 5)
            Else
                g.DrawString((vpp * 1000).ToString("0.00") & " mVpp", textFont, Brushes.Azure, 5, 5)
            End If
            ' ----------------------------------------------------------------- show
            pbox.Invalidate()
            'ReportString = "Scope time: " & (sw1.Elapsed.TotalMilliseconds * 1000).ToString("0") & " uS"
        End Sub

    End Module
End Namespace
