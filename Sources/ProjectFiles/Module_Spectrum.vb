Module Spectrum

    Private m_FourierSamples As Int32 = 16384 ' 8192 with 44100 or 16384 with 88200 or 32768 with 176400
    Private m_SamplesPerSec As Int32 = 88200 ' 44100 or 176400 or 192000 or 198450
    Private m_Channels As Int32 = 1
    Private m_NumBuffers As Int32 = 4
    Private m_LenBuffers As Int32 = 4096

    Private m_BackColor As Color = Color.FromArgb(30, 90, 120)
    Private m_scalePen As Pen = New Pen(Color.FromArgb(35, 48, 60))
    Private m_tracePen As Pen = New Pen(Color.FromArgb(240, 255, 255))
    Private m_TextBrush As Brush = Brushes.Azure
    Private m_leftborder As Int32 = 28
    Private m_downborder As Int32 = 16

    Private m_ddx As Int32
    Private m_ddy As Int32
    Private m_dispx As Int32
    Private m_dispy As Int32
    Private m_fmin As Single
    Private m_fmax As Single
    Private m_dbmin As Single
    Private m_dbmax As Single
    Private m_xlog As Boolean
    Private m_ylog As Boolean
    Private m_windowtype As FHT.WindowTypes

    Private Const kdb As Double = 8.685889638

    Private m_fifo As FIFO = New FIFO(200000)
    Private m_fht As FHT = New FHT
    Private m_Recorder As AudioIn
    Private m_InputBuffer() As Int16

    Private Sub DataArrived(ByRef data() As Int16)
        m_fifo.AddArray(data)
    End Sub

    Friend Sub Initialize(ByVal pbox As PictureBox, _
                          Optional ByVal InputDeviceIndex As Int32 = -1)
        CloseAll()
        m_Recorder = New AudioIn()
        m_Recorder.InputOn(InputDeviceIndex, _
                           m_SamplesPerSec, _
                           m_Channels, _
                           m_LenBuffers, _
                           m_NumBuffers, _
                           New AudioIn.BufferDoneEventHandler(AddressOf DataArrived))
        ReDim m_InputBuffer(m_FourierSamples - 1)
    End Sub

    Friend Sub CloseAll()
        If m_Recorder IsNot Nothing Then
            m_Recorder.InputOff()
            m_Recorder = Nothing
        End If
    End Sub

    Friend Sub Update(ByVal pbox As PictureBox, _
                      ByVal fmin As Single, _
                      ByVal fmax As Single, _
                      ByVal dbmin As Single, _
                      ByVal dbmax As Single, _
                      ByVal xlog As Boolean, _
                      ByVal ylog As Boolean, _
                      ByVal speed As Single, _
                      ByVal windowtype As FHT.WindowTypes)

        'Dim sw1 As Diagnostics.Stopwatch = New Diagnostics.Stopwatch : sw1.Start()

        InitPictureboxImage(pbox)
        If pbox.Image Is Nothing Then Return

        If pbox.Image.Width <> m_ddx OrElse _
                                pbox.Image.Height <> m_ddy OrElse _
                                fmin <> m_fmin OrElse _
                                fmax <> m_fmax OrElse _
                                dbmin <> m_dbmin OrElse _
                                dbmax <> m_dbmax OrElse _
                                xlog <> m_xlog OrElse _
                                ylog <> m_ylog OrElse _
                                windowtype <> m_windowtype Then

            m_ddx = pbox.Image.Width
            m_ddy = pbox.Image.Height
            m_fmin = fmin
            m_fmax = fmax
            m_dbmin = dbmin
            m_dbmax = dbmax
            m_xlog = xlog
            m_ylog = ylog
            m_dispx = m_ddx - m_leftborder
            m_dispy = m_ddy - m_downborder
            m_windowtype = windowtype
            speed = 1
        End If

        Dim g As Graphics = Graphics.FromImage(pbox.Image)
        g.Clear(m_BackColor)

        If m_dbmin >= m_dbmax Or m_fmin >= m_fmax Or fmax >= m_SamplesPerSec \ 2 Then
            g.DrawString("Invalid params", New Font("Arial", 11), m_TextBrush, 10, 10)
            pbox.Invalidate()
            Return
        End If

        CreateGrid(g)

        m_fifo.SetReadIndex(m_fifo.GetWriteIndex - m_InputBuffer.Length)
        m_fifo.GetArray(m_InputBuffer)
        m_fht.ExecuteFHT(m_InputBuffer, m_windowtype)
        m_fht.ExtractBands(m_dispx, m_fmin, m_fmax, m_dbmin, m_dbmax, m_xlog, m_ylog, m_SamplesPerSec, speed)

        Dim ix As Int32
        Dim x As Single
        Dim y As Single
        Dim oldx As Single = -1
        Dim oldy As Single
        For ix = 0 To m_dispx - 1
            x = ix + m_leftborder
            y = m_fht.BandsBuffer(ix)
            y = m_dispy - m_dispy * y
            If oldx >= 0 Then
                g.DrawLine(m_tracePen, oldx, oldy, x, y)
            End If
            oldx = x
            oldy = y
        Next

        pbox.Invalidate()

        'ReportString = "Spectrum Update time: " & (sw1.Elapsed.TotalMilliseconds * 1000).ToString("0") & " uS"
    End Sub


    Private Sub CreateGrid(ByVal g As Graphics)
        '
        g.DrawLine(m_scalePen, m_leftborder, m_dispy, m_ddx, m_dispy)
        g.DrawLine(m_scalePen, m_leftborder, 0, m_leftborder, m_dispy)
        '
        Dim s As String
        Dim max, min As Double
        Dim px As Int32
        Dim klog As Double = (m_dispx - 1) / Math.Log(m_fmax / m_fmin)
        Dim fontsize As Int32 = m_ddx \ 40
        If fontsize > 7 Then fontsize = 7
        If fontsize < 4 Then fontsize = 4
        Dim xFont As Font = New Font("Arial", fontsize)
        fontsize = m_ddy \ 40
        If fontsize > 7 Then fontsize = 7
        If fontsize < 4 Then fontsize = 4
        Dim yFont As Font = New Font("Arial", fontsize)
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        '
        If m_xlog And m_fmax / m_fmin >= 10 Then
            ' ---------------------------------------------------------------------- X AXIS LOG
            Dim decmin As Double = 100000
            Dim decmax As Double = 0.01
            While decmin > m_fmin : decmin /= 10 : End While
            While decmax < m_fmax : decmax *= 10 : End While
            Dim f As Double = decmin
            While f <= decmax
                For i As Int32 = 1 To 9
                    Dim x As Double = klog * Math.Log(f * i / m_fmin)
                    If x >= 0 AndAlso x <= m_dispx Then
                        px = CInt(Math.Floor(m_leftborder + x + 0.5))
                        If i = 1 Then
                            g.DrawLine(m_scalePen, px, 0, px, m_ddy - m_downborder + 4)
                            If (px > m_ddx - 18) Then px = m_ddx - 18
                            If (f = 0.1) Then
                                g.DrawString(".1 Hz", xFont, m_TextBrush, px - 8, m_ddy - 12)
                            ElseIf f = 1 Then
                                g.DrawString("1 Hz", xFont, m_TextBrush, px - 4, m_ddy - 12)
                            ElseIf (f = 10) Then
                                g.DrawString("10 Hz", xFont, m_TextBrush, px - 8, m_ddy - 12)
                            ElseIf (f = 100) Then
                                g.DrawString("100 Hz", xFont, m_TextBrush, px - 12, m_ddy - 12)
                            ElseIf (f = 1000) Then
                                g.DrawString("1 KHz", xFont, m_TextBrush, px - 10, m_ddy - 12)
                            ElseIf (f = 10000) Then
                                g.DrawString("10 KHz", xFont, m_TextBrush, px - 14, m_ddy - 12)
                            End If
                        ElseIf i = 2 Then
                            g.DrawLine(m_scalePen, px, 0, px, m_ddy - m_downborder + 4)
                            If m_fmax / m_fmin < 200 Then
                                If (px > m_ddx - 18) Then px = m_ddx - 18
                                If (f = 0.1) Then
                                    g.DrawString(".2 Hz", xFont, m_TextBrush, px - 8, m_ddy - 12)
                                ElseIf f = 1 Then
                                    g.DrawString("2 Hz", xFont, m_TextBrush, px - 4, m_ddy - 12)
                                ElseIf (f = 10) Then
                                    g.DrawString("20 Hz", xFont, m_TextBrush, px - 8, m_ddy - 12)
                                ElseIf (f = 100) Then
                                    g.DrawString("200 Hz", xFont, m_TextBrush, px - 12, m_ddy - 12)
                                ElseIf (f = 1000) Then
                                    g.DrawString("2 KHz", xFont, m_TextBrush, px - 10, m_ddy - 12)
                                ElseIf (f = 10000) Then
                                    g.DrawString("20 KHz", xFont, m_TextBrush, px - 14, m_ddy - 12)
                                End If
                            End If
                        ElseIf i = 5 Then
                            g.DrawLine(m_scalePen, px, 0, px, m_ddy - m_downborder + 4)
                            If m_fmax / m_fmin < 100 Then
                                If (px > m_ddx - 18) Then px = m_ddx - 18
                                If (f = 0.1) Then
                                    g.DrawString(".5 Hz", xFont, m_TextBrush, px - 8, m_ddy - 12)
                                ElseIf f = 1 Then
                                    g.DrawString("5 Hz", xFont, m_TextBrush, px - 4, m_ddy - 12)
                                ElseIf (f = 10) Then
                                    g.DrawString("50 Hz", xFont, m_TextBrush, px - 8, m_ddy - 12)
                                ElseIf (f = 100) Then
                                    g.DrawString("500 Hz", xFont, m_TextBrush, px - 12, m_ddy - 12)
                                ElseIf (f = 1000) Then
                                    g.DrawString("5 KHz", xFont, m_TextBrush, px - 10, m_ddy - 12)
                                End If
                            End If
                        Else
                            g.DrawLine(m_scalePen, px, 0, px, m_ddy - m_downborder)
                        End If
                    End If
                Next
                f *= 10
            End While
        Else
            ' ---------------------------------------------------------------------- X AXIS LIN
            Dim nstep As Int32
            Dim deltaf As Double = m_fmax - m_fmin
            Dim stepsize As Double = 0.1
            '
            If deltaf > 0 Then
                nstep = 13
                If m_fmax >= 10000 Then nstep = 10
                If m_xlog AndAlso m_fmax / m_fmin > 2 AndAlso m_fmax >= 1000 Then nstep = 8
                '
                While deltaf / stepsize > nstep
                    stepsize *= 2
                    If deltaf / stepsize > nstep Then stepsize *= 2.5
                    If deltaf / stepsize > nstep Then stepsize *= 2
                End While
                '
                max = Math.Ceiling(m_fmax / stepsize + 0.5) * stepsize
                min = Math.Floor(m_fmin / stepsize + 0.5) * stepsize
                Dim f As Double = min
                While f <= max
                    If f >= m_fmin And f <= m_fmax Then
                        If m_xlog Then
                            px = m_leftborder + CInt(Math.Floor(klog * Math.Log(f / m_fmin)))
                        Else
                            px = m_leftborder + CInt(Math.Floor((m_dispx - 1) * (f - m_fmin) / deltaf))
                        End If

                        g.DrawLine(m_scalePen, px, 0, px, m_ddy - m_downborder + 4)
                        If px < m_ddx - 10 Then
                            If (Int(f) = f) Then
                                g.DrawString(f.ToString("0"), xFont, m_TextBrush, px - 10, m_ddy - 12)
                            Else
                                g.DrawString(f.ToString("0.0"), xFont, m_TextBrush, px - 10, m_ddy - 12)
                            End If
                        End If
                    End If
                    f += stepsize
                End While
            End If
            '
            nstep = 4
            If m_xlog Then nstep = 3
            If deltaf / stepsize < nstep Then
                s = m_fmax.ToString("0.0")
                If m_fmax >= 10000 Then
                    g.DrawString(s, xFont, m_TextBrush, m_ddx - 36, m_ddy - 12)
                ElseIf m_fmax >= 1000 Then
                    g.DrawString(s, xFont, m_TextBrush, m_ddx - 28, m_ddy - 12)
                Else
                    g.DrawString(s, xFont, m_TextBrush, m_ddx - 24, m_ddy - 12)
                End If
            End If
        End If
        '
        If m_ylog Then
            ' ---------------------------------------------------------------------- Y AXIS LOG
            Dim deltadb As Double = m_dbmax - m_dbmin
            Dim stepsize As Double = 0.1
            Dim nstep As Int32 = 20
            While deltadb / stepsize > nstep
                stepsize *= 2
                If deltadb / stepsize > nstep Then stepsize *= 2.5
                If deltadb / stepsize > nstep Then stepsize *= 2
            End While
            max = Math.Ceiling(m_dbmax / stepsize + 0.5) * stepsize
            min = Math.Floor(m_dbmin / stepsize + 0.5) * stepsize
            Dim db As Double = min
            While db <= max
                If db >= m_dbmin And db <= m_dbmax Then
                    Dim y As Int32 = CInt(Math.Floor(m_dispy - (m_dispy * (db - m_dbmin) / (m_dbmax - m_dbmin)) + 0.5))
                    g.DrawLine(m_scalePen, m_leftborder - 4, y, m_ddx, y)
                    If y > 14 Then
                        If y > m_dispy - 2 Then y = m_dispy - 2
                        If Math.Floor(db) = db Then
                            s = db.ToString("0")
                        Else
                            s = db.ToString("0.0")
                        End If
                        g.DrawString(s, yFont, m_TextBrush, 3, y - 6)
                    End If
                End If
                db += stepsize
            End While
            g.DrawString("dB", yFont, m_TextBrush, 5, 0)
        Else
            ' ---------------------------------------------------------------------- Y AXIS LIN
            Dim mvmin As Double = 2 * Math.Exp(m_dbmin / kdb) * 1000
            Dim mvmax As Double = 2 * Math.Exp(m_dbmax / kdb) * 1000
            Dim deltamv As Double = mvmax - mvmin
            Dim stepsize As Double = 0.001
            Dim nstep As Int32 = 20
            While deltamv / stepsize > nstep
                stepsize *= 2
                If deltamv / stepsize > nstep Then stepsize *= 2.5
                If deltamv / stepsize > nstep Then stepsize *= 2
            End While
            '
            max = Math.Ceiling(mvmax / stepsize + 0.5) * stepsize
            min = Math.Floor(mvmin / stepsize + 0.5) * stepsize
            Dim mv As Double = min
            While mv <= max
                If mv >= mvmin And mv <= mvmax Then
                    Dim y As Int32 = CInt(Math.Floor((m_dispy - (m_dispy * (mv - mvmin) / (mvmax - mvmin)) + 0.5)))
                    g.DrawLine(m_scalePen, m_leftborder - 4, y, m_ddx, y)
                    If y >= 16 Then
                        If y > m_dispy - 2 Then y = m_dispy - 2
                        If mvmax >= 1 Then
                            If Math.Floor(mv) = mv Then
                                s = mv.ToString("0")
                            ElseIf Math.Abs(Math.Floor(mv * 10 + 0.5) - mv * 10) < 0.0001 Then
                                s = mv.ToString("0.0")
                            ElseIf Math.Abs(Math.Floor(mv * 100 + 0.5) - mv * 100) < 0.0001 Then
                                s = mv.ToString("0.00")
                            Else
                                s = mv.ToString("0.000")
                            End If
                        Else
                            s = (mv * 1000).ToString("0")
                        End If
                        g.DrawString(s, yFont, m_TextBrush, 3, y - 6)
                    End If
                End If
                mv += stepsize
            End While
            If mvmax >= 1 Then
                g.DrawString("mV", yFont, m_TextBrush, 4, 0)
            Else
                g.DrawString("uV", yFont, m_TextBrush, 4, 0)
            End If
        End If
    End Sub

End Module
