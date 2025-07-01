Imports System.Math

Class FHT

    ' ===============================================================================================
    '   FHT - FAST HARTLEY TRANSFORM
    ' ===============================================================================================
    ' FHT is more efficient than FFT (Fast Fourier Transform) because transforms real inputs 
    ' to real outputs with no intrinsic involvement of complex numbers. 
    ' This implementation is the son of the DAA (Digital Audio Analyzer) implementation.
    ' There are 20 years of knowledge here, there is nothing so simple and so efficient. 
    ' This class uses native VbNet arrays and works faster than the C++ DAA implementation. 
    ' No array is allocated and destroyed at run time, all is done in the initialization function.
    ' A 1024 samples FHT takes about 60 uSec. This is 5 times faster than the DAA C++ version!
    ' ===============================================================================================

    Friend OutputBuffer() As Single
    Friend BandsBuffer() As Single
    Friend BandsGains() As Single

    Private hTab() As Double
    Private senTab() As Double
    Private cosTab() As Double
    Private permTab() As Int32
    Private aTab() As Double
    Private bTab() As Double

    Friend Enum WindowTypes As Int32
        Rectangular
        Sin
        Sin2
        Sin5
        Sin10
        Sin20
        Sin50
        Sin100
        Sin200
        Sin500
        Sin1000
        Sin2000
        Hanning
        BartlettHann
        Hamming
        Blackman
        NuttallNarrow
        Nuttall
        BlackmanNuttall
        BlackmanHarris
        FlatTop
    End Enum
    Private m_windowtype As WindowTypes = WindowTypes.Rectangular

    Private initSamples As Int32 = 0
    Private Const kdb As Double = 8.685889638
    Private Const PI As Double = 3.14159265359

    Private Function IsPowerOfTwo(ByVal n As Int32) As Boolean
        If n = 0 Then Return False
        Return (n And (n - 1)) = 0
    End Function

    ' ------------------------------------------------------------------------- FHT init
    Private Sub FHTinit(ByVal nn As Int32, ByVal m_WindowType As WindowTypes)
        '
        Dim index, j, s As Int32
        Dim ang As Double = 0
        Dim omega As Double = 2 * PI / nn
        Dim lg2 As Int32 = CInt(Math.Log(nn) / Log(2.0))
        Dim knn As Double = 2.0F / nn
        '
        ReDim hTab(nn - 1)
        ReDim senTab(nn - 1)
        ReDim cosTab(nn - 1)
        ReDim permTab(nn - 1)
        ReDim aTab(nn)
        ReDim bTab(nn)
        ReDim OutputBuffer(nn \ 2 - 1)
        '
        For i As Int32 = 0 To nn - 1
            ' ----------------------------------------------------------------- Window
            Select Case m_WindowType
                Case WindowTypes.Rectangular
                    hTab(i) = knn

                Case WindowTypes.Sin
                    hTab(i) = Sin(ang / 2) * knn * 1.4
                Case WindowTypes.Sin2
                    hTab(i) = Pow(Sin(ang / 2), 2) * knn * 1.7
                Case WindowTypes.Sin5
                    hTab(i) = Pow(Sin(ang / 2), 5) * knn * 2.4
                Case WindowTypes.Sin10
                    hTab(i) = Pow(Sin(ang / 2), 10) * knn * 3.3
                Case WindowTypes.Sin20
                    hTab(i) = Pow(Sin(ang / 2), 20) * knn * 4.6
                Case WindowTypes.Sin50
                    hTab(i) = Pow(Sin(ang / 2), 50) * knn * 7.1
                Case WindowTypes.Sin100
                    hTab(i) = Pow(Sin(ang / 2), 100) * knn * 10.0
                Case WindowTypes.Sin200
                    hTab(i) = Pow(Sin(ang / 2), 200) * knn * 14.1
                Case WindowTypes.Sin500
                    hTab(i) = Pow(Sin(ang / 2), 500) * knn * 22.3
                Case WindowTypes.Sin1000
                    hTab(i) = Pow(Sin(ang / 2), 1000) * knn * 31.5
                Case WindowTypes.Sin2000
                    hTab(i) = Pow(Sin(ang / 2), 2000) * knn * 44.7

                Case WindowTypes.Hanning
                    hTab(i) = (1 - Cos(ang)) * knn * 0.84
                Case WindowTypes.BartlettHann
                    hTab(i) = (0.62 - _
                               0.48 * Abs(nn / (nn - 1) - 0.5) - _
                               0.38 * Cos(ang)) * knn * 2.25
                Case WindowTypes.Hamming
                    hTab(i) = (0.53836 - _
                               0.46164 * Cos(ang)) * knn * 1.6
                Case WindowTypes.Blackman
                    hTab(i) = (0.42659 - _
                               0.49656 * Cos(ang) + _
                               0.076849 * Cos(ang * 2)) * knn * 1.95
                Case WindowTypes.NuttallNarrow
                    hTab(i) = (0.40217 - _
                               0.49703 * Cos(ang) + _
                               0.09892 * Cos(ang * 2) - _
                               0.00188 * Cos(ang * 3)) * knn * 2.05
                Case WindowTypes.Nuttall
                    hTab(i) = (0.355768 - _
                               0.487396 * Cos(ang) + _
                               0.144232 * Cos(ang * 2) - _
                               0.012604 * Cos(ang * 3)) * knn * 2.25
                Case WindowTypes.BlackmanNuttall
                    hTab(i) = (0.3635819 - _
                               0.4891775 * Cos(ang) + _
                               0.1365995 * Cos(ang * 2) - _
                               0.0106411 * Cos(ang * 3)) * knn * 2.2
                Case WindowTypes.BlackmanHarris
                    hTab(i) = (0.35875 - _
                               0.48829 * Cos(ang) + _
                               0.14128 * Cos(ang * 2) - _
                               0.01168 * Cos(ang * 3)) * knn * 2.25
                Case WindowTypes.FlatTop
                    hTab(i) = (0.215578948 - _
                               0.41663158 * Cos(ang) + _
                               0.277263158 * Cos(ang * 2) - _
                               0.083578947 * Cos(ang * 3) + _
                               0.006947368 * Cos(ang * 4)) * knn * 3.5
            End Select
            ' ----------------------------------------------------------------- SIN & COS tables
            senTab(i) = Sin(ang)
            cosTab(i) = Cos(ang)
            ang += omega
            ' ----------------------------------------------------------------- Permutation table
            index = i
            j = 0
            For k As Int32 = 0 To lg2 - 1
                s = index >> 1
                j = j + j + index - s - s
                index = s
            Next
            permTab(i) = j
        Next
        initSamples = nn
    End Sub




    ' ------------------------------------------------------------------------- FHT execute
    Friend Sub ExecuteFHT(ByVal InputBuffer() As Int16, _
                          Optional ByVal WindowType As FHT.WindowTypes = WindowTypes.Sin10)
        '
        'Dim sw1 As Diagnostics.Stopwatch = New Diagnostics.Stopwatch : sw1.Start()
        '
        Dim nn As Int32 = InputBuffer.Length
        If Not IsPowerOfTwo(nn) Then
            MsgBox("FHT - N must be a power of two")
            Return
        End If
        '
        If nn <> initSamples OrElse WindowType <> m_windowtype Then
            m_windowtype = WindowType
            FHTinit(nn, m_windowtype)
        End If
        '
        Dim i, j, le, le2, k, mk, trig_ind, trig_inc As Int32
        Dim ua, ub, v As Double
        Dim swaptab() As Double

        ' --------------------------------------------------------------------- zero level correction using the sample area values
        Static correction As Double = 0
        Dim deltasum As Double = 0
        For i = 0 To nn - 1
            deltasum += (InputBuffer(i) - correction) * hTab(i)
        Next
        Static SmoothedCorr As Double
        SmoothedCorr += (deltasum - SmoothedCorr) * 0.3
        correction += SmoothedCorr

        ' --------------------------------------------------------------------- from circular buffer to aTab with window 
        For i = 0 To nn - 1
            aTab(permTab(i)) = (InputBuffer(i) - correction) * hTab(i)
        Next
        ' --------------------------------------------------------------------- FHT
        le = 1
        trig_inc = nn >> 1
        Do
            trig_ind = 0
            j = 0
            le2 = le * 2
            Do
                ua = cosTab(trig_ind)
                ub = senTab(trig_ind)
                trig_ind += trig_inc
                i = j
                mk = le2 - i
                Do
                    k = le + i
                    v = ua * aTab(k) + ub * aTab(mk)
                    bTab(k) = aTab(i) - v
                    bTab(i) = aTab(i) + v
                    mk += le2
                    i += le2
                Loop While i < nn
                j += 1
            Loop While j < le
            swaptab = aTab
            aTab = bTab
            bTab = swaptab
            le *= 2
            trig_inc = trig_inc >> 1
        Loop While le < nn
        ' --------------------------------------------------------------------- Results to OutputBuffer   
        For j = 0 To nn \ 2 - 1
            ua = aTab(j + 1)
            ub = aTab(nn - (j + 1))
            OutputBuffer(j) = CSng(Sqrt(ua * ua + ub * ub))
        Next
        '
        'ReportString = "FHT time: " & (sw1.Elapsed.TotalMilliseconds * 1000).ToString("0") & " uS"
    End Sub

    Friend Sub ExtractBands(ByVal nBands As Int32, _
                            Optional ByVal freqmin As Single = 10, _
                            Optional ByVal freqmax As Single = 20000, _
                            Optional ByVal dbmin As Single = -120, _
                            Optional ByVal dbmax As Single = 0, _
                            Optional ByVal xlog As Boolean = False, _
                            Optional ByVal ylog As Boolean = False, _
                            Optional ByVal SampleFreq As Int32 = 44100, _
                            Optional ByVal ResponseSpeed As Single = 0.2, _
                            Optional ByVal AGC As Boolean = False)
        '
        'Dim sw1 As Diagnostics.Stopwatch = New Diagnostics.Stopwatch : sw1.Start()
        '
        Static m_agc As Boolean = False
        If BandsBuffer Is Nothing OrElse BandsBuffer.Length <> nBands Then
            ReDim BandsBuffer(nBands - 1)
            ReDim BandsGains(nBands - 1)
            For i As Int32 = 0 To nBands - 1
                BandsGains(i) = 1
            Next
            m_agc = Not AGC
        End If
        If AGC <> m_agc Then
            For i As Int32 = 0 To nBands - 1
                BandsGains(i) = 1
            Next
            m_agc = AGC
        End If
        ' ---------------------------------------------------------------------
        Dim nn As Int32 = OutputBuffer.Length * 2
        ' --------------------------------------------------------------------- X scale - log / lin
        Dim kf, kk As Single
        If xlog Then
            kf = freqmin * nn / SampleFreq
            kk = CSng(Log(freqmax / freqmin) / nBands)
        Else
            kf = CSng(nn / SampleFreq)
            kk = (freqmax - freqmin) / nBands
        End If
        ' --------------------------------------------------------------------- Y scale - log / lin
        Dim kv, vmin, vmax As Single
        If ylog Then
        Else
            kv = 2.0 / 32768
            vmin = CSng(2 * Exp(dbmin / kdb))
            vmax = CSng(2 * Exp(dbmax / kdb))
        End If
        ' --------------------------------------------------------------------- vars
        Dim max As Single
        Dim index As Int32
        Dim oldindex As Int32
        If xlog Then
            oldindex = CInt(kf)
        Else
            oldindex = CInt(kf * freqmin)
        End If
        Dim oldband As Int32 = -1
        Dim val As Single
        Dim db As Single
        Dim oldVal As Single = 0
        ' ---------------------------------------------------------------------
        For band As Int32 = 0 To nBands - 1
            ' ----------------------------------------------------------------- Y scale - log / lin
            If xlog Then
                index = CInt(-0.5 + kf * Exp(kk * (band + 1)))
            Else
                index = CInt(-0.5 + kf * (freqmin + kk * (band + 1)))
            End If
            If oldindex < 0 Then oldindex = index
            If index > oldindex Then
                ' ------------------------------------------------------------- find max value in the band
                max = 0
                For j As Int32 = oldindex To index - 1
                    max = Math.Max(OutputBuffer(j), max)
                Next
                ' ------------------------------------------------------------- find mid value in the band
                'max = 0
                'For j As Int32 = oldindex To index - 1
                '    max += OutputBuffer(j)
                'Next
                'max /= (index - oldindex)
                ' ------------------------------------------------------------- bands gains (for AGC and equalizers) 
                max *= BandsGains(band)
                ' ------------------------------------------------------------- Y scale - log / lin
                If ylog Then
                    db = CSng(kdb * Log(max / 32768))
                    val = (db - dbmin) / (dbmax - dbmin)
                Else
                    val = (max * kv - vmin) / (vmax - vmin)
                End If
                ' ------------------------------------------------------------- limit from 0 to 1
                If val < 0 Then val = 0
                If val > 1 Then val = 1
                ' ------------------------------------------------------------- create a smoot interpolation
                If oldband >= 0 Then
                    Dim k As Single = (val - oldVal) / (band - oldband)
                    For j As Int32 = oldband + 1 To band
                        val = oldVal + (j - oldband) * k
                        SetBandsBufferValue(val, j, ResponseSpeed)
                    Next
                Else
                    ' --------------------------------------------------------- first band
                    For j As Int32 = oldband + 1 To band
                        SetBandsBufferValue(val, j, ResponseSpeed)
                    Next
                End If
                oldband = band
                oldindex = index
                oldVal = val
            End If
            ' ----------------------------------------------------------------- last band 
            If band = nBands - 1 Then
                For j As Int32 = oldband + 1 To band
                    SetBandsBufferValue(val, j, ResponseSpeed)
                Next
            End If
        Next
        ' --------------------------------------------------------------------- AGC
        If AGC Then
            For j As Int32 = 0 To nBands - 1
                If BandsBuffer(j) < 0.8 Then
                    If BandsGains(j) < 100 Then
                        BandsGains(j) *= 1.01F
                    End If
                Else
                    If BandsGains(j) > 0.01 Then
                        BandsGains(j) *= 0.9F
                    End If
                End If
            Next
            'Form1.Text = BandsGains(0).ToString("0") & "  " & _
            '             BandsGains(1).ToString("0") & "  " & _
            '             BandsGains(2).ToString("0")
        End If
        '
        'ReportString = "ExtractBands time: " & (sw1.Elapsed.TotalMilliseconds * 1000).ToString("0") & " uS"
    End Sub

    Private Sub SetBandsBufferValue(ByVal val As Single, ByVal j As Int32, ByVal speed As Single)
        If speed >= 1 Then
            BandsBuffer(j) = val
            Return
        End If
        If val > BandsBuffer(j) Then
            speed *= 10
            If speed > 1 Then speed = 1
            If speed < 0.1 Then speed = 0.1
        End If
        BandsBuffer(j) += (val - BandsBuffer(j)) * speed
    End Sub

End Class
