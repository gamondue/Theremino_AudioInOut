Class WavePlayer

    ' =====================================================================================
    '  PRIVATE VARS AND FUNCTIONS
    ' =====================================================================================

    Private m_SamplesPerSec As Int32 = 44100 ' 44100 or 176400 or 192000 or 198450
    Private m_Channels As Int32 = 1
    Private m_NumBuffers As Int32 = 4
    Private m_LenBuffers As Int32 = 2048
    Private m_OutputDeviceIndex As Int32 = -1

    Private m_Playing As Boolean = False
    Private m_Frequency As Single
    Private m_oldfreq As Single
    Private m_OutLevel As Single
    Private m_oldlevel As Single

    Private m_kdb As Double = 8.685889638
    Private m_phase As Int32
    Private m_WaveTable() As Single
    Private m_WaveType As WaveTypes
    Private m_randomInt As Int32
    Private m_randomFloat As Single
    Private ao As AudioOut = New AudioOut
    Private Const ShiftBits As Int32 = 12

    Private Sub InitWaveTable()
        If m_SamplesPerSec = 0 Then Return
        If m_WaveTable Is Nothing OrElse m_SamplesPerSec <> m_WaveTable.Length Then
            ReDim m_WaveTable(m_SamplesPerSec - 1)
            m_oldfreq = m_Frequency
            m_oldlevel = m_OutLevel
            m_phase = 0
        End If
        Dim i As Int32
        Select Case m_WaveType
            Case WaveTypes.Sine
                Dim k_pi As Double = (Math.PI * 2) / m_WaveTable.Length
                For i = 0 To m_WaveTable.Length - 1
                    m_WaveTable(i) = CSng(Math.Sin(i * k_pi))
                Next
            Case WaveTypes.Triangle
                For i = 0 To m_WaveTable.Length \ 4 - 1
                    m_WaveTable(i) = CSng(4 * i / m_WaveTable.Length)
                Next
                For i = m_WaveTable.Length \ 4 To m_WaveTable.Length * 3 \ 4 - 1
                    m_WaveTable(i) = CSng(2 - (4 * i / m_WaveTable.Length))
                Next
                For i = m_WaveTable.Length * 3 \ 4 To m_WaveTable.Length - 1
                    m_WaveTable(i) = CSng(4 * i / m_WaveTable.Length - 4)
                Next
            Case WaveTypes.RampPos
                For i = 0 To m_WaveTable.Length - 1
                    m_WaveTable(i) = -CSng(2 * (1 - i / m_WaveTable.Length) - 1)
                Next
            Case WaveTypes.RampNeg
                For i = 0 To m_WaveTable.Length - 1
                    m_WaveTable(i) = -CSng(2 * i / m_WaveTable.Length - 1)
                Next
            Case WaveTypes.Square
                For i = 0 To m_WaveTable.Length \ 2 - 1
                    m_WaveTable(i) = 1
                Next
                For i = m_WaveTable.Length \ 2 To m_WaveTable.Length - 1
                    m_WaveTable(i) = -1
                Next
            Case WaveTypes.PmtPulses
                For i = 0 To 150
                    m_WaveTable(i) = i * 0.96F / 150
                Next
                For i = 0 To 25
                    m_WaveTable(151 + i) = 0.96F + i * 0.04F / 25
                Next
                For i = 0 To 25
                    m_WaveTable(176 + i) = 1.0F - i * 0.04F / 25
                Next
                For i = 0 To 250
                    m_WaveTable(201 + i) = 0.96F - i * 0.96F / 250
                Next
                For i = 0 To 200
                    m_WaveTable(451 + i) = -i * 0.34F / 200
                Next
                For i = 0 To 100
                    m_WaveTable(651 + i) = -0.34F - i * 0.06F / 100
                Next
                For i = 0 To 100
                    m_WaveTable(751 + i) = -0.4F + i * 0.03F / 100
                Next
                For i = 0 To 600
                    m_WaveTable(851 + i) = -0.37F + i * 0.32F / 600
                Next
                For i = 0 To 200
                    m_WaveTable(1451 + i) = -0.05F + i * 0.05F / 200
                Next
                For i = 1651 To m_WaveTable.Length - 1
                    m_WaveTable(i) = 0
                Next
        End Select
    End Sub

    Private Sub FillBuffer(ByRef data() As Int16)
        Select Case m_WaveType
            Case WaveTypes.Noise
                Dim kvol As Single = CSng(Math.Exp(m_OutLevel / m_kdb))
                If m_randomInt = 0 Then m_randomInt = 1
                For i As Int32 = 0 To data.Length - 1
                    ' --------------------------------------------------------- DAA random method
                    If (m_randomInt And &H800000) <> 0 Then
                        m_randomInt = (m_randomInt << 1) Xor &H1D872B41
                    Else
                        m_randomInt = m_randomInt << 1
                    End If
                    ' --------------------------------------------------------- limit randomFloat value
                    m_randomFloat += m_randomInt * 0.0002F
                    If Math.Abs(m_randomFloat) > 1000000000.0 Then m_randomFloat = 0
                    m_randomFloat = CInt(m_randomFloat) And 32767
                    ' --------------------------------------------------------- set the value
                    data(i) = CShort(kvol * m_randomFloat)
                Next
            Case Else
                ' ------------------------------------------------------------- limit frequency
                If m_Frequency < 0 Then m_Frequency = 0
                If m_Frequency > m_SamplesPerSec / 2.0F Then m_Frequency = m_SamplesPerSec / 2.0F
                '
                If m_Frequency = m_oldfreq And m_OutLevel = m_oldlevel Then
                    ' --------------------------------------------------------- fast exec if constant freq and volume
                    Dim f As Int32 = CInt(m_oldfreq * (1 << ShiftBits))
                    Dim kvol As Single = -CSng(32767 * Math.Exp(m_OutLevel / m_kdb))
                    Dim pmax As Int32 = (m_SamplesPerSec << ShiftBits)
                    For i As Int32 = 0 To data.Length - 1
                        data(i) = CShort(kvol * m_WaveTable(m_phase >> ShiftBits))
                        m_phase += f
                        If m_phase >= pmax Then
                            m_phase -= pmax
                        End If
                    Next
                Else
                    ' --------------------------------------------------------- prepare for smooth frequency change
                    Dim f As Int32 = CInt(m_oldfreq * (1 << ShiftBits))
                    Dim df As Int32 = CInt(m_Frequency - m_oldfreq) * (1 << ShiftBits) \ data.Length
                    m_oldfreq = m_Frequency
                    Dim pmax As Int32 = m_SamplesPerSec << ShiftBits
                    ' --------------------------------------------------------- prepare for smooth volume change
                    Dim kvol As Single = -CSng(32767 * Math.Exp(m_OutLevel / m_kdb))
                    Dim oldkvol As Single = -CSng(32767 * Math.Exp(m_oldlevel / m_kdb))
                    Dim dv As Single = (kvol - oldkvol) / data.Length
                    ' --------------------------------------------------------- fill buffer
                    For i As Int32 = 0 To data.Length - 1
                        data(i) = CShort(oldkvol * m_WaveTable(m_phase >> ShiftBits))
                        m_phase += f
                        f += df
                        oldkvol += dv
                        If m_phase >= pmax Then
                            m_phase -= pmax
                        End If
                    Next
                    m_oldlevel = m_OutLevel
                End If
        End Select
    End Sub


    ' =====================================================================================
    '  PUBLIC FUNCTIONS
    ' =====================================================================================
    Friend Enum WaveTypes As Int32
        Sine = 0
        Triangle = 1
        RampPos = 2
        RampNeg = 3
        Square = 4
        Noise = 5
        PmtPulses = 6
    End Enum

    Friend Sub SetWaveType(ByVal wavetype As WaveTypes)
        If wavetype <> m_WaveType Then
            m_WaveType = wavetype
            InitWaveTable()
        End If
    End Sub

    Friend Sub SetFrequency(ByVal freq As Single)
        m_Frequency = freq
    End Sub

    Friend Sub SetOutLevel(ByVal dB As Single)
        m_OutLevel = dB
    End Sub

    Friend Sub PlayStart(ByVal OutputDeviceIndex As Int32)
        If OutputDeviceIndex <> m_OutputDeviceIndex Then
            PlayStop()
            m_OutputDeviceIndex = OutputDeviceIndex
        End If
        If ao Is Nothing Then Return
        If m_Playing Then Return
        InitWaveTable()
        m_oldlevel = -120
        ao.OutputOn(m_OutputDeviceIndex, _
                    m_SamplesPerSec, _
                    m_Channels, _
                    m_LenBuffers, _
                    m_NumBuffers, _
                    AddressOf FillBuffer)
        m_Playing = True
    End Sub

    Friend Sub PlayStart(ByVal OutputDeviceIndex As Int32, _
                         ByVal freq As Single, _
                         ByVal db As Single, _
                         ByVal WaveType As WaveTypes)
        If OutputDeviceIndex <> m_OutputDeviceIndex Then
            PlayStop()
            m_OutputDeviceIndex = OutputDeviceIndex
        End If
        If ao Is Nothing Then Return
        If m_Playing Then Return
        m_Frequency = freq
        m_OutLevel = db
        m_oldlevel = -120
        m_WaveType = WaveType
        InitWaveTable()
        ao.OutputOn(m_OutputDeviceIndex, _
                    m_SamplesPerSec, _
                    m_Channels, _
                    m_LenBuffers, _
                    m_NumBuffers, _
                    AddressOf FillBuffer)
        m_Playing = True
    End Sub

    Friend Sub PlayStop()
        If ao Is Nothing Then Return
        If Not m_Playing Then Return
        Dim old As Single = m_OutLevel
        m_OutLevel = -120
        Threading.Thread.Sleep(ao.GetTotalDelayMillisec() + 20)
        'Debug.Print(ao.GetTotalDelayMilliseconds().ToString)
        ao.OutputOff()
        m_OutLevel = old
        m_Playing = False
    End Sub

    Friend Function GetTotalDelayMillisec() As Int32
        Return ao.GetTotalDelayMillisec()
    End Function

    Friend Sub PlayPrepareForStop()
        If ao Is Nothing Then Return
        If Not m_Playing Then Return
        m_OutLevel = -120
    End Sub

    Friend Sub PlayStopImmediate()
        If ao Is Nothing Then Return
        If Not m_Playing Then Return
        ao.OutputOff()
        m_Playing = False
    End Sub

End Class

