
Imports System.Runtime.InteropServices

Public Class AudioOut

    ' ===============================================================================================
    '  AUDIO OUT - PRIVATE PROPS AND FUNCTIONS
    ' ===============================================================================================
    Private Structure Buffer
        Public Header As WaveNative.WAVEHDR
        Public Data() As Int16
        Public gcHandleHeader As GCHandle
        Public gcHandleData As GCHandle
        Public HeaderPointer As IntPtr
    End Structure
    Private m_Buffers() As Buffer

    Private m_device As Int32
    Private m_rate As Int32
    Private m_channels As Int32
    Private m_BufferSize As Int32
    Private m_BuffersCount As Int32
    Private m_FillProc As BufferFillEventHandler
    Private m_OutputActive As Boolean
    Private m_hWaveOut As Int32
    Private m_WavehdrLen As Int32
    Private m_WavehdrFlagOffset As Int32
    Private m_Result As Int32
    Private m_NextBuffer As Int32
    Private m_Timer As MultimediaTimer = New MultimediaTimer

    Private Sub FillBuffers()
        SyncLock Me
            While m_OutputActive
                ' ----------------------------------------------------------------- test the flag WHDR_INQUEUE
                Dim dwFlags As Int16
                dwFlags = Marshal.ReadInt16(m_Buffers(m_NextBuffer).HeaderPointer, m_WavehdrFlagOffset)
                If (dwFlags And WaveNative.WHDR_INQUEUE) <> 0 Then Exit While
                ' ----------------------------------------------------------------- fill buffer
                m_FillProc(m_Buffers(m_NextBuffer).Data)
                ' ----------------------------------------------------------------- send the buffer
                m_Result = WaveNative.waveOutWrite(m_hWaveOut, _
                                                  m_Buffers(m_NextBuffer).HeaderPointer, _
                                                  m_WavehdrLen)
                WaveNative.DecodeError("waveOutWrite", m_Result)
                If m_Result > 0 Then OutputOff()
                ' ----------------------------------------------------------------- select next buffer
                m_NextBuffer += 1
                If m_NextBuffer >= m_BuffersCount Then m_NextBuffer = 0
            End While
        End SyncLock
    End Sub

    ' ===============================================================================================
    '  AUDIO OUT - PUBLIC PROPS AND FUNCTIONS
    ' ===============================================================================================
    Friend Delegate Sub BufferFillEventHandler(ByRef data() As Int16)

    Friend Function OutputOn(ByVal device As Int32, _
                             ByVal rate As Int32, _
                             ByVal channels As Int32, _
                             ByVal bufferSize As Int32, _
                             ByVal bufferCount As Int32, _
                             ByVal fillProc As BufferFillEventHandler) As Boolean
        If fillProc Is Nothing Then
            MsgBox("Function OutputOn - You must provide a fillProc.")
            Return False
        End If
        If m_OutputActive AndAlso m_device = device AndAlso _
                                  m_rate = rate AndAlso _
                                  m_channels = channels AndAlso _
                                  m_BufferSize = bufferSize AndAlso _
                                  m_BuffersCount = bufferCount AndAlso _
                                  m_FillProc = fillProc Then
            Return True
        End If
        OutputOff()
        SyncLock Me
            m_device = device
            m_rate = rate
            m_channels = channels
            m_BufferSize = bufferSize
            m_BuffersCount = bufferCount
            m_FillProc = fillProc
            Dim format As WaveNative.WAVEFORMATEX
            format = New WaveNative.WAVEFORMATEX(m_rate, 16, m_channels)
            ' --------------------------------------------------------------------- Wave Out Open
            m_Result = WaveNative.waveOutOpen(m_hWaveOut, m_device, format, 0, 0, 0)
            If m_Result <> WaveNative.MMSYSERR_NOERROR Then
                WaveNative.DecodeError("WaveOutOpen", m_Result)
                Return False
            End If
            ' --------------------------------------------------------------------- prepare buffer array 
            ReDim m_Buffers(m_BuffersCount - 1)
            ' --------------------------------------------------------------------- prepare the WAVEHDR len and offset 
            m_WavehdrLen = Marshal.SizeOf(m_Buffers(0).Header)
            m_WavehdrFlagOffset = If(m_WavehdrLen = 32, 16, 24)
            ' --------------------------------------------------------------------- init buffers
            For i As Int32 = 0 To m_BuffersCount - 1
                '
                ReDim m_Buffers(i).Data(bufferSize - 1)
                m_Buffers(i).gcHandleHeader = GCHandle.Alloc(m_Buffers(i).Header, GCHandleType.Pinned)
                m_Buffers(i).gcHandleData = GCHandle.Alloc(m_Buffers(i).Data, GCHandleType.Pinned)
                m_Buffers(i).HeaderPointer = m_Buffers(i).gcHandleHeader.AddrOfPinnedObject()
                m_Buffers(i).Header.lpData = m_Buffers(i).gcHandleData.AddrOfPinnedObject()
                m_Buffers(i).Header.dwBufferLength = m_BufferSize * 2
                m_Buffers(i).Header.dwFlags = 0
                m_Buffers(i).Header.dwLoops = 0
                Marshal.StructureToPtr(m_Buffers(i).Header, m_Buffers(i).HeaderPointer, True)
                '
                m_Result = WaveNative.waveOutPrepareHeader(m_hWaveOut, _
                                                           m_Buffers(i).HeaderPointer, _
                                                           m_WavehdrLen)
                WaveNative.DecodeError("waveOutPrepareHeader", m_Result)
            Next
            ' --------------------------------------------------------------------- start all
            m_OutputActive = True
            m_NextBuffer = 0
            ' --------------------------------------------------------------------- the pause eliminates starting-clicks
            m_Result = WaveNative.waveOutPause(m_hWaveOut)
            WaveNative.DecodeError("waveOutPause", m_Result)
            m_Timer.Timer_Start(10, AddressOf FillBuffers)
            FillBuffers()
            m_Result = WaveNative.waveOutRestart(m_hWaveOut)
            WaveNative.DecodeError("waveOutRestart", m_Result)
            Return True
        End SyncLock
    End Function

    Friend Sub OutputOff()
        SyncLock Me
            m_Timer.Timer_Stop()
            If Not m_OutputActive Then Return
            ' --------------------------------------------------------------------- stop sending buffers
            m_OutputActive = False
            ' --------------------------------------------------------------------- reset
            m_Result = WaveNative.waveOutReset(m_hWaveOut)
            WaveNative.DecodeError("waveOutReset", m_Result)
            ' --------------------------------------------------------------------- wait all the buffers
            For m_NextBuffer = 0 To m_BuffersCount - 1
                ' ----------------------------------------------------------------- wait up to 1 Sec (10 * 100mS)
                For j As Int32 = 1 To 10
                    ' ------------------------------------------------------------- test the flag WHDR_INQUEUE
                    Dim dwFlags As Int16
                    dwFlags = Marshal.ReadInt16(m_Buffers(m_NextBuffer).HeaderPointer, m_WavehdrFlagOffset)
                    If (dwFlags And WaveNative.WHDR_INQUEUE) = 0 Then Exit For
                    System.Threading.Thread.Sleep(100)
                Next
            Next
            ' --------------------------------------------------------------------- unprepare all the headers 
            For i As Int32 = 0 To m_BuffersCount - 1
                m_Result = WaveNative.waveOutUnprepareHeader(m_hWaveOut, _
                                                             m_Buffers(i).HeaderPointer, _
                                                             m_WavehdrLen)
                WaveNative.DecodeError("waveOutUnprepareHeader", m_Result)
            Next
            ' --------------------------------------------------------------------- close wave out
            m_Result = WaveNative.waveOutClose(m_hWaveOut)
            WaveNative.DecodeError("waveOutClose", m_Result)
            ' --------------------------------------------------------------------- free GC Handlers
            For i As Int32 = 0 To m_BuffersCount - 1
                m_Buffers(i).gcHandleHeader.Free()
                m_Buffers(i).gcHandleData.Free()
            Next
        End SyncLock
    End Sub

    ' -----------------------------------------------------------------------------
    '  The time required to reflect volume and freq changes to audio-out
    '  and also the time required (after volume fadeout) before to close wave out 
    ' -----------------------------------------------------------------------------
    Friend Function GetTotalDelayMillisec() As Int32
        If m_rate > 0 Then
            Return (1000 * m_BufferSize * m_BuffersCount) \ (m_rate * m_channels)
        Else
            Return 0
        End If
    End Function

End Class




