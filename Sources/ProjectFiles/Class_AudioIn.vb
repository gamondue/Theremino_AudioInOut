Imports System.Runtime.InteropServices

Public Class AudioIn

    ' ===============================================================================================
    '  AUDIO IN - PRIVATE PROPS AND FUNCTIONS
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
    Private m_DoneProc As BufferDoneEventHandler
    Private m_InputActive As Boolean
    Private m_hWaveIn As Int32
    Private m_WavehdrLen As Int32
    Private m_WavehdrFlagOffset As Int32
    Private m_Result As Int32
    Private m_NextBuffer As Int32
    Private m_Timer As MultimediaTimer = New MultimediaTimer

    Private Sub AddBuffers()
        SyncLock Me
            While m_InputActive
                ' ----------------------------------------------------------------- test the flag WHDR_INQUEUE
                Dim dwFlags As Int16
                dwFlags = Marshal.ReadInt16(m_Buffers(m_NextBuffer).HeaderPointer, m_WavehdrFlagOffset)
                If (dwFlags And WaveNative.WHDR_INQUEUE) <> 0 Then Exit While
                ' ----------------------------------------------------------------- read buffer
                m_DoneProc(m_Buffers(m_NextBuffer).Data)
                ' ----------------------------------------------------------------- send the buffer
                m_Result = WaveNative.waveInAddBuffer(m_hWaveIn, _
                                                      m_Buffers(m_NextBuffer).HeaderPointer, _
                                                      m_WavehdrLen)
                WaveNative.DecodeError("waveInAddBuffer", m_Result)
                If m_Result > 0 Then InputOff()
                ' ----------------------------------------------------------------- select next buffer
                m_NextBuffer += 1
                If m_NextBuffer >= m_BuffersCount Then m_NextBuffer = 0
            End While
        End SyncLock
    End Sub


    ' ===============================================================================================
    '  AUDIO IN - PUBLIC PROPS AND FUNCTIONS
    ' ===============================================================================================
    Friend Delegate Sub BufferDoneEventHandler(ByRef data() As Int16)

    Friend Function InputOn(ByVal InputDeviceIndex As Int32, _
                            ByVal rate As Int32, _
                            ByVal channels As Int32, _
                            ByVal bufferSize As Int32, _
                            ByVal bufferCount As Int32, _
                            ByVal readProc As BufferDoneEventHandler) As Boolean
        If readProc Is Nothing Then
            MsgBox("Function InputOn - You must provide a readProc.")
            Return False
        End If
        If m_InputActive AndAlso m_device = InputDeviceIndex AndAlso _
                                 m_rate = rate AndAlso _
                                 m_channels = channels AndAlso _
                                 m_BufferSize = bufferSize AndAlso _
                                 m_BuffersCount = bufferCount AndAlso _
                                 m_DoneProc = readProc Then
            Return True
        End If
        InputOff()
        SyncLock Me
            m_device = InputDeviceIndex
            m_rate = rate
            m_channels = channels
            m_BufferSize = bufferSize
            m_BuffersCount = bufferCount
            m_DoneProc = readProc
            Dim format As WaveNative.WAVEFORMATEX
            format = New WaveNative.WAVEFORMATEX(m_rate, 16, m_channels)
            ' --------------------------------------------------------------------- Wave In Open
            m_Result = WaveNative.waveInOpen(m_hWaveIn, m_device, format, 0, 0, 0)
            If m_Result <> WaveNative.MMSYSERR_NOERROR Then
                WaveNative.DecodeError("WaveInOpen", m_Result)
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
                m_Result = WaveNative.waveInPrepareHeader(m_hWaveIn, _
                                                          m_Buffers(i).HeaderPointer, _
                                                          m_WavehdrLen)
                WaveNative.DecodeError("waveInPrepareHeader", m_Result)
            Next
            ' --------------------------------------------------------------------- start all
            m_InputActive = True
            m_NextBuffer = 0
            m_Timer.Timer_Start(10, AddressOf AddBuffers)
            m_Result = WaveNative.waveInStart(m_hWaveIn)
            WaveNative.DecodeError("waveInStart", m_Result)
            Return True
        End SyncLock
    End Function

    Friend Sub InputOff()
        SyncLock Me
            m_Timer.Timer_Stop()
            If Not m_InputActive Then Return
            ' --------------------------------------------------------------------- stop sending buffers
            m_InputActive = False
            ' --------------------------------------------------------------------- reset
            m_Result = WaveNative.waveInReset(m_hWaveIn)
            WaveNative.DecodeError("waveInReset", m_Result)
            ' --------------------------------------------------------------------- wait all the buffers
            ' --------------------------------------------------------------------- wait up to 1 Sec (10 * 100mS)
            For m_NextBuffer = 0 To m_BuffersCount - 1
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
                m_Result = WaveNative.waveInUnprepareHeader(m_hWaveIn, _
                                                            m_Buffers(i).HeaderPointer, _
                                                            m_WavehdrLen)
                WaveNative.DecodeError("waveInUnprepareHeader", m_Result)
            Next
            ' --------------------------------------------------------------------- close wave in
            m_Result = WaveNative.waveInClose(m_hWaveIn)
            WaveNative.DecodeError("waveInClose", m_Result)
            ' --------------------------------------------------------------------- free GC Handlers
            For i As Int32 = 0 To m_BuffersCount - 1
                m_Buffers(i).gcHandleHeader.Free()
                m_Buffers(i).gcHandleData.Free()
            Next
        End SyncLock
    End Sub

End Class
