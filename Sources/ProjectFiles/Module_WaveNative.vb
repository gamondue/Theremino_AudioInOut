
Imports System.Runtime.InteropServices

Namespace WaveNative
    Friend Module WaveNative

        ' ===================================================================================================
        '   >>> WINMM STABILITY SECRETS <<<
        ' ===================================================================================================
        ' When using winmm.dll functions use the WINMM timer and do not use other timers or threads.
        ' Ensure that headers and data buffer memory do not moves using GCHandles and AddrOfPinnedObject.
        ' Prepare all the buffers in the start function and unprepare all the buffers closing.
        ' Do not use the WaveCallBack test the flag WHDR_INQUEUE is more efficient and less dangerous.
        ' Before unpreparing headers wait for all buffers not INQUEUE for a long time.
        ' Free the GCHandlers when closing.
        ' Include FillBuffers/OutputOn/OutputOff and AddBuffers/InputOn/InputOff with a SyncLock
        '  because FillBuffers and AddBuffers are called from a different thread by the multimedia timer. 
        ' ==================================================================================================

        ' --------------------------------------------------------------------- errors
        Public Const MMSYSERR_NOERROR As Int32 = 0

        Public Function MMSYSERR(ByVal n As Int32) As String
            Select Case n
                Case 0 : Return "MMSYSERR_NOERROR"
                Case 1 : Return "MMSYSERR_UNDEFINED_ERROR"
                Case 2 : Return "MMSYSERR_BADDEVICEID - Specified device identifier is out of range."
                Case 3 : Return "MMSYSERR_NOTENABLED"
                Case 4 : Return "MMSYSERR_ALLOCATED - Specified resource is already allocated."
                Case 5 : Return "MMSYSERR_INVALHANDLE - Specified device handle is invalid."
                Case 6 : Return "MMSYSERR_NODRIVER - No device driver is present."
                Case 7 : Return "MMSYSERR_NOMEM - Unable to allocate or lock memory."
                Case 8 : Return "MMSYSERR_NOTSUPPORTED"
                Case 9 : Return "MMSYSERR_BADERRNUM"
                Case 10 : Return "MMSYSERR_INVALFLAG"
                Case 11 : Return "MMSYSERR_INVALPARAM"
                Case 12 : Return "MMSYSERR_HANDLEBUSY"
                Case 13 : Return "MMSYSERR_INVALIDALIAS"
                Case 14 : Return "MMSYSERR_BADDB"
                Case 15 : Return "MMSYSERR_KEYNOTFOUND"
                Case 16 : Return "MMSYSERR_READERROR"
                Case 17 : Return "MMSYSERR_WRITEERROR"
                Case 18 : Return "MMSYSERR_DELETEERROR"
                Case 19 : Return "MMSYSERR_VALNOTFOUND"
                Case 20 : Return "MMSYSERR_NODRIVERCB"
                Case 32 : Return "WAVERR_BADFORMAT - Attempted to open with an unsupported waveform-audio format."
                Case 33 : Return "WAVERR_STILLPLAYING - The data block pointed to by the pwh parameter is still in the queue."
                Case 34 : Return "WAVERR_UNPREPARED - The data block pointed to by the pwh parameter hasn't been prepared."
                Case 35 : Return "WAVERR_SYNC - The device is synchronous but waveOutOpen was called without using the WAVE_ALLOWSYNC flag."
                Case Else : Return "MMSYSERR UNDEFINED - ERROR NUMBER = " & n.ToString
            End Select
        End Function

        Public Sub DecodeError(ByVal FunctionName As String, ByVal errnum As Int32)
            If errnum = MMSYSERR_NOERROR Then Return
            ' ----------------------------------------------- do not show errors - the show must go on!
            MsgBox(FunctionName & " - " & MMSYSERR(errnum))
        End Sub

        ' --------------------------------------------------------------------- device count
        Public Function OutputDevicesCount() As Int32
            Return waveOutGetNumDevs()
        End Function
        Public Function InputDevicesCount() As Int32
            Return waveInGetNumDevs()
        End Function

        ' --------------------------------------------------------------------- constants
        Public Const MM_WOM_OPEN As Int32 = &H3BB
        Public Const MM_WOM_CLOSE As Int32 = &H3BC
        Public Const MM_WOM_DONE As Int32 = &H3BD

        Public Const MM_WIM_OPEN As Int32 = &H3BE
        Public Const MM_WIM_CLOSE As Int32 = &H3BF
        Public Const MM_WIM_DATA As Int32 = &H3C0

        Public Const WAVE_FORMAT_PCM As Int32 = 1

        Public Const WHDR_DONE As Int32 = 1
        Public Const WHDR_PREPARED As Int32 = 2
        Public Const WHDR_BEGINLOOP As Int32 = 4
        Public Const WHDR_ENDLOOP As Int32 = 8
        Public Const WHDR_INQUEUE As Int32 = 16

        ' --------------------------------------------------------------------- structs 
        <StructLayout(LayoutKind.Sequential)> _
        Public Structure WAVEFORMATEX
            Public wFormatTag As Int16
            Public nChannels As Int16
            Public nSamplesPerSec As Int32
            Public nAvgBytesPerSec As Int32
            Public nBlockAlign As Int16
            Public wBitsPerSample As Int16
            Public cbSize As Int16
            Public Sub New(ByVal rate As Int32, ByVal bits As Int32, ByVal channels As Int32)
                wFormatTag = WAVE_FORMAT_PCM
                nChannels = CShort(channels)
                nSamplesPerSec = rate
                nBlockAlign = CShort(channels * (bits \ 8))
                nAvgBytesPerSec = nSamplesPerSec * nBlockAlign
                wBitsPerSample = CShort(bits)
                cbSize = 0
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Public Structure WAVEHDR
            Public lpData As IntPtr             ' pointer to locked data buffer
            Public dwBufferLength As Int32      ' length of data buffer
            Public dwBytesRecorded As Int32     ' used for input only
            Public dwUser As IntPtr             ' for client's use
            Public dwFlags As Int32             ' assorted flags (see defines)
            Public dwLoops As Int32             ' loop control counter
            Public lpNext As IntPtr             ' PWaveHdr, reserved for driver
            Public reserved As IntPtr           ' reserved for driver
        End Structure

        ' --------------------------------------------------------------------- WaveOut
        <DllImport("winmm.dll")> _
        Public Function waveOutGetNumDevs() As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutOpen(ByRef hWaveOut As Int32, _
                                    ByVal uDeviceID As Int32, _
                                    ByRef lpFormat As WAVEFORMATEX, _
                                    ByVal dwCallback As Int32, _
                                    ByVal dwInstance As Int32, _
                                    ByVal dwFlags As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutPrepareHeader(ByVal hWaveOut As Int32, _
                                             ByVal lpWaveOutHdr As IntPtr, _
                                             ByVal uSize As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutUnprepareHeader(ByVal hWaveOut As Int32, _
                                               ByVal lpWaveOutHdr As IntPtr, _
                                               ByVal uSize As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutWrite(ByVal hWaveOut As Int32, _
                                     ByVal lpWaveOutHdr As IntPtr, _
                                     ByVal uSize As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutReset(ByVal hWaveOut As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutClose(ByVal hWaveOut As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutPause(ByVal hWaveOut As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutRestart(ByVal hWaveOut As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveOutGetPosition(ByVal hWaveOut As Int32, _
                                           ByRef lpInfo As Int32, _
                                           ByVal uSize As Int32) As Int32
        End Function

        ' --------------------------------------------------------------------- WaveIn
        <DllImport("winmm.dll")> _
        Public Function waveInGetNumDevs() As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInOpen(ByRef hWaveIn As Int32, _
                                   ByVal uDeviceID As Int32, _
                                   ByRef lpFormat As WAVEFORMATEX, _
                                   ByVal dwCallback As Int32, _
                                   ByVal dwInstance As Int32, _
                                   ByVal dwFlags As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInPrepareHeader(ByVal hWaveIn As Int32, _
                                            ByVal lpWaveInHdr As IntPtr, _
                                            ByVal uSize As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInUnprepareHeader(ByVal hWaveIn As Int32, _
                                              ByVal lpWaveInHdr As IntPtr, _
                                              ByVal uSize As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInAddBuffer(ByVal hwi As Int32, _
                                        ByVal pwh As IntPtr, _
                                        ByVal cbwh As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInClose(ByVal hwi As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInReset(ByVal hwi As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInStart(ByVal hwi As Int32) As Int32
        End Function
        <DllImport("winmm.dll")> _
        Public Function waveInStop(ByVal hwi As Int32) As Int32
        End Function

        ' --------------------------------------------------------------------- 
        '  GET WaveIn NAMES
        ' --------------------------------------------------------------------- 
        <StructLayout(LayoutKind.Sequential)> _
        Private Structure WAVEINCAPS
            Dim wMid As Int16
            Dim wPid As Int16
            Dim vDriverVersion As Int32
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public szPname As String
            Dim dwFormats As Int32
            Dim wChannels As Int16
            Dim wReserved1 As Int16
        End Structure

        ' DeviceID is a value from 1 to wInGetNumDevs (0 = default device)
        Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
                                                                (ByVal DeviceID As Int32, _
                                                                 ByRef Caps As WAVEINCAPS, _
                                                                 ByVal Size As Int32) As Int32

        Public Function GetInputDevicesNameList() As String
            Dim wcap As WAVEINCAPS = New WAVEINCAPS
            Dim ret As String = ""
            For i As Int32 = 0 To InputDevicesCount() - 1
                waveInGetDevCaps(i, wcap, Marshal.SizeOf(wcap))
                ret = ret + wcap.szPname.Trim + vbCrLf
            Next
            Return ret
        End Function

        Public Function GetInputDevicesNames() As String()
            Dim Names(InputDevicesCount() - 1) As String
            Dim wcap As WAVEINCAPS = New WAVEINCAPS
            For i As Int32 = 0 To InputDevicesCount() - 1
                waveInGetDevCaps(i, wcap, Marshal.SizeOf(wcap))
                Names(i) = wcap.szPname.Trim
            Next
            Return Names
        End Function


        ' --------------------------------------------------------------------- 
        '  GET WaveOut NAMES
        ' --------------------------------------------------------------------- 
        <StructLayout(LayoutKind.Sequential)> _
        Private Structure WAVEOUTCAPS
            Dim wMid As Int16
            Dim wPid As Int16
            Dim vDriverVersion As Int32
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public szPname As String
            Dim dwFormats As Int32
            Dim wChannels As Int16
            Dim wReserved1 As Int16
            Dim dwSupport As Int32
        End Structure

        ' DeviceID is a value from 1 to wInGetNumDevs (0 = default device)
        Private Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" _
                                                                (ByVal DeviceID As Int32, _
                                                                 ByRef Caps As WAVEOUTCAPS, _
                                                                 ByVal Size As Int32) As Int32

        Public Function GetOutputDevicesNameList() As String
            Dim wcap As WAVEOUTCAPS = New WAVEOUTCAPS
            Dim ret As String = ""
            For i As Int32 = 0 To OutputDevicesCount() - 1
                waveOutGetDevCaps(i, wcap, Marshal.SizeOf(wcap))
                ret = ret + wcap.szPname.Trim + vbCrLf
            Next
            Return ret
        End Function

        Public Function GetOutputDevicesNames() As String()
            Dim Names(OutputDevicesCount() - 1) As String
            Dim wcap As WAVEOUTCAPS = New WAVEOUTCAPS
            For i As Int32 = 0 To OutputDevicesCount() - 1
                waveOutGetDevCaps(i, wcap, Marshal.SizeOf(wcap))
                Names(i) = wcap.szPname.Trim
            Next
            Return Names
        End Function

    End Module
End Namespace