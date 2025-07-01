Public Class FIFO

    Private m_FifoLength As Int32
    Private m_WriteIndex As Int32
    Private m_ReadIndex As Int32
    Private m_CircularBuffer() As Int16

    Friend Sub New(ByVal FifoLength As Int32)
        m_FifoLength = FifoLength
        ReDim m_CircularBuffer(m_FifoLength - 1)
        Flush()
    End Sub

    Friend Sub Flush()
        m_ReadIndex = 0
        m_WriteIndex = 0
    End Sub

    Friend Function GetWriteIndex() As Int32
        Return m_WriteIndex
    End Function

    Friend Sub SetReadIndex(ByVal position As Int32)
        m_ReadIndex = position
        While m_ReadIndex < 0
            m_ReadIndex += m_FifoLength
        End While
    End Sub

    Friend Sub AddArray(ByVal Values() As Int16)
        'SyncLock Me
        For i As Int32 = 0 To Values.Length - 1
            m_CircularBuffer(m_WriteIndex) = Values(i)
            m_WriteIndex += 1
            If m_WriteIndex >= m_FifoLength Then m_WriteIndex = 0
        Next
        'End SyncLock
    End Sub

    Friend Sub GetArray(ByRef Values() As Int16)
        'SyncLock Me
        For i As Int32 = 0 To Values.Length - 1
            Values(i) = m_CircularBuffer(m_ReadIndex)
            m_ReadIndex += 1
            If m_ReadIndex >= m_FifoLength Then m_ReadIndex = 0
        Next
        'End SyncLock
    End Sub

    Friend Sub AddElement(ByVal Value As Int16)
        'SyncLock Me
        m_CircularBuffer(m_WriteIndex) = Value
        m_WriteIndex += 1
        If m_WriteIndex >= m_FifoLength Then m_WriteIndex = 0
        'End SyncLock
    End Sub

    Friend Function GetElement() As Int16
        'SyncLock Me
        GetElement = m_CircularBuffer(m_ReadIndex)
        m_ReadIndex += 1
        If m_ReadIndex >= m_FifoLength Then m_ReadIndex = 0
        'End SyncLock
    End Function

    Friend Function GetElement(ByVal Index As Int32) As Int16
        'SyncLock Me
        While Index < 0 : Index += m_FifoLength : End While
        While Index >= m_FifoLength : Index -= m_FifoLength : End While
        Return m_CircularBuffer(Index)
        'End SyncLock
    End Function

End Class
