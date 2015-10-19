VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsByteQueue.cls - FIFO list of bytes.
'Creates and manipulates byte arrays to be sent and received by both client and server
'Used to create and manipulate the byte arrays to be sent and received by both client and server

Option Explicit

Option Base 0       'It's the default, but we make it explicit just in case...

'The Error number thrown when there is not enough data in
'the buffer to read the specified data type.
'It's 9 (subscript out of range) + the object Error constant
Private Const NOT_ENOUGH_DATA As Long = vbObjectError + 9

'The Error number thrown when there is not enough space in
'the buffer to write.
Private Const NOT_ENOUGH_SPACE As Long = vbObjectError + 10

'Default size of a data buffer (10 Kbs)
Private Const DATA_BUFFER As Long = 10240

'The byte data
Dim data() As Byte

'How big the data array is
Dim queueCapacity As Long

'How far into the data array have we written
Dim queueLength As Long

Private Sub Class_Initialize()
'Initializes the queue with the default queueCapacity
    ReDim data(DATA_BUFFER - 1) As Byte
    
    queueCapacity = DATA_BUFFER
End Sub

Private Sub Class_Terminate()
'Clean up
    Erase data
End Sub

Public Sub CopyBuffer(ByRef source As clsByteQueue)
'A Visual Basic equivalent of a Copy Contructor

    If source.length = 0 Then
        'Clear the list and exit
        Call RemoveData(length)
        Exit Sub
    End If
    
    'Set capacity and resize array - make sure all data is lost
    queueCapacity = source.Capacity
    
    ReDim data(queueCapacity - 1) As Byte
    
    'Read buffer
    Dim buf() As Byte
    ReDim buf(source.length - 1) As Byte
    
    Call source.PeekBlock(buf, source.length)
    
    queueLength = 0
    
    'Write buffer
    Call WriteBlock(buf, source.length)
End Sub

Private Function Min(ByVal val1 As Long, ByVal val2 As Long) As Long
'It's faster than iif and I like it better
    If val1 < val2 Then
        Min = val1
    Else
        Min = val2
    End If
End Function

Private Function WriteData(ByRef buf() As Byte, ByVal dataLength As Long) As Long

'If the queueCapacity allows it copyes a byte buffer to the queue, if not it throws NOT_ENOUGH_DATA
    'Check if there is enough free space
    If queueCapacity - queueLength - dataLength < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function
    End If
    
    'Copy data from buffer
    Call CopyMemory(data(queueLength), buf(0), dataLength)
    
    'Update length of data
    queueLength = queueLength + dataLength
    WriteData = dataLength
End Function

Private Function ReadData(ByRef buf() As Byte, ByVal dataLength As Long) As Long
'If enough memory is available, it copies the requested number of bytes to the buffer
    'Check if we can read the number of bytes requested
        
    If dataLength > queueLength Then
    'Call Err.Raise(NOT_ENOUGH_DATA)
    'exit function
    End If
    
    'Copy data to buffer
    Call CopyMemory(buf(0), data(0), dataLength)
    ReadData = dataLength
    
End Function

Public Function RemoveData(ByVal dataLength As Long) As Long
'Moves forward the queue overwriting the first dataLength bytes
    'Figure out how many bytes we can remove
    RemoveData = Min(dataLength, queueLength)
    
    'Remove data - prevent rt9 when cleaning a full queue
    If RemoveData <> queueCapacity Then
        Call CopyMemory(data(0), data(RemoveData), queueLength - RemoveData)
    End If
    
    'Update length
    queueLength = queueLength - RemoveData
End Function

Public Function WriteByte(ByVal Value As Byte) As Long
'Writes a byte to the queue

    Dim buf(0) As Byte
    
    buf(0) = Value
    
    WriteByte = WriteData(buf, 1)
    
    'Debug.Print "1 WriteByte " & Value
End Function

Public Function WriteInteger(ByVal Value As Integer) As Long
'Writes an integer to the queue
    Dim buf(1) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), Value, 2)
    
    WriteInteger = WriteData(buf, 2)

    'Debug.Print "2 WriteInteger " & Value
End Function

Public Function WriteLong(ByVal Value As Long) As Long
'Writes a long to the queue
    Dim buf(3) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), Value, 4)
    
    WriteLong = WriteData(buf, 4)

    'Debug.Print "4 WriteLong " & Value
End Function

Public Function WriteSingle(ByVal Value As Single) As Long
'Writes a single to the queue
    Dim buf(3) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), Value, 4)
    
    WriteSingle = WriteData(buf, 4)

    'Debug.Print "4 WriteSingle " & Value
End Function

Public Function WriteDouble(ByVal Value As Double) As Long
'Writes a double to the queue
    Dim buf(7) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), Value, 8)
    
    WriteDouble = WriteData(buf, 8)

    'Debug.Print "4 WriteDouble " & Value
End Function

Public Function WriteBoolean(ByVal Value As Boolean) As Long
'Writes a byte to the queue
    Dim buf(0) As Byte
    
    If Value Then
        buf(0) = 1
    End If
    
    WriteBoolean = WriteData(buf, 1)
    
    'Debug.Print "4 WriteBoolean " & Value
End Function

Public Function WriteASCIIStringFixed(ByVal Value As String) As Long
'Writes a fixed length ASCII string to the queue
    Dim buf() As Byte
    ReDim buf(Len(Value) - 1) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), ByVal StrPtr(StrConv(Value, vbFromUnicode)), Len(Value))
    
    WriteASCIIStringFixed = WriteData(buf, Len(Value))

    'Debug.Print Len(Value) & " WriteASCIIStringFixed " & Value
End Function

Public Function WriteUnicodeStringFixed(ByVal Value As String) As Long
'Writes a fixed length UNICODE string to the queue
    Dim buf() As Byte
    ReDim buf(LenB(Value)) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), ByVal StrPtr(Value), LenB(Value))
    
    WriteUnicodeStringFixed = WriteData(buf, LenB(Value))

    'Debug.Print LenB(Value) & " WriteUnicodeStringFixed " & Value
End Function

Public Function WriteASCIIString(ByVal Value As String) As Long
'Writes a variable length ASCII string to the queue

    Dim buf() As Byte
    ReDim buf(Len(Value) + 1) As Byte
    
    'Copy length to temp buffer
    Call CopyMemory(buf(0), CInt(Len(Value)), 2)
    
    If Len(Value) > 0 Then
        'Copy data to temp buffer
        Call CopyMemory(buf(2), ByVal StrPtr(StrConv(Value, vbFromUnicode)), Len(Value))
    End If
    
    WriteASCIIString = WriteData(buf, Len(Value) + 2)
    
    'Debug.Print Len(Value) + 2 & " WriteASCIIString " & Value
End Function

Public Function WriteUnicodeString(ByVal Value As String) As Long
'Writes a variable length UNICODE string to the queue

    Dim buf() As Byte
    ReDim buf(LenB(Value) + 1) As Byte
    
    'Copy length to temp buffer
    Call CopyMemory(buf(0), CInt(Len(Value)), 2)
    
    If Len(Value) > 0 Then
        'Copy data to temp buffer
        Call CopyMemory(buf(2), ByVal StrPtr(Value), LenB(Value))
    End If
    
    WriteUnicodeString = WriteData(buf, LenB(Value) + 2)

    'Debug.Print LenB(Value) + 2 & " WriteUnicodeString " & Value
End Function

Public Function WriteBlock(ByRef Value() As Byte, Optional ByVal length As Long = -1) As Long
'Writes a byte array to the queue
    'Prevent from copying memory outside the array
    If length > UBound(Value()) + 1 Or length < 0 Then
        length = UBound(Value()) + 1
    End If
    
    WriteBlock = WriteData(Value, length)

    'Debug.Print length & " WriteBlock " & Value
End Function

Public Function ReadByte() As Byte
'Reads a byte from the queue and removes it

    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 1))
    
    ReadByte = buf(0)
    
    'Debug.Print "1 ReadByte " & ReadByte
End Function

Public Function ReadInteger() As Integer
'Reads an integer from the queue and removes it
    Dim buf(1) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 2))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadInteger, buf(0), 2)
    
    'Debug.Print "2 ReadInteger " & ReadInteger
End Function

Public Function ReadLong() As Long
'Reads a long from the queue and removes it
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 4))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadLong, buf(0), 4)
    
    'Debug.Print "4 ReadLong " & ReadLong
End Function

Public Function ReadSingle() As Single
'Reads a single from the queue and removes it
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 4))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadSingle, buf(0), 4)
    
    'Debug.Print "4 ReadSingle " & ReadSingle
End Function

Public Function ReadDouble() As Double
'Reads a double from the queue and removes it
    Dim buf(7) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 8))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadDouble, buf(0), 8)
    
    'Debug.Print "8 ReadDouble " & ReadDouble
End Function

Public Function ReadBoolean() As Boolean
'Reads a Boolean from the queue and removes it
    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 1))
    
    If buf(0) = 1 Then
        ReadBoolean = True
    End If
    
    'Debug.Print "1 ReadBoolean " & ReadBoolean
End Function

Public Function ReadASCIIStringFixed(ByVal length As Long) As String
'Reads a fixed length ASCII string from the queue and removes it
    If length < 1 Then
        Exit Function
    End If
    
    If queueLength >= length Then
        Dim buf() As Byte
        ReDim buf(length - 1) As Byte
        
        'Read the data and remove it
        Call RemoveData(ReadData(buf, length))
        
        ReadASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
    
    'Debug.Print Length & " ReadASCIIStringFixed " & ReadASCIIStringFixed
End Function

Public Function ReadUnicodeStringFixed(ByVal length As Long) As String
'Reads a fixed length UNICODE string from the queue and removes it
    If length < 1 Then
        Exit Function
    End If
    
    If queueLength >= length + length Then
        Dim buf() As Byte
        ReDim buf(length + length - 1) As Byte
        
        'Read the data and remove it
        Call RemoveData(ReadData(buf, length + length))
        
        ReadUnicodeStringFixed = buf
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
    
    'Debug.Print Length + Length & " ReadUnicodeStringFixed " & ReadUnicodeStringFixed
End Function

Public Function ReadASCIIString() As String
'Reads a variable length ASCII string from the queue and removes it

    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length + 2 Then
            'Remove the length
            Call RemoveData(2)
            
            If length > 0 Then
                Dim buf2() As Byte
                ReDim buf2(length - 1) As Byte
                
                
                'Read the data and remove it
                Call RemoveData(ReadData(buf2, length))
                
                ReadASCIIString = StrConv(buf2, vbUnicode)
            End If
        Else
'Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
'        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
        
    'Debug.Print Length & " ReadASCIIString " & ReadASCIIString
End Function

Public Function ReadUnicodeString() As String
'Reads a variable length UNICODE string from the queue and removes it

    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length + length + 2 Then
            'Remove the length
            Call RemoveData(2)
            
            Dim buf2() As Byte
            ReDim buf2(length + length - 1) As Byte
            
            'Read the data and remove it
            Call RemoveData(ReadData(buf2, length + length))
            
            ReadUnicodeString = buf2
        Else
            Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
    
    'Debug.Print Length + Length & " ReadUnicodeString " & ReadUnicodeString
End Function

Public Function ReadBlock(ByRef block() As Byte, ByVal dataLength As Long) As Long
'Reads a byte array from the queue and removes it
    'Read the data and remove it
    If dataLength > 0 Then
        ReadBlock = RemoveData(ReadData(block(), dataLength))
    End If
        
    'Debug.Print dataLength & " ReadBlock " & ReadBlock

End Function

Public Function PeekByte() As Byte
'Reads a byte from the queue but doesn't removes it
    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 1)
    
    PeekByte = buf(0)
End Function
Public Function PeekInteger() As Integer
'Reads an integer from the queue but doesn't removes it
    Dim buf(1) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 2)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekInteger, buf(0), 2)
End Function

Public Function PeekLong() As Long
'Reads a long from the queue but doesn't removes it
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 4)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekLong, buf(0), 4)
End Function

Public Function PeekSingle() As Single
'Reads a single from the queue but doesn't removes it
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 4)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekSingle, buf(0), 4)
End Function

Public Function PeekDouble() As Double
'Reads a double from the queue but doesn't removes it
    Dim buf(7) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 8)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekDouble, buf(0), 8)
End Function

Public Function PeekBoolean() As Boolean
'Reads a Boolean from the queue but doesn't removes it
    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 1)
    
    If buf(0) = 1 Then
        PeekBoolean = True
    End If
End Function

Public Function PeekASCIIStringFixed(ByVal length As Long) As String
'Reads a fixed length ASCII string from the queue but doesn't removes it
    If length < 1 Then
        Exit Function
    End If
    
    If queueLength >= length Then
        Dim buf() As Byte
        ReDim buf(length - 1) As Byte
        
        'Read the data and remove it
        Call ReadData(buf, length)
        
        PeekASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekUnicodeStringFixed(ByVal length As Long) As String
'Reads a fixed length UNICODE string from the queue but doesn't removes it

    If length < 1 Then
        Exit Function
    End If
    
    If queueLength >= length + length Then
        Dim buf() As Byte
        ReDim buf(length + length - 1) As Byte
        
        'Read the data and remove it
        Call ReadData(buf, length + length)
        
        PeekUnicodeStringFixed = buf
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekASCIIString() As String
'Reads a variable length ASCII string from the queue but doesn't removes it

    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length + 2 Then
            Dim buf2() As Byte
            ReDim buf2(length + 1) As Byte
            
            'Read the data (we have to read the length again)
            Call ReadData(buf2, length + 2)
            
            If length > 0 Then
                'Get rid of the length
                Dim buf3() As Byte
                ReDim buf3(length - 1) As Byte
                Call CopyMemory(buf3(0), buf2(2), length)
                
                PeekASCIIString = StrConv(buf3, vbUnicode)
            End If
        Else
            Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekUnicodeString() As String
'Reads a variable length UNICODE string from the queue but doesn't removes it

    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length + length + 2 Then
            Dim buf2() As Byte
            ReDim buf2(length + length + 1) As Byte
            
            'Read the data (we need to read the length again)
            Call ReadData(buf2, length + length + 2)
            
            'Get rid of the length bytes
            Dim buf3() As Byte
            ReDim buf3(length + length - 1) As Byte
            Call CopyMemory(buf3(0), buf2(2), length + length)
            
            PeekUnicodeString = buf3
        Else
            Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekBlock(ByRef block() As Byte, ByVal dataLength As Long) As Long
'Reads a byte array from the queue but doesn't removes it
    'Read the data
    If dataLength > 0 Then
        PeekBlock = ReadData(block(), dataLength)
    End If
End Function

Public Property Get Capacity() As Long
'Retrieves the current capacity of the queue
    Capacity = queueCapacity
End Property

Public Property Let Capacity(ByVal Value As Long)
'Sets the current capacity of the queue.
'All data in the queue exceeding the new capacity is lost
    'Upate capacity
    queueCapacity = Value
    
    'All extra data is lost
    If length > Value Then
        queueLength = Value
    End If
    
    'Resize the queue
    ReDim Preserve data(queueCapacity - 1) As Byte
End Property

Public Property Get length() As Long
'Retrieves the current number of bytes in the queue
    length = queueLength
End Property

Public Property Get NotEnoughDataErrCode() As Long
'Retrieves the NOT_ENOUGH_DATA Error code
    NotEnoughDataErrCode = NOT_ENOUGH_DATA
End Property

Public Property Get NotEnoughSpaceErrCode() As Long
'Retrieves the NOT_ENOUGH_SPACE Error code
    NotEnoughSpaceErrCode = NOT_ENOUGH_SPACE
End Property

