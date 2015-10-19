Attribute VB_Name = "modAudio"
Option Explicit

'DirectMusic's Performance object
Dim Performance As DirectMusicPerformance8

'Currently loaded segment
Dim Segment As DirectMusicSegment8

'The one and only DirectMusic Loader
Dim Loader As DirectMusicLoader8

'State of the currently loaded segment
Dim SegState As DirectMusicSegmentState8

'General Sound Volume
Dim SoundVolume As Long
 
'Maximum number of sound buffers to hold in memory at any time.
Private Const BufferSize As Byte = 30
 
'The distance at which the sound is inaudible.
Private Const MAX_DISTANCE_TO_SOURCE As Integer = 150
 
'Custom sound buffer structure.
Private Type SoundBuffer
    FileName As String
    Looping As Boolean
    X As Byte
    Y As Byte
    normalFq As Long
    Buffer As clsAudio
End Type

'States how to set a sound's looping state.
Public Enum LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum
 
'States the last position where the listener was in the X-Y axis
Dim lastPosX As Integer, lastPosY As Integer
 
'Array of all existing sound buffers
Dim SoundBuffers(1 To BufferSize) As SoundBuffer
 
'Destructor. Releases all created objects assuring no memory-leaks.
Public Sub Sound_Destroy()
    Dim loopc As Long
   
    'Stop every channel being used and destroy the buffer
    For loopc = 1 To BufferSize
        If Not SoundBuffers(loopc).Buffer Is Nothing Then
            SoundBuffers(loopc).Buffer.Stopping
            SoundBuffers(loopc).Buffer.Destroy
            Set SoundBuffers(loopc).Buffer = Nothing
        End If
    Next loopc
End Sub

Public Sub Sound_Play(ByVal FileName As String, Optional ByVal srcX As Byte = 0, Optional ByVal srcY As Byte = 0, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default)
On Error GoTo ErrHandler

    Dim BufferIndex As Long

    'Get the buffer index were wave was loaded
    BufferIndex = Sound_Load(FileName, LoopSound)
    
    If BufferIndex = 0 Then
        Exit Sub
    End If
    
    With SoundBuffers(BufferIndex)
        
        SoundVolume = 1000
        
        'Apply volume
        Call .Buffer.SetVolume(SoundVolume)

        'If .Looping Then
            .Buffer.Play 'DSBPLAY_LOOPING
        'Else
        '   .Buffer.Play 'DSBPLAY_DEFAULT
        'End If

        'Store position
        .X = srcX
        .Y = srcY

    End With

    If srcX <> 0 And srcY <> 0 Then
        Call Sound_Update3D(BufferIndex, 0, 0)
    End If
Exit Sub
 
ErrHandler:
End Sub

Public Function Sound_PlayF(ByVal FileName As String, Optional ByVal srcX As Byte = 0, Optional ByVal srcY As Byte = 0, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long
On Error GoTo ErrHandler
    Dim BufferIndex As Long

    'Get the buffer index were wave was loaded
    BufferIndex = Sound_Load(FileName, LoopSound)
    If BufferIndex = 0 Then Exit Function   'If an error ocurred abort

    With SoundBuffers(BufferIndex)
        
        SoundVolume = 1000
        
        'Apply volume
        Call .Buffer.SetVolume(SoundVolume)

        If .Looping Then
            .Buffer.Play 'DSBPLAY_LOOPING
        Else
            .Buffer.Play 'DSBP``LAY_DEFAULT
        End If

        'Store position
        .X = srcX
        .Y = srcY

    End With

    Sound_PlayF = BufferIndex

    If srcX <> 0 And srcY <> 0 Then
        Call Sound_Update3D(BufferIndex, 0, 0)
    End If
Exit Function
 
ErrHandler:
End Function
 
Private Sub Sound_Update3D(ByVal BufferIndex As Long, ByVal deltaX As Integer, ByVal deltaY As Integer)
    Dim linearDistanceOld As Single
    Dim linearDistanceNew As Single
    Dim deltaDistance As Single
    Dim distanceXOld As Integer
    Dim distanceYOld As Integer
    Dim distanceXNew As Integer
    Dim distanceYNew As Integer
   
    With SoundBuffers(BufferIndex)
        distanceXOld = .X - lastPosX
        distanceYOld = .Y - lastPosY
       
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
       
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
       
        deltaDistance = linearDistanceNew - linearDistanceOld
  
        'Set volumen amortiguation according to distance
        Call .Buffer.SetVolume(SoundVolume * (1 - linearDistanceNew / MAX_DISTANCE_TO_SOURCE))
  
        'Prevent division by zero
        If linearDistanceNew = 0 Then linearDistanceNew = 1
  
        'Set panning according to relative position of the source to the listener
        Call .Buffer.SetBalance((distanceXNew / linearDistanceNew) * DSBPAN_RIGHT)
    End With
End Sub
 
Public Sub Sound_MoveListener(ByVal X As Integer, ByVal Y As Integer)
'Updates 3D sounds based on the movement of the listener.
        
On Error GoTo Error

    Dim i As Long
    Dim deltaX As Integer
    Dim deltaY As Integer
   
    deltaX = X - lastPosX
    deltaY = Y - lastPosY
   
    For i = 1 To BufferSize
        If Not SoundBuffers(i).Buffer Is Nothing Then
            If SoundBuffers(i).Buffer.status = 1 Then
                If SoundBuffers(i).X <> 0 And SoundBuffers(i).Y <> 0 Then
                    Call Sound_Update3D(i, deltaX, deltaY)
                End If
            End If
        End If
    Next i
   
    lastPosX = X
    lastPosY = Y

Error:
End Sub

Private Function Sound_Load(ByVal FileName As String, ByVal Looping As LoopStyle) As Long
On Error GoTo ErrHandler

    Dim i As Long

    FileName = UCase$(FileName)

    If Not FileExist(SfxPath & FileName, vbArchive) Then
        Exit Function
    End If
    
    'Check if the buffer is in memory and not playing
    For i = 1 To BufferSize
        If SoundBuffers(i).FileName = FileName Then
            If Not SoundBuffers(i).Buffer Is Nothing Then
                If SoundBuffers(i).Buffer.status = 0 Then
                    'Found it!!! We just play this one :)
                    Sound_Load = i
    
                    'Set looping if needed
                    If Looping <> LoopStyle.Default Then
                        SoundBuffers(i).Looping = (Looping = LoopStyle.Enabled)
                    End If
                    
                    Exit Function
                End If
            End If
        End If
    Next i

    'Not in memory, search for an empty buffer
    For i = 1 To BufferSize
        If SoundBuffers(i).Buffer Is Nothing Then
        
            With SoundBuffers(i)
                Set .Buffer = Nothing   'Get rid of any previous data
        
                .FileName = FileName
                .Looping = (Looping = LoopStyle.Enabled)
        
                Set .Buffer = New clsAudio
                .Buffer.Load (SfxPath & FileName)
            End With
        
            Sound_Load = i
            
            Exit Function
        End If
    Next i

Exit Function

ErrHandler:
End Function
 
'Stops a given sound or all of them.
Public Sub Sound_Stop(Optional ByVal BufferIndex As Long = 0)

    If BufferIndex > 0 And BufferIndex <= BufferSize Then
        If SoundBuffers(BufferIndex).Buffer.status = 1 Then
            Call SoundBuffers(BufferIndex).Buffer.Stopping
        End If
    ElseIf BufferIndex = 0 Then
        Dim i As Long
        For i = 1 To BufferSize
            If Not SoundBuffers(i).Buffer Is Nothing Then
                If SoundBuffers(i).Buffer.status = 0 Then
                    Call SoundBuffers(i).Buffer.Stopping
                End If
            End If
        Next i
    End If
    
End Sub
 
'Retrieves wether there are sounds currentyl playing or not.
Public Function Sound_Playing() As Boolean
    Dim i As Long
   
    For i = 1 To BufferSize
        If SoundBuffers(i).Buffer.status = 1 Then
            Sound_Playing = True
            Exit Function
        End If
    Next i
    
End Function
 
'Sets the volume of sound.
Public Function Sound_Volume(Volume As Long)
    Dim i As Long
    
    'Take percentage to actual value
    SoundVolume = Volume
    
    For i = 1 To BufferSize
        If Not SoundBuffers(i).Buffer Is Nothing Then
            If SoundBuffers(i).Buffer.status = 1 Then
                Call SoundBuffers(i).Buffer.SetVolume(SoundVolume)
            End If
        End If
    Next i
    
End Function

'Creates and configures all DirectMusic objects.
Public Function Music_Initialize() As Boolean
On Error GoTo ErrHandler
    
    Dim musParams As DMUS_AUDIOPARAMS

    Set Loader = DirectX_8.DirectMusicLoaderCreate()
    
    Set Performance = DirectX_8.DirectMusicPerformanceCreate()
    Performance.InitAudio frmMain.hWnd, DMUS_AUDIOF_ALL, musParams, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    Performance.SetMasterAutoDownload True 'Enable auto download of instruments
        
    'Set tempo to 0 and volume of music
    Music_Tempo 0
    Performance.SetMasterVolume 200
    
    Music_Initialize = True
Exit Function

ErrHandler:
    Debug.Print "Error in Music_Play"
End Function

'Plays a new MIDI file.
Public Sub Music_Play(Optional ByVal file As String = vbNullString, Optional ByVal Loops As Long = -1)
'On Error GoTo ErrHandler

    If Music_Playing() Then Music_Stop
    
    If LenB(file) > 0 Then
        If Not Music_Load(file) Then
            Exit Sub
        End If
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then
            Exit Sub
        End If
    End If
    
    'Play it
    Segment.SetRepeats Loops
    
    Set SegState = Performance.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0)
    
Exit Sub

ErrHandler:
    Debug.Print "Error in Music_Play"
End Sub

'Loads a new MIDI file.
Private Function Music_Load(ByVal file As String) As Boolean

On Error GoTo ErrHandler

    If Not FileExist(MusicPath & file, vbArchive) Then
        Exit Function
    End If
    
    Music_Stop
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(MusicPath & file)
    
    If Segment Is Nothing Then
        Exit Function
    End If
    
    Segment.SetStandardMidiFile
    
    Music_Load = True
Exit Function

ErrHandler:
    Debug.Print "Error in Music_Load"
End Function

''
' Stops playing the currently loaded MIDI file.

Public Sub Music_Stop()
On Error GoTo ErrHandler

    If Music_Playing Then
        Performance.StopEx Segment, 0, DMUS_SEGF_DEFAULT
    End If
    
Exit Sub
ErrHandler:
    Debug.Print "Error in Music_Stop"
End Sub

''
' Checks wether there is music currently playing.

Public Function Music_Playing() As Boolean

    If Segment Is Nothing Then Exit Function
    
    Music_Playing = Performance.IsPlaying(Segment, SegState)
End Function

''
' Retrieves the music's length.

Public Function Music_Lenght() As Long

    Music_Lenght = Segment.GetLength()
    
End Function

''
' Sets the music's volume.

Public Sub Music_Volume(ByVal Volume As Long)
    
    If Volume < 0 Or Volume > 100 Then Exit Sub
    
    ' Volume ranges from -10000 to 10000
    Performance.SetMasterVolume Volume * 200 - 10000
End Sub

''
' Sets the music's tempo.

Public Sub Music_Tempo(ByVal Tempo As Single)
    
    If Tempo < 0.25 Or Tempo > 2# Then Exit Sub
    
    Performance.SetMasterTempo Tempo
End Sub


