VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Controls all audio using DirectSound and DirectMusic 7.
Option Explicit

'Maximum number of sound buffers to hold in memory at any time.
Private Const NumSoundBuffers As Byte = 30

'The distance at which the sound is inaudible.
Private Const MAX_DISTANCE_TO_SOURCE As Integer = 150

'The frequency variation coefficient for Doppler effect simulation.
Private Const DELTA_FQ As Integer = 75

'Custom sound buffer structure.
Private Type SoundBuffer
    FileName As String
    looping As Boolean
    x As Integer
    y As Integer
    normalFq As Long
    Buffer As DirectSoundBuffer
End Type

'States how to set a sound's looping state.
Public Enum LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum

'Set to true once the object is created to prevent multiple initializations of the same object.
Dim inited As Boolean

'The one and only DirectSound object
Dim DirectSound As DirectSound

'Array of all existing sound buffers
Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer

'Directory from which to load sound files by default
Dim SoundPath As String

'The volume for sounds (FXs)
Dim SndVolume As Long

'DirectMusic's Performance object
Dim Performance As DirectMusicPerformance

'Currently loaded segment
Dim Segment As DirectMusicSegment

'The one and only DirectMusic Loader
Dim Loader As DirectMusicLoader

'State of the currently loaded segment
Dim SegState As DirectMusicSegmentState

'States wether music is supported by hardware or not
Dim Music As Boolean

'Last loaded music file. Used to be able to load it when music gets activated
Dim CurrentMusicFile As String

'States the last position where the listener was in the X axis
Dim lastPosX As Integer

'States the last position where the listener was in the Y axis
Dim lastPosY As Integer

Private LastPlayMP3 As Boolean

'Used By MP3 Playing
Dim IMC   As IMediaControl
Dim IBA   As IBasicAudio
Dim IME   As IMediaEvent
Dim IMPos As IMediaPosition

'Constructor. Performs no action at all.
Private Sub Class_Initialize()
End Sub

'Destructor. Releases all created objects assuring no memory-leaks.
Private Sub Class_Terminate()
'Destroy's all DirectMusic and DirectSound objects

On Error Resume Next
    'Stop every channel being used and destroy the buffer
    Call StopWave
    
    'Stop music
    Call Performance.Stop(Segment, SegState, 0, 0)
    
    Call Performance.CloseDown
    
    Set DirectSound = Nothing
    
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
    
    inited = False
End Sub

Public Function Initialize(ByRef dX As DirectX7, ByVal hWnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean
'Initializes the clsAudio object.
'Creates all DS and DM objects

On Error GoTo ErrHandler
    
    'Prevent multiple creations
    If inited Then
        Exit Function
    End If
    
    'Validate DX object
    If dX Is Nothing Then
        GoTo ErrHandler
    End If
    
    'Initialize DirectSound
    Set DirectSound = dX.DirectSoundCreate(vbNullString)
    
    Call DirectSound.SetCooperativeLevel(hWnd, DSSCL_PRIORITY)
    
    'Set sound path
    SoundPath = sound_path
    
    'Set default volume
    SndVolume = 0
    
    'Initialize DirectMusic
    Music = InitializedirectMusic(dX, hWnd)
    
    CurrentMusicFile = vbNullString
    
    'Set tempo to 0 and volume of music
    MusicTempo = 0
    Call Performance.SetMasterVolume(200)
    
    inited = True
    
    Initialize = True
Exit Function

ErrHandler:
End Function

Private Function InitializedirectMusic(ByRef dX As DirectX7, ByVal hWnd As Long) As Boolean
'Creates and configures all DirectMusic objects.
'Initialize DM. Returns false if music isn't supported by HW

On Error GoTo ErrHandler
    Set Loader = DirectX.DirectMusicLoaderCreate()
    
    Set Performance = dX.DirectMusicPerformanceCreate()
    Call Performance.Init(DirectSound, hWnd)
    Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
    
    'Set to port default, 1 channel group is enough for MIDIs
    Call Performance.SetPort(-1, 1)
    
    InitializedirectMusic = True
Exit Function

ErrHandler:
End Function

Public Function Play(ByVal FileName As String, Optional ByVal srcX As Integer = 0, Optional ByVal srcY As Integer = 0, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long
'Plays a new WAVE file.

On Error GoTo ErrHandler
    Dim bufferIndex As Long
    
    If Not SoundEffectsActivated Then
        Exit Function
    End If
    
    'Get the buffer index were wave was loaded
    bufferIndex = LoadWave(FileName, LoopSound)
    
    If bufferIndex = 0 Then
        Exit Function   'If an Error ocurRed abort
    End If
    
    With DSBuffers(bufferIndex)
        If .looping Then
            .Buffer.Play DSBPLAY_LOOPING
        Else
            .Buffer.Play DSBPLAY_DEFAULT
        End If
        
        'Apply volume
        Call .Buffer.SetVolume(SndVolume)
        
        'Store position
        .x = srcX
        .y = srcY
        
        'Get standard frequency
        .normalFq = .Buffer.GetFrequency()
    End With
    
   Play = bufferIndex
    
    If SoundEffectsActivated Then
        If srcX > 0 And srcY > 0 Then
            Call Update3DSound(bufferIndex, 0, 0)
        End If
    End If
Exit Function

ErrHandler:
End Function

Private Sub ResetSoundValues()
'Reset the frequency, volume and pan.

    Dim i As Long
    
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            With DSBuffers(i)
                Call .Buffer.SetFrequency(.normalFq)
                Call .Buffer.SetVolume(SndVolume)
                Call .Buffer.SetPan(0)
            End With
        End If
    Next i
End Sub

Private Sub Update3DSound(ByVal bufferIndex As Long, ByVal deltaX As Integer, ByVal deltaY As Integer)
'Refreshes the sound's arguments.

    Dim linearDistanceOld As Single
    Dim linearDistanceNew As Single
    Dim deltaDistance As Single
    Dim distanceXOld As Long
    Dim distanceYOld As Long
    Dim distanceXNew As Long
    Dim distanceYNew As Long
    Dim perceivedFq As Long
    
    With DSBuffers(bufferIndex)
        distanceXOld = .x - lastPosX
        distanceYOld = .y - lastPosY
        
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
        
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
        
        deltaDistance = linearDistanceNew - linearDistanceOld
        
        'Update frequency emulating Doppler effect
        perceivedFq = .normalFq * (1 - deltaDistance / DELTA_FQ)
        
        If perceivedFq < DSBFREQUENCY_MIN Then
            perceivedFq = DSBFREQUENCY_MIN
        End If
        
        Call .Buffer.SetFrequency(perceivedFq)
        
        'Set volumen amortiguation according to distance
        If linearDistanceNew > MAX_DISTANCE_TO_SOURCE Then
            Call .Buffer.SetVolume(SndVolume + (DSBVOLUME_MIN - SndVolume))
        Else
            If UserMoving Then
                Call .Buffer.SetVolume(SndVolume + (linearDistanceNew / MAX_DISTANCE_TO_SOURCE) * (DSBVOLUME_MIN - SndVolume))
            Else
                Call .Buffer.SetVolume(SndVolume + (linearDistanceNew / MAX_DISTANCE_TO_SOURCE) * (DSBVOLUME_MIN - SndVolume))
            End If
        End If

        'Prevent division by zero
        If linearDistanceNew = 0 Then
            linearDistanceNew = 1
        End If
        
        'Set panning according to relative position of the source to the listener
        Call .Buffer.SetPan((distanceXNew / linearDistanceNew) * DSBPAN_RIGHT)
    End With
End Sub

Public Sub MoveListener(ByVal x As Integer, ByVal y As Integer)
'Updates 3D sounds based on the movement of the listener.
    
    Dim i As Byte
    Dim deltaX As Integer
    Dim deltaY As Integer
    
    If Not SoundEffectsActivated Then
        Exit Sub
    End If
    
    deltaX = x - lastPosX
    
    deltaY = y - lastPosY
    
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                If DSBuffers(i).x > 0 And DSBuffers(i).y > 0 Then
                    Call Update3DSound(i, deltaX, deltaY)
                End If
            End If
        End If
    Next i
    
    lastPosX = x
    lastPosY = y
End Sub

Private Function LoadWave(ByVal FileName As String, ByVal looping As LoopStyle) As Long
'Loads a new wave into memory (if it's not already loaded) and returns
'the buffer index

On Error GoTo ErrHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim i As Long
    
    FileName = UCase$(FileName)
    
    If Not FileExist(SoundPath & FileName, vbArchive) Then
        Exit Function
    End If
    
    'Check if the buffer is in memory and not playing
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).FileName = FileName Then
            If DSBuffers(i).Buffer.GetStatus = 0 Or DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                'Found it! We just play this one :)
                LoadWave = i
                
                'Set looping if needed
                If looping <> LoopStyle.Default Then
                    DSBuffers(i).looping = (looping = LoopStyle.Enabled)
                End If
                
                'Reset default sound frecuency to avoid issues.
                Call DSBuffers(i).Buffer.SetFrequency(DSBuffers(i).normalFq)
                
                Exit Function
            End If
        End If
    Next i
    
    'Not in memory, search for an empty buffer
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).Buffer Is Nothing Then
            Exit For
        End If
    Next i
    
    'No empty buffers?? Search for one not being played and use that Slot
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or DSBuffers(i).Buffer.GetStatus() = 0 Then
                Exit For
            End If
        Next i
    End If
    
    'Still no free buffers?? Get the first non-looping sound
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If Not (DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Call DSBuffers(i).Buffer.Stop
                Exit For
            End If
        Next i
    End If
    
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first Slot by default)
    If i > NumSoundBuffers Then
        If Not looping Then
            Exit Function
        End If
        i = 1
    End If
    
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
    With DSBuffers(i)
        Set .Buffer = Nothing   'Get rid of any previous data
        
        .FileName = FileName
        .looping = (looping = LoopStyle.Enabled)
        
        Set .Buffer = DirectSound.CreateSoundBufferFromFile(SoundPath & FileName, bufferDesc, waveFormat)
    End With
    
    LoadWave = i
Exit Function

ErrHandler:
End Function

Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)
'Stops the sound in the given buffer index, or every sound if
'index is 0

    If bufferIndex > 0 And bufferIndex <= NumSoundBuffers Then
        If DSBuffers(bufferIndex).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).Buffer.Stop
        End If
    ElseIf bufferIndex = 0 Then
        Dim i As Long
        For i = 1 To NumSoundBuffers
            If Not DSBuffers(i).Buffer Is Nothing Then
                If DSBuffers(i).Buffer.GetStatus() > 0 And DSBuffers(i).Buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(i).Buffer.Stop
                End If
            End If
        Next i
    End If
End Sub

Public Property Get PlayingSound() As Boolean
'Returns true if there is at least one sound playing

    Dim i As Long
    
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
                PlayingSound = True
                Exit Property
            End If
        End If
    Next i
End Property

Public Property Let SoundVolume(ByVal Volume As Byte)
'Sets the sound's volume as a percentage

    Dim i As Long
    
    If Volume > 100 Then
        Exit Property
    End If
    
    'Take percentage to actual value
    SndVolume = Volume * (DSBVOLUME_MAX - DSBVOLUME_MIN) * 0.01 + DSBVOLUME_MIN
    
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                Call DSBuffers(i).Buffer.SetVolume(SndVolume)
            End If
        End If
    Next i
End Property

Public Property Get SoundVolume() As Byte
'Sets the music's volume as a percentage

    SoundVolume = CByte((SndVolume - DSBVOLUME_MIN) * 100 / (DSBVOLUME_MAX - DSBVOLUME_MIN))
End Property

Public Function PlayMIDI(Optional ByVal FileNumber As Byte = 0, Optional ByVal Loops As Long = -1) As Boolean
'Loads a midi and start to play it

On Error GoTo ErrHandler
    
    If Not MusicActivated Then
        Exit Function
    End If
    
    If Not Music Then
        Exit Function
    End If
    
    If PlayingMusic Then
        Call StopMidi
    End If
    
    If FileNumber < 1 Then
        Exit Function
    End If
        
    CurrentMusicFile = FileNumber

    LastPlayMP3 = False
    
    If FileNumber > 0 Then
        If Not LoadMidi(FileNumber) Then
            Exit Function
        End If
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then
            Exit Function
        End If
    End If
    
    Call MusicMP3Stop
    
    'Play it
    Call Segment.SetRepeats(Loops)
    
    Set SegState = Performance.PlaySegment(Segment, 0, 0)
    
    PlayMIDI = True
Exit Function

ErrHandler:
End Function

Private Function LoadMidi(ByVal FileNumber As Byte) As Boolean
'Loads a midi to memory

On Error GoTo ErrHandler
    If Not FileExist(MusicPath & FileNumber & ".mid", vbArchive) Then
        Exit Function
    End If
    
    Call StopMidi
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(MusicPath & FileNumber & ".mid")
    
    If Segment Is Nothing Then
        Exit Function
    End If
    
    Call Segment.SetStandardMidiFile
    
    LoadMidi = True
Exit Function

ErrHandler:
End Function

Public Sub StopMidi()
'Stops playing a midi
On Error GoTo ErrHandler
    If PlayingMusic Then
        Call Performance.Stop(Segment, SegState, 0, 0)
    End If
Exit Sub

ErrHandler:
End Sub

Public Property Let MusicTempo(ByVal Tempo As Single)
'Sets the music's tempo

    If Not Music Then
        Exit Property
    End If
    
    If Tempo < 0.25 Or Tempo > 2# Then
        Exit Property
    End If
    
    Call Performance.SetMasterTempo(Tempo)
End Property

Public Property Let MusicVolume(ByVal Volume As Long)
'Sets the music's volume as a percentage

    If Not Music Then
        Exit Property
    End If
    
    If Volume < 0 Or Volume > 100 Then
        Exit Property
    End If
    
    'Volume ranges from -10000 to 10000
    Call Performance.SetMasterVolume(Volume * 200 - 10000)
End Property

Public Property Get MusicTempo() As Single
'Returns the music's tempo

    If Not Music Then
        Exit Property
    End If
    
    MusicTempo = Performance.GetMasterTempo()
End Property

Public Property Get MusicVolume() As Long
'Returns the music's volume as a percentage
    
    If Not Music Then
        Exit Property
    End If
    
    MusicVolume = (Performance.GetMasterVolume() + 10000) / 200
End Property

Public Property Get PlayingMusic() As Boolean
'Returns True if a midi is being played
    
    If Not Music Or Segment Is Nothing Then
        Exit Property
    End If
    
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property

Public Property Get MusicLenght() As Long
'Returns the lenght of the current midi
    
    If Not Music Then
        Exit Property
    End If
    
    MusicLenght = Segment.GetLength()
End Property

Public Function MusicMP3Empty() As Boolean
'Destroy the MP3 objects.

On Error GoTo ErrorHandl
    
    Call MusicMP3Stop
    
    Set IBA = Nothing
    Set IME = Nothing
    Set IMPos = Nothing
    Set IMC = Nothing
    
    MusicMP3Empty = True
    Exit Function

ErrorHandl:
    MusicMP3Empty = False
    End
End Function

Public Function MusicMP3Load(ByVal FilePath As String, Optional ByVal Volume As Long = 0, Optional ByVal Balance As Long = 0) As Boolean 'Loads a MP3 *
    
    On Error GoTo ErrorHandl
    
    If Not FileExist(FilePath, vbArchive) Then
        Exit Function
    End If
    
    If Not MusicMP3Empty() Then
        GoTo ErrorHandl
    End If
    
    Set IMC = New FilgraphManager
    Call IMC.RenderFile(FilePath)
    
    Set IBA = IMC
    
    If Volume > 0 Then
        Volume = 0
    End If
    
    If Volume < -4000 Then
        Volume = -4000
    End If
    
    IBA.Volume = Volume
    
    If Balance > 5000 Then
        Balance = 5000
    End If
    
    If Balance < -5000 Then
        Balance = -5000
    End If
    
    IBA.Balance = Balance
    
    Set IME = IMC
    
    Set IMPos = IMC
    
    If ObjPtr(IMPos) Then
        IMPos.Rate = 1#
    End If
    
    IMPos.CurrentPosition = 0
    
    MusicMP3Load = True
    Exit Function

ErrorHandl:
    MusicMP3Load = False
End Function

Public Function MusicMP3Play(ByVal MusicFile As Byte) As Boolean

On Error GoTo ErrorHandl
    
    If Not MusicActivated Then
        Exit Function
    End If
    
    If Not Music Then
        Exit Function
    End If
    
    If Not MusicMP3Load(MusicPath & MusicFile & ".mp3") Then
        Exit Function
    End If
    
    Call IMC.Run
    
    MusicMP3Play = True
    Exit Function

ErrorHandl:
    MusicMP3Play = False
End Function

Public Sub MusicMP3GetLoop()
'Check if the song need looping.
    On Error GoTo ErrorHandl
    
    If IMPos.StopTime = IMPos.CurrentPosition Then
        MusicMP3Stop
        'MusicMP3Play
    End If
ErrorHandl:
End Sub

Public Function MusicMP3Stop() As Boolean

On Error GoTo ErrorHandl
    
    If ObjPtr(IMC) > 0 Then
        IMC.Stop
        IMPos.CurrentPosition = 0
    End If
    
    MusicMP3Stop = True
    Exit Function

ErrorHandl:
    MusicMP3Stop = False
End Function

Public Function MusicMP3Pause() As Boolean
'Pause a the current MP3 file.

    On Error GoTo ErrorHandl
    
    Call IMC.Pause
    
    MusicMP3Pause = True
    Exit Function

ErrorHandl:
    MusicMP3Pause = False
End Function

Public Function MusicMP3VolumeSet(ByVal Volume As Long) As Boolean
'Set the volume to the current MP3.

    On Error GoTo ErrorHandl
    
    If Volume > 0 Then
        Volume = 0
    End If
    
    If Volume < -4000 Then
        Volume = -4000
    End If
    
    IBA.Volume = Volume

    MusicMP3VolumeSet = True
    Exit Function

ErrorHandl:
    MusicMP3VolumeSet = False
End Function

