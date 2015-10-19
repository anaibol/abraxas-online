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
'Modulo de clase by •Parra

Option Explicit
 
''
' Maximum number of sound buffers to hold in memory at any time.
Private Const NumSoundBuffers As Byte = 30

Public Musica As Boolean

''
' Custom sound buffer structure.
'
' @param    FileName    The name of the file loaded in this buffer
' @param    looping     Set to true if the sound will loop until stopped
' @param    X           The X position of the sound in map coordinates
' @param    Y           The Y position of the sound in map coordinates
' @param    normalFq    The frequency of the sound when not affected by Doppler effect.
' @param    Buffer      The DirectSoundBuffer object to wich the file is loaded
Private Type SoundBuffer
    FileName As String
    looping As Boolean
    X As Byte
    Y As Byte
    normalFq As Long
    Buffer As DirectSoundSecondaryBuffer8
End Type
''

''
' Set to true once the object is created to prevent multiple initializations of the same object.
Dim inited As Boolean
 

''
'States wether music is supported by hardware or not
Dim Music As Boolean

'**** Direct Sound ********
 
Private dX As DirectX8
Private DS As DirectSound8
Private DS_Enum As DirectSoundEnum8
 
'**** Direct Music ********
''
'DirectMusic's Performance object
Private Performance As DirectMusicPerformance8
 
''
'Currently loaded segment
Private Segment As DirectMusicSegment8
 
''
'The one and only DirectMusic Loader
Private Loader As DirectMusicLoader8
 
''
'State of the currently loaded segment
Private SegState As DirectMusicSegmentState8
 
''
'Array of all existing sound buffers
Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer
 
''
''
'Destructor. Releases all created objects assuring no memory-leaks.
 
Private Sub Class_Terminate()
On Local Error Resume Next
    Dim loopc As Long
   
    'Stop every channel being used and destroy the buffer
    For loopc = 1 To NumSoundBuffers
        Call DSBuffers(loopc).Buffer.Stop
        Set DSBuffers(loopc).Buffer = Nothing
    Next loopc
   
    'Stop music
    Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
   
    Call Performance.CloseDown
   
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
   
    inited = False
End Sub
 
''
'Initializes the clsAudio object.
Public Function Initialize(ByVal hwnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean
On Error GoTo ErrHandler
   
    ' Prevent multiple creations
    If inited Then Exit Function
   
     'Cargamos
    Set dX = New DirectX8
   
    'Validate DX object
    If dX Is Nothing Then GoTo ErrHandler
   
    Set DS_Enum = dX.GetDSEnum
    Set DS = dX.DirectSoundCreate(DS_Enum.GetGuid(1))
 
    DS.SetCooperativeLevel hwnd, DSSCL_PRIORITY
   
    'Initialize DirectMusic
    Music = InitializedirectMusic(dX, hwnd)
   
    'Set tempo to 0 and volume of music
    MusicTempo = 0
    Call Performance.SetMasterVolume(200)
   
    inited = True
   
    Initialize = True
Exit Function
 
ErrHandler:
End Function
 
''
' Creates and configures all DirectMusic objects.
 
Private Function InitializedirectMusic(ByRef dX As DirectX8, ByVal hwnd As Long) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Initialize DM. Returns false if music isn't supported by HW
'***************************************************
On Error GoTo ErrHandler
    Dim mus_Params As DMUS_AUDIOPARAMS
    Set Loader = dX.DirectMusicLoaderCreate()
   
    Set Performance = dX.DirectMusicPerformanceCreate()
    Call Performance.InitAudio(hwnd, DMUS_AUDIOF_ALL, mus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128)
    Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
   
    Performance.SetMasterTempo 1
    Performance.SetMasterVolume 1
   
    InitializedirectMusic = True
Exit Function
 
ErrHandler:
End Function
 
''
' Plays a new WAVE file.
 
 
 
''
' Updates 3D sounds based on the movement of the listener.
 
 
''
' Loads a new WAVE file.
 

''
' Retrieves wether there are sounds currentyl playing or not.
 
 
''
' Sets the volume of sound.
 
 
''
' Gets the volume of sound.
 
 
''
' Plays a new MIDI file.
 
Public Function PlayMIDI(Optional ByVal file As String = "", Optional ByVal Loops As Long = -1) As Boolean
On Error GoTo ErrHandler
    If Not Musica Then Exit Function
   
    If PlayingMusic Then StopMidi
   
    If LenB(file) > 0 Then
        If Not LoadMidi(file) Then Exit Function
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Function
    End If
   
    'Play it
    Call Segment.SetRepeats(Loops)
   
    Set SegState = Performance.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0)
   
    PlayMIDI = True
Exit Function
 
ErrHandler:
End Function
 
''
' Loads a new MIDI file.
 
Private Function LoadMidi(ByVal file As String) As Boolean
On Error GoTo ErrHandler

    If Not FileExist(file, vbArchive) Then Exit Function
   
    Call StopMidi
   
    'Destroy old object
    Set Segment = Nothing
   
    Set Segment = Loader.LoadSegment(file)
   
    If Segment Is Nothing Then
        Exit Function
    End If
   
    Call Segment.SetStandardMidiFile
   
    LoadMidi = True
    
Exit Function
 
ErrHandler:

End Function
 
''
' Stops playing the currently loaded MIDI file.
 
Public Sub StopMidi()
On Error GoTo ErrHandler
    If PlayingMusic Then
        Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    End If
Exit Sub
 
ErrHandler:
End Sub
 
''
' Sets the music's tempo.
 
Public Property Let MusicTempo(ByVal Tempo As Single)
    If Not Music Then Exit Property
   
    If Tempo < 0.25 Or Tempo > 2# Then Exit Property
   
    Call Performance.SetMasterTempo(Tempo)
End Property
 
''
' Sets the music's volume.
 
 
''
' Gets the music's tempo.
 
 
''
' Gets the music's volume.
  
''
' Checks wether there is music currently playing.
 
Public Property Get PlayingMusic() As Boolean
    If Not Music Or Segment Is Nothing Then Exit Property
   
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property
 
''
' Retrieves the music's length.
 
 
''
' Retrieves wether music is activated or not.
 
 
''
' Sets wether music is activated or not.
 
 
''
' Retrieves wether sound is activated or not.
 
 
''
' Sets wether sound is activated or not.
 

