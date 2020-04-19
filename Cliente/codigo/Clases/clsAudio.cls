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
'**************************************************************
' clsAudio.cls - Controls all audio using DirectSound and DirectMusic
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@hotmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

'16/03/2004 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - First Release
'
'03/03/2007 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: Sound and Music activation / deactivation from within the class.
'   - Add: Documentation
'   - Change: Volume is no longer the dB amplification / atenuation to be
'           applied but a percentage for easier interaction with user.
'
'30/03/2008 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: 3D-like sound effects by altering frequency to emulate Doppler effect,
'           volume to emulate distance and panning for position.
'
'28/07/2008 - Franco Zeoli (Noich) - (zeolifranco@gmail.com)
'   -Add: Optional sound effects
'
'23/03/2010 - Torres Patricio(Pato) - (paatoo.torres@gmail.com)
'   -Add: ORE code to reproduce MP3

''
' Controls all audio using DirectSound and DirectMusic 7.
'
' @file     clsAudio.cls
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.2.1
' @date     20070303

Option Explicit

''
' Maximum number of sound buffers to hold in memory at any time.
Private Const NumSoundBuffersMax     As Byte = 30
Private NumSoundBuffers              As Byte

''
' The distance at which the sound is inaudible.
Private Const MAX_DISTANCE_TO_SOURCE As Integer = 150

''
' The frequency variation coefficient for Doppler effect simulation.
Private Const DELTA_FQ               As Integer = 75

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
    buffer As DirectSoundSecondaryBuffer8

End Type

''
' States how to set a sound's looping state.
'
' @param    Default     If wave is in memory use it's looping setting, otherwise false
' @param    Disabled    Never loop
' @param    Enabled     Loop until stopped
Public Enum LoopStyle

    Default = 0
    Disabled = 1
    Enabled = 2

End Enum

''
' Set to true once the object is created to prevent multiple initializations of the same object.
Dim inited                             As Boolean

''
'The one and only DirectSound object
Dim DirectSound                        As DirectSound8

'Array of all existing sound buffers
Dim DSBuffers(1 To NumSoundBuffersMax) As SoundBuffer

''
'Directory from which to load sound files by default
Dim SoundPath                          As String

''
'The volume for sounds (FXs)
Dim SndVolume                          As Long

''
'DirectMusic's Performance object
Dim Performance                        As DirectMusicPerformance8

''
'Currently loaded segment
Dim Segment                            As DirectMusicSegment8

''
'The one and only DirectMusic Loader
Dim Loader                             As DirectMusicLoader8

''
'State of the currently loaded segment
Dim SegState                           As DirectMusicSegmentState8

''
'States wether music is supported by hardware or not
Dim Music                              As Boolean

''
'Directory from which to load music files by default
Dim MusicPath                          As String

''
'Last loaded music file. Used to be able to load it when music gets activated
Dim CurrentMusicFile                   As String

''
'States wether music is currently activated or not
Dim mActivated                         As Boolean

''
'States wether sound is currently activated or not
Dim sActivated                         As Boolean

''
'States the last position where the listener was in the X axis
Dim lastPosX                           As Integer

''
'States the last position where the listener was in the Y axis
Dim lastPosY                           As Integer

''
'States wether sound effects are currently activated or not
Dim sEffectsActivated                  As Boolean

''
'Constructor. Performs no action at all.

Private Sub Class_Initialize()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    '
    '***************************************************
End Sub

''
'Destructor. Releases all created objects assuring no memory-leaks.

Private Sub Class_Terminate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/31/09
    'Destroy's all DirectMusic and DirectSound objects
    '12/31/09 - Pato: Use StopWave method to prevent errors.
    '***************************************************
    On Local Error Resume Next

    Dim LoopC As Long
   
    'Stop every channel being used and destroy the buffer
    For LoopC = 1 To NumSoundBuffers

        If LenB(DSBuffers(LoopC).FileName) <> 0 Then ' GSZAO
            Call DSBuffers(LoopC).buffer.Stop

        End If

        Set DSBuffers(LoopC).buffer = Nothing
    Next LoopC
   
    'Stop every channel being used and destroy the buffer
    Call StopWave
    
    'Stop music
    If Not Segment Is Nothing Then Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    
    Call Performance.CloseDown
    
    Set DirectSound = Nothing
    
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
    
    inited = False

End Sub

''
'Initializes the clsAudio object.
'
' @param    dX          DirectX7 object to be used when creating the DirectSound and Directmusic objects.
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
' @param    sound_path  Directory in wich to search for sound files.
' @param    music_path  Directory in wich to search for music files.
'
' @return   True if no errors occured, false otherwise.

Public Function Initialize(ByVal HWnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Creates all DS and DM objects
    '***************************************************
    On Error GoTo ErrHandler
    
    ' Prevent multiple creations
    If inited Then Exit Function
    
    'Validate DX object
    If DirectX Is Nothing Then GoTo ErrHandler
    
    'Initialize DirectSound
    Set DirectSound = DirectX.DirectSoundCreate(vbNullString)
    
    Call DirectSound.SetCooperativeLevel(HWnd, DSSCL_PRIORITY)
    
    'Set sound path
    SoundPath = sound_path
    
    'Set default volume
    SndVolume = 0
    
    'Initialize DirectMusic
    Music = InitializedirectMusic(HWnd)
    
    'Set music path
    MusicPath = music_path
    
    mActivated = True
    sActivated = True
        
    CurrentMusicFile = vbNullString
    
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
'
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
'
' @return   True if no errors occured, false otherwise.

Private Function InitializedirectMusic(ByVal HWnd As Long) As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Initialize DM. Returns false if music isn't supported by HW
    '***************************************************
    On Error GoTo ErrHandler

    Dim mus_Params As DMUS_AUDIOPARAMS
    Set Loader = DirectX.DirectMusicLoaderCreate()
    
    Set Performance = DirectX.DirectMusicPerformanceCreate()
    Call Performance.InitAudio(HWnd, DMUS_AUDIOF_ALL, mus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128)
    Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
    
    'Set to port default, 1 channel group is enough for MIDIs
    Performance.SetMasterTempo 1
    Performance.SetMasterVolume 1
    
    InitializedirectMusic = True
    Exit Function

ErrHandler:
    
End Function

''
' Plays a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    srcX        The X position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
' @param    srcY        The Y position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
' @param    LoopSound   The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Function PlayWave(ByVal FileName As String, _
    Optional ByVal SrcX As Byte = 0, _
    Optional ByVal SrcY As Byte = 0, _
    Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 28/07/08
    'Last Modified by: Franco Zeoli (Noich)
    'Added support for 3D sound effects.
    '***************************************************
    On Error GoTo ErrHandler
    
    If Not SoundActivated Then Exit Function
    
    Dim bufferIndex  As Long
    Dim Data()       As Byte
    Dim TemporalFile As String
    Dim handle       As Integer
    
    TemporalFile = SoundPath & FileName
    
    If Not Get_File_Data(DirRecursos, UCase$(FileName), Data, WAV_RESOURCE_FILE) Then Exit Function

    handle = FreeFile
    Open TemporalFile For Binary Access Write As handle
    Put handle, , Data
    Close handle
    
    'Get the buffer index were wave was loaded
    bufferIndex = LoadWave(FileName, LoopSound)

    If bufferIndex = 0 Then
        If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)

        Exit Function   'If an error ocurred abort

    End If

    With DSBuffers(bufferIndex)

        If .looping Then
            .buffer.Play DSBPLAY_LOOPING
        Else
            .buffer.Play DSBPLAY_DEFAULT

        End If
        
        'Apply volume
        Call .buffer.SetVolume(SndVolume)
        
        'Store position
        .X = SrcX
        .Y = SrcY
        
        'Get standard frequency
        .normalFq = .buffer.GetFrequency()

    End With
    
    PlayWave = bufferIndex
    
    
    'CRAW; 18/03/2020 --> ARREGLO SONIDO 3D
    If True Then
        If SrcX <> 0 And SrcY <> 0 Then
            Call Update3DSound(bufferIndex, 0, 0)

        End If

    End If

    If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)

    Exit Function

ErrHandler:

    If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)

End Function

Private Sub ResetSoundValues()

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 03/04/10
    'Reset the frequency, volume and pan.
    '***************************************************
    Dim i As Long
    
    For i = 1 To NumSoundBuffersMax

        If Not DSBuffers(i).buffer Is Nothing Then

            With DSBuffers(i)
                Call .buffer.SetFrequency(.normalFq)
                Call .buffer.SetVolume(SndVolume)
                Call .buffer.SetPan(0)

            End With

        End If

    Next i

End Sub

Private Sub Update3DSound(ByVal bufferIndex As Long, ByVal deltaX As Integer, ByVal deltaY As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/09/09
    'Last Modified by: Marco Vanotti (MarKoxX)
    'Fixed Dopplers formula
    'Refreshes the sound's arguments.
    'Fixed the sndVolume formula
    '***************************************************
    Dim linearDistanceOld As Single
    Dim linearDistanceNew As Single
    Dim deltaDistance     As Single
    Dim distanceXOld      As Long
    Dim distanceYOld      As Long
    Dim distanceXNew      As Long
    Dim distanceYNew      As Long
    Dim perceivedFq       As Long
    
    With DSBuffers(bufferIndex)
        distanceXOld = .X - lastPosX
        distanceYOld = .Y - lastPosY
        
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
        
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
        
        deltaDistance = linearDistanceNew - linearDistanceOld
        
        'Update frequency emulating Doppler effect
        perceivedFq = .normalFq * (1 - deltaDistance / DELTA_FQ)

        If perceivedFq < DSBFREQUENCY_MIN Then perceivedFq = DSBFREQUENCY_MIN
        
        Call .buffer.SetFrequency(perceivedFq)
        
        'Set volumen amortiguation according to distance
        If linearDistanceNew > MAX_DISTANCE_TO_SOURCE Then
            Call .buffer.SetVolume(SndVolume + (DSBVOLUME_MIN - SndVolume))
        Else
            Call .buffer.SetVolume(SndVolume + (linearDistanceNew / MAX_DISTANCE_TO_SOURCE) * (DSBVOLUME_MIN - SndVolume))

        End If
        
        'Prevent division by zero
        If linearDistanceNew = 0 Then linearDistanceNew = 1
        
        'Set panning according to relative position of the source to the listener
        Call .buffer.SetPan((distanceXNew / linearDistanceNew) * DSBPAN_RIGHT)

    End With

End Sub

''
' Updates 3D sounds based on the movement of the listener.
'
' @param    x   The new coordinate of the listener in the x axis (consider 0 as the furthest point to the left).
' @param    y   The new coordinate of the listener in the y axis (consider 0 as the furthest point to the front).
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Sub MoveListener(ByVal X As Integer, ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 28/07/08
    'Last Modification by: Franco Zeoli (Noich)
    'Moves the position of the listener updating all 3D sounds.
    '***************************************************
    Dim i      As Long
    Dim deltaX As Integer
    Dim deltaY As Integer
    
    If Not True Then Exit Sub
    
    deltaX = X - lastPosX
    deltaY = Y - lastPosY
    
    For i = 1 To NumSoundBuffersMax

        If Not DSBuffers(i).buffer Is Nothing Then
            If DSBuffers(i).buffer.GetStatus And DSBSTATUS_PLAYING Then
                If DSBuffers(i).X <> 0 And DSBuffers(i).Y <> 0 Then
                    Call Update3DSound(i, deltaX, deltaY)

                End If

            End If

        End If

    Next i
    
    lastPosX = X
    lastPosY = Y

End Sub

''
' Loads a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    looping     The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Private Function LoadWave(ByVal FileName As String, ByVal looping As LoopStyle) As Long

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 30/10/2012 - ^[GS]^
    'Loads a new wave into memory (if it's not already loaded) and returns the buffer index
    '***************************************************
    On Error GoTo ErrHandler

    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim i          As Long
    
    If Not FileExist(SoundPath & FileName, vbArchive) Then Exit Function
    
    'Check if the buffer is in memory and not playing
    For i = 1 To NumSoundBuffersMax

        If DSBuffers(i).FileName = FileName Then
            If DSBuffers(i).buffer.GetStatus = 0 Or DSBuffers(i).buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                'Found it!!! We just play this one :)
                LoadWave = i
                
                'Set looping if needed
                If looping <> LoopStyle.Default Then DSBuffers(i).looping = (looping = LoopStyle.Enabled)
                
                'Reset default sound frecuency to avoid issues.
                Call DSBuffers(i).buffer.SetFrequency(DSBuffers(i).normalFq)
                
                Exit Function

            End If

        End If

    Next i
    
    'Not in memory, search for an empty buffer
    For i = 1 To NumSoundBuffersMax

        If DSBuffers(i).buffer Is Nothing Then
            NumSoundBuffers = CByte(i)
            Exit For

        End If

    Next i
    
    'No empty buffers?? Search for one not being played and use that slot
    If i > NumSoundBuffersMax Then

        For i = 1 To NumSoundBuffersMax

            If DSBuffers(i).buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or DSBuffers(i).buffer.GetStatus() = 0 Then
                Exit For

            End If

        Next i

    End If
    
    'Still no free buffers?? Get the first non-looping sound
    If i > NumSoundBuffersMax Then

        For i = 1 To NumSoundBuffersMax

            If Not (DSBuffers(i).buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Call DSBuffers(i).buffer.Stop
                Exit For

            End If

        Next i

    End If
    
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
    If i > NumSoundBuffersMax Then
        If Not looping Then Exit Function
        i = 1

    End If
    
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 44100 '22050 (44k GSZAO TEST)
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
    With DSBuffers(i)
        Set .buffer = Nothing   'Get rid of any previous data
        
        .FileName = FileName
        .looping = (looping = LoopStyle.Enabled)
        
        Set .buffer = DirectSound.CreateSoundBufferFromFile(SoundPath & FileName, bufferDesc)

    End With
    
    LoadWave = i
    Exit Function

ErrHandler:

End Function

''
' Stops a given sound or all of them.
'
' @param    bufferIndex     The index of the sound file to stop playing. If 0, all sounds are stopped.

Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Stops the sound in the given buffer index, or every sound if
    'index is 0
    '***************************************************
    If bufferIndex > 0 And bufferIndex <= NumSoundBuffersMax Then
        If DSBuffers(bufferIndex).buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).buffer.Stop

        End If

    ElseIf bufferIndex = 0 Then
        Dim i As Long

        For i = 1 To NumSoundBuffersMax

            If Not DSBuffers(i).buffer Is Nothing Then
                If DSBuffers(i).buffer.GetStatus() <> 0 And DSBuffers(i).buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(i).buffer.Stop

                End If

            End If

        Next i

    End If

End Sub

''
' Retrieves wether there are sounds currentyl playing or not.
'
' @return   true if there is at least one sound playing, false otherwise.

Private Property Get PlayingSound() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns true if there is at least one sound playing
    '***************************************************
    Dim i As Long
    
    For i = 1 To NumSoundBuffersMax

        If Not DSBuffers(i).buffer Is Nothing Then
            If DSBuffers(i).buffer.GetStatus() And DSBSTATUS_PLAYING Then
                PlayingSound = True
                Exit Property

            End If

        End If

    Next i

End Property

''
' Sets the volume of sound.
'
' @param    Volume  New value to be set, in percentage (0 to 100).

Public Property Let SoundVolume(ByVal Volume As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the sound's volume as a percentage
    '***************************************************
    Dim i As Long
    
    If Volume > 100 Then Exit Property
    
    'Take percentage to actual value
    SndVolume = Volume * (DSBVOLUME_MAX - DSBVOLUME_MIN) / 100 + DSBVOLUME_MIN
    
    For i = 1 To NumSoundBuffersMax

        If Not DSBuffers(i).buffer Is Nothing Then
            If DSBuffers(i).buffer.GetStatus And DSBSTATUS_PLAYING Then
                Call DSBuffers(i).buffer.SetVolume(SndVolume)

            End If

        End If

    Next i

End Property

''
' Gets the volume of sound.
'
' @return   The current sound volume in percentage.

Public Property Get SoundVolume() As Byte

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the music's volume as a percentage
    '***************************************************
    SoundVolume = CByte((SndVolume - DSBVOLUME_MIN) * 100 / (DSBVOLUME_MAX - DSBVOLUME_MIN))

End Property

''
' Plays a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    Loops   Number of times to loop sound. For eternal loops set it to -1.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Public Function PlayMIDI(Optional ByVal File As String = vbNullString, Optional ByVal Loops As Long = -1) As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 22/08/2013 - ^[GS]^
    'Loads a midi and start to play it
    '***************************************************
    On Error GoTo ErrHandler

    If CurrentMusicFile = File Then Exit Function ' No cargar nuevamnete el tema que ya estaba

    If LenB(File) > 0 Then CurrentMusicFile = File
        
    If Not MusicActivated Then Exit Function
    
    If Not Music Then Exit Function
    
    If PlayingMusic Then StopMidi
    
    If LenB(File) > 0 Then
        Dim Data()       As Byte
        Dim TemporalFile As String
        Dim handle       As Integer
        
        If Not Get_File_Data(DirRecursos, UCase$(File), Data, MIDI_RESOURCE_FILE) Then Exit Function

        TemporalFile = MusicPath & PlayMIDI
        handle = FreeFile
        Open TemporalFile For Binary Access Write As handle
        Put handle, , Data
        Close handle
    
        If Not LoadMidi(PlayMIDI) Then
            If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)
            Exit Function

        End If

        If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)

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

    If FileExist(TemporalFile, vbArchive) Then Call Kill(TemporalFile)

End Function

''
' Loads a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Private Function LoadMidi(ByVal File As String) As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Loads a midi to memory
    '***************************************************
    On Error GoTo ErrHandler

    If Not FileExist(MusicPath & File, vbArchive) Then Exit Function
    
    Call StopMidi
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(MusicPath & File)
    
    If Segment Is Nothing Then Exit Function
    
    Call Segment.SetStandardMidiFile
    
    LoadMidi = True
    Exit Function

ErrHandler:

End Function

''
' Stops playing the currently loaded MIDI file.

Public Sub StopMidi()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Stops playing a midi
    '***************************************************
    On Error GoTo ErrHandler

    If PlayingMusic Then Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)

    Exit Sub

ErrHandler:

End Sub

''
' Sets the music's tempo.
'
' @param    Tempo   New Tempo value. Must be in the range of 0.25 to 2.0

Public Property Let MusicTempo(ByVal Tempo As Single)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the music's tempo
    '***************************************************
    If Not Music Then Exit Property
    
    If Tempo < 0.25 Or Tempo > 2# Then Exit Property
    
    Call Performance.SetMasterTempo(Tempo)

End Property

''
' Sets the music's volume.
'
' @param    Volume  Percentage to wich to set the music volume.

Public Property Let MusicVolume(ByVal Volume As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the music's volume as a percentage
    '***************************************************
    If Not Music Then Exit Property
    
    If Volume < 0 Or Volume > 100 Then Exit Property
    
    ' Volume ranges from -10000 to 10000
    Call Performance.SetMasterVolume(Volume * 200 - 10000)

End Property

''
' Gets the music's tempo.
'
' @return   Current music'tempo.

Private Property Get MusicTempo() As Single

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns the music's tempo
    '***************************************************
    If Not Music Then Exit Property
    
    MusicTempo = Performance.GetMasterTempo()

End Property

''
' Gets the music's volume.
'
' @return   Current music's volume as a percentage.

Public Property Get MusicVolume() As Long

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns the music's volume as a percentage
    '***************************************************
    If Not Music Then Exit Property
    
    MusicVolume = (Performance.GetMasterVolume() + 10000) / 200

End Property

''
' Checks wether there is music currently playing.
'
' @return   True if there is a MIDI playing, false otherwise.

Public Property Get PlayingMusic() As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns True if a midi is being played
    '***************************************************
    If Not Music Or Segment Is Nothing Then Exit Property
    
    PlayingMusic = Performance.IsPlaying(Segment, SegState)

End Property

''
' Retrieves wether music is activated or not.
'
' @return   True if music is activated, false otherwise.

Public Property Get MusicActivated() As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/03/07
    'Retrieves wether the music is activated or not
    '***************************************************
    MusicActivated = mActivated

End Property

''
' Sets wether music is activated or not.
'
' @param    nValue  True if music should be activated, false otherwise.

Public Property Let MusicActivated(ByVal nValue As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/11/2011 - ^[GS]^
    'Sets wether the music is activated or not
    '***************************************************
    If mActivated = nValue Then Exit Property
    
    mActivated = nValue
    
    If nValue Then
        Call PlayMIDI(CurrentMusicFile)
    Else
        Call StopMidi

    End If

End Property

''
' Retrieves wether sound is activated or not.
'
' @return   True if sound is activated, false otherwise.

Public Property Get SoundActivated() As Boolean

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/03/07
    'Retrieves wether the sound is activated or not
    '***************************************************
    SoundActivated = sActivated

End Property

''
' Sets wether sound is activated or not.
'
' @param    nValue  True if sound should be activated, false otherwise.

Public Property Let SoundActivated(ByVal nValue As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/03/07
    'Sets wether the sound is activated or not
    '***************************************************
    If sActivated = nValue Then Exit Property
    
    sActivated = nValue
    
    If Not nValue Then
        Call StopWave

    End If

End Property

''
' Sets wether sound effects are activated or not.
'
' @param nValue True if sound effects should be activated, false otherwise.

Public Property Let SoundEffectsActivated(ByVal nValue As Boolean)

    '***************************************************
    'Author: Franco Zeoli (Noich)
    'Last Modification: 03/04/10
    'Sets wether the sound effects are activated or not
    '03/04/10 - Pato: If sound effects are disable, then reset the frequency, volume and pan.
    '***************************************************
    sEffectsActivated = nValue
    
    If Not nValue Then
        Call ResetSoundValues

    End If

End Property

''
' Retrieves wether sound effects are activated or not.
'
' @return   True if sound effects are activated, false otherwise.

Public Property Get SoundEffectsActivated() As Boolean

    '***************************************************
    'Author: Franco Zeoli (Noich)
    'Last Modification: 28/07/08
    'Retrieves wether the sound effects are activated or not
    '***************************************************
    SoundEffectsActivated = sEffectsActivated

End Property


