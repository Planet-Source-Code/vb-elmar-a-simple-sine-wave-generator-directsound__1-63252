Attribute VB_Name = "directSnd"
Public samPles As Long, myByte(360) As Byte
Public DX7 As New DirectX7, DS As DirectSound, DSB(1) As DirectSoundBuffer
Public DESC As DSBUFFERDESC, PCM As WAVEFORMATEX
Sub Init_DX7(Hwnd As Long)
Set DS = DX7.DirectSoundCreate("") 'Create the DirectSound Object

'DSSCL_NORMAL other applications can use the sound card
'DSSCL_EXCLUSIVE only i can use the sound card
DS.SetCooperativeLevel Hwnd, DSSCL_NORMAL 'Set the Cooperative Level

'Fill WaveFormat Structure
PCM.nFormatTag = WAVE_FORMAT_PCM
PCM.nChannels = 1
PCM.lSamplesPerSec = 11025
PCM.nBitsPerSample = 8
PCM.nBlockAlign = 1
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
PCM.nSize = 0

'DSBCAPS_CTRLPAN enables Pan, DSBCAPS_CTRLVOLUME enables volume control
DESC.lFlags = DSBCAPS_STATIC Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN
'set Bytes
DESC.lBufferBytes = samPles

'********************************************************
'Create Buffers   (buffer "0" and buffer "1")
Set DSB(0) = DS.CreateSoundBuffer(DESC, PCM)
Set DSB(1) = DS.CreateSoundBuffer(DESC, PCM)
        
DSB(0).Play DSBPLAY_LOOPING 'Initialize to play loop  DSBPLAY_LOOPING
DSB(1).Play DSBPLAY_LOOPING '( DSBPLAY_DEFAULT would play the sound only one time )
End Sub
Sub Term_DX7() 'Clear the created DX7 Objects.
    Set DSB(0) = Nothing: Set DSB(1) = Nothing: Set DS = Nothing: Set DX7 = Nothing
End Sub
Sub DSBWRITE(Num As Integer, ByRef Buffer() As Byte)
    'Writing an array of bytes to a given DirectSoundBuffer.
    DSB(Num).WriteBuffer 0, 0, Buffer(0), DSBLOCK_ENTIREBUFFER
End Sub
