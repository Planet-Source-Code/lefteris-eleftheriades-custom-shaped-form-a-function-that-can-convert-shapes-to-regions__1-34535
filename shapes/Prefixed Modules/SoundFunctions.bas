Attribute VB_Name = "SoundFunctions"
' A MODULE WITH THE BEST DECLAIRS TO PLAY SOUNDS
'        WITH OUT LAGGING THE PROGRAM.
''''''''''''''''GET DOS FILE PATH'''''''''''''
'MCI send string requires dos names for paths:
'From  C:\Program files
'To    C:\Progra~1\
'Inpput's The Filepath\Name
'Returns:
'lpszShortPath = The short file path & some vbNull values
'GetShortPathName = ShortPath's Length
Public Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Function Folows at the end
'''''''''''''''''''''''''''''''''''''''''''''''''
'''            Sound    Functions             '''
'''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''Play Sound Part'''''''''''''''''
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
     ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'See PlaySoundPart Sub
'''''''''''''''''''''''''''''''''''''''''''''''''
                   
''''''''''''''' Play Sound File '''''''''''''''''
Public Declare Function PlaySoundFile Lib "winmm.dll" _
       Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
       ByVal uFlags As Long) As Long
       
Public Const SND_ALIAS = &H10000
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ALIAS_START = 0
Public Const SND_ASYNC = &H1
Public Const SND_APPLICATION = &H80
Public Const SND_FILENAME = &H20000
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000
Public Const SND_PURGE = &H40
Public Const SND_RESERVED = &HFF000000
Public Const SND_RESOURCE = &H40004
Public Const SND_SYNC = &H0
Public Const SND_TYPE_MASK = &H170007
Public Const SND_VALID = &H1F
Public Const SND_VALIDFLAGS = &H17201F
       
'Enum of mcisendstring audio type
Public Enum AudioTypes
 AVI_Video = 0 '"AVIVideo"
 CD_Sound = 1 '"CDAudio"
 DigitalCasette = 2 '"DAT"
 DigitalVideo = 3 '"DigitalVideo"
 MultimediaVideo = 4 '"MMMovie"
 VedioOverlay = 5 '"Overlay"
 Scanner = 6 '"Scanner"
 MIDI_Sequence = 7 '"Sequencer"
 VideoCasette = 8 '"VCR"
 VideoDisk = 9 '"VideoDisk"
 WaveFiles = 10 '"WaveAudio"
 Custom = 11 '"Other"
End Enum
'I got this form an other control
'Microsoft multimedia control
'MMControl.DeviceType (Press F1 on the device type)
''''''''''''''''Dos File path function'''''''''''''''

Public Function GetDosFileName(ByVal FileName As String) As String
    Dim rc As Long
    Dim ShortPath As String
    Const PATH_LEN& = 164
    ShortPath = String$(PATH_LEN + 1, 0)
    rc = GetShortPathName(FileName, ShortPath, PATH_LEN)
    GetDosFileName = Left$(ShortPath, rc)
    'From  C:\Program files
    'To    C:\Progra~1\
End Function
''''''''''''''''''''Decode the Enum''''''''''''''''''
Public Function AudioTypesDecoader$(Selection&)
    'Converts the AudioTypes-Enum to the string necessary
    'For Mci to work
     Select Case Selection&
       Case 0: AudioTypesDecoader$ = "AVIVideo"
       Case 1: AudioTypesDecoader$ = "CDAudio"
       Case 2: AudioTypesDecoader$ = "DAT"
       Case 3: AudioTypesDecoader$ = "DigitalVideo"
       Case 4: AudioTypesDecoader$ = "MMMovie"
       Case 5: AudioTypesDecoader$ = "Overlay"
       Case 6: AudioTypesDecoader$ = "Scanner"
       Case 7: AudioTypesDecoader$ = "Sequencer"
       Case 8: AudioTypesDecoader$ = "VCR"
       Case 9: AudioTypesDecoader$ = "VideoDisk"
      Case 10: AudioTypesDecoader$ = "WaveAudio"
      Case 11: AudioTypesDecoader$ = "Other"
     End Select
End Function
''''''''''''''''''''Sound Functions''''''''''''''''''
'Public Function PlaySoundPart(ByVal File_Name As String, Start_&, Stop_&, Audio_Type As AudioTypes, Optional Alias_ As String = "SndPrt")
 Public Sub PlaySoundPart(ByVal File_Name As String, Start_&, Stop_&, Audio_Type As AudioTypes, Optional Alias_ As String = "SndPrt")
    Dim RetStr As String, CallBack As Long, ShortName
    Dim AT As String
    AT = AudioTypesDecoader$(Audio_Type)
    RetStr = Space$(128)
    ShortName = GetDosFileName(File_Name)
    PlayMidiPart = mciSendString("open " & AT & "!" & ShortName & " alias " & Alias_, RetStr, 128, CallBack)
    PlayMidiPart = mciSendString("play " & Alias_ & " from " & Start_& & " to " & Stop_&, RetStr, 128, 1)
End Sub
'End Function
'Inputs:
'Alias_ = a unique ID for the sound so you can order it to stop etc.
'Outputs:
'Plays a part of the sound through the speakers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function PlayLargeSound(ByVal StrName As String, Audio_Type As AudioTypes, Optional Alias_ As String = "Sndfle") As Long
Public Sub PlayLargeSound(ByVal StrName As String, Audio_Type As AudioTypes, Optional Alias_ As String = "Sndfle")
    Dim RetStr As String, CallBack As Long, ShortName
    Dim AT As String
    AT = AudioTypesDecoader$(Audio_Type)
    RetStr = Space$(128)
    ShortName = GetDosFileName(StrName)
    PlayMidi = mciSendString("open " & AT & "!" & ShortName & " alias " & Alias_, RetStr, 128, CallBack)
    PlayMidi = mciSendString("play " & Alias_, RetStr, 128, 1)
End Sub
'End Function
'Plays a sound until you tell it to stop.
'If you don't stop it your self it will keep
'your audio device busy and you won't be able to use it
'again, till you reboot.
'To stop it, use StopLargeSound.
'A great way to stop it when the song is played
'is to check every 1000 miliseconds
'With a timer the PersendagePlayed
'Using the function PersendagePlayedOfAnOpenedSound
'If it is 100% then StopLargeSound

'Example
'Private Sub Form_Load()
' PlayLargeSound "C:\WINDOWS\MEDIA\Ding.wav", WaveFiles, "BellSnd"
'End Sub
'Private Sub CheckIfSongEnded_Timer()
'  If PersendagePlayedOfAnOpenedSound("BellSnd") = 100 then '100%
'     StopLargeSound "BellSnd"
'End if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function StopLargeSound(Optional Alias_ As String = "Sndfle") As Long
 Public Sub StopLargeSound(Optional Alias_ As String = "Sndfle")
    Dim RetStr As String, CallBack As Long
    RetStr = Space$(128)
    StopMidi = mciSendString("stop " & Alias_, RetStr, 128, CallBack)
    StopMidi = mciSendString("close " & Alias_, RetStr, 128, CallBack)
End Sub
'End Function
'Stops a sound
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PersendagePlayedOfAnOpenedSound(Alias_ As String) As Double
 On Error GoTo Erro
 PersendagePlayedOfAnOpenedSound = (OpenedSoundPosition(Alias_) / OpenedSoundLenith(Alias_) * 100)
 Exit Function
Erro:
PersendagePlayedOfAnOpenedSound = 100
End Function
'Returns how much of the sound was played
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenedSoundLenith(Alias_ As String) As Long
 Dim tmp As Variant
 Dim RetStr As String, CallBack As Long
 RetStr = Space$(128)
 tmp = mciSendString("status " & Alias_ & " length", RetStr, 128, CallBack)
 OpenedSoundLenith = Val(RetStr)
End Function
'Returns the sound(alias)'s length in milisecs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenedSoundPosition(Alias_ As String) As Long
 Dim tmp As Variant
 Dim RetStr As String, CallBack As Long
 RetStr = Space$(128)
 tmp = mciSendString("status " & Alias_ & " position", RetStr, 128, CallBack)
 OpenedSoundPosition = Val(RetStr)
End Function

Public Function PlaySound(sFileName As String, Optional ModalSound As Boolean = False)
  'Modal = pauses the program until the sound is stopped
  If ModalSound Then
     PlaySoundFile sFileName, SND_NODEFAULT
  Else
     PlaySoundFile sFileName, SND_ASYNC + SND_NODEFAULT
  End If
End Function

'Returns the number of sound(alias)'s miliseconds played
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'An old function I had which pauses the game from
'working until the sound was done
'NEVER USE THIS PLEASE I HAVE PUT IT THERE
'FOR EDUCATIONAL PERPOSES.
Public Sub PlayLaggedSoundPart(ByVal File_Name As String, Start As Long, _
       Stop_ As Long, Audio_Type As AudioTypes, Key As String)
     Dim errorCode As Integer, returnStr As Integer
     Dim cmd As String * 255
     Dim AT As String
     File_Name = GetDosFileName(File_Name)
     AT = AudioTypesDecoader$(Audio_Type)
     cmd = "open " & """" & File_Name & """" & " type " & AT & " alias " & Key
     errorCode = mciSendString(cmd, returnStr, 255, 0)
     errorCode = mciSendString("play " & Key & " from " & Start & " to " & Stop_ & " wait", returnStr, 255, 0)
     errorCode = mciSendString("Close " & Key, returnStr, 255, 0)
End Sub

