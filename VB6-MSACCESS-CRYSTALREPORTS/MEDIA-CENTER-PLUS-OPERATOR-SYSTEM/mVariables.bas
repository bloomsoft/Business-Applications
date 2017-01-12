Attribute VB_Name = "mVariables"
Public iRandomCount As Integer         'If random is selected, this holds the value of how many Files have been played.
                                       'Once this value reaches the same as frmMain.lstFiles.ListItems.Count, we know that all
                                       'Files have been played.
                                        
Private m_sComputerName As String      'Holds ComputerName.
Private m_byCommandLight As Byte       'Holds the index number of the light which is switched on. (frmMain.speCommandLight(0-5).
Private m_lTrackLength As Long         'Hold the length of the current track.
Private m_bTrackIsPlaying As Boolean   'Holds the value of TRUE if a track is playing.
Private m_bAudioAllOff As Boolean      'Holds TRUE when the Mute button has been depressed.
Private m_bRandomSet As Boolean        'Holds TRUE if random selection is set.
Private m_iVolumeSetting As Integer    'Holds the current value of the volume.


'm_sComputerName.
Property Get sComputerName() As String
    sComputerName = m_sComputerName
End Property
Property Let sComputerName(newValue As String)
    m_sComputerName = newValue
End Property

'm_byCommandLight.
Property Get byCommandLight() As Byte
    byCommandLight = m_byCommandLight
End Property
Property Let byCommandLight(newValue As Byte)
    m_byCommandLight = newValue
End Property

'm_lTrackLength.
Property Get lTrackLength() As Long
    lTrackLength = m_lTrackLength
End Property
Property Let lTrackLength(newValue As Long)
    m_lTrackLength = newValue
End Property

'm_TrackIsPlaying.
Property Get bTrackIsPlaying() As Boolean
    bTrackIsPlaying = m_bTrackIsPlaying
End Property
Property Let bTrackIsPlaying(newValue As Boolean)
    m_bTrackIsPlaying = newValue
End Property

'm_bAudioAllOff.
Property Get bAudioAllOff() As Boolean
    bAudioAllOff = m_bAudioAllOff
End Property
Property Let bAudioAllOff(newValue As Boolean)
    m_bAudioAllOff = newValue
End Property

'm_bRandomSet.
Property Get bRandomSet() As Boolean
    bRandomSet = m_bRandomSet
End Property
Property Let bRandomSet(newValue As Boolean)
    m_bRandomSet = newValue
End Property

'm_iVolumeSetting.
Property Get iVolumeSetting() As Integer
    iVolumeSetting = m_iVolumeSetting
End Property
Property Let iVolumeSetting(newValue As Integer)
    m_iVolumeSetting = newValue
End Property
