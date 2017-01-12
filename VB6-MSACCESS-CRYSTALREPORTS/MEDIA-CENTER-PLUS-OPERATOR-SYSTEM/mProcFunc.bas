Attribute VB_Name = "mProcFunc"
'***********************************************************************************
'Strips the NullCharacter from sInput.
'***********************************************************************************
Public Function ftnStripNullChar(sInput As String) As String

    Dim x As Integer
    
    x = InStr(1, sInput$, Chr$(0))

    If x > 0 Then
        ftnStripNullChar = Left(sInput$, x - 1)
    End If

End Function


'***********************************************************************************
'Return Correct Drive Path when Clicking on Node in Explorer Tree eg. "My Computer\C:\WinNt"
'would return "C:\WinNt".
'***********************************************************************************
Public Function ftnReturnNodePath(sExplorerPath As String) As String

    Dim iSearch(1) As Integer
    Dim sRootPath As String
    
    iSearch%(0) = InStr(1, sExplorerPath$, "(", vbTextCompare)
    iSearch%(1) = InStr(1, sExplorerPath$, ")", vbTextCompare)
    
    If iSearch%(0) > 0 Then
        sRootPath$ = Mid(sExplorerPath$, iSearch%(0) + 1, 2)
    End If
    
    If iSearch%(1) > 0 Then
        ftnReturnNodePath$ = sRootPath$ & Mid(sExplorerPath$, iSearch%(1) + 1, Len(sExplorerPath$)) & "\"
    End If
    
End Function

'***********************************************************************************
'Set all frmMain.speCommandLight(x) to vbBlack.
'***********************************************************************************
Public Sub subSetLightColour(sLightSource As String)

    Dim x As Integer

    Select Case sLightSource$

        Case "speCommandLight"
            For x = 0 To 5
                frmMain.speCommandLight(x).FillColor = vbBlack
            Next
            
    End Select
End Sub

'***********************************************************************************
'Displays an MCI message response to the command passed to it.
'***********************************************************************************
Public Sub subSendMCIMessage(sCommand As String)
    
    Dim lERReturn(1) As Long
    Dim sError As String * 256
    
    'Used to return status information of the mp3 being played.---------------------
'    Dim sStatus As String * 256
'    lERReturn(0) = mciSendString(sCommand$, sStatus$, 256, 0)
    '-------------------------------------------------------------------------------
    
    lERReturn&(1) = mciSendString(sCommand$, 0, 0, hwnd)
    
    mciGetErrorString lERReturn&(1), sError$, Len(sError$)
    frmMain.Caption = sError$
    
End Sub

'***********************************************************************************
'Returns an index number of a file that has not already played.
'***********************************************************************************
Public Function ftnRandomSelect() As Integer
       
    Dim iSearch As Integer
    Dim iRandomNo As Integer
       
    'Initialize random-number generator.--------------------------------------------
    Randomize
    '-------------------------------------------------------------------------------
                      
    'Select a random File to play, if the File has already been played, do nothing.-
    'Else play File. Add the Val(1) to the amount of Files that have been played to
    'mVariables.iRandomCount.
    Do
        
        With frmMain.lstFiles
            
            iRandomNo% = Int((.ListItems.Count * Rnd) + 1)
            
            If .ListItems(iRandomNo%).ListSubItems(3).Text = "P" Then
            Else
                ftnRandomSelect% = iRandomNo%
                mVariables.iRandomCount = mVariables.iRandomCount + 1
            Exit Function
            End If
        
        End With

    Loop Until mVariables.iRandomCount >= frmMain.lstFiles.ListItems.Count

End Function


'***********************************************************************************
'Increase or Decrease the volume, update volume indicator controls.
'***********************************************************************************
Public Sub subSetVolume(sVolumeSet As String)

    Select Case sVolumeSet$
                
        Case "Increase"
            
            If mVariables.iVolumeSetting < 995 Then
                
                mVariables.iVolumeSetting = mVariables.iVolumeSetting + 166
    
                'Set the playing volume.--------------------------------------------------------
                mciSendString "setaudio mp3 volume to " & mVariables.iVolumeSetting, 0, 0, 0
                '-------------------------------------------------------------------------------

            End If
        
        Case "Decrease"
        
            If mVariables.iVolumeSetting > 331 Then
                
                mVariables.iVolumeSetting = mVariables.iVolumeSetting - 166
                                
                'Set the playing volume.--------------------------------------------------------
                mciSendString "setaudio mp3 volume to " & mVariables.iVolumeSetting, 0, 0, 0
                '-------------------------------------------------------------------------------
            
            End If

    End Select

    Call subVolumeInd(mVariables.iVolumeSetting)

End Sub


'***********************************************************************************
'Updates the volume indicator controls.
'***********************************************************************************
Private Sub subVolumeInd(iVolume As Integer)
        
        
    With frmMain
        
        If iVolume% <= 166 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = &H80&
            .VolumeInd(2).FillColor = &H80&
            .VolumeInd(3).FillColor = &H80&
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 166 And iVolume% <= 332 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = &H80&
            .VolumeInd(3).FillColor = &H80&
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 332 And iVolume% <= 498 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = &H80&
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 498 And iVolume% <= 664 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = RGB(250, 0, 0)
            .VolumeInd(4).FillColor = &H80&
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 664 And iVolume% <= 830 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = RGB(250, 0, 0)
            .VolumeInd(4).FillColor = RGB(250, 0, 0)
            .VolumeInd(5).FillColor = &H80&
        ElseIf iVolume% > 830 And iVolume% <= 996 Then
            .VolumeInd(0).FillColor = RGB(250, 0, 0)
            .VolumeInd(1).FillColor = RGB(250, 0, 0)
            .VolumeInd(2).FillColor = RGB(250, 0, 0)
            .VolumeInd(3).FillColor = RGB(250, 0, 0)
            .VolumeInd(4).FillColor = RGB(250, 0, 0)
            .VolumeInd(5).FillColor = RGB(250, 0, 0)
        End If
        
    End With

End Sub

