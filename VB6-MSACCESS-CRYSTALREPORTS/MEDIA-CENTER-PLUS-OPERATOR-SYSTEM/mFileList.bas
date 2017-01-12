Attribute VB_Name = "mFileList"
'***********************************************************************************
'All that were doing here is using the FindFirstFile, FindNextFile and FindClose API
'adding all Files to the ListView as we go. Once we`ve added all the Files, we close
'the search Handle lngReturn&
'***********************************************************************************

Option Explicit
                                                                                    
            
Public Sub subFileList(sFolderPath As String)

    Dim lReturn As Long                    'Search Handle of specified Path.
    Dim lNextFile As Long                  'Search Handle of specified File.
    Dim sPath As String                    'Path to search.
    Dim WFD As WIN32_FIND_DATA             'Set Variable WFD as Structure(VBType) WIN32_FIND_DATA.
    Dim LstItem As ListItem                'lstItem = A ListView ListItem.
    Dim lstSubItem As ListSubItem          'lstSubItem = A ListView ListSubItem.
    Dim sFileName As String                'Filename (WFD.cFileName).
    Dim oFileList As ListView              'Set oFileList as Control being used.
        Set oFileList = frmExplore.FileList
        sPath$ = sFolderPath$ & "*.mp3"
    Dim lFileLoop As Long                  'Loop for setting ForeColour of specific Files. eg(*.exe).
   
    With oFileList
        
        .Visible = False
        .ListItems.Clear
    
        lReturn& = FindFirstFile(sPath$, WFD) & Chr$(0)
        frmExplore.MousePointer = 11
                
        Do
                       
            'If we find a Directory do nothing, else List Files taking off the Chr$(0)
            'Loop until lNextFile& = val(0), no more Files to List
            If Not (WFD.dwFileAttributes And vbDirectory) = vbDirectory Then
        
                sFileName$ = mProcFunc.ftnStripNullChar(WFD.cFileName)
            
                If sFileName > Trim("") Then
                        Set LstItem = .ListItems.Add(, , sFileName$)
                        Set lstSubItem = LstItem.ListSubItems.Add(, , Format(WFD.nFileSizeLow, "#,0"))
'                        Set lstSubItem = lstItem.ListSubItems.Add(, , mProcFunc.ftnReturnAttributes(WFD.dwFileAttributes))
                End If
            
            End If
        
            lNextFile& = FindNextFile(lReturn&, WFD)
        
        Loop Until lNextFile& <= Val(0)

        frmExplore.MousePointer = 0
    
        'Close Search Handle.-------------------------------------------------------
        lNextFile& = FindClose(lReturn&)
        '---------------------------------------------------------------------------
        
        'Set ForeColor of specified Files in FileList.------------------------------
        For lFileLoop = 1 To .ListItems.Count

            If InStrRev(LCase(.ListItems(lFileLoop).Text), ".mp3", , vbTextCompare) Then
                .ListItems(lFileLoop).ForeColor = RGB(60, 60, 140)
''            ElseIf InStrRev(LCase(.ListItems(lFileLoop).Text), ".zip", , vbTextCompare) Then
''                .ListItems(lFileLoop).ForeColor = vbBlue
''            ElseIf InStrRev(LCase(.ListItems(lFileLoop).Text), ".txt", , vbTextCompare) Then
''                .ListItems(lFileLoop).ForeColor = RGB(50, 100, 50)
            
            End If
        
        Next
        
        .Visible = True
    
    End With

End Sub

