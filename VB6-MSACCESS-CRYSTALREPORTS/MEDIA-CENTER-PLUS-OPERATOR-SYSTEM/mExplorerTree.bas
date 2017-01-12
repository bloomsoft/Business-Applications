Attribute VB_Name = "mExplorerTree"
'***********************************************************************************
'Populates a TreeView with the sub Folder`s of the selected Drive or Folder.
'***********************************************************************************

Option Explicit

Public Sub subShowFolderList(oFolderList As ListBox, oExplorerTree As TreeView, sDriveLetter As String, vParentID As Variant)
    
    Dim nNode As Node                           'Node object for DirTree.
    Dim lReturn As Long                         'Holds Search Handle of File.
    Dim lNextFile As Long                       'Return Search Handle of next Folder.
    Dim sPath As String                         'Path to search.
    Dim WFD As WIN32_FIND_DATA                  'Win32 Structure (VB Type).
    Dim sFolderName As String                   'Name of Folder.
    Dim x As Long                               'Used to loop through Folders in frmMain.List1).
    Set oFolderList = frmExplore.List1          'Set Object oFolderList as frmMain.List1.
    Set oExplorerTree = frmExplore.Explorer     'Set Object oExplorerTree as source Explorer Tree.
       
        
    'Return all Folders from selected Drive.----------------------------------------
    sPath$ = (sDriveLetter & "*.*") & Chr$(0)
    '-------------------------------------------------------------------------------
    
    'Search for First Folder Handle.------------------------------------------------
    lReturn& = FindFirstFile(sPath$, WFD)
    '-------------------------------------------------------------------------------
    
    'Loop through all Folders (One level).------------------------------------------
    Do

        'If a Folder is found add to oFolderList.------------------------------------
        If (WFD.dwFileAttributes And vbDirectory) Then
            
            'Strip vbNullChar from Folder Name.-------------------------------------
            sFolderName$ = mProcFunc.ftnStripNullChar(WFD.cFileName)
            '-----------------------------------------------------------------------
            
            If sFolderName$ <> "." And sFolderName$ <> ".." Then
                
                'If the Folder has an Attribute <> 16 then add "~A~" to it.---------
                If WFD.dwFileAttributes <> 16 Then
                    oFolderList.AddItem sFolderName$ & "~A~"
                Else
                    oFolderList.AddItem sFolderName$ & "~~~"
                End If
                '-------------------------------------------------------------------

            End If
        End If
        '---------------------------------------------------------------------------
        
        'Search for Handle of next Folder.------------------------------------------
        lNextFile& = FindNextFile(lReturn&, WFD)
        '---------------------------------------------------------------------------
 
    Loop Until lNextFile& = False
    '-------------------------------------------------------------------------------
  
    'Close Handle of Folder.--------------------------------------------------------
    lNextFile& = FindClose(lReturn&)
    '-------------------------------------------------------------------------------

    'Loop through oFolderList which has it`s sorted property set to True, then add---
    'Folder Path to DirTree
    For x = 0 To oFolderList.ListCount - 1

        'If the Folder has an Attribute, set ForeColor to Grey----------------------
        If Right(oFolderList.List(x), 3) = "~A~" Then
            Set nNode = oExplorerTree.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
            nNode.ForeColor = RGB(120, 120, 120)
        Else
            Set nNode = oExplorerTree.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
        End If
        '---------------------------------------------------------------------------
    
    Next x

    'Clear frmMain.List1.-----------------------------------------------------------
    oFolderList.Clear
    '-------------------------------------------------------------------------------

End Sub




