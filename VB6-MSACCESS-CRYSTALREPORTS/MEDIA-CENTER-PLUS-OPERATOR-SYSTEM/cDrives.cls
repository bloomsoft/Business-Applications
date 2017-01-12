VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************
'Return all Drives installed on the Users Local System using the FileSystemObject and
'add to frmExplorer.Explorer(TreeView).
'***********************************************************************************

Option Explicit

Public Sub subLoadTreeView()

    Dim TreeList As TreeView                    'Explorer Tree.
        Set TreeList = frmExplore.Explorer
    Dim iDriveNum As Integer                    'Key Index in DirTree.
    Dim sDriveType As String                    'Holds DriveType.
    Dim fso, d As Object                        'Used to return DriveType.
        Set fso = CreateObject("Scripting.FileSystemObject")
    Dim x As Integer                            'Loop through Drives.
    Dim RetStr(1) As String                     'Holds Drive letters.
    Dim nNode As Node                           'Node object for ExplorerTree.
    Dim sComputerName As String                 'Hold Computer Name.
        sComputerName$ = mVariables.sComputerName
                        
    'Return Drive structure from XFile.Dll.-----------------------------------------
    RetStr$(0) = ftnShowDriveList
    '-------------------------------------------------------------------------------

    With TreeList

        'Add Computer name to DirTree-----------------------------------------------
        Set nNode = .Nodes.Add(, , sComputerName$, sComputerName$, "mycomputer", "mycomputer")
        '---------------------------------------------------------------------------
        
        'Add Drive A:\ to DirTree---------------------------------------------------
'        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & "0", "3.5 Floppy (A:)", "drvremove") 'Add Drive "A:"
        '---------------------------------------------------------------------------
        
        'Loop through RetStr$(0) to retrieve Drives. eg."ACDEF".--------------------
        For x = 1 To Len(RetStr$(0))
            
            'Strip Driveinfo eg "A"-------------------------------------------------
            RetStr$(1) = Mid(RetStr$(0), x, 1)
            '-----------------------------------------------------------------------
            
            'Get DriveType information.---------------------------------------------
            Set d = fso.GetDrive(RetStr$(1))
            '-----------------------------------------------------------------------
            
            'Used to make unique Key Index in DirTree.------------------------------
            iDriveNum% = x
            '-----------------------------------------------------------------------
                       
            'Determine Drive type and add to TreeList.------------------------------
            Select Case d.drivetype
                
                'Unknown Drive.-----------------------------------------------------
                Case 0: sDriveType = "Unknown"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvunknown")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvunknown")
                    End If

                'Removable Drive.---------------------------------------------------
                Case 1: sDriveType = "Removable"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":", "drvremove")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvremove")

                    End If
                    
                'Fixed Drive.-------------------------------------------------------
                Case 2: sDriveType = "Fixed"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvfixed")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvfixed")
                    End If

                'Network Drive.-----------------------------------------------------
                Case 3: sDriveType = "Network"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvremote")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvremote")
                    End If
                    
                'CD-Rom.------------------------------------------------------------
                Case 4: sDriveType = "CD-ROM"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvcd")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvcd")
                    End If
                    
                'Ram Disk.----------------------------------------------------------
                Case 5: sDriveType = "Ram Disk"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvram")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvram")
                    End If
            
            End Select
            '-----------------------------------------------------------------------
            
        Next x

    End With

End Sub


'***********************************************************************************
'Return all Drives installed on the Users Local System using the FileSystemObject
'***********************************************************************************
Private Function ftnShowDriveList()
  
    Dim fso, d, dc As Object
    Dim sDriveLetter As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set dc = fso.Drives
    
    For Each d In dc
        sDriveLetter$ = sDriveLetter$ & d.driveletter
    Next
    
    ftnShowDriveList = sDriveLetter$

End Function

