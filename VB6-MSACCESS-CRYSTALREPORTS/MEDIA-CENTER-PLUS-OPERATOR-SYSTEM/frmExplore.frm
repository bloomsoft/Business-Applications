VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorer."
   ClientHeight    =   6120
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSComctlLib.ListView FileList 
      Height          =   6015
      Left            =   4440
      TabIndex        =   0
      Top             =   110
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName."
         Object.Width           =   6421
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FileSize."
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   210
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":0000
            Key             =   "cldfolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":015A
            Key             =   "opnfolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":02B4
            Key             =   "drvcd"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":040E
            Key             =   "drvremove"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":0568
            Key             =   "drvfixed"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":06C2
            Key             =   "drvremote"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":081C
            Key             =   "mycomputer"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":0976
            Key             =   "drvunknown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplore.frx":0AD0
            Key             =   "drvmemory"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Explorer 
      Height          =   5985
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10557
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Line lneMenu 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   0
      X2              =   10440
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line lneMenu 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   10440
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu h2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileCOpenFiles 
         Caption         =   "Open Files"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu h1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMp3 
      Caption         =   "Mp3"
      Begin VB.Menu h3 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuMp3CSelectAll 
         Caption         =   "Select All"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "frmExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'***********************************************************************************
'Form Events.
'***********************************************************************************
'Form_Load.
Private Sub Form_Load()
    
    Dim sComputerName As String * 255
    Dim lAPIReturn As Long
    Dim cDrives As cDrives
        Set cDrives = New cDrives
    
    'Use API to retrieve ComputerName, strip NullChar and Hold ComputerName in.-----
    'Private Variable.
    lAPIReturn& = GetComputerName(sComputerName$, Len(sComputerName$))
        
    mVariables.sComputerName = mProcFunc.ftnStripNullChar(sComputerName$)
    '-------------------------------------------------------------------------------
    
    'Iterate all Local Drives into Explorer Tree.-----------------------------------
    cDrives.subLoadTreeView
    '-------------------------------------------------------------------------------
    
    'Expand first Node in Explorer Tree.--------------------------------------------
    Explorer.Nodes(1).Expanded = True
    '-------------------------------------------------------------------------------
    
End Sub

'Form_Unload.
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frmExplore = Nothing
End Sub


'***********************************************************************************
'Explorer Events.
'***********************************************************************************
'Explorer_Expand.
Private Sub Explorer_Expand(ByVal Node As MSComctlLib.Node)

    DoEvents
    Dim x As Long
    
    'Branch all sub Folders if mVariables.BranchFolders = True.---------------------
    Me.MousePointer = 11
    For x = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
        Explorer_NodeClick Explorer.Nodes(x)
    Next x
    Me.MousePointer = 0
    '-------------------------------------------------------------------------------
    
End Sub

'Explorer_NodeClick.
Private Sub Explorer_NodeClick(ByVal Node As MSComctlLib.Node)
   
    Dim sNodePath As String
        sNodePath$ = mProcFunc.ftnReturnNodePath(Node.FullPath)
      
    'If Not Children list Folders.--------------------------------------------------
    If Not Node.Children > 0 Then
        mExplorerTree.subShowFolderList List1, Explorer, sNodePath$, Node.Index
    End If
    '-------------------------------------------------------------------------------
    
    If Node.Selected = True Then
            
        'List Files if Node is selected.--------------------------------------------
        Call mFileList.subFileList(sNodePath$)
        '---------------------------------------------------------------------------
    
        'If Filecount > 0 then enable menu items.-----------------------------------
        If FileList.ListItems.Count > 0 Then
            mnuMp3CSelectAll.Enabled = True
            mnuFileCOpenFiles.Enabled = True
        Else
            mnuFileCOpenFiles.Enabled = False
            mnuMp3CSelectAll.Enabled = False
        End If
        '---------------------------------------------------------------------------
                
    End If
    '-------------------------------------------------------------------------------

End Sub


'***********************************************************************************
'Menu File Commands.
'***********************************************************************************
'Open.
Private Sub mnuFileCOpenFiles_Click()
    
    Dim x As Long
    Dim LstItem As ListItem
    Dim lstSubItem As ListSubItem
    Dim lMp3Length As Long
    Dim sCurrentPath As String
        sCurrentPath$ = mProcFunc.ftnReturnNodePath(Explorer.SelectedItem.FullPath)
    
    'If item is selected, add to frmMain.lstFiles.----------------------------------
    With FileList
        For x = 1 To .ListItems.Count
            If .ListItems(x).Selected = True Then
                Set LstItem = frmMain.lstFiles.ListItems.Add(, , sCurrentPath$)
                
                'Add a 0 if frmMain.lstFiles.ListItems.Count < 10, sets the column to
                'start from 01 instead of 1 etc.
                If Val(frmMain.lstFiles.ListItems.Count) < Val(10) Then
                    Set lstSubItem = LstItem.ListSubItems.Add(, , "0" & Val(frmMain.lstFiles.ListItems.Count) & ".")
                Else
                    Set lstSubItem = LstItem.ListSubItems.Add(, , Val(frmMain.lstFiles.ListItems.Count) & ".")
                End If
                '-------------------------------------------------------------------
                
                Set lstSubItem = LstItem.ListSubItems.Add(, , .ListItems(x).Text)
                Set lstSubItem = LstItem.ListSubItems.Add(, , "U")
'                MsgBox "Test"
            End If
        Next
    End With
    '-------------------------------------------------------------------------------

    'Form_Unload (0)

End Sub

'Exit.
Private Sub mnuFileExit_Click()

    Form_Unload (0)

End Sub


'***********************************************************************************
'Menu Mp3 Commands.
'***********************************************************************************
'Select All.
Private Sub mnuMp3CSelectAll_Click()
    
    Dim x As Long
    
    'Select all Files in FileList.--------------------------------------------------
    With FileList
        For x = 1 To .ListItems.Count
            .ListItems(x).Selected = True
        Next
    End With
    '-------------------------------------------------------------------------------
    
End Sub
