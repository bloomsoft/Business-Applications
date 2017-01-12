VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "BloomSoft Media Center"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "Apna Radio"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4230
      TabIndex        =   23
      Top             =   1995
      Width           =   1155
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Desi Radio"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3090
      TabIndex        =   21
      Top             =   1980
      Width           =   1155
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2565
      Left            =   -285
      TabIndex        =   20
      Top             =   3465
      Visible         =   0   'False
      Width           =   5835
      ExtentX         =   10292
      ExtentY         =   4524
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "SunRise Radio"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   1665
      TabIndex        =   19
      Top             =   1995
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Audio Player"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   315
      TabIndex        =   18
      Top             =   1965
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   3150
      Top             =   4920
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   3225
      Top             =   5880
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   1950
      Top             =   6060
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4455
      Top             =   5895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   5295
   End
   Begin VB.PictureBox ProgressBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      DrawWidth       =   40
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   80
      Left            =   735
      ScaleHeight     =   45
      ScaleWidth      =   3585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3135
      Width           =   3615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   5775
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   2550
      Left            =   240
      TabIndex        =   0
      Top             =   3450
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4498
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   8454143
      BackColor       =   0
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Path."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Count."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Name."
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Played"
         Object.Width           =   0
      EndProperty
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP1 
      Height          =   2445
      Left            =   30
      TabIndex        =   22
      Top             =   3555
      Width           =   5460
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9631
      _cy             =   4313
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BloomSoft Media Center [http://www.bloomsoft.net]      "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   45
      TabIndex        =   17
      Top             =   60
      Width           =   5490
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1770
      TabIndex        =   16
      Top             =   615
      Width           =   3270
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   8
      Left            =   1800
      Top             =   2730
      Width           =   105
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   15
      Left            =   4080
      Picture         =   "frmMain.frx":014A
      Stretch         =   -1  'True
      ToolTipText     =   "Normal Play"
      Top             =   2730
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   14
      Left            =   1920
      Picture         =   "frmMain.frx":06B4
      ToolTipText     =   "Clear Playlist"
      Top             =   2730
      Width           =   510
   End
   Begin VB.Label lblPlayMode 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Normal play"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   1575
      Width           =   825
   End
   Begin VB.Label lblPlayMode 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Play Mode."
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1575
      Width           =   795
   End
   Begin VB.Image imgControl 
      Height          =   150
      Index           =   13
      Left            =   4305
      Picture         =   "frmMain.frx":0D0E
      Top             =   3105
      Width           =   180
   End
   Begin VB.Image imgControl 
      Height          =   150
      Index           =   12
      Left            =   570
      Picture         =   "frmMain.frx":11F8
      Top             =   3105
      Width           =   180
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   4695
      X2              =   4695
      Y1              =   2490
      Y2              =   3270
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   4680
      X2              =   4680
      Y1              =   2490
      Y2              =   3270
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   5
      Left            =   5040
      Top             =   2490
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   4
      Left            =   5040
      Top             =   2625
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   3
      Left            =   5040
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   2
      Left            =   5040
      Top             =   2910
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   1
      Left            =   5040
      Top             =   3045
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   0
      Left            =   5040
      Top             =   3180
      Width           =   135
   End
   Begin VB.Image imgControl 
      Height          =   375
      Index           =   11
      Left            =   4800
      Picture         =   "frmMain.frx":16E2
      ToolTipText     =   "Decrease Volume"
      Top             =   2895
      Width           =   180
   End
   Begin VB.Image imgControl 
      Height          =   375
      Index           =   10
      Left            =   4800
      Picture         =   "frmMain.frx":1CBC
      ToolTipText     =   "Increase Volume"
      Top             =   2490
      Width           =   180
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   7
      Left            =   1080
      Top             =   2730
      Width           =   105
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   9
      Left            =   1200
      Picture         =   "frmMain.frx":2296
      ToolTipText     =   "Random Select"
      Top             =   2730
      Width           =   510
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   6
      Left            =   360
      Top             =   2730
      Width           =   105
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   8
      Left            =   480
      Picture         =   "frmMain.frx":28F0
      ToolTipText     =   "Mute"
      Top             =   2730
      Width           =   510
   End
   Begin VB.Label lblArtistA 
      BackColor       =   &H00400000&
      Caption         =   "No Track loaded."
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      Top             =   615
      Width           =   3360
   End
   Begin VB.Label lblPositionA 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   12
      Top             =   1095
      Width           =   90
   End
   Begin VB.Label lblLengthA 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      Top             =   855
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   ":"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2070
      TabIndex        =   10
      Top             =   1335
      Width           =   45
   End
   Begin VB.Label lblELTimeD 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   1335
      Width           =   90
   End
   Begin VB.Label lblELTimeC 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   1335
      Width           =   90
   End
   Begin VB.Label lblELTimeB 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   1335
      Width           =   90
   End
   Begin VB.Label lblELTimeA 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   1335
      Width           =   90
   End
   Begin VB.Label lblElTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Elapsed Time."
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1335
      Width           =   1005
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00400000&
      Caption         =   "Position."
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1095
      Width           =   615
   End
   Begin VB.Label lblArtist 
      BackColor       =   &H00400000&
      Caption         =   "Artist."
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   615
      Width           =   4815
   End
   Begin VB.Label lblLength 
      BackColor       =   &H00400000&
      Caption         =   "Length."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   855
      Width           =   555
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   690
      X2              =   4340
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   690
      X2              =   4340
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   7
      Left            =   4320
      Picture         =   "frmMain.frx":2F4A
      ToolTipText     =   "Fast Play"
      Top             =   2730
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   6
      Left            =   3840
      Picture         =   "frmMain.frx":34B4
      Stretch         =   -1  'True
      ToolTipText     =   "Slow Play"
      Top             =   2730
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   480
      X2              =   4560
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   480
      X2              =   4560
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   5
      Left            =   3960
      Top             =   2490
      Width           =   105
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   4
      Left            =   3240
      Top             =   2490
      Width           =   105
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   3
      Left            =   2520
      Top             =   2490
      Width           =   105
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   2
      Left            =   1800
      Top             =   2490
      Width           =   105
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   1
      Left            =   1080
      Top             =   2490
      Width           =   105
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   0
      Left            =   360
      Top             =   2490
      Width           =   105
   End
   Begin VB.Shape speBorder 
      BorderColor     =   &H00E0E0E0&
      Height          =   975
      Index           =   0
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2355
      Width           =   5055
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   0
      Left            =   480
      Picture         =   "frmMain.frx":3A1E
      ToolTipText     =   "Previous"
      Top             =   2490
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   1
      Left            =   1185
      Picture         =   "frmMain.frx":4078
      ToolTipText     =   "Play"
      Top             =   2490
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   2
      Left            =   1920
      Picture         =   "frmMain.frx":46D2
      ToolTipText     =   "Pause"
      Top             =   2490
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   3
      Left            =   2655
      Picture         =   "frmMain.frx":4D2C
      ToolTipText     =   "Stop"
      Top             =   2490
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   4
      Left            =   3360
      Picture         =   "frmMain.frx":5386
      ToolTipText     =   "Next"
      Top             =   2490
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   5
      Left            =   4080
      Picture         =   "frmMain.frx":59E0
      ToolTipText     =   "Open"
      Top             =   2490
      Width           =   510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   240
      X2              =   5310
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Shape InfoWindow 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   5055
   End
   Begin VB.Shape speBorder 
      BorderColor     =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   1
      Left            =   255
      Shape           =   4  'Rounded Rectangle
      Top             =   2355
      Width           =   5055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mniSetup 
         Caption         =   "Setup"
      End
      Begin VB.Menu mniHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mniExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private sReturnBuffer As String * 30 'Holds return information when using mciSendString.
Private fso As New FileSystemObject
Dim CurrentPlayType As String
Dim CurrentFilePath As String
Dim CurrentFilePathAzan As String
Dim CurrentPosition As String
Private Sub GetPlayMode()
Dim DB As Database
Dim TB As Recordset
Set DB = OpenDatabase(App.Path & "\Main.mdb")
Set TB = DB.OpenRecordset("Select * from Setup")
If Not TB.EOF Then
    If TB.Fields("Mode") = 1 Then
        Option1 = True
    End If
    If TB.Fields("Mode") = 2 Then
        Option2 = True
        lstFiles.Visible = False
        WebBrowser1.Visible = True
        WebBrowser1.Navigate2 "http://www.sunriseradio.com/player/audioplayer.asp"
        CurrentPlayType = "RADIO"
        
    End If
End If
TB.Close
DB.Close
End Sub
Private Sub PlayAzan(A As String)
    Dim lShortPath As Long          'Used for API GetShortPathName.
    Dim sShortPath As String * 260  'Used for API GetShortPathName.
    Dim sShortPathName As String    'Used for holding correct ShortPathName after stripping NullChar.
    lShortPath& = GetShortPathName(App.Path & "\Azan\" & A, sShortPath$, 260)
    sShortPathName$ = mProcFunc.ftnStripNullChar(sShortPath$)
    'MsgBox "Test"
     '---------------------------------------------------------------
    '1.Open File, 2.Set time format, 3.Play File. 4.Retrieve length of File.
    mciSendString "open " & sShortPathName$ & " type MPEGVideo alias mp3", 0, 0, 0
    mciSendString "set mp3 time format tmsf", 0, 0, 0
    mciSendString "play mp3", 0, 0, 0
    mciSendString "status mp3 length", sReturnBuffer$, Len(sReturnBuffer$), 0
    mciSendString "setaudio mp3 volume to " & mVariables.iVolumeSetting, 0, 0, 0
    mVariables.lTrackLength = Val(sReturnBuffer)
                    '---------------------------------------------------------------
    mVariables.bTrackIsPlaying = True
                    '---------------------------------------------------------------
                    
End Sub
'***********************************************************************************
'Form Events.
'***********************************************************************************
'Form_Load.
Private Sub SaveTrack(TrackName As String, TrackPosition As String)
On Error GoTo ErrHand

If Len(TrackName) <= 0 Then Exit Sub
If Len(TrackPosition) <= 0 Then Exit Sub
Dim DB As Database
Dim Ssql As String
Set DB = OpenDatabase(App.Path & "\Main.mdb")
Ssql = "Update Tracks Set "
If CurrentPlayType = "QURAN" Then
    Ssql = Ssql & " Quran='" & TrackName & "',QuranPosition='" & TrackPosition & "'"
End If
If CurrentPlayType = "NAAT" Then
    Ssql = Ssql & " Naat='" & TrackName & "',NaatPosition='" & TrackPosition & "'"
End If
If CurrentPlayType = "SONGS" Then
    Ssql = Ssql & " Song='" & TrackName & "',SongPosition='" & TrackPosition & "'"
End If
DB.Execute Ssql
DB.Close
Exit Sub
ErrHand:
End Sub
Private Sub DeciedPlayList()
'MsgBox "PlayList Decision"
If Option2 = True Then
    If CurrentPlayType <> "RADIO1" Then
'        MsgBox "Going to Play Radio"
        WebBrowser1.Visible = True
        WMP1.Visible = False
        WebBrowser1.Navigate2 "http://www.sunriseradio.com/player/audioplayer.asp"
        CurrentPlayType = "RADIO1"
        
        Exit Sub
    End If
End If
If Option3 = True Then
    If CurrentPlayType <> "RADIO2" Then
'        MsgBox "Going to Play Radio"
        WMP1.URL = "http://208.53.138.213:80/"
        WMP1.Controls.Play
        WebBrowser1.Visible = False
        WMP1.Visible = True
        CurrentPlayType = "RADIO2"
        Exit Sub
    End If
End If

If Option4 = True Then
    If CurrentPlayType <> "RADIO3" Then
'        MsgBox "Going to Play Radio"
        WMP1.URL = "http://www.apnaradio.com/live/media24/ApnaRadio.asx"
        'MsgBox "Test"
        WMP1.Controls.Play
        WebBrowser1.Visible = False
        WMP1.Visible = True
        CurrentPlayType = "RADIO3"
        Exit Sub
    End If
End If

Dim B As Boolean
Dim R As Integer
Dim EndDate As Date
Dim StartDate As Date
MyTimes(1) = DateAdd("h", GPlayTime, MyTimes(0))

If MyTimes(0) > Now Then
        imgControl_Click 2
        lblArtistA.Caption = ""
        'lblArtist.Caption = ""
        Label2.Caption = "It's Stoppage Time!"
'        MsgBox "Test"
        Exit Sub
ElseIf MyTimes(1) < Now Then
        imgControl_Click 2
        lblArtistA.Caption = ""
        'lblArtist.Caption = ""
        Label2.Caption = "It's Stoppage Time!"
 '       MsgBox "Test"
        Exit Sub
Else
        lblArtistA.Caption = ""
        'lblArtist.Caption = ""
        Label2.Caption = ""
End If


EndDate = DateAdd("h", 1, MyTimes(0))
'MsgBox "Test"
If MyTimes(0) <= Now And EndDate > Now Then
    'QuranTime
    If CurrentPlayType <> "QURAN" Then
        imgControl_Click 3
        DoEvents
        B = LoadPlayList("Quran")
        DoEvents
        If B = True Then
            imgControl_Click 1
        End If
        
    End If
    
    Exit Sub
End If
StartDate = DateAdd("h", 1, MyTimes(0))
EndDate = DateAdd("h", 2, MyTimes(0))
If StartDate <= Now And EndDate > Now Then
    'NaatTime
    If CurrentPlayType <> "NAAT" Then
        imgControl_Click 3
        DoEvents
        B = LoadPlayList("Naat")
        DoEvents
        If B = True Then
            imgControl_Click 1
        End If
    End If
    Exit Sub
End If
    If CurrentPlayType <> "SONGS" Then
        If Option1 = True Then
        'MsgBox "Going to Play Songs"
        imgControl_Click 3
        DoEvents
        B = LoadPlayList("Songs")
        DoEvents
        If B = True Then
            imgControl_Click 1
        End If
        End If
    End If

End Sub
Private Function LoadPlayList(FolderType As String) As Boolean
On Error Resume Next
    If Option2 = True Then
        LoadPlayList = False
        Exit Function
    End If
    Dim LstItem As ListItem
    Dim MFile As File
    Dim CurrentP As String
    
    Dim lstSubItem As ListSubItem
    Dim lMp3Length As Long
    
    If FolderType = "Quran" Then
        
        CurrentP = QuranFolder
        CurrentPlayType = "QURAN"
    End If
    If FolderType = "Naat" Then
        
        CurrentP = NaatFolder
        CurrentPlayType = "NAAT"
    End If
    If FolderType = "Songs" Then
        
        CurrentP = SongsFolder
        CurrentPlayType = "SONGS"
    End If
    Dim R As Integer
        
    lstFiles.ListItems.Clear
    For Each MFile In fso.GetFolder(CurrentP).Files
    
                Set LstItem = lstFiles.ListItems.Add(, , CurrentP & "\")
                
                'Add a 0 if frmMain.lstFiles.ListItems.Count < 10, sets the column to
                'start from 01 instead of 1 etc.
                If Val(lstFiles.ListItems.Count) < Val(10) Then
                    Set lstSubItem = LstItem.ListSubItems.Add(, , "0" & Val(lstFiles.ListItems.Count) & ".")
                Else
                    Set lstSubItem = LstItem.ListSubItems.Add(, , Val(lstFiles.ListItems.Count) & ".")
                End If
                '-------------------------------------------------------------------
                
                Set lstSubItem = LstItem.ListSubItems.Add(, , MFile.Name)
                Set lstSubItem = LstItem.ListSubItems.Add(, , "U")
    Next
    
    DoEvents
    DoSelection
    LoadPlayList = True
End Function
Private Sub DoSelection()
Dim DB As Database
Dim Ssql As String
Dim TB As Recordset
Dim LItem As ListItem
Ssql = "Select * from Tracks"
Set DB = OpenDatabase(App.Path & "\Main.mdb")
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        For Each LItem In lstFiles.ListItems
            If CurrentPlayType = "QURAN" Then
                If LItem.SubItems(2) = TB.Fields("Quran").Value Then
                    CurrentPosition = TB.Fields("QuranPosition").Value & ""
                    LItem.Selected = True
                    Exit For
                End If
            End If
            If CurrentPlayType = "NAAT" Then
                If LItem.SubItems(2) = TB.Fields("Naat").Value Then
                    CurrentPosition = TB.Fields("NaatPosition").Value & ""
                    LItem.Selected = True
                    Exit For
                End If
            End If
            If CurrentPlayType = "SONGS" Then
                If LItem.SubItems(2) = TB.Fields("Song").Value Then
                    CurrentPosition = TB.Fields("SongPosition").Value & ""
                    LItem.Selected = True
                    Exit For
                End If
            End If

        Next
End If
TB.Close
DB.Close
End Sub

Private Sub Command1_Click()
PlayAzan "Azan1.mp3"
End Sub

Private Sub Form_Load()
    GetPlayMode
    Module1.GetTimes
    'Create visible region of frmMain.----------------------------------------------
    Dim lCreateRegion As Long
'        lCreateRegion& = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 30, 30)

'    SetWindowRgn Me.hwnd, lCreateRegion&, True
    '-------------------------------------------------------------------------------

    'Set mciSendString time format.-------------------------------------------------
    'mciSendString "set mp3 time format tmsf wait", True
    '-------------------------------------------------------------------------------

    'Set the current playing volume to 50% of max value.----------------------------
    'mVariables.iVolumeSetting = 498
    mVariables.iVolumeSetting = Val(oldVolume)
'    MsgBox Val(oldVolume)
    If Val(oldVolume) > 498 Then
        mProcFunc.subSetVolume ("Increase")
    ElseIf Val(oldVolume) < 498 Then
        mProcFunc.subSetVolume ("Decrease")
        
    End If
    '-------------------------------------------------------------------------------

    'Update Volume indicator controls.----------------------------------------------
    VolumeInd(0).FillColor = RGB(250, 0, 0)
    VolumeInd(1).FillColor = RGB(250, 0, 0)
    VolumeInd(2).FillColor = RGB(250, 0, 0)
    '-------------------------------------------------------------------------------
DeciedPlayList

        With SysIcon
        .cbSize = LenB(SysIcon)
        .hWnd = Me.hWnd
        .uFlags = NIF_DOALL
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .sTip = "BloomSoft Media Center"
    End With
    Shell_NotifyIcon NIM_ADD, SysIcon
    Me.Visible = False

End Sub

'Form_MouseMove.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
                
    'Enable and disable notification lights when mouse moves over a button.---------
    speCommandLight(mVariables.byCommandLight).FillColor = vbBlack
    '-------------------------------------------------------------------------------
        Dim ShellMsg As Long
    
    ShellMsg = x / Screen.TwipsPerPixelX
    Select Case ShellMsg
    Case WM_LBUTTONDBLCLK
        Me.Visible = True
    Case WM_RBUTTONUP
        PopupMenu mnuFile
    End Select

    
End Sub

'Form_QueryUnload.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Check to see if a File is being played, if so, close MCI Device.---------------
    If mVariables.bTrackIsPlaying = True Then
        Call imgControl_Click(3)
    End If
    '-------------------------------------------------------------------------------

End Sub

'Form_Resize.
Private Sub Form_Resize()
        
    'Stops runtime Error 384 when the minimize button is depressed.-----------------
    On Error Resume Next
    '-------------------------------------------------------------------------------
    
    'Don`t allow the Form to be resized.--------------------------------------------
    If Me.Width <> 5670 Then
        Me.Width = 5670
    ElseIf Me.Height <> 6900 Then
        Me.Height = 6900
    End If
    '-------------------------------------------------------------------------------
    
End Sub

'***********************************************************************************
'imgControl Events.
'***********************************************************************************
'imgControl_Click.
Private Sub imgControl_Click(Index As Integer)

    Dim lShortPath As Long          'Used for API GetShortPathName.
    Dim sShortPath As String * 260  'Used for API GetShortPathName.
    Dim sShortPathName As String    'Used for holding correct ShortPathName after stripping NullChar.
    
'    If lstFiles.ListItems.Count > 0 Then
        
        Select Case Index

            'Previous Track.--------------------------------------------------------
            Case 0
                With lstFiles
                    If .ListItems.Count > 1 Then
                        If .SelectedItem.Index > 1 Then
                            
                            .ListItems(.SelectedItem.Index - 1).Selected = True
                        
                            'Stop and close mp3.------------------------------------
                            Call imgControl_Click(3)
                            '-------------------------------------------------------
                            'Play mp3.----------------------------------------------
                            Call imgControl_Click(1)
                            '-------------------------------------------------------
                        
                        End If
                    End If
                End With
                                                                                    
            'Play Track.------------------------------------------------------------
            Case 1
                If lstFiles.ListItems.Count > 1 And Option1 = True Then
                    'MsgBox "Going to Play"
                    Me.MousePointer = 11
                    ProgressBar.Cls
                    
                    'Use API GetShortPathName to retrieve the Path of the selected File.
                    With lstFiles
                        lShortPath& = GetShortPathName(.ListItems(.SelectedItem.Index).Text & .ListItems(.SelectedItem.Index).ListSubItems(2).Text, sShortPath$, 260)
                        sShortPathName$ = mProcFunc.ftnStripNullChar(sShortPath$)
                        'MsgBox "Test"
                        SaveTrack .ListItems(.SelectedItem.Index).ListSubItems(2).Text, 0
                        CurrentFilePath = .ListItems(.SelectedItem.Index).ListSubItems(2).Text
                    End With
                    '---------------------------------------------------------------
                    
                    '1.Open File, 2.Set time format, 3.Play File. 4.Retrieve length of File.
                    mciSendString "open " & sShortPathName$ & " type MPEGVideo alias mp3", 0, 0, 0
                    mciSendString "set mp3 time format tmsf", 0, 0, 0
                    If Len(CurrentPosition) <= 0 Then
                        mciSendString "play mp3", 0, 0, 0
                    Else
                        mciSendString "play mp3 from " & CurrentPosition, 0, 0, 0
                        CurrentPosition = ""
                    End If
                    mciSendString "status mp3 length", sReturnBuffer$, Len(sReturnBuffer$), 0
                    mciSendString "setaudio mp3 volume to " & mVariables.iVolumeSetting, 0, 0, 0
                    lblLengthA.Caption = sReturnBuffer$
                    mVariables.lTrackLength = Val(sReturnBuffer)
                    '---------------------------------------------------------------
                    
                    'Update Controls.-----------------------------------------------
                    With lstFiles
                        lblArtistA.Caption = "(" & Mid(.ListItems(.SelectedItem.Index).ListSubItems(1).Text, 1, 2) & ") " & .ListItems(.SelectedItem.Index).ListSubItems(2).Text
                        .ListItems(.SelectedItem.Index).ListSubItems(3).Text = "P"
                    End With
                    '---------------------------------------------------------------
                    
                    'Set bTrackIsPlaying to TRUE.-----------------------------------
                    mVariables.bTrackIsPlaying = True
                    '---------------------------------------------------------------
                    
                    Timer1.Enabled = True
                    Me.MousePointer = 0
                    
                End If
                
            'Pause Track.-----------------------------------------------------------
            Case 2
                If lstFiles.ListItems.Count > 0 Then
                    mciSendString "pause mp3", 0, 0, 0
                    Timer1.Enabled = False
                    mVariables.bTrackIsPlaying = False
                End If
                    
            'Stop Track.------------------------------------------------------------
            Case 3
                '1.Disable Timer, 2.Stop mp3, 3.Close MCI Device.-------------------
'                MsgBox "Going to Stop"
                WMP1.Controls.Stop
                WebBrowser1.Stop
                Timer1.Enabled = False
                mciSendString "stop mp3", 0, 0, 0
                mciSendString "close all", 0, 0, 0
                '-------------------------------------------------------------------
                
                'Update Controls and Variables.-------------------------------------
                lblLengthA.Caption = "0"
                lblPositionA.Caption = "0"
                lblELTimeA.Caption = "0"
                lblELTimeB.Caption = "0"
                lblELTimeC.Caption = "0"
                lblELTimeD.Caption = "0"
                lblArtistA.Caption = "No Track loaded."
                ProgressBar.Cls
                mVariables.bTrackIsPlaying = False
                '-------------------------------------------------------------------
                
            'Next Track.------------------------------------------------------------
            Case 4
                If lstFiles.ListItems.Count > 0 Then
                    If lstFiles.SelectedItem.Index < lstFiles.ListItems.Count Then
                        
                        lstFiles.ListItems(lstFiles.SelectedItem.Index + 1).Selected = True
                        
                        'Stop and close mp3.----------------------------------------
                        Call imgControl_Click(3)
                        '-----------------------------------------------------------
                        
                        'Play mp3.--------------------------------------------------
                        Call imgControl_Click(1)
                        '-----------------------------------------------------------
                    
                    End If
                End If
                        
            'Open Track`s.----------------------------------------------------------
            Case 5
                Me.MousePointer = 11
                frmExplore.Show (1)
                Me.MousePointer = 0
                
            'Slow Play.-------------------------------------------------------------
            Case 6
                mciSendString "set mp3 speed 500", 0, 0, 0
                lblPlayMode(1).Caption = "Slow play"
            
            'Fast Play.-------------------------------------------------------------
            Case 7
                mciSendString "set mp3 speed 1500", 0, 0, 0
                lblPlayMode(1).Caption = "Fast play"
        
            'Mute.------------------------------------------------------------------
            Case 8
                'If a track is playing and mute is FALSE, then set mute to TRUE.----
                If mVariables.bTrackIsPlaying = True Then
                    If mVariables.bAudioAllOff = False Then
                        mciSendString "set mp3 audio all off", 0, 0, 0
                        mVariables.bAudioAllOff = True
                        speCommandLight(6).FillColor = vbRed
                    Else
                        mciSendString "set mp3 audio all on", 0, 0, 0
                        mVariables.bAudioAllOff = False
                    End If
                End If
                '-------------------------------------------------------------------
                    
            'Random.----------------------------------------------------------------
            Case 9
            'Set Rand light to red if Rand = TRUE and yellow if Rand = FALSE.-------
            If mVariables.bRandomSet = True Then
                mVariables.bRandomSet = False
                speCommandLight(7).FillColor = vbYellow
            Else
                mVariables.bRandomSet = True
                speCommandLight(7).FillColor = vbRed
            End If
            
            'Volume Increase.-------------------------------------------------------
            Case 10
                
                 mProcFunc.subSetVolume ("Increase")
                 Timer2.Enabled = True
            
            'Volume Decrease.-------------------------------------------------------
            Case 11
            
                mProcFunc.subSetVolume ("Decrease")
                Timer2.Enabled = True
            
            'Left hand image on progressbar.----------------------------------------
            Case 12
            
            'Right hand image on progressbar.---------------------------------------
            Case 13
            
            'Clear playlist.--------------------------------------------------------
            Case 14
                lstFiles.ListItems.Clear
                
            'Normal Play.-----------------------------------------------------------
            Case 15
                mciSendString "set mp3 speed 1000", 0, 0, 0
                lblPlayMode(1).Caption = "Normal play"
        
        End Select

'    End If

End Sub

'imgControl_MouseDown.
Private Sub imgControl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Timer4.Enabled = False
    Select Case Index
            
        'Previous Track.------------------------------------------------------------
        Case 0
            speCommandLight(0).FillColor = vbGreen
                    
        'Play Track.----------------------------------------------------------------
        Case 1
            speCommandLight(1).FillColor = vbGreen
        
        'Pause Track.---------------------------------------------------------------
        Case 2
            speCommandLight(2).FillColor = vbGreen
        
        'Stop Track.----------------------------------------------------------------
        Case 3
            speCommandLight(3).FillColor = vbGreen
        
        'Next Track.----------------------------------------------------------------
        Case 4
            speCommandLight(4).FillColor = vbGreen
        
        'Open Track`s.--------------------------------------------------------------
        Case 5
            speCommandLight(5).FillColor = vbGreen
        
        'Scan Back.-----------------------------------------------------------------
        Case 6
        
        'Scan Forward.--------------------------------------------------------------
        Case 7
                
        'Mute.----------------------------------------------------------------------
        Case 8
            
        'Random.--------------------------------------------------------------------
        Case 9
        
        'Volume Increase.-----------------------------------------------------------
        Case 10
        
        'Volume Decrease.-----------------------------------------------------------
        Case 11
        
        'Left hand image on progressbar.--------------------------------------------
        Case 12
        
        'Right hand image on progressbar.-------------------------------------------
        Case 13
        
        'Clear Playlist.------------------------------------------------------------
        Case 14
            speCommandLight(8).FillColor = vbGreen
            
        'Normal Play.---------------------------------------------------------------
        Case 15
    
    End Select

End Sub

'imgControl_MouseMove.
Private Sub imgControl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    mProcFunc.subSetLightColour ("speCommandLight")
    
    Select Case Index
                                                                                    
        'Previous Track.------------------------------------------------------------
        Case 0
            If lstFiles.ListItems.Count > 1 Then
                speCommandLight(0).FillColor = vbYellow
                mVariables.byCommandLight = 0
            End If
                        
        'Play Track.----------------------------------------------------------------
        Case 1
            If lstFiles.ListItems.Count > 0 Then
                speCommandLight(1).FillColor = vbYellow
                mVariables.byCommandLight = 1
            End If
            
        'Pause Track.---------------------------------------------------------------
        Case 2
            If mVariables.bTrackIsPlaying = True Then
                speCommandLight(2).FillColor = vbYellow
                mVariables.byCommandLight = 2
            End If
        
        'Stop Track.----------------------------------------------------------------
        Case 3
            If mVariables.bTrackIsPlaying = True Then
                speCommandLight(3).FillColor = vbYellow
                mVariables.byCommandLight = 3
            End If
        
        'Next Track.----------------------------------------------------------------
        Case 4
            If lstFiles.ListItems.Count > 1 Then
                speCommandLight(4).FillColor = vbYellow
                mVariables.byCommandLight = 4
            End If
    
        'Open Track`s.--------------------------------------------------------------
        Case 5
            speCommandLight(5).FillColor = vbYellow
            mVariables.byCommandLight = 5
        
        'Slow Play.-----------------------------------------------------------------
        Case 6
        
        'Fast Play.-----------------------------------------------------------------
        Case 7
        
        'Mute.----------------------------------------------------------------------
        Case 8
            'Set Mute light to red if Mute = TRUE and Green if Mute = FALSE.--------
            If mVariables.bTrackIsPlaying = True Then
                If mVariables.bAudioAllOff = True Then
                    speCommandLight(6).FillColor = vbRed
                Else
                    speCommandLight(6).FillColor = vbYellow
                End If
            End If
            '-----------------------------------------------------------------------
            mVariables.byCommandLight = 6
        
        'Random.--------------------------------------------------------------------
        Case 9
            'Set Rand light to red if Rand = TRUE and green if Rand = FALSE.--------
            If mVariables.bRandomSet = True Then
                speCommandLight(7).FillColor = vbRed
            Else
                speCommandLight(7).FillColor = vbYellow
            End If
            '-----------------------------------------------------------------------
            mVariables.byCommandLight = 7
            
        'Volume Increase.-----------------------------------------------------------
        Case 10
        
        'Volume Decrease.-----------------------------------------------------------
        Case 11
        
        'Left hand image on progressbar.--------------------------------------------
        Case 12
        
        'Right hand image on progressbar.-------------------------------------------
        Case 13
        
        'Clear Playlist.------------------------------------------------------------
        Case 14
            speCommandLight(8).FillColor = vbYellow
            mVariables.byCommandLight = 8
            
        'Normal Play.---------------------------------------------------------------
        Case 15
    
    End Select

End Sub


Private Sub imgControl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

If Index >= 10 And Index <= 11 Then
Dim DB As Database
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\Main.mdb")
Ssql = "Update Setup Set Volume='" & mVariables.iVolumeSetting & "'"
DB.Execute Ssql
DB.Close
End If
End Sub

'***********************************************************************************
'lstFiles Events.
'***********************************************************************************
'lstFiles_DblClick.
Private Sub lstFiles_DblClick()
    
    'Play selected File.------------------------------------------------------------
    Call imgControl_Click(3)
    Call imgControl_Click(1)
    '-------------------------------------------------------------------------------
    
End Sub


Private Sub mniExit_Click()
WebBrowser1.Stop
SaveTrack CurrentFilePath, lblPositionA.Caption
DoEvents
imgControl_Click 3
End
End Sub

Private Sub mniHide_Click()
Me.Visible = False
End Sub

Private Sub mniSetup_Click()
Setup.Show
End Sub

Private Sub Option1_Click()
WebBrowser1.Stop
WebBrowser1.Visible = False
WMP1.Visible = False
DoEvents
Dim DB As Database
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\main.mdb")
Ssql = "Update Setup Set Mode=1"
DB.Execute Ssql
DB.Close
WebBrowser1.Visible = False
lstFiles.Visible = True
DeciedPlayList
DoEvents
imgControl_Click 1
End Sub

Private Sub Option2_Click()
imgControl_Click 3
DoEvents
Dim DB As Database
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\main.mdb")
Ssql = "Update Setup Set Mode=2"
DB.Execute Ssql
DB.Close
WMP1.Visible = False
WebBrowser1.Visible = True
lstFiles.Visible = False
DeciedPlayList
End Sub

Private Sub Option3_Click()
imgControl_Click 3
DoEvents
Dim DB As Database
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\main.mdb")
Ssql = "Update Setup Set Mode=2"
DB.Execute Ssql
DB.Close
WebBrowser1.Visible = False
lstFiles.Visible = False
WMP1.Visible = True
DeciedPlayList

End Sub

Private Sub Option4_Click()
imgControl_Click 3
DoEvents
Dim DB As Database
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\main.mdb")
Ssql = "Update Setup Set Mode=2"
DB.Execute Ssql
DB.Close
WebBrowser1.Visible = False
lstFiles.Visible = False
WMP1.Visible = True
DeciedPlayList

End Sub

'***********************************************************************************
'ProgressBar Events.
'***********************************************************************************
'ProgressBar_Click.
Private Sub ProgressBar_Click()


End Sub


'ProgressBar_MouseDown.
Private Sub ProgressBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim lPosition As Long

    lPosition& = Val(lblLengthA.Caption) / ProgressBar.Width * x

    mciSendString "play mp3 from " & lPosition&, 0, 0, 0

    lPosition& = 0
        
    lstFiles.SetFocus

End Sub


'***********************************************************************************
'Update ProgressBar by passing SelCount of Files and LoopCounter.
'***********************************************************************************
Private Sub subPicScan(lMaxValue As Long, sPercent As Single)
    
    On Error Resume Next
    
    With ProgressBar
        .Cls
        .DrawMode = 13
        .CurrentX = .Width / 2 - .TextWidth("   ") / 2
        .CurrentY = .Height - 255
        'ProgressBar.Print Format(sPercent / lMaxValue * 100, "##,00") & " %"
        .DrawMode = 10
        ProgressBar.Line (-200, 30)-Step(.Width * sPercent \ lMaxValue + 200, 0), RGB(0, 0, 110), BF
        .Refresh
    End With

End Sub


Private Sub Timer1_Timer()
If Left(CurrentPlayType, 5) = "RADIO" Then Exit Sub
    'Retrieve position during the playing of an mp3 File.---------------------------
    mciSendString "status mp3 position", sReturnBuffer$, Len(sReturnBuffer$), 0
    lblPositionA.Caption = sReturnBuffer$
    SaveTrack CurrentFilePath, lblPositionA.Caption
    '-------------------------------------------------------------------------------

    'Update progressBar.------------------------------------------------------------
    subPicScan (mVariables.lTrackLength), (Val(sReturnBuffer$))
    '-------------------------------------------------------------------------------
    
    'Display elapsed time.----------------------------------------------------------
    lblELTimeD.Caption = Val(lblELTimeD.Caption) + 1
    If Val(lblELTimeD.Caption) > Val(9) Then
        lblELTimeC.Caption = Val(lblELTimeC.Caption) + 1
        lblELTimeD.Caption = "0"
    End If
    If lblELTimeC.Caption > 5 Then
        lblELTimeB.Caption = Val(lblELTimeB.Caption) + 1
        lblELTimeC.Caption = "0"
    End If
    '-------------------------------------------------------------------------------

    'It is possible to know if the Device is playing by using this MCI command.-----
    'Dim lReturn As Long, sReturn As String * 20
    'lReturn& = mciSendString("status mp3 mode", sReturn$, Len(sReturn$), hwnd)
    'MsgBox sReturn$. Returns playing if playing etc.
    '-------------------------------------------------------------------------------
    
    'If current position is => track length then update Variables, Controls.--------
    If Val(lblPositionA.Caption) >= Val(lblLengthA.Caption) Then
        mVariables.bTrackIsPlaying = False
        lblPositionA.Caption = "0"
        lblLengthA.Caption = "0"
                
        With lstFiles
            
            'If Random is FALSE then use this If Then statement.--------------------
            If .ListItems.Count > 1 And .SelectedItem.Index < .ListItems.Count And mVariables.bRandomSet = False Then
                
                .ListItems(.SelectedItem.Index + 1).Selected = True
                
                'Stop and close mp3.------------------------------------------------
                Call imgControl_Click(3)
                '-------------------------------------------------------------------
                DoEvents
                DeciedPlayList
                DoEvents
                
                'play mp3.----------------------------------------------------------
                Call imgControl_Click(1)
                '-------------------------------------------------------------------
            
            Else
                'Stop and close mp3.------------------------------------------------
                Call imgControl_Click(3)
                DoEvents
                DeciedPlayList
                DoEvents
                '-------------------------------------------------------------------
            End If
            
'            'If Random is TRUE then use this If Then statement.---------------------
            If mVariables.bRandomSet = True Then

                Dim iRandomSelect As Integer

                iRandomSelect% = mProcFunc.ftnRandomSelect

                .ListItems(iRandomSelect%).Selected = True
                
                'Stop and close mp3.------------------------------------------------
                Call imgControl_Click(3)
                '-------------------------------------------------------------------
                DoEvents
                DeciedPlayList
                DoEvents
                'play mp3.----------------------------------------------------------
                Call imgControl_Click(1)
                '-------------------------------------------------------------------

            End If
            
        End With
        
    End If

End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
'MsgBox mVariables.iVolumeSetting
Dim DB As Database
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\Main.mdb")
Ssql = "Update Setup Set Volume='" & mVariables.iVolumeSetting & "'"
DB.Execute Ssql
DB.Close
End Sub

Private Sub Timer3_Timer()
Dim R As Integer
Dim EndDate As Date
Dim StartDate As Date
SaveTrack CurrentFilePath, lblPositionA.Caption
For R = 2 To 7
    EndDate = DateAdd("n", 20, MyTimes(R))
    
    If MyTimes(R) <= Now And EndDate >= Now Then
        
        
        DoEvents
        If R = 2 And CurrentFilePathAzan <> "Azan1.mp3" Then
            imgControl_Click 3
            Label2.Caption = "Fajar Time"
            CurrentFilePathAzan = "Azan1.mp3"
            PlayAzan "Azan1.mp3"
            lblArtistA.Caption = ""
            
            Exit Sub
        End If
        If R = 3 And CurrentFilePathAzan <> "Azan2.mp3" Then
'            MsgBox Weekday(Date, vbFriday)
            If Weekday(Date, vbFriday) <> 1 Then
                imgControl_Click 3
                DoEvents
                Label2.Caption = "Duhar Time"
                CurrentFilePathAzan = "Azan2.mp3"
                PlayAzan "Azan2.mp3"
                lblArtistA.Caption = ""
                Exit Sub
            Else
                imgControl_Click 3
                DoEvents
                Label2.Caption = "Juma Prayer Time"
                CurrentFilePathAzan = "Azan2.mp3"
                PlayAzan "Azan1.mp3"
                lblArtistA.Caption = ""
                Exit Sub
            End If
            

        End If
        If R = 4 And CurrentFilePathAzan <> "Azan3.mp3" Then
            imgControl_Click 3
            Label2.Caption = "Asr Time"
            CurrentFilePathAzan = "Azan3.mp3"
            PlayAzan "Azan3.mp3"
            lblArtistA.Caption = ""
        End If
        If R = 5 And CurrentFilePathAzan <> "Azan4.mp3" Then
            imgControl_Click 3
            Label2.Caption = "Maghrib Time"
            CurrentFilePathAzan = "Azan4.mp3"
            PlayAzan "Azan4.mp3"
            lblArtistA.Caption = ""
        End If
        If R = 6 And CurrentFilePathAzan <> "Azan5.mp3" Then
            imgControl_Click 3
            Label2.Caption = "Isha Time"
            CurrentFilePathAzan = "Azan5.mp3"
            PlayAzan "Azan5.mp3"
            lblArtistA.Caption = ""
        End If
'        If R = 7 And CurrentFilePathAzan <> "Azan6.mp3" Then
'            If Weekday(Date, vbFriday) = vbFriday Then
'                imgControl_Click 3
'                Label2.Caption = "Juma Prayer Time"
'                CurrentFilePathAzan = "Azan6.mp3"
'                PlayAzan "Azan6.mp3"
'                lblArtistA.Caption = ""
'            End If
'        End If
        Exit Sub
    End If
Next R
If MyTimes(0) > Now Then
        imgControl_Click 2
        lblArtistA.Caption = ""
        'lblArtist.Caption = ""
        CurrentPlayType = ""
        Label2.Caption = "It's Stoppage Time!"
        'MsgBox "Test"
        Exit Sub
ElseIf MyTimes(1) < Now Then
        imgControl_Click 2
        lblArtistA.Caption = ""
        'lblArtist.Caption = ""
        CurrentPlayType = ""
        Label2.Caption = "It's Stoppage Time!"
        
        Exit Sub
End If
  
DeciedPlayList
DoEvents
imgControl_Click 1
  
End Sub

Private Sub Timer4_Timer()
  mVariables.iVolumeSetting = Val(oldVolume)
    
    If Val(oldVolume) > 498 Then
        mProcFunc.subSetVolume ("Increase")
    ElseIf Val(oldVolume) < 498 Then
        mProcFunc.subSetVolume ("Decrease")
        
    End If
End Sub

Private Sub Timer5_Timer()
Dim J As String
J = Left(Label3.Caption, 1)
Label3.Caption = Right(Label3.Caption, Len(Label3.Caption) - 1) & J

End Sub

