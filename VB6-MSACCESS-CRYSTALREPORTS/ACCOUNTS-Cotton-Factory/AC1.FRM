VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ac1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lists"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "ac1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   600
      TabIndex        =   3
      Top             =   1860
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   609
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1455
      Left            =   2160
      Picture         =   "ac1.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin Crystal.CrystalReport r1 
      Left            =   3000
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1455
      Left            =   600
      Picture         =   "ac1.frx":325C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "ac1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As bloom_r
Private Blm1 As New bloom1





Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then

R1.ReportFileName = Blm.report_path & "chart.rpt"

R1.DataFiles(1) = Blm1.SettingsPath
R1.ReportTitle = Blm1.orgname
R1.DataFiles(0) = Blm1.patHmain
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 2 Then
R1.ReportFileName = Blm.report_path & "item1.rpt"
R1.ReportTitle = Blm1.orgname
R1.DataFiles(0) = Blm1.patHmain
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 3 Then
R1.ReportFileName = Blm.report_path & "Addresses.rpt"
R1.ReportTitle = Blm1.orgname
R1.DataFiles(0) = Blm1.patHmain
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 4 Then
    
    Blm.PurchaseJobLedger ProgressBar1
    R1.ReportFileName = Blm.report_path & "PJobLedger.rpt"
    R1.DataFiles(0) = App.path & "\Book.mdb"
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1

End If

If Val(Text1.Text) = 5 Then

R1.ReportFileName = Blm.report_path & "empchart.rpt"

R1.DataFiles(1) = Blm1.SettingsPath
R1.ReportTitle = Blm1.orgname
R1.DataFiles(0) = Blm1.patHmain
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

Screen.MousePointer = vbDefault

End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Set Blm = New bloom_r
'Me.Icon = LoadPicture(blm.report_path & "earth.ico")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Blm = Nothing
End Sub
