VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form ac1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lists"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "ac1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
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
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
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
Private blm As bloom_r
Private blm1 As New bloom1





Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then

r1.ReportFileName = blm.report_path & "chart.rpt"

r1.DataFiles(1) = blm1.SettingsPath
r1.ReportTitle = blm1.orgname
r1.DataFiles(0) = blm1.patHmain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 2 Then
r1.ReportFileName = blm.report_path & "item1.rpt"
r1.ReportTitle = blm1.orgname
r1.DataFiles(0) = blm1.patHmain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 3 Then
r1.ReportFileName = blm.report_path & "Addresses.rpt"
r1.ReportTitle = blm1.orgname
r1.DataFiles(0) = blm1.patHmain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

Screen.MousePointer = vbDefault

End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Set blm = New bloom_r
'Me.Icon = LoadPicture(blm.report_path & "earth.ico")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub
