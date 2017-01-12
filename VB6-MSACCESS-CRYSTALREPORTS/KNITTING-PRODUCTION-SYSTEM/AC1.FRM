VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form ac1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lists"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "AC1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
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
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   1335
      Left            =   1200
      Picture         =   "AC1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1695
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
Private blm2 As bloom1

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
r1.LogOnServer "pdsodbc.dll", "Knitting", "beq-local", "BLOOMNEW", "mlb"

If Val(Text1.Text) = 1 Then
r1.ReportFileName = blm.report_path & "chart.rpt"
r1.DataFiles(0) = "BLOOMNEW.ACCHART"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If
If Val(Text1.Text) = 2 Then

r1.ReportFileName = blm.report_path & "Cloth.rpt"
r1.DataFiles(0) = "BLOOMNEW.CLOTH"

r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If

If Val(Text1.Text) = 3 Then
r1.ReportFileName = blm.report_path & "Yarn.rpt"
r1.DataFiles(0) = "BLOOMNEW.YARN"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If

If Val(Text1.Text) = 4 Then
r1.ReportFileName = blm.report_path & "Machine.rpt"
r1.DataFiles(0) = "BLOOMNEW.MACHINE"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If

If Val(Text1.Text) = 5 Then
r1.ReportFileName = blm.report_path & "Emp1.rpt"
r1.DataFiles(0) = "BLOOMNEW.EMP1"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If

If Val(Text1.Text) = 6 Then
r1.ReportFileName = blm.report_path & "ContD.rpt"
r1.DataFiles(0) = "BLOOMNEW.CONT_VIEW"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 7 Then
r1.ReportFileName = App.path & "\parts.rpt"
r1.DataFiles(0) = "BLOOMNEW.PARTS"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100
End If


Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Set blm = New bloom_r
Set blm2 = New bloom1

'Me.Icon = LoadPicture(blm.report_path & "earth.ico")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub
