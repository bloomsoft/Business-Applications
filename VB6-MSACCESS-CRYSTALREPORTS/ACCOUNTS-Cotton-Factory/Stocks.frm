VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Stocks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Wise Stocks"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   3285
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2160
      Top             =   2400
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
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   2640
      Picture         =   "Stocks.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1095
      Left            =   705
      Picture         =   "Stocks.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1785
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   3660
         TabIndex        =   5
         Top             =   1260
         Visible         =   0   'False
         Width           =   315
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56492035
         CurrentDate     =   37366
      End
      Begin VB.Label Label1 
         Caption         =   "End Date"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Stocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New bloom_r
Private Blm1 As New bloom1
Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then
'Blmr.ClosingStocks date1.Value,Date2 ProgressBar1
'R1.ReportFileName = Blmr.report_path & "Stocks.Rpt"
'R1.ReportTitle = Blm1.orgname
'R1.DataFiles(0) = App.path & "\Book.mdb"
'R1.WindowTop = 0
'R1.WindowLeft = 0
'R1.WindowState = crptMaximized
'R1.Action = 1
'R1.PageZoom 100
End If

If Val(Text1.Text) = 2 Then
Blmr.PurchaseJobs date1.Value, ProgressBar1
R1.ReportFileName = Blmr.report_path & "PurJobs.Rpt"
R1.ReportTitle = Blm1.orgname
R1.DataFiles(0) = App.path & "\Book.mdb"
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
R1.PageZoom 100
End If

If Val(Text1.Text) = 3 Then
Blmr.SaleJobs date1.Value, ProgressBar1
R1.ReportFileName = Blmr.report_path & "SaleJobs.Rpt"
R1.ReportTitle = Blm1.orgname
R1.DataFiles(0) = App.path & "\Book.mdb"
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
R1.PageZoom 100
End If

If Val(Text1.Text) = 4 Then
Blmr.Aging date1.Value, ProgressBar1, StatusBar1
R1.ReportFileName = Blmr.report_path & "Aging.Rpt"
R1.ReportTitle = Blm1.orgname
R1.DataFiles(0) = App.path & "\Book.mdb"
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
R1.PageZoom 100
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me

End Sub

Private Sub Form_Load()
date1.Value = Date
End Sub
