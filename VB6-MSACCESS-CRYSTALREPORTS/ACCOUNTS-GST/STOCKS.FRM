VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Stocks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Wise Stocks"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport r1 
      Left            =   2160
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
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
      Left            =   720
      Picture         =   "Stocks.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
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
         Format          =   24707075
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
Private blmr As New bloom_r
Private blm1 As New bloom1
Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
blmr.OverAllStockNew Date1.Value

r1.ReportFileName = blmr.report_path & "StockOverAll.Rpt"
r1.ReportTitle = blm1.orgname
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me

End Sub

Private Sub Form_Load()
Date1.Value = Date
End Sub
