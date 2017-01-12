VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Stock1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks of All Items"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport r1 
      Left            =   2520
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   1095
      Left            =   3000
      Picture         =   "Stock1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Prev"
      Height          =   1095
      Left            =   240
      Picture         =   "Stock1.frx":0561
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Date"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   37412
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   37412
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Stock1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New Bloom_r
Private Blm1 As New bloom1
Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then
    Blmr.CreateStock DTPicker1.Value
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.ReportFileName = App.Path & "\StocksQty.rpt"
    r1.WindowTop = 0
    r1.ReportTitle = "To : " & Format(DTPicker1.Value, "dd/MM/yyyy")
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100

End If
If Val(Text1.Text) = 2 Then
Blmr.CreateStockBetweenDates DTPicker1.Value, DTPicker2.Value
    r1.ReportFileName = App.Path & "\StocksQtyPeriod.rpt"
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.WindowTop = 0
    r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100

End If
If Val(Text1.Text) = 3 Then
Blmr.TrialBalance DTPicker1.Value
    r1.ReportFileName = App.Path & "\Trial.rpt"
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.WindowTop = 0
    r1.ReportTitle = "To : " & Format(DTPicker1.Value, "dd/MM/yyyy")
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100

End If
If Val(Text1.Text) = 4 Then
Blmr.CreateStock DTPicker1.Value
    r1.ReportFileName = App.Path & "\StocksQtyValues.rpt"
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.WindowTop = 0
    r1.ReportTitle = "To : " & Format(DTPicker1.Value, "dd/MM/yyyy")
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100

End If


If Val(Text1.Text) = 5 Then
Blmr.DayBook DTPicker1.Value
r1.ReportFileName = App.Path & "\DayBook.rpt"
r1.DataFiles(0) = App.Path & "\Book.mdb"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100

End If

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub
