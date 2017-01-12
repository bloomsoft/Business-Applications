VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Stock1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks of All Items"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   135
      TabIndex        =   9
      Top             =   3075
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   3390
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   635
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
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
      Left            =   225
      Picture         =   "Stock1.frx":0561
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1905
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
         Format          =   53018627
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
         Format          =   53018627
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
Blmr.day_due2_TEMP DTPicker1.Value, ProgressBar1, StatusBar1
r1.ReportFileName = App.Path & "\recable.rpt"
r1.DataFiles(0) = App.Path & "\Book.mdb"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 6 Then
Blmr.shortfallPurchase DTPicker1.Value, ProgressBar1
r1.ReportFileName = App.Path & "\ShortFall.rpt"
r1.DataFiles(0) = App.Path & "\Book.mdb"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 7 Then
Blmr.shortfallSale DTPicker1.Value, ProgressBar1
r1.ReportFileName = App.Path & "\ShortFall.rpt"
r1.DataFiles(0) = App.Path & "\Book.mdb"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
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
