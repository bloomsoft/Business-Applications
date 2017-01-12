VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form itmledger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Wise Ledger"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport R1 
      Left            =   4080
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ProgressBar P1 
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1095
      Left            =   720
      Picture         =   "itmledger.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   2640
      Picture         =   "itmledger.frx":0CE2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker Date2 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19660803
         CurrentDate     =   37366
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19660803
         CurrentDate     =   37366
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Item"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "itmledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Blmr As New bloom_r
Private ReportFile As New CRAXDRT.Report

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Blmr.ItemLedger Date1.Value, Date2.Value, Combo1.ItemData(Combo1.ListIndex), P1
R1.ReportFileName = Blmr.report_path & "ItmLedger.Rpt"
R1.DataFiles(0) = App.path & "\Book.Mdb"
R1.ReportTitle = Blm1.orgname
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
R1.PageZoom 100

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Dim ssql As String

ssql = "Select * from Items Order By Name"
Blm1.fill_comb ssql, Combo1, "Name", "Code"
Date1.Value = Date
Date2.Value = Date
End Sub
