VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form DayRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day Book"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3960
   Begin VB.CommandButton Command1 
      Caption         =   "&Prev"
      Height          =   975
      Left            =   150
      Picture         =   "DayRep.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1095
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   975
      Left            =   2490
      Picture         =   "DayRep.frx":03F5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1095
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   270
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   50397187
      CurrentDate     =   39125
   End
   Begin Crystal.CrystalReport r1 
      Left            =   1785
      Top             =   1095
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   210
      Left            =   285
      TabIndex        =   0
      Top             =   315
      Width           =   1200
   End
End
Attribute VB_Name = "DayRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New Bloom_r
Private Sub Command1_Click()
    Blmr.DayBook DTPicker1.Value
    r1.ReportFileName = App.Path & "\DayBook.rpt"
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.WindowTop = 0
    'r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub
