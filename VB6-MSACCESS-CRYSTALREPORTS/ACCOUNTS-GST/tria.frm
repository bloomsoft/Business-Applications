VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form trial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trial Balance"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "tria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3225
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
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
   Begin ComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "With Sub Heads Totals"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   1035
         Left            =   2520
         Picture         =   "tria.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   36771
      End
      Begin Crystal.CrystalReport r1 
         Left            =   480
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         DiscardSavedData=   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Preview"
         Height          =   1035
         Left            =   960
         Picture         =   "tria.frx":325C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   36619
      End
      Begin Crystal.CrystalReport r2 
         Left            =   480
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         DiscardSavedData=   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   135
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   15
      End
   End
End
Attribute VB_Name = "trial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blmr As bloom_r
Private Blm1 As New bloom1

Private Sub Command1_Click()
Dim u As Integer
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then
blmr.trial date1.Value, P1
If Check1.Value <> 1 Then
    r1.ReportFileName = blmr.report_path & "trial.rpt"
Else
    r1.ReportFileName = blmr.report_path & "trialNew.rpt"
End If
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport

End If

If Val(Text1.Text) = 3 Then
'blmr.day_due date1.Value
r1.ReportFileName = blmr.report_path & "payable.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport
End If

If Val(Text1.Text) = 4 Then
'blmr.day_due2 date1.Value
r1.ReportFileName = blmr.report_path & "recable.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport
End If

If Val(Text1.Text) = 5 Then
blmr.open_trial P1
r1.ReportFileName = blmr.report_path & "trial2.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport

End If

If Val(Text1.Text) = 6 Then
blmr.trial date1.Value, P1
r1.ReportFileName = blmr.report_path & "trial4.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport

End If

If Val(Text1.Text) = 7 Then
blmr.trial2 Date2.Value, date1.Value, P1
r1.ReportFileName = blmr.report_path & "trial3.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport

End If

If Val(Text1.Text) = 8 Then
blmr.day_due2_TEMP date1.Value, P1, sb1
r1.ReportFileName = blmr.report_path & "recable.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport

End If

If Val(Text1.Text) = 9 Then
blmr.day_due_TEMP date1.Value, P1, sb1
r1.ReportFileName = blmr.report_path & "payable.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport

End If

If Val(Text1.Text) = 10 Then
blmr.ProfitLossState date1.Value, P1
blmr.BalanceSheetState date1.Value, P1
r1.ReportFileName = blmr.report_path & "PL7.rpt"
r1.DataFiles(0) = App.path & "\Finance.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1

r2.ReportFileName = blmr.report_path & "BaL7.rpt"
r1.DataFiles(0) = App.path & "\Finance.mdb"
r2.ReportTitle = Blm1.orgname
r2.WindowTop = 0
r2.WindowLeft = 0
r2.WindowState = crptMaximized
r2.Action = 1

End If

If Val(Text1.Text) = 11 Then
blmr.PeriodicStockNew Date2.Value, date1.Value
r1.ReportFileName = blmr.report_path & "PeriodicStock.Rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100

End If


Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Set blmr = New bloom_r
'Me.Icon = LoadPicture(blmr.report_path & "earth.ico")
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Date2.Value = Date - 30
date1.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blmr = Nothing
End Sub

