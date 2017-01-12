VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form book 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books Print"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   180
      TabIndex        =   6
      Top             =   1110
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Ref. No."
         Height          =   315
         Left            =   2310
         TabIndex        =   9
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Voucher No."
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   2280
      Picture         =   "book.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin Crystal.CrystalReport r1 
      Left            =   3840
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1095
      Left            =   960
      Picture         =   "book.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1905
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3720
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1110
         TabIndex        =   2
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60162051
         CurrentDate     =   36770
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom_r
Private Blm1 As New bloom1

Private Sub Command1_Click()
Dim F As String
If Val(Text1.Text) = 1 Then

r1.ReportFileName = Blm.report_path & "jb.rpt"
F = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 2 Then

r1.ReportFileName = Blm.report_path & "bb.rpt"
F = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If
If Val(Text1.Text) = 3 Then

r1.ReportFileName = Blm.report_path & "cb.rpt"
F = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 4 Then

r1.ReportFileName = Blm.report_path & "sjb.rpt"
F = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 5 Then

r1.ReportFileName = Blm.report_path & "pjb.rpt"
F = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 6 Then

r1.ReportFileName = Blm.report_path & "DailyIssue.rpt"
F = "{issue_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
F = F & " and {Issue_View.VNo}=" & Text2.Text
F = F & " and {Issue_View.RefNo}=" & Text3.Text
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 7 Then

r1.ReportFileName = Blm.report_path & "Arrivals.rpt"
F = "{Arrivals.Adate} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
F = F & " and {Arrivals.VNo}=" & Text2.Text
F = F & " and {Arrivals.RefNo}=" & Text3.Text

r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 8 Then

r1.ReportFileName = Blm.report_path & "Dispatches.rpt"
F = "{Dispatches.Ddate} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
F = F & " and {Dispatches.VNo}=" & Text2.Text
F = F & " and {Dispatches.RefNo}=" & Text3.Text
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 9 Then

r1.ReportFileName = Blm.report_path & "Expences.rpt"
F = "{Expences.Edate} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
F = F & " and {Expences.VNo}=" & Text2.Text
F = F & " and {Expences.RefNo}=" & Text3.Text
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 10 Then

r1.ReportFileName = Blm.report_path & "GeneralData.rpt"
'f = "{Expences.Edate} = Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.SelectionFormula = F
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If


If Val(Text1.Text) = 11 Then
    
End If
End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Date1.Value = Date

End Sub
