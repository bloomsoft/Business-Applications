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
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1095
      Left            =   960
      Picture         =   "book.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
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
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   54853635
         CurrentDate     =   36770
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom_r
Private Blm1 As New bloom1

Private Sub Command1_Click()
Dim f As String
If Val(Text1.Text) = 1 Then

r1.ReportFileName = blm.report_path & "jb.rpt"
f = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 2 Then

r1.ReportFileName = blm.report_path & "bb.rpt"
f = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If
If Val(Text1.Text) = 3 Then

r1.ReportFileName = blm.report_path & "cb.rpt"
f = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 4 Then

r1.ReportFileName = blm.report_path & "sjb.rpt"
f = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 5 Then

r1.ReportFileName = blm.report_path & "pjb.rpt"
f = "{vou_view.v_date} = Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Date1.Value = Date

End Sub
