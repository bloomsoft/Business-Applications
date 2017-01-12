VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PeriodicRep 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   1065
         TabIndex        =   13
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         TabIndex        =   12
         Top             =   600
         Width           =   2835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Emp Name"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "Emp Code"
         Height          =   255
         Left            =   135
         TabIndex        =   14
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   360
      TabIndex        =   8
      Top             =   990
      Visible         =   0   'False
      Width           =   3915
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1740
         TabIndex        =   10
         Top             =   150
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Ref. No."
         Height          =   225
         Left            =   840
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   330
      TabIndex        =   7
      Top             =   3120
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   345
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker Date1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   60555267
      CurrentDate     =   39098
   End
   Begin MSComCtl2.DTPicker Date2 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   60555267
      CurrentDate     =   39098
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2760
      Top             =   2790
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PeriodicRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Blm As New bloom_r
Private Sub Command1_Click()
Dim F As String
If Val(Text1.Text) = 1 Then
    R1.ReportFileName = Blm.report_path & "Purchases.rpt"
    F = "{Purchaseview.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    R1.DataFiles(0) = Blm1.patHmain
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If
If Val(Text1.Text) = 2 Then
    R1.ReportFileName = Blm.report_path & "Sales.rpt"
    F = "{Saleview.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    R1.DataFiles(0) = Blm1.patHmain
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 3 Then
    Blm.AtdRegister date1.Value, Date2.Value, ProgressBar1
    R1.ReportFileName = Blm.report_path & "AtdReg.rpt"
    R1.DataFiles(0) = App.path & "\Book.mdb"
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 4 Then
    F = "{SalVoucher.SDate}=Date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
    F = F & " and {SalVoucher.EDate}=Date(" & Format(Date2.Value, "yyyy,MM,dd") & ")"
    R1.ReportFileName = Blm.report_path & "SalSheet.rpt"
    R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 5 Then
    F = "{Vou_View.V_Date} in Date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
    F = F & " to Date(" & Format(Date2.Value, "yyyy,MM,dd") & ")"
    R1.ReportFileName = Blm.report_path & "AdvSummary.rpt"
    R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 6 Then
'Blm.MonthlyProduction Date1.Value, Date2.Value, ProgressBar1, Val(Text2.Text)
'Blm.BankBalances Date2.Value
Blm.DailyProduction date1.Value, Date2.Value, , Val(Text2.Text)
    R1.ReportFileName = Blm.report_path & "DailyProduction.rpt"
    R1.DataFiles(0) = App.path & "\Book.mdb"
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 7 Then
    Blm.MonthlyPurchases date1.Value, Date2.Value, ProgressBar1
    Blm.BankBalances Date2.Value
    R1.ReportFileName = Blm.report_path & "MonthlyPurchases.rpt"
    R1.DataFiles(0) = App.path & "\Book.mdb"
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 8 Then
    R1.ReportFileName = Blm.report_path & "PurchasesSummary.rpt"
    F = "{Purchaseview.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    R1.DataFiles(0) = Blm1.patHmain
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If
If Val(Text1.Text) = 9 Then
    R1.ReportFileName = Blm.report_path & "SalesSummary.rpt"
    F = "{Saleview.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    R1.DataFiles(0) = Blm1.patHmain
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 10 Then

F = "{OverTimeVW.ADate}  in Date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
F = F & " To Date(" & Format(Date2.Value, "yyyy,MM,dd") & ")"
If Len(Text4.Text) > 0 Then
    F = F & " and {OverTimeVW.AcCode}=" & Val(Text4.Text)
End If
R1.ReportFileName = Blm.report_path & "OverTime.rpt"
R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
R1.SelectionFormula = F
R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1

End If

If Val(Text1.Text) = 11 Then

    R1.ReportFileName = Blm.report_path & "DayBook.rpt"
    F = "{vou_view.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    R1.DataFiles(0) = Blm1.patHmain
    R1.ReportTitle = "Day Book" & vbCrLf & "From : " & Format(date1.Value, "dd-MMM-yyyy") & " " & "To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 12 Then

    R1.ReportFileName = Blm.report_path & "OverAllPurchasesGST.rpt"
    F = "{Purchaseview.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    R1.DataFiles(0) = Blm1.patHmain
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " " & "To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 13 Then

    R1.ReportFileName = Blm.report_path & "OverAllSalesGST.rpt"
    F = "{Saleview.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    R1.DataFiles(0) = Blm1.patHmain
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " " & "To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 14 Then
    F = "{Vou_View.V_Date} in Date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
    F = F & " to Date(" & Format(Date2.Value, "yyyy,MM,dd") & ")"
    If Len(Text4.Text) > 0 Then
        F = F & " and {Vou_View.Party}=" & Val(Text4.Text)
    End If
    R1.ReportFileName = Blm.report_path & "AdvanceBook.rpt"
    R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If

If Val(Text1.Text) = 15 Then
    F = "{Vou_View.V_Date} in Date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
    F = F & " to Date(" & Format(Date2.Value, "yyyy,MM,dd") & ")"
    F = F & " and {Vou_View.V_type}=" & SelectedVType
    If SelectedVType = 4 Then
        R1.ReportFileName = Blm.report_path & "PurchaseBook.rpt"
        R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    ElseIf SelectedVType = 5 Then
        R1.ReportFileName = Blm.report_path & "SaleBook.rpt"
        R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    ElseIf SelectedVType = 15 Then
        R1.ReportFileName = Blm.report_path & "SalBook.rpt"
        R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    ElseIf SelectedVType = 18 Then
        R1.ReportFileName = Blm.report_path & "TranferBook.rpt"
        R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    ElseIf SelectedVType = 20 Then
        R1.ReportFileName = Blm.report_path & "IssueBook.rpt"
        R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    ElseIf SelectedVType = 21 Then
        R1.ReportFileName = Blm.report_path & "ProdBook.rpt"
        R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    Else
        R1.ReportFileName = Blm.report_path & "DayBook.rpt"
        R1.ReportTitle = SelectedReportTitle
        R1.ReportTitle = R1.ReportTitle & vbCrLf & "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    End If
    R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
    
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
'    MsgBox "Test"
End If


If Val(Text1.Text) = 16 Then

Blm.DailyIssueHeadWise date1.Value, Date2.Value, , Val(Text2.Text)

R1.ReportFileName = Blm.report_path & "DailyIssueHeadWise.rpt"
R1.DataFiles(0) = App.path & "\Book.mdb"
R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
R1.SelectionFormula = F
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 17 Then
R1.ReportFileName = Blm.report_path & "Expences.rpt"
F = "{Expences.Edate} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
If Val(Text2.Text) > 0 Then
F = F & " and {Expences.RefNo}=" & Val(Text2.Text)
End If
R1.DataFiles(0) = Blm1.patHmain
R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
R1.SelectionFormula = F
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 18 Then
R1.ReportFileName = Blm.report_path & "DailyIssue.rpt"
F = "{issue_view.v_date} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
If Val(Text2.Text) > 0 Then
    F = F & " and {Issue_View.RefNo}=" & Text2.Text
End If
R1.DataFiles(0) = Blm1.patHmain
R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
R1.SelectionFormula = F
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 19 Then

R1.ReportFileName = Blm.report_path & "Arrivals.rpt"
F = "{Arrivals.Adate} in Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
If Val(Text2.Text) > 0 Then F = F & " and {Arrivals.RefNo}=" & Text2.Text

R1.DataFiles(0) = Blm1.patHmain
R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
R1.SelectionFormula = F
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 20 Then

R1.ReportFileName = Blm.report_path & "Dispatches.rpt"
F = "{Dispatches.Ddate} = Date(" & date1.Year & ", " & date1.Month & ", " & date1.Day & ")"
F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
If Val(Text2.Text) > 0 Then F = F & " and {Dispatches.RefNo}=" & Text2.Text
R1.DataFiles(0) = Blm1.patHmain
R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
R1.SelectionFormula = F
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1
End If

If Val(Text1.Text) = 21 Then
'Blm.MonthlyProduction Date1.Value, Date2.Value, ProgressBar1, Val(Text2.Text)
'Blm.BankBalances Date2.Value
Blm.DailyProduction date1.Value, Date2.Value, , Val(Text2.Text)
    R1.ReportFileName = Blm.report_path & "DailyProduction2.rpt"
    R1.DataFiles(0) = App.path & "\Book.mdb"
    R1.ReportTitle = "From : " & Format(date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    R1.SelectionFormula = F
    R1.WindowTop = 0
    R1.WindowLeft = 0
    R1.WindowState = crptMaximized
    R1.Action = 1
End If
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
date1.Value = FStartDate
Date2.Value = Date
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text4.Text = SelectedAccountCode
    Text3.Text = SelectedAccountName
End If

End Sub
