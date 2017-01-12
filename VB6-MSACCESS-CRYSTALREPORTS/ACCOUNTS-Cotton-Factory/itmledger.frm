VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form itmledger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Wise Ledger"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4140
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2415
      Visible         =   0   'False
      Width           =   555
   End
   Begin Crystal.CrystalReport R1 
      Left            =   4080
      Top             =   2880
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
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2880
         TabIndex        =   14
         Top             =   330
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1110
         TabIndex        =   12
         Top             =   750
         Width           =   3105
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1110
         TabIndex        =   0
         Top             =   345
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker Date2 
         Height          =   375
         Left            =   1110
         TabIndex        =   6
         Top             =   1575
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59965443
         CurrentDate     =   37366
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   1110
         TabIndex        =   4
         Top             =   1140
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59965443
         CurrentDate     =   37366
      End
      Begin VB.Label Label5 
         Caption         =   "Ref #"
         Height          =   255
         Left            =   2250
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Item Code"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1575
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Item Name"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   765
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

Private Sub Combo1_Click()
If Combo1.ListIndex > -1 Then
    Text2.Text = Text2.Text
End If
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then
Blmr.ItemLedger Date1.Value, Date2.Value, Text2.Text, p1
r1.ReportFileName = Blmr.report_path & "ItmLedger.Rpt"
r1.DataFiles(0) = App.path & "\Book.Mdb"
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If
If Val(Text1.Text) = 2 Then
r1.ReportFileName = Blmr.report_path & "DailyIssue2.rpt"
F = "{issue_view.v_date} in Date(" & Format(Date1.Value, "yyyy,MM,dd") & ") To Date(" & Format(Date2.Value, "yyyy,MM,dd") & ")"
If Val(Text2.Text) > 0 Then F = F & " and {Issue_View.ItemCode}=" & Text2.Text
If Val(Text4.Text) > 0 Then F = F & " and {Issue_View.RefNo}=" & Text4.Text
r1.SelectionFormula = F
'MsgBox F
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1

'Blmr.ItemWiseIssue Text2.Text, Date1.Value, Date2.Value, P1, Val(Text4.Text)
'R1.ReportFileName = Blmr.report_path & "ItemIssue.Rpt"
'R1.DataFiles(0) = App.path & "\Book.Mdb"
'R1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
'R1.WindowTop = 0
'R1.WindowLeft = 0
'R1.WindowState = crptMaximized
'R1.Action = 1
'R1.PageZoom 100
End If

If Val(Text1.Text) = 3 Then
    r1.ReportFileName = Blmr.report_path & "PurchasesItem.rpt"
    F = "{Purchaseview.v_date} in Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    F = F & " and {Purchaseview.Item}=" & Text2.Text
    r1.DataFiles(0) = Blm1.patHmain
    r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    r1.SelectionFormula = F
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
End If

If Val(Text1.Text) = 4 Then
    r1.ReportFileName = Blmr.report_path & "SalesItem.rpt"
    F = "{SaleView.v_date} in Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    F = F & " and {Saleview.Item}=" & Text2.Text
    r1.DataFiles(0) = Blm1.patHmain
    r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    r1.SelectionFormula = F
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
End If

If Val(Text1.Text) = 5 Then
Blmr.ItemLedgerNew Date1.Value, Date2.Value, Text2.Text, p1
r1.ReportFileName = Blmr.report_path & "ItmLedgerNew.Rpt"
r1.DataFiles(0) = App.path & "\Book.Mdb"
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 6 Then

    Blmr.DailyProduction Date1.Value, Date2.Value, , Val(Text4.Text), Val(Text2.Text)
    r1.ReportFileName = Blmr.report_path & "DailyProduction.rpt"
    r1.DataFiles(0) = App.path & "\Book.mdb"
    r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    r1.SelectionFormula = F
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
Date1.Value = FStartDate
Date2.Value = Date
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    'Combo1.SetFocus
    Load Search1
    Search1.Show vbModal
    Text2.Text = SelectedItemCode
    Text3.Text = SelectedItemName
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

