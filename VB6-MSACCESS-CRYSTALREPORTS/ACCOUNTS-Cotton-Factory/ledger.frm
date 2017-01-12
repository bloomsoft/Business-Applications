VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ledger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "ledger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6135
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   2880
      Picture         =   "ledger.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2850
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1095
      Left            =   1440
      Picture         =   "ledger.frx":325C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2850
      Width           =   1335
   End
   Begin Crystal.CrystalReport r1 
      Left            =   4320
      Top             =   2850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
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
   Begin ComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3975
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2490
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   5895
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   1980
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Rate Wise"
         Height          =   255
         Left            =   1290
         TabIndex        =   19
         Top             =   1980
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5460
         TabIndex        =   18
         Text            =   "7"
         Top             =   450
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1290
         TabIndex        =   16
         Top             =   1620
         Width           =   4395
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Only Checked for GST Invoice"
         Height          =   315
         Left            =   3120
         TabIndex        =   15
         Top             =   1950
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   1305
         TabIndex        =   2
         Top             =   1260
         Width           =   1395
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60358659
         CurrentDate     =   36611
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60358659
         CurrentDate     =   36611
      End
      Begin VB.Label Label6 
         Caption         =   "Aging Interval"
         Height          =   255
         Left            =   4410
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   375
         TabIndex        =   14
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select A/c"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom_r
Private Blm1 As New bloom1

Private Sub comb_type()
    Combo2.clear
        Combo2.AddItem "Sale"
        Combo2.ItemData(Combo2.NewIndex) = 1
        Combo2.AddItem "Purchase"
        Combo2.ItemData(Combo2.NewIndex) = 2
    Combo2.ListIndex = 0
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex > -1 Then
    Text2.Text = Val(Text2.Text)
End If
End Sub

Private Sub Command1_Click()

Dim F As String
Dim l As String
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then
sb1.SimpleText = "Creating Ledger For " & Text3.Text
Blm.ledger Val(Text2.Text), p1
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = Blm.report_path & "LEDGER1.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 2 Then
sb1.SimpleText = "Creating Ledger For " & Text3.Text
Blm.ledger2 Val(Text2.Text), p1, Date1.Value, Date2.Value
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = Blm.report_path & "LEDGER4.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"

r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")

r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 3 Then
sb1.SimpleText = "Creating Ledger For " & Text3.Text
Blm.ledger3 p1, Date1.Value, Date2.Value
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = Blm.report_path & "LEDGER3.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text1.Text) = 4 Then
sb1.SimpleText = "Creating Cash Book For " & Text3.Text
Blm.CashBook p1, Date1.Value, Date2.Value, Text2.Text
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = Blm.report_path & "CashBook.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If


If Val(Text1.Text) = 5 Then
    r1.ReportFileName = Blm.report_path & "PurchasesParty.rpt"
    F = "{Purchaseview.v_date} in Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    F = F & " and {Purchaseview.Seller}=" & Text2.Text
    If Check1.Value = 1 Then
        F = F & " and {PurchaseView.GSTINV}=1"
    End If
    If Check2.Value = 1 Then
        F = F & " and {PurchaseView.Rate}=" & Val(Text5.Text)
    End If
    r1.DataFiles(0) = Blm1.patHmain
    r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    r1.SelectionFormula = F
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
End If

If Val(Text1.Text) = 6 Then
    r1.ReportFileName = Blm.report_path & "SalesParty.rpt"
    F = "{Saleview.v_date} in Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
    F = F & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
    F = F & " and {Saleview.Seller}=" & Text2.Text
    If Check1.Value = 1 Then
        F = F & " and {SaleView.GSTINV}=1"
    End If
    If Check2.Value = 1 Then
        F = F & " and {SaleView.Rate}=" & Val(Text5.Text)
    End If
    
    r1.DataFiles(0) = Blm1.patHmain
    r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")
    r1.SelectionFormula = F
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
End If

If Val(Text1.Text) = 7 Then
sb1.SimpleText = "Creating Ledger For " & Text3.Text
Blm.ledger2 Val(Text2.Text), p1, Date1.Value, Date2.Value, Val(Text4.Text)
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = Blm.report_path & "LEDGER44.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"

r1.ReportTitle = "From : " & Format(Date1.Value, "dd-MMM-yyyy") & " To : " & Format(Date2.Value, "dd-MMM-yyyy")

r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If


Screen.MousePointer = vbDefault
'Combo1.ListIndex = 0

End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()


'Me.Icon = LoadPicture(blm1.report_path & "earth.ico")
Dim Ssql As String
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2




Date1.Value = FStartDate
Date2.Value = Date

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text2.Text = SelectedAccountCode
    Text3.Text = SelectedAccountName
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

Private Sub Text2_Validate(Cancel As Boolean)
Dim R As Long

Text3.Text = Blm1.party1(Val(Text2.Text))
End Sub
