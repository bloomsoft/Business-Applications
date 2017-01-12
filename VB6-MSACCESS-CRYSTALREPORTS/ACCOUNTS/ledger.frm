VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ledger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "ledger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6135
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   2880
      Picture         =   "ledger.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1095
      Left            =   1440
      Picture         =   "ledger.frx":325C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
   Begin Crystal.CrystalReport r1 
      Left            =   4320
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3735
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
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1320
         Width           =   4455
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
         Format          =   20709379
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
         Format          =   20709379
         CurrentDate     =   36611
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
         Top             =   1320
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
Private blm As New bloom_r
Private Blm1 As New bloom1

Private Sub comb_type()
    Combo2.clear
        Combo2.AddItem "Sale"
        Combo2.ItemData(Combo2.NewIndex) = 1
        Combo2.AddItem "Purchase"
        Combo2.ItemData(Combo2.NewIndex) = 2
    Combo2.ListIndex = 0
End Sub
Private Sub Command1_Click()

Dim f As String
Dim l As String
Screen.MousePointer = vbHourglass
l = Mid(Combo1.ItemData(Combo1.ListIndex), 1, 5)
    If ledgerhide = False Then
        
        If l = "11002" Then
        'If Then
        Else
        If l = "11004" Then
        Else
            MsgBox "You Don't Have Rights to View this Ledger..."
            Screen.MousePointer = vbDefault
            Exit Sub
        
        End If
        End If
    End If
If Val(Text1.Text) = 1 Then
sb1.SimpleText = "Creating Ledger For " & Combo1.Text
blm.ledger Combo1.ItemData(Combo1.ListIndex), P1
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = blm.report_path & "LEDGER1.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 2 Then
sb1.SimpleText = "Creating Ledger For " & Combo1.Text
blm.ledger2 Combo1.ItemData(Combo1.ListIndex), P1, date1.Value, Date2.Value
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = blm.report_path & "LEDGER4.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 3 Then
sb1.SimpleText = "Creating Ledger For " & Combo1.Text
blm.ledger3 Combo1.ItemData(Combo1.ListIndex), P1, date1.Value, Date2.Value
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = blm.report_path & "LEDGER3.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

If Val(Text1.Text) = 4 Then
sb1.SimpleText = "Creating Cash Book For " & Combo1.Text
blm.CashBook P1, date1.Value, Date2.Value
DoEvents
sb1.SimpleText = "Printing the Ledger"
r1.ReportFileName = blm.report_path & "CashBook.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.ReportTitle = Blm1.orgname
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If




Screen.MousePointer = vbDefault
Combo1.ListIndex = 0

End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()


'Me.Icon = LoadPicture(blm1.report_path & "earth.ico")
Dim ssql As String
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

ssql = "select * from acchart order by name"
Blm1.fill_comb ssql, Combo1, "name", "code"



date1.Value = Date
Date2.Value = Date

End Sub

