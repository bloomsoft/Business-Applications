VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Ledger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ledgers"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4215
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "(F2) to Search Party"
            TextSave        =   "(F2) to Search Party"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "(F3) to Search Item"
            TextSave        =   "(F3) to Search Item"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Item Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2760
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   855
         Left            =   3240
         Picture         =   "Ledger.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   1800
         Picture         =   "Ledger.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Prev"
         Height          =   855
         Left            =   360
         Picture         =   "Ledger.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Account Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   4335
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Account Title"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Account Code"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dates"
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56557571
         CurrentDate     =   37415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56557571
         CurrentDate     =   37415
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blmr As New Bloom_r
Private Blm1 As New bloom1
Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Val(Text3.Text) = 1 Then
    
    blmr.LedgerBetweenDates DTPicker1.Value, DTPicker2.Value, Val(Text1.Text)
    r1.ReportFileName = App.Path & "\Ledger.rpt"
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.WindowTop = 0
    r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If

If Val(Text3.Text) = 2 Then
    
    blmr.ItemLedger DTPicker1.Value, DTPicker2.Value, Val(Text6.Text)
    r1.ReportFileName = App.Path & "\ItemLedger.rpt"
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.WindowTop = 0
    r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
    
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
    Search2.Text3.Text = 3
    Search2.Show
End If
If KeyCode = vbKeyF3 Then
    Search1.Text3.Text = 4
    Search1.Show
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    Text2.Text = Blm1.party1(Val(Text1.Text))
    If Text2.Text = "NOT" Then
        MsgBox "Invalid Account Code Press (F2) to Select Account"
        Cancel = True
    End If
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Val(Text6.Text) > 0 Then
    Text5.Text = Blm1.Item1(Val(Text6.Text))
End If
End Sub
