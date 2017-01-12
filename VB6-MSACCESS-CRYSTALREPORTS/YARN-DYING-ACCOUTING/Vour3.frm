VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form vour3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4275
   Begin VB.Frame Frame2 
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2400
         Picture         =   "Vour3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   360
         Picture         =   "Vour3.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin Crystal.CrystalReport r1 
         Left            =   3360
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   1200
         TabIndex        =   2
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   60096515
         CurrentDate     =   39506
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1200
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   60096515
         CurrentDate     =   39506
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Give Starting and Last Date For Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "vour3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New Bloom_r
Private Blm As New bloom1
Private Sub Command1_Click()
Dim f As String
Screen.MousePointer = vbHourglass
    If Val(Text1.Text) = 1 Then
        f = "{YarnIssueVW.ISSUE_DATE} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\Reports\YarnIssue.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text1.Text) = 2 Then
        f = "{ClothRecVW.DATE_RECIEVE} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\Reports\KoraRec.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text1.Text) = 3 Then
        f = "{PACKINGVW.DATE} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\Reports\LOTPACKING.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text1.Text) = 4 Then
        f = "{PaymentLoomVW.DATE} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\Reports\VouLoom.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text1.Text) = 5 Then
        f = "{PaymentDyingVW.DATE} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\Reports\VouDying.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date - 30
DTPicker2.Value = Date

Me.Top = 10
Me.Left = 10
End Sub
