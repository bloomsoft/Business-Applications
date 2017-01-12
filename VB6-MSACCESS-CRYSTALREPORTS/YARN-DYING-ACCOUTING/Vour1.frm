VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form vour1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Preview"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4335
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
      Top             =   3120
      Width           =   4095
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
         Picture         =   "Vour1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1275
      End
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
         Picture         =   "Vour1.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin Crystal.CrystalReport r1 
         Left            =   3240
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   1320
         Width           =   1815
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
         Top             =   2160
         Width           =   855
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
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Give the Voucher No's For Preview"
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
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "vour1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New Bloom_r
Private Blm As New bloom1
Private Sub Command1_Click()
Dim f As String
Dim b As Boolean
Screen.MousePointer = vbHourglass

    If Val(Text3.Text) = 1 Then
        f = "{YarnIssueVW.ISSUE_NO} >= " & Val(Text1.Text)
        'f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        f = f & " and {YarnIssueVW.ISSUE_NO} <= " & Val(Text2.Text)
        r1.ReportFileName = App.Path & "\Reports\YarnIssue.rpt"
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 2 Then
        f = "{ClothRecVW.Rec_NO} >= " & Val(Text1.Text)
        'f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        f = f & " and {ClothRecVW.Rec_NO} <= " & Val(Text2.Text)
        r1.ReportFileName = App.Path & "\Reports\KoraRec.rpt"
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 3 Then
        f = "{PackingVW.Vou_NO} >= " & Val(Text1.Text)
        'f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        f = f & " and {PackingVW.Vou_NO} <= " & Val(Text2.Text)
        r1.ReportFileName = App.Path & "\Reports\LotPacking.rpt"
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 4 Then
        f = "{PaymentLoomVW.Vou_NO} >= " & Val(Text1.Text)
        'f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        f = f & " and {PaymentLoomVW.Vou_NO} <= " & Val(Text2.Text)
        r1.ReportFileName = App.Path & "\Reports\VouLoom.rpt"
        r1.DataFiles(0) = Blm.pathMain
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 5 Then
        f = "{PaymentDyingVW.Vou_NO} >= " & Val(Text1.Text)
        'f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        f = f & " and {PaymentDyingVW.Vou_NO} <= " & Val(Text2.Text)
        r1.ReportFileName = App.Path & "\Reports\VouDying.rpt"
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

Private Sub Form_Load()
'Text1.SetFocus
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub
