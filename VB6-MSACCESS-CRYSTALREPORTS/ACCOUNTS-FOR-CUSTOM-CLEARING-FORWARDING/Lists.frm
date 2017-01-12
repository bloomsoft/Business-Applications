VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Lists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lists"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin Crystal.CrystalReport r1 
         Left            =   1920
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         DiscardSavedData=   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   975
         Left            =   2640
         Picture         =   "Lists.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Prev"
         Height          =   975
         Left            =   480
         Picture         =   "Lists.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Select a City"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Lists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Blm1 As New bloom1
Private blmr As New Bloom_r

Private Sub Command1_Click()
Dim f As String
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then
    r1.ReportFileName = App.Path & "\Chart.Rpt"
    r1.DataFiles(0) = App.Path & "\Bloom.mdb"
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If

If Val(Text1.Text) = 2 Then
    r1.ReportFileName = App.Path & "\Item.Rpt"
    r1.DataFiles(0) = App.Path & "\Bloom.mdb"
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If

If Val(Text1.Text) = 6 Then
    r1.ReportFileName = App.Path & "\Item.Rpt"
    r1.DataFiles(0) = App.Path & "\Bloom.mdb"
    f = "{Item_Vw.Code}=" & Combo1.ItemData(Combo1.ListIndex)
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

Private Sub Form_Activate()
Dim ssql As String
If Val(Text1.Text) = 3 Then
        ssql = "Select * from City Order By Name"
        Blm1.fill_comb ssql, Combo1, "Name", "Code"
End If

If Val(Text1.Text) = 6 Then
        ssql = "Select * from Groups Order By Name"
        Blm1.fill_comb ssql, Combo1, "Name", "Code"
End If

End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
End Sub
