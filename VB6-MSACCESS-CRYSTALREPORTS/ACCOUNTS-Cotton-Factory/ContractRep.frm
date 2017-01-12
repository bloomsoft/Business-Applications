VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ContractRep 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6330
   Begin VB.CheckBox Check3 
      Caption         =   "Contract No"
      Height          =   315
      Left            =   480
      TabIndex        =   11
      Top             =   1020
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2250
      TabIndex        =   10
      Top             =   600
      Width           =   3525
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1290
      TabIndex        =   9
      Top             =   600
      Width           =   885
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2250
      TabIndex        =   8
      Top             =   270
      Width           =   3525
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1290
      TabIndex        =   7
      Top             =   270
      Width           =   885
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Item"
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   540
      Width           =   1065
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Party"
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   210
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   1050
      Width           =   915
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   1170
      TabIndex        =   4
      Top             =   2130
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2610
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1140
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   615
      Left            =   1185
      TabIndex        =   1
      Top             =   1425
      Width           =   1455
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2760
      Top             =   2730
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
End
Attribute VB_Name = "ContractRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Blm As New bloom_r
Private Sub Command1_Click()
Dim F As String
If Val(Text1.Text) = 1 Then
    If Check3.Value = 1 Then
        Blm.PurchaseJobLedger ProgressBar1, Val(Text2.Text)
    ElseIf Check1.Value = 1 And Check2.Value = 0 Then
        Blm.PurchaseJobLedger ProgressBar1, , Val(Text3.Text)
    ElseIf Check1.Value = 0 And Check2.Value = 1 Then
        Blm.PurchaseJobLedger ProgressBar1, , , Val(Text5.Text)
    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
        Blm.PurchaseJobLedger ProgressBar1, , Val(Text3.Text), Val(Text5.Text)
    ElseIf Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 Then
        Blm.PurchaseJobLedger ProgressBar1
    End If
    
    r1.ReportFileName = Blm.report_path & "PJobLedger.rpt"
    r1.DataFiles(0) = App.path & "\Book.mdb"
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
End If

If Val(Text1.Text) = 2 Then
    If Check3.Value = 1 Then
        Blm.SaleJobLedger ProgressBar1, Val(Text2.Text)
    ElseIf Check1.Value = 1 And Check2.Value = 0 Then
        Blm.SaleJobLedger ProgressBar1, , Val(Text3.Text)
    ElseIf Check1.Value = 0 And Check2.Value = 1 Then
        Blm.SaleJobLedger ProgressBar1, , , Val(Text5.Text)
    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
        Blm.SaleJobLedger ProgressBar1, , Val(Text3.Text), Val(Text5.Text)
    ElseIf Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 Then
        Blm.SaleJobLedger ProgressBar1
    End If
    r1.ReportFileName = Blm.report_path & "SJobLedger.rpt"
    r1.DataFiles(0) = App.path & "\Book.mdb"
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
End If

End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text3.Text = SelectedAccountCode
    Text4.Text = SelectedAccountName
End If

End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    'Combo1.SetFocus
    Load Search1
    Search1.Show vbModal
    Text5.Text = SelectedItemCode
    Text6.Text = SelectedItemName
End If

End Sub
