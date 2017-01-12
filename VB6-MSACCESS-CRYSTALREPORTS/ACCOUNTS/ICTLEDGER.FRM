VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form ICTLedger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICT Ledger"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport R1 
      Left            =   4080
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ProgressBar P1 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Prev"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Select Ref#"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "ICTLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New bloom_r
Private Blm1 As New bloom1
Private Sub Command1_Click()


Blmr.ICTLedger Combo1.Text, P1
DoEvents

R1.ReportFileName = Blmr.report_path & "LEDGER1ICT.rpt"
R1.DataFiles(0) = App.path & "\Book.mdb"
R1.ReportTitle = Blm1.orgname
R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.PrintReport


End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Dim ssql As String
Dim Tb As Recordset
Dim db As Database
ssql = "Select Distinct ICT from voudtl"
Set db = OpenDatabase(Blm1.patHmain)
Set Tb = db.OpenRecordset(ssql)
If Not Tb.EOF Then
    Combo1.clear
        Do While Not Tb.EOF
            Combo1.AddItem Tb.Fields("ICT").Value & ""
            Tb.MoveNext
        
        Loop
        Combo1.ListIndex = 0
End If
Tb.Close
db.Close

End Sub
