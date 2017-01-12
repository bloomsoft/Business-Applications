VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PeriodicRep 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4680
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
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1200
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
      Format          =   20709379
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
      Format          =   20709379
      CurrentDate     =   39098
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2760
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
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
Dim f As String
If Val(Text1.Text) = 1 Then

r1.ReportFileName = Blm.report_path & "Purchases.rpt"
f = "{Purchaseview.v_date} in Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
f = f & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If
If Val(Text1.Text) = 2 Then

r1.ReportFileName = Blm.report_path & "Sales.rpt"
f = "{Saleview.v_date} in Date(" & Date1.Year & ", " & Date1.Month & ", " & Date1.Day & ")"
f = f & " To Date(" & Date2.Year & ", " & Date2.Month & ", " & Date2.Day & ")"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.PrintReport
End If

End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub
