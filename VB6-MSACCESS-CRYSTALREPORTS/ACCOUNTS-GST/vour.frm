VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form vour 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher Print"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "vour.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport r1 
      Left            =   3840
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   2400
      Picture         =   "vour.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   1095
      Left            =   960
      Picture         =   "vour.frx":325C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   4455
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "To Voucher No."
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "From Voucher No."
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "vour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom_r
Private Blm1 As New bloom1
Private Function CheckGST(VNo As Long) As Boolean
Dim RS As Recordset
Dim DB As Database
Dim S As String

Set DB = OpenDatabase(App.path & "\Bloom.mdb")
S = "Select * from Sales where V_No=" & VNo & " and STRate>0"
Set RS = DB.OpenRecordset(S)
If Not RS.EOF Then
    CheckGST = True
Else
    CheckGST = False
End If
RS.Close
DB.Close
End Function
Private Sub Command1_Click()
Dim f As String
Dim B As Boolean
Dim PF As Parameter
Screen.MousePointer = vbHourglass
If Val(Text2.Text) = 1 Then
r1.ReportFileName = blm.report_path & "jv.rpt"
f = "{vou_view.v_no} in " & Val(Text1.Text)
f = f & " to " & Val(Text3.Text)

r1.DataFiles(0) = Blm1.patHmain
r1.SelectionFormula = f
r1.ReportTitle = Blm1.orgname

r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport
End If
If Val(Text2.Text) = 2 Then
r1.ReportFileName = blm.report_path & "bv.rpt"
f = "{vou_view.v_no} in " & Val(Text1.Text)
f = f & " to " & Val(Text3.Text)
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport
End If

If Val(Text2.Text) = 3 Then
r1.ReportFileName = blm.report_path & "cv.rpt"
f = "{vou_view.v_no} in " & Val(Text1.Text)
f = f & " to " & Val(Text3.Text)
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport
End If

If Val(Text2.Text) = 4 Then
'PF.Name = "Address"
B = CheckGST(Val(Text1.Text))
If B = True Then
    r1.ReportFileName = blm.report_path & "SJVGST.rpt"
Else
    r1.ReportFileName = blm.report_path & "SJV.rpt"
End If
r1.DataFiles(0) = Blm1.patHmain
r1.SelectionFormula = "{SaleView.V_no} = " & Text1.Text
r1.ReportTitle = Blm1.orgname
r1.ParameterFields(0) = "Address;" & Blm1.orgAddress & ";true"
r1.ParameterFields(1) = "Phone;" & Blm1.orgPhone & ";true"
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

If Val(Text2.Text) = 5 Then
r1.ReportFileName = blm.report_path & "pjv.rpt"
r1.ReportTitle = Blm1.orgname
f = "{Purchaseview.v_no} in " & Val(Text1.Text)
f = f & " to " & Val(Text3.Text)
r1.DataFiles(0) = Blm1.patHmain
'r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport
End If

If Val(Text2.Text) = 6 Then
r1.ReportFileName = blm.report_path & "Cjv.rpt"
r1.ReportTitle = Blm1.orgname
f = "{Consumeview.v_no} in " & Val(Text1.Text)
f = f & " to " & Val(Text3.Text)
r1.DataFiles(0) = Blm1.patHmain
'r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport
End If

If Val(Text2.Text) = 7 Then
r1.ReportFileName = blm.report_path & "SJob.rpt"
r1.DataFiles(0) = Blm1.patHmain
r1.ReportTitle = Blm1.orgname
r1.ParameterFields(0) = "Address;" & Blm1.orgAddress & ";true"
r1.ParameterFields(1) = "Phone;" & Blm1.orgPhone & ";true"

r1.SelectionFormula = "{SaleView.V_no} = " & Text1.Text
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
