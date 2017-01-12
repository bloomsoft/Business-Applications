VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form vour 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher Print"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "VOUR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   4416
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gst Invoice Info"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   4416
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   44236803
         CurrentDate     =   37103
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Sr #"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin ComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1950
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1950
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   8
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
      Begin VB.Label Label1 
         Caption         =   "Voucher No."
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport R1 
      Left            =   4032
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport R2 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "vour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom_r
Private Blm1 As New bloom1
Private Sub invoice()
Dim db As Database
Dim tb As Recordset

Set db = OpenDatabase(Blm1.pathMain)
db.Execute "Delete from GST"
Set tb = db.OpenRecordset("GST", dbOpenTable)
tb.AddNew
    tb.Fields("SR_no").Value = Val(Text3.Text)
    tb.Fields("Sr_Date").Value = date1.Value
    tb.Fields("Out_no").Value = Val(Text1.Text)
tb.Update
tb.Close
db.Close
End Sub
Private Sub Command1_Click()
Dim f As String
r1.LogOnServer "pdsodbc.dll", "Knitting", "beq-local", "BLOOMNEW", "mlb"

Screen.MousePointer = vbHourglass
If Val(Text2.Text) = 1 Then
r1.ReportFileName = blm.report_path & "inward1.rpt"
r1.DataFiles(0) = "BLOOMNEW.INWARD_VIEW"
r1.DataFiles(1) = "BLOOMNEW.YARN"
f = "{inward_view.in_no}= " & Val(Text1.Text)
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If
If Val(Text2.Text) = 11 Then
'r1.ReportFileName = blm.report_path & "in2.rpt"
'r1.DataFiles(0) = blm.Userpath
'f = "{in_view_final.a.a.p_no}= " & Val(Text1.Text)
'r1.SelectionFormula = f
'r1.WindowTop = 0
'r1.WindowLeft = 0
'r1.WindowState = crptMaximized
''r1.PrintReport
'r1.Action = 1
'r1.PageZoom 100

End If


If Val(Text2.Text) = 2 Then

'r1.ReportFileName = blm.report_path & "out1.rpt"
''r1.DataFiles(0) = blm.Userpath
'f = "{outward_view.out_no}= " & Val(Text1.Text)
'r1.SelectionFormula = f
'r1.WindowTop = 0
'r1.WindowLeft = 0
'r1.WindowState = crptMaximized
''r1.PrintReport
'r1.Action = 1
r1.PageZoom 100

r1.ReportFileName = blm.report_path & "Inward2.rpt"
r1.DataFiles(0) = "BLOOMNEW.OUTWARD_VIEW"
f = "{outward_view.out_no}= " & Val(Text1.Text)
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If

If Val(Text2.Text) = 22 Then
'r1.ReportFileName = blm.report_path & "sale_2.rpt"
'r1.DataFiles(0) = blm.Userpath
'f = "{sale_view_final.a.a.sale_no}= " & Val(Text1.Text)
'r1.SelectionFormula = f
'r1.WindowTop = 0
'r1.WindowLeft = 0
'r1.WindowState = crptMaximized
''r1.PrintReport
'r1.Action = 1
'r1.PageZoom 100

End If
If Val(Text2.Text) = 3 Then
r1.ReportFileName = blm.report_path & "out3.rpt"
r1.DataFiles(0) = "BLOOMNEW.OUTWARD_VIEW"
f = "{outward_view.out_no}= " & Val(Text1.Text)
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If
If Val(Text2.Text) = 33 Then
'r1.ReportFileName = blm.report_path & "sent2.rpt"
'r1.DataFiles(0) = blm.Userpath
'f = "{sent_view_final.a.a.issue_no}= " & Val(Text1.Text)
'r1.SelectionFormula = f
'r1.WindowTop = 0
'r1.WindowLeft = 0
'r1.WindowState = crptMaximized
''r1.PrintReport
'r1.Action = 1
'r1.PageZoom 100

End If

If Val(Text2.Text) = 4 Then
'invoice
r1.ReportFileName = blm.report_path & "out2.rpt"

r1.DataFiles(0) = "BLOOMNEW.OUTWARD_VIEW"
f = "{outward_view.out_no}= " & Val(Text1.Text)
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

'r2.ReportFileName = blm.report_path & "GSTINV.rpt"

'r1.DataFiles(0) = blm.Userpath
'f = "{outward_view.out_no}= " & Val(Text1.Text)
'r2.SelectionFormula = f
'r2.WindowTop = 0
'r2.WindowLeft = 0
'r2.WindowState = crptMaximized
'r1.PrintReport
'r2.Action = 1
'r1.PageZoom 100

End If

If Val(Text2.Text) = 5 Then
'r1.ReportFileName = blm.report_path & "Inward2.rpt"
''r1.DataFiles(0) = blm.Userpath
'f = "{inward_view.in_no}= " & Val(Text1.Text)
'r1.SelectionFormula = f
'r1.WindowTop = 0
'r1.WindowLeft = 0
'r1.WindowState = crptMaximized
''r1.PrintReport
'r1.Action = 1
'r1.PageZoom 150
r1.ReportFileName = blm.report_path & "out1.rpt"
r1.DataFiles(0) = "BLOOMNEW.INWARD_VIEW"
r1.DataFiles(1) = "BLOOMNEW.YARN"
r1.DataFiles(2) = "BLOOMNEW.YARN"

f = "{inward_view.in_no}= " & Val(Text1.Text)
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100
End If


If Val(Text2.Text) = 6 Then
r1.ReportFileName = blm.report_path & "Inward3.rpt"
r1.DataFiles(0) = "BLOOMNEW.INWARD_VIEW"
f = "{inward_view.in_no}= " & Val(Text1.Text)
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
'r1.PrintReport
r1.Action = 1
r1.PageZoom 100

End If

If Val(Text2.Text) = 51 Then
    blm.cont_summery Val(Text1.Text), 1
    r1.ReportFileName = blm.report_path & "Cont_1_Sum.rpt"
    r1.DataFiles(0) = Blm1.pathMain
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.ReportTitle = "Knitting Purchase Contract Summary"
    r1.Action = 1
  r1.PageZoom 100
End If

If Val(Text2.Text) = 52 Then
    blm.cont_summery Val(Text1.Text), 2
    r1.ReportFileName = blm.report_path & "Cont_1_Sum.rpt"
    r1.DataFiles(0) = Blm1.pathMain
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.ReportTitle = "Knitting Sale Contract Summary"
    r1.Action = 1
    r1.PageZoom 100
End If

If Val(Text2.Text) = 53 Then
    blm.Cont_Sum Val(Text1.Text), 3, p1
    r1.ReportFileName = blm.report_path & "Cont_3_Sum.rpt"
    r1.DataFiles(0) = blm.path
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If

If Val(Text2.Text) = 54 Then
    blm.Cont_Sum_Party Combo1.ItemData(Combo1.ListIndex), 1, p1
    r1.ReportFileName = blm.report_path & "Cont_1_Sumold.rpt"
    r1.DataFiles(0) = blm.path
    
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If

If Val(Text2.Text) = 55 Then
    blm.Cont_Sum_Party Combo1.ItemData(Combo1.ListIndex), 2, p1
    r1.ReportFileName = blm.report_path & "Cont_2_Sum.rpt"
    r1.DataFiles(0) = blm.path
    
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    'r1.PageZoom 150
End If

If Val(Text2.Text) = 56 Then
    blm.Cont_Sum_Party Combo1.ItemData(Combo1.ListIndex), 3, p1
    r1.ReportFileName = blm.report_path & "Cont_3_Sum.rpt"
    r1.DataFiles(0) = blm.path
    
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If

If Val(Text2.Text) = 61 Then
    r1.ReportFileName = blm.report_path & "Cont_1.rpt"
    r1.DataFiles(0) = "BLOOMNEW.CONT_VIEW"
    
    f = "{cont_view.cont_no} = " & Val(Text1.Text)
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End If


If Val(Text2.Text) = 62 Then
    r1.ReportFileName = blm.report_path & "Cont_2.rpt"
    r1.DataFiles(0) = "BLOOMNEW.CONT_VIEW"
    
    f = "{cont_view.cont_no} = " & Val(Text1.Text)
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
   r1.PageZoom 100
End If

If Val(Text2.Text) = 63 Then
    r1.ReportFileName = blm.report_path & "Cont_3.rpt"
    r1.DataFiles(0) = "BLOOMNEW.CONT_VIEW"
    f = "{cont_view.cont_no} = " & Val(Text1.Text)
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
 r1.PageZoom 100
End If

If Val(Text2.Text) = 64 Then
    r1.ReportFileName = blm.report_path & "needlein.rpt"
    r1.DataFiles(0) = "BLOOMNEW.NEEDLEIN_VW"
    f = "{NEEDLEIN_VW.INNO} =" & Val(Text1.Text)
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
 r1.PageZoom 100
End If

If Val(Text2.Text) = 65 Then
    r1.ReportFileName = blm.report_path & "needleout.rpt"
    r1.DataFiles(0) = "BLOOMNEW.NEEDLEOUT_VW"
    f = "{NEEDLEOUT_VW.OUTNO} =" & Val(Text1.Text)
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
 r1.PageZoom 100
End If

If Val(Text2.Text) = 66 Then
    r1.ReportFileName = blm.report_path & "issueyarn.rpt"
    r1.DataFiles(0) = "BLOOMNEW.YARNISSUE_VW"
    f = "{YARNISSUE_VW.IN_NO} =" & Val(Text1.Text)
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
 r1.Action = 1
 r1.PageZoom 100
End If

If Val(Text2.Text) = 67 Then

    R2.ReportFileName = blm.report_path & "ins.rpt"
    R2.DataFiles(0) = "BLOOMNEW.INS_VW"
    f = "{INS_VW.INO} =" & Val(Text1.Text)
    R2.SelectionFormula = f
    R2.WindowTop = 0
    R2.WindowLeft = 0
    R2.WindowState = crptMaximized
    R2.Action = 1
 R2.PageZoom 100

    r1.ReportFileName = blm.report_path & "receivefabric.rpt"
    r1.DataFiles(0) = "BLOOMNEW.FABRICRECEIVE_VW"
    f = "{FABRICRECEIVE_VW.OUT_NO} =" & Val(Text1.Text)
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
 r1.PageZoom 100

End If

If Val(Text2.Text) = 68 Then

    R2.ReportFileName = blm.report_path & "ins.rpt"
    R2.DataFiles(0) = "BLOOMNEW.INS_VW"
    f = "{INS_VW.INO} =" & Val(Text1.Text)
    R2.SelectionFormula = f
    R2.WindowTop = 0
    R2.WindowLeft = 0
    R2.WindowState = crptMaximized
    R2.Action = 1
 R2.PageZoom 100

End If


'Text1.Text = vbNullString
'Text1.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If

End Sub

Private Sub Form_Load()
Dim ssql As String
ssql = "select * from Acchart Order by Name"
Blm1.fill_comb ssql, Combo1, "Name", "COde"
End Sub

