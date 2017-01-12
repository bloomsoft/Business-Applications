VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form vou_r 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Books"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "VOU_R.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4695
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   240
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   4230
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1170
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   315
         Width           =   2865
      End
      Begin VB.Label Label6 
         Caption         =   "Select Fabric"
         Height          =   336
         Left            =   132
         TabIndex        =   18
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   4230
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1170
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   315
         Width           =   2865
      End
      Begin VB.Label Label5 
         Caption         =   "Select Yarn"
         Height          =   330
         Left            =   225
         TabIndex        =   16
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   225
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   4230
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1170
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   315
         Width           =   2865
      End
      Begin VB.Label Label4 
         Caption         =   "Select Party"
         Height          =   330
         Left            =   225
         TabIndex        =   14
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1710
      Left            =   240
      TabIndex        =   8
      Top             =   855
      Width           =   4248
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         Height          =   345
         Left            =   1440
         TabIndex        =   3
         Top             =   1245
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   330
         Left            =   1155
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20381699
         CurrentDate     =   36986
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   330
         Left            =   1155
         TabIndex        =   1
         Top             =   270
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20381699
         CurrentDate     =   36611
      End
      Begin Crystal.CrystalReport r1 
         Left            =   -144
         Top             =   36
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
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   315
         Width           =   615
      End
   End
   Begin ComctlLib.ProgressBar p1 
      Height          =   270
      Left            =   255
      TabIndex        =   7
      Top             =   2640
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   492
      Left            =   0
      TabIndex        =   6
      Top             =   2988
      Width           =   4692
      _ExtentX        =   8281
      _ExtentY        =   873
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
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   -630
      TabIndex        =   10
      Top             =   495
      Width           =   855
   End
End
Attribute VB_Name = "vou_r"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1

Private Sub Command1_Click()
Dim blm As New bloom_r
Dim f As String
r1.LogOnServer "pdsodbc.dll", "Knitting", "beq-local", "BLOOMNEW", "mlb"
Screen.MousePointer = vbHourglass
If Val(Text1.Text) = 1 Then
r1.ReportFileName = blm.report_path & "inward1d.rpt"
r1.DataFiles(0) = "BLOOMNEW.INWARD_VIEW"
f = "{inward_view.in_date} in Date (" & Format(date1.Value, "yyyy,MM,dd") & ")"
f = f & " to Date (" & Format(date2.Value, "yyyy,MM,dd") & ")"
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100

'r1.PrintReport
End If
If Val(Text1.Text) = 2 Then
'r1.ReportFileName = blm.report_path & "out1d.rpt"
'f = "{outward_view.V_date} in Date(" & date1.Year & "," & date1.Month & "," & date1.Day & ")"
'f = f & " to Date(" & date2.Year & "," & date2.Month & "," & date2.Day & ")"
'r1.SelectionFormula = f
'r1.WindowTop = 0
'r1.WindowLeft = 0
'r1.WindowState = crptMaximized
'r1.Action = 1
'r1.PageZoom 100

r1.ReportFileName = blm.report_path & "inward2d.rpt"
r1.DataFiles(0) = "BLOOMNEW.OUTWARD_VIEW"

f = "{outward_view.v_date} in Date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
f = f & " to Date(" & Format(date2.Value, "yyyy,MM,dd") & ")"

r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100

'r1.PrintReport
    
End If
If Val(Text1.Text) = 3 Then
r1.ReportFileName = blm.report_path & "out3d.rpt"
r1.DataFiles(0) = "BLOOMNEW.OUTWARD_VIEW"

f = "{outward_view.V_date} in Date(" & date1.Year & "," & date1.Month & "," & date1.Day & ")"
f = f & " to Date(" & date2.Year & "," & date2.Month & "," & date2.Day & ")"
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100

'r1.PrintReport
    
End If

If Val(Text1.Text) = 4 Then
r1.ReportFileName = blm.report_path & "out2d.rpt"
r1.DataFiles(0) = "BLOOMNEW.OUTWARD_VIEW"

f = "{outward_view.V_date} in Date(" & date1.Year & "," & date1.Month & "," & date1.Day & ")"
f = f & " to Date(" & date2.Year & "," & date2.Month & "," & date2.Day & ")"
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100

'r1.PrintReport
    
End If

If Val(Text1.Text) = 5 Then
'r1.ReportFileName = blm.report_path & "inward2d.rpt"
'f = "{inward_view.in_date} in Date(" & date1.Year & "," & date1.Month & "," & date1.Day & ")"
'f = f & " to Date(" & date2.Year & "," & date2.Month & "," & date2.Day & ")"
'r1.SelectionFormula = f
'r1.WindowTop = 0
'r1.WindowLeft = 0
'r1.WindowState = crptMaximized
'r1.Action = 1
'r1.PageZoom 150

r1.ReportFileName = blm.report_path & "out1d.rpt"
r1.DataFiles(0) = "BLOOMNEW.INWARD_VIEW"

f = "{inward_view.in_date} in Date(" & date1.Year & "," & date1.Month & "," & date1.Day & ")"
f = f & " to Date(" & date2.Year & "," & date2.Month & "," & date2.Day & ")"
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
    
End If

If Val(Text1.Text) = 6 Then
r1.ReportFileName = blm.report_path & "inward3d.rpt"
r1.DataFiles(0) = "BLOOMNEW.INWARD_VIEW"

f = "{inward_view.in_date} in Date(" & date1.Year & "," & date1.Month & "," & date1.Day & ")"
f = f & " to Date(" & date2.Year & "," & date2.Month & "," & date2.Day & ")"
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 7 Then
blm.inventoryledger Combo1.ItemData(Combo1.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "inventoryledger.rpt"
'f = "{inventory.date} in date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
'f = f & "to date (" & Format(date2.Value, "yyyy,MM,dd") & ")"
r1.DataFiles(0) = Blm1.pathMain
'r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.ReportTitle = Combo1.Text & "        " & "From : " & Format(date1.Value, "dd/MM/yyyy") & "  To : " & Format(date2.Value, "dd/MM/yyyy")
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 8 Then
blm.inventoryinward Combo1.ItemData(Combo1.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "inventoryinward.rpt"
r1.DataFiles(0) = Blm1.pathMain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 9 Then
blm.inventoryoutward Combo1.ItemData(Combo1.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "inventoryoutward.rpt"
r1.DataFiles(0) = Blm1.pathMain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 10 Then
blm.yinventoryledger Combo2.ItemData(Combo2.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "yinventoryledger.rpt"
f = "{inventory.date} in date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
f = f & "to date (" & Format(date2.Value, "yyyy,MM,dd") & ")"
r1.DataFiles(0) = Blm1.pathMain
r1.ReportTitle = Combo2.Text
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 11 Then
blm.finventoryledger Combo3.ItemData(Combo3.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "finventoryledger.rpt"
f = "{inventory.date} in date(" & Format(date1.Value, "yyyy,MM,dd") & ")"
f = f & "to date (" & Format(date2.Value, "yyyy,MM,dd") & ")"
r1.ReportTitle = Combo3.Text
r1.DataFiles(0) = Blm1.pathMain
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 12 Then
blm.yinventoryinward Combo2.ItemData(Combo2.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "yinventoryinward.rpt"
r1.ReportTitle = Combo2.Text
r1.DataFiles(0) = Blm1.pathMain
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 13 Then
blm.finventoryinward Combo3.ItemData(Combo3.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "finventoryinward.rpt"
r1.ReportTitle = Combo3.Text
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 14 Then
blm.yinventoryoutward Combo2.ItemData(Combo2.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "yinventoryoutward.rpt"
r1.ReportTitle = Combo2.Text
r1.DataFiles(0) = Blm1.pathMain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 15 Then
blm.finventoryoutward Combo3.ItemData(Combo3.ListIndex), date1.Value, date2.Value
r1.ReportFileName = blm.report_path & "finventoryoutward.rpt"
r1.ReportTitle = Combo3.Text
r1.DataFiles(0) = Blm1.pathMain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 16 Then
blm.TrialStock date1.Value
r1.ReportFileName = blm.report_path & "TrialStock.rpt"
r1.DataFiles(0) = Blm1.pathMain
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 17 Then
r1.ReportFileName = blm.report_path & "needlein.rpt"
r1.DataFiles(0) = "BLOOMNEW.NEEDLEIN_VW"
f = "{NEEDLEIN_VW.INDATE} IN DATE('" & Format(date1.Value, "yyyy,MM,dd") & "')"
f = f & " to date('" & Format(date2.Value, "yyyy,MM,dd") & "')"
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If

If Val(Text1.Text) = 18 Then
r1.ReportFileName = blm.report_path & "needleout.rpt"
r1.DataFiles(0) = "BLOOMNEW.NEEDLEOUT_VW"
f = "{NEEDLEOUT_VW.OUTDATE} IN DATE('" & Format(date1.Value, "yyyy,MM,dd") & "')"
f = f & " to date('" & Format(date2.Value, "yyyy,MM,dd") & "')"
r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.Action = 1
r1.PageZoom 100
End If


If Val(Text1.Text) = 19 Then
    r1.ReportFileName = blm.report_path & "issueyarn.rpt"
    r1.DataFiles(0) = "BLOOMNEW.YARNISSUE_VW"
    f = "{YARNISSUE_VW.IN_DATE} IN DATE('" & Format(date1.Value, "yyyy,MM,dd") & "')"
    f = f & " to date('" & Format(date2.Value, "yyyy,MM,dd") & "')"
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
 r1.PageZoom 100
End If

If Val(Text1.Text) = 20 Then
    r1.ReportFileName = blm.report_path & "receivefabric.rpt"
    r1.DataFiles(0) = "BLOOMNEW.FABRICRECEIVE_VW"
    f = "{FABRICRECEIVE_VW.V_DATE} IN DATE('" & Format(date1.Value, "yyyy,MM,dd") & "')"
    f = f & " to date('" & Format(date2.Value, "yyyy,MM,dd") & "')"
    r1.SelectionFormula = f
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
 r1.PageZoom 100
End If

Set blm = Nothing
Screen.MousePointer = vbDefault
'Unload Me
'Me.Hide
End Sub

Private Sub Form_Activate()
If Frame3.Visible = True Then
Combo1.SetFocus
End If
If Frame2.Visible = True Then
Combo2.SetFocus
End If
If Frame4.Visible = True Then
Combo3.SetFocus
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If

End Sub

Private Sub Form_Load()
Dim ssql As String

'Set blm = New bloom_r
'Me.Icon = LoadPicture(blm.report_path & "earth.ico")
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
date1.Value = Date
date2.Value = Date

ssql = "select * from Acchart Order by Name"
Blm1.fill_comb ssql, Combo1, "Name", "COde"

ssql = "select * from Yarn order by code"
Blm1.fill_comb ssql, Combo2, "name", "code" ', "wIDTH"

ssql = "select * from Cloth order by code"
Blm1.fill_comb ssql, Combo3, "name", "code" ', "wIDTH"


End Sub

