VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Setup1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "P/&L Details"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   3960
      Width           =   1935
   End
   Begin Crystal.CrystalReport r1 
      Left            =   4560
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510464
      CurrentDate     =   36969
   End
   Begin ComctlLib.ProgressBar p1 
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker Date1 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510464
      CurrentDate     =   36953
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&P/L Notes"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selectd For Profit && Loss"
      Height          =   4695
      Left            =   6480
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
      Begin VB.ListBox List2 
         Height          =   4335
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available Heads"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
      Begin VB.ListBox List1 
         Height          =   4335
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Setup1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private blmr As New bloom_r
Private Sub PLCalc(v_Date As Date, cntl As Control, e_date As Date)
Dim db_m As Database
Dim tb As Recordset
Dim tb_p As Recordset
Dim tb_t As Recordset
Dim db_t As Database
Dim ssql As String
Dim q As Long
Dim f As Byte
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(blmr.path)
ssql = "Delete from trial"
db_t.Execute ssql
ssql = "delete from pl"
db_t.Execute ssql
ssql = "select * from Heads order by code"
Set tb_p = db_m.OpenRecordset(ssql)
Set tb_t = db_t.OpenRecordset("pl", dbOpenDynaset)
ssql = "select HeadCode,sum(debit) as Deb,sum(credit) as Cred,Sum(Debit)-Sum(Credit) as bal from vou_view where v_date  between#" & v_Date & "# and #" & e_date & "# Group by HeadCOde"
'MsgBox ssql
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    tb.MoveLast
    cntl.Value = 0
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
    f = 0
        For q = 0 To List2.ListCount - 1
            If List2.ItemData(q) = tb.Fields("HeadCOde") Then
            
            f = 1
                Exit For
                
            End If
        Next q
        If Not tb_p.EOF Then
        
                tb_p.FindFirst "Code = " & tb.Fields("HeadCOde").Value
            
            If tb_p.NoMatch Then
                p$ = ""
            Else
                p$ = tb_p.Fields("Name").Value
            End If
        End If
        If f = 1 Then
     
        tb_t.AddNew
            tb_t.Fields("V_date").Value = v_Date
            tb_t.Fields("s_date").Value = e_date
            tb_t.Fields("Code").Value = tb.Fields("HeadCOde").Value
            tb_t.Fields("Name").Value = p$ '
        
                tb_t.Fields("Debit").Value = tb.Fields("Deb").Value
                tb_t.Fields("Credit").Value = 0
        
        tb_t.Update
        tb_t.AddNew
            tb_t.Fields("V_date").Value = v_Date
            tb_t.Fields("Code").Value = tb.Fields("HeadCOde").Value
            tb_t.Fields("Name").Value = p$ '
                tb_t.Fields("Debit").Value = 0
                tb_t.Fields("Credit").Value = tb.Fields("Cred").Value '* -1
        
        tb_t.Update
       
       End If
        tb.MoveNext
        cntl.Value = cntl.Value + 1
        f = 0
    
    Loop
End If
tb.Close
If Not tb_t.EOF Then

tb_t.MoveFirst
Dim cred As Currency
Dim deb As Currency
Dim b As Currency
Do While Not tb_t.EOF
cred = cred + tb_t.Fields("Credit").Value
deb = deb + tb_t.Fields("Debit").Value
tb_t.MoveNext
Loop
tb_t.Close
Set tb_t = db_t.OpenRecordset("Trial", dbOpenTable)
    tb_t.AddNew
        tb_t.Fields("s_date").Value = v_Date
        tb_t.Fields("Code").Value = 1
        tb_t.Fields("Name").Value = "Net Profit & Loss"
        b = deb - (cred * -1)
        If b > 0 Then
            tb_t.Fields("Debit").Value = b
            tb_t.Fields("Credit").Value = 0
        End If
        If b < 0 Then
            tb_t.Fields("Debit").Value = 0
            tb_t.Fields("Credit").Value = b
        End If
        If b = 0 Then
            tb_t.Fields("Debit").Value = 0
            tb_t.Fields("Credit").Value = 0
        End If
        tb_t.Update
End If
tb_t.Close
tb_p.Close
db_t.Close
db_m.Close
End Sub
Private Sub PLCalc_Detail(v_Date As Date, cntl As Control, e_date As Date)
Dim db_m As Database
Dim tb As Recordset
Dim tb_p As Recordset
Dim tb_t As Recordset
Dim db_t As Database
Dim ssql As String
Dim q As Long
Dim f As Byte
Dim hcode As Long
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(blmr.path)
ssql = "Delete from trial"
db_t.Execute ssql
ssql = "delete from pl"
db_t.Execute ssql
ssql = "select * from Heads order by code"
Set tb_p = db_m.OpenRecordset(ssql)
Set tb_t = db_t.OpenRecordset("pl", dbOpenDynaset)
ssql = "select SubCode,sum(debit) as Deb,sum(credit) as Cred,Sum(Debit)-Sum(Credit) as bal from vou_view where v_date  between#" & v_Date & "# and #" & e_date & "# Group by SubCode"
'MsgBox ssql
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    tb.MoveLast
    cntl.Value = 0
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
    f = 0
        For q = 0 To List2.ListCount - 1
            If List2.ItemData(q) = Val(Mid(CStr(tb.Fields("SubCOde").Value), 1, 2)) Then
                f = 1
                Exit For
                
            End If
        Next q
        If Not tb_p.EOF Then
        
                tb_p.FindFirst "Code = " & Val(Mid(CStr(tb.Fields("SubCOde").Value), 1, 2))
            
            If tb_p.NoMatch Then
                p$ = ""
            Else
                p$ = tb_p.Fields("Name").Value
            End If
            tb_p.FindFirst "Code = " & tb.Fields("SubCOde").Value
            
            If tb_p.NoMatch Then
                ps$ = ""
            Else
                ps$ = tb_p.Fields("Name").Value
            End If
        End If
        
        If f = 1 Then
     
        tb_t.AddNew
            tb_t.Fields("V_date").Value = v_Date
            tb_t.Fields("s_date").Value = e_date
            tb_t.Fields("Code").Value = Val(Mid(CStr(tb.Fields("SubCOde").Value), 1, 2))
            tb_t.Fields("Name").Value = p$ '
            tb_t.Fields("HCode").Value = tb.Fields("SubCOde").Value
            tb_t.Fields("HName").Value = ps$ '
        
                tb_t.Fields("Debit").Value = tb.Fields("Deb").Value
                tb_t.Fields("Credit").Value = tb.Fields("Cred").Value '* -1
        
        tb_t.Update
        'tb_t.AddNew
        '    tb_t.Fields("V_date").Value = v_Date
        '    tb_t.Fields("Code").Value = Val(Mid(CStr(tb.Fields("SubCOde").Value), 1, 2))
        '    tb_t.Fields("Name").Value = p$ '
        '    tb_t.Fields("HCode").Value = tb.Fields("SubCOde").Value
        '    tb_t.Fields("HName").Value = ps$ '
        '    tb_t.Fields("Debit").Value = 0
        '    tb_t.Fields("Credit").Value = tb.Fields("Cred").Value '* -1
        
        'tb_t.Update
       
       End If
        tb.MoveNext
        cntl.Value = cntl.Value + 1
        f = 0
    
    Loop
End If
tb.Close
If Not tb_t.EOF Then

tb_t.MoveFirst
Dim cred As Currency
Dim deb As Currency
Dim b As Currency
Do While Not tb_t.EOF
cred = cred + tb_t.Fields("Credit").Value
deb = deb + tb_t.Fields("Debit").Value
tb_t.MoveNext
Loop
tb_t.Close
Set tb_t = db_t.OpenRecordset("Trial", dbOpenTable)
    tb_t.AddNew
        tb_t.Fields("s_date").Value = v_Date
        tb_t.Fields("Code").Value = 1
        tb_t.Fields("Name").Value = "Net Profit & Loss"
        b = deb - (cred * -1)
        If b > 0 Then
            tb_t.Fields("Debit").Value = b
            tb_t.Fields("Credit").Value = 0
        End If
        If b < 0 Then
            tb_t.Fields("Debit").Value = 0
            tb_t.Fields("Credit").Value = b
        End If
        If b = 0 Then
            tb_t.Fields("Debit").Value = 0
            tb_t.Fields("Credit").Value = 0
        End If
        tb_t.Update
End If
tb_t.Close
tb_p.Close
db_t.Close
db_m.Close
End Sub

Private Sub FillList()
Dim ssql As String
List2.Clear

ssql = "Select * from heads where Code < 100 order by COde"
blm.fill_comb ssql, List1, "Name", "COde"
End Sub

Private Sub Command1_Click()
Dim i As Long
If List1.ListCount > 0 Then
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
        List2.AddItem List1.List(i)
        List2.ItemData(List2.NewIndex) = List1.ItemData(i)
    End If
Next i
Dim j As Long
j = List1.ListCount - 1
For i = j To 0 Step -1
    If List1.Selected(i) Then
        List1.RemoveItem (i)
    End If
Next i

End If
End Sub

Private Sub Command2_Click()
Dim i As Long
If List2.ListCount > 0 Then
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) Then
        List1.AddItem List2.List(i)
        List1.ItemData(List1.NewIndex) = List2.ItemData(i)
    End If
Next i
Dim j As Long
j = List2.ListCount - 1
For i = j To 0 Step -1
    If List2.Selected(i) Then
        List2.RemoveItem (i)
    End If
Next i

End If
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Command4_Click()
Screen.MousePointer = vbHourglass
PLCalc Date1.Value, p1, date2.Value
Screen.MousePointer = vbDefault
r1.ReportFileName = blmr.report_path & "pl.rpt"
r1.WindowState = crptMaximized
r1.WindowTop = 0
r1.WindowLeft = 0
r1.Action = 1
r1.PageZoom 100
End Sub

Private Sub Command5_Click()
Screen.MousePointer = vbHourglass
PLCalc_Detail Date1.Value, p1, date2.Value
Screen.MousePointer = vbDefault
r1.ReportFileName = blmr.report_path & "pl2.rpt"
r1.WindowState = crptMaximized
r1.WindowTop = 0
r1.WindowLeft = 0
r1.Action = 1
r1.PageZoom 100

End Sub

Private Sub Form_Load()
FillList
End Sub

