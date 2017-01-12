VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CityBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "City Wise Balances"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5325
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   5865
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   3150
      TabIndex        =   0
      Top             =   30
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20709379
      CurrentDate     =   39141
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1170
      Left            =   135
      TabIndex        =   6
      Top             =   4650
      Width           =   5025
      Begin VB.CommandButton Command1 
         Caption         =   "&Prev"
         Height          =   855
         Left            =   360
         Picture         =   "CityBal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   2085
         Picture         =   "CityBal.frx":03CC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   855
         Left            =   3735
         Picture         =   "CityBal.frx":08D3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ListBox List1 
      Height          =   4335
      Left            =   165
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   315
      Width           =   4980
   End
   Begin Crystal.CrystalReport r1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      Caption         =   "Upto"
      Height          =   210
      Left            =   2610
      TabIndex        =   7
      Top             =   75
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Cities"
      Height          =   210
      Left            =   150
      TabIndex        =   5
      Top             =   75
      Width           =   1455
   End
End
Attribute VB_Name = "CityBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bLm1 As New bloom1
Private Blmr As New Bloom_r

Public Sub CityBalance(e_date As Date, C As String, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
ssql = "delete from CityBal"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("CityBal", dbOpenTable)
    
'Vouchers

    ssql = "select a.party,b.name as AcName,c.name as CityName,sum(a.debit-a.credit) as bal from voucher a,Parties b,City c where a.Party=b.Code and b.CCode=c.Code and a.v_date <= #" & e_date & "# and b.CCode in (" & C & ") group by a.party,b.name,c.name"
    Set tb = db_m.OpenRecordset(ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("party").Value = tb.Fields("party").Value
        tb_t.Fields("Acname").Value = tb.Fields("AcName").Value
        tb_t.Fields("Bal").Value = tb.Fields("Bal").Value
        tb_t.Fields("CityName").Value = tb.Fields("CityName").Value
        tb_t.Update
        tb.MoveNext
        DoEvents
        cntl.Value = cntl.Value + 1
        Loop
        End If
    
    tb.Close

tb_t.Close

Set TBT = db_t.OpenRecordset("CityBal", dbOpenTable)

If Not TBT.EOF Then
    Do While Not TBT.EOF
        ssql = "Select V_Date,Credit from Voucher where Party=" & TBT.Fields("Party").Value & " and V_Date = (Select Max(V_Date) from Voucher where Party=" & TBT.Fields("Party").Value & ")"
        Set TBM = db_m.OpenRecordset(ssql)
        If Not TBM.EOF Then
            TBT.Edit
                TBT.Fields("LastDate").Value = TBM.Fields("V_Date").Value
                TBT.Fields("LastBal").Value = TBM.Fields("Credit").Value
            TBT.Update
        End If
        TBM.Close
    TBT.MoveNext
    Loop
End If
TBT.Close

db_t.Close
db_m.Close
End Sub

Private Sub Combs()
Dim ssql As String

ssql = "Select * from City Order by Name"
bLm1.fill_comb ssql, List1, "Name", "Code"
End Sub

Private Sub Command1_Click()
Dim CityCodes As String
Dim R As Integer

For R = 0 To List1.ListCount - 1
    If List1.Selected(R) = True Then
        CityCodes = CityCodes & List1.ItemData(R) & ","
    End If
Next R
CityCodes = Left(CityCodes, Len(CityCodes) - 1)
CityBalance DTPicker1.Value, CityCodes, ProgressBar1
    r1.ReportFileName = App.Path & "\CityBal.rpt"
    r1.DataFiles(0) = App.Path & "\Book.mdb"
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
    r1.PageZoom 100
End Sub

Private Sub Command2_Click()
Combs
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
Combs
DTPicker1.Value = Date

End Sub
