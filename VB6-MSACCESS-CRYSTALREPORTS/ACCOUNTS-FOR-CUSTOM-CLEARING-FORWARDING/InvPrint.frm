VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form InvPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice Print"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1425
      Left            =   180
      TabIndex        =   7
      Top             =   60
      Width           =   5580
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   3300
         TabIndex        =   3
         Top             =   885
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   39142
      End
      Begin VB.TextBox txtbillNo 
         Height          =   285
         Left            =   1305
         TabIndex        =   2
         Top             =   900
         Width           =   1320
      End
      Begin VB.TextBox txtPartyname 
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Top             =   540
         Width           =   4065
      End
      Begin VB.TextBox txtSerialNo 
         Height          =   285
         Left            =   1290
         TabIndex        =   0
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "(F1) to Search Document"
         Height          =   255
         Left            =   2820
         TabIndex        =   12
         Top             =   210
         Width           =   2490
      End
      Begin VB.Label Label4 
         Caption         =   "Date"
         Height          =   210
         Left            =   2820
         TabIndex        =   11
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Bill No."
         Height          =   240
         Left            =   225
         TabIndex        =   10
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Party Name"
         Height          =   225
         Left            =   225
         TabIndex        =   9
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Serial No"
         Height          =   240
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Width           =   885
      End
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2025
      Top             =   2115
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2685
      TabIndex        =   6
      Top             =   1875
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   975
      Left            =   4335
      Picture         =   "InvPrint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1590
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Prev"
      Default         =   -1  'True
      Height          =   975
      Left            =   285
      Picture         =   "InvPrint.frx":0561
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1620
      Width           =   1335
   End
End
Attribute VB_Name = "InvPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrintStatus As Integer
Dim EditMode As Boolean
Private blm As New bloom1
Private Blmr As New Bloom_r
Private Sub ExpTransfer(SerialNo As Double)
Dim db As Database
Dim DBT As Database
Dim tb As Recordset
Dim TBItm As Recordset
Dim TBT As Recordset
Dim ssql As String
Dim I As Integer
Set DBT = OpenDatabase(App.Path & "\Book.mdb")
ssql = "Delete from BillExp"
DBT.Execute ssql
Set TBT = DBT.OpenRecordset("BillExp", dbOpenTable)
Set db = OpenDatabase(App.Path & "\Bloom.mdb")
ssql = "Select * from Item Order by Code"
Set TBItm = db.OpenRecordset(ssql)
ssql = "Select * from Voucher where SerialNo=" & SerialNo
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    tb.MoveLast
    tb.MoveFirst
End If
If Not TBItm.EOF Then
    I = 0
    Do While Not TBItm.EOF
        TBT.AddNew
            TBT.Fields("SerialNo").Value = SerialNo
            TBT.Fields("ExpName").Value = TBItm.Fields("Name").Value
        
        If tb.RecordCount > 0 Then
            tb.FindFirst "ExPCode=" & TBItm.Fields("Code").Value
            If tb.NoMatch = False Then
                    TBT.Fields("ExpRemarks").Value = tb.Fields("ExpRemarks").Value
                    TBT.Fields("Amount").Value = tb.Fields("Debit").Value
            End If
        End If
        
        TBT.Update
        I = I + 1
    TBItm.MoveNext
    Loop
End If
TBItm.Close
tb.Close
TBT.Close
ssql = "Select Sum(Credit) as C from Voucher where SerialNo=" & SerialNo
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) And I > 0 Then
    ssql = "Update BillExp Set Credit=" & tb.Fields("C").Value / I
    DBT.Execute ssql
End If
tb.Close
db.Close

DBT.Close
End Sub
Private Function Max1() As Double
Dim db As Database
Dim TB1 As Recordset
Dim ssql As String


Set db = OpenDatabase(App.Path & "\Bloom.mdb")
ssql = "Select Max(BillNo) as B from Invoices"
Set TB1 = db.OpenRecordset(ssql)
If Not IsNull(TB1.Fields("B")) Then
    Max1 = TB1.Fields("B").Value + 1
Else
    Max1 = 1
End If
TB1.Close
db.Close
End Function
Private Sub ShowRecord()
Dim db As Database
Dim TB1 As Recordset
Dim ssql As String


Set db = OpenDatabase(App.Path & "\Bloom.mdb")
ssql = "Select * from Invoices where SerialNo=" & Val(txtSerialNo.Text)
Set TB1 = db.OpenRecordset(ssql)
txtPartyname.Text = blm.PartyofCase(Val(txtSerialNo.Text))
If Not TB1.EOF Then
    EditMode = True
    txtPartyname.Text = TB1.Fields("partyName").Value & ""
    txtbillNo.Text = TB1.Fields("Billno").Value
    DTPicker1.Value = TB1.Fields("BillDate").Value
    If TB1.Fields("Printed").Value = 1 Then PrintStatus = 1
    If TB1.Fields("Duplicate").Value = 1 Then PrintStatus = 2
    If TB1.Fields("Triplicate").Value = 1 Then PrintStatus = 3
Else
    EditMode = False
    txtbillNo.Text = Max1
End If
TB1.Close
db.Close
End Sub
Private Sub Save()
Dim db As Database
Dim TB1 As Recordset
Dim ssql As String


Set db = OpenDatabase(App.Path & "\Bloom.mdb")
If EditMode = True Then
    ssql = "Delete from Invoices where SerialNo=" & Val(txtSerialNo.Text)
    db.Execute ssql
End If
Set TB1 = db.OpenRecordset("Invoices", dbOpenTable)
TB1.AddNew
    TB1.Fields("SerialNo").Value = Val(txtSerialNo.Text)
    TB1.Fields("partyname").Value = txtPartyname.Text
    TB1.Fields("BillNo").Value = Val(txtbillNo.Text)
    TB1.Fields("BillDate").Value = DTPicker1.Value
    If PrintStatus = 0 Then
        TB1.Fields("Printed").Value = 1
        TB1.Fields("Duplicate").Value = 0
        TB1.Fields("Triplicate").Value = 0
        
    End If
    If PrintStatus = 1 Then
        TB1.Fields("Printed").Value = 1
        TB1.Fields("Duplicate").Value = 1
        TB1.Fields("Triplicate").Value = 0
        
    End If
    If PrintStatus >= 2 Then
        TB1.Fields("Printed").Value = 1
        TB1.Fields("Duplicate").Value = 1
        TB1.Fields("Triplicate").Value = 1
    End If
TB1.Update
TB1.Close
db.Close
EditMode = False
End Sub

Private Sub Command1_Click()
Dim f As String
Dim b As Boolean
Screen.MousePointer = vbHourglass
    If Val(Text2.Text) = 1 Then
        Save
        ExpTransfer Val(txtSerialNo.Text)
        f = "{InvoicesVW.BillNo}=" & Val(txtbillNo.Text)
        r1.ReportFileName = App.Path & "\Bill.rpt"
        r1.DataFiles(0) = App.Path & "\BLOOM.MDB"
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    If Val(Text2.Text) = 2 Then
        Blmr.CaseLedger Val(txtSerialNo.Text)
        r1.ReportFileName = App.Path & "\CaseLedger.rpt"
        r1.DataFiles(0) = App.Path & "\BOOK.MDB"
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
DTPicker1.Value = Date

Me.Top = 10
Me.Left = 10
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtSerialNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load DocsList
    DocsList.Show vbModal
    txtSerialNo.Text = SelSerialNo
End If
End Sub

Private Sub txtSerialNo_Validate(Cancel As Boolean)
If Val(txtSerialNo.Text) > 0 Then
    ShowRecord
End If
End Sub
