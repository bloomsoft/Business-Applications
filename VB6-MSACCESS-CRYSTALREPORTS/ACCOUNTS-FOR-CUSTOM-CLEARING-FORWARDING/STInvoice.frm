VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form STInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Tax Invoice"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8055
   Begin VB.CommandButton Command3 
      Caption         =   "&Reset"
      Height          =   975
      Left            =   6615
      Picture         =   "STInvoice.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1395
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Prev"
      Default         =   -1  'True
      Height          =   975
      Left            =   6600
      Picture         =   "STInvoice.frx":0507
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   975
      Left            =   6615
      Picture         =   "STInvoice.frx":08FC
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2565
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   6180
      Left            =   135
      TabIndex        =   25
      Top             =   75
      Width           =   6330
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2970
         Top             =   2940
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   4860
         TabIndex        =   22
         Top             =   5820
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53870595
         CurrentDate     =   39145
      End
      Begin VB.TextBox txtIndex 
         Height          =   285
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2745
         Width           =   1020
      End
      Begin VB.TextBox txtCustomBENo 
         Height          =   285
         Left            =   4995
         TabIndex        =   21
         Top             =   5445
         Width           =   1170
      End
      Begin VB.TextBox txtSTRNo 
         Height          =   300
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1650
         Width           =   4935
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   1395
         TabIndex        =   20
         Top             =   5490
         Width           =   1050
      End
      Begin VB.TextBox txtAddSTAmount 
         Height          =   285
         Left            =   4995
         TabIndex        =   19
         Top             =   5055
         Width           =   1185
      End
      Begin VB.TextBox txtAddSTRate 
         Height          =   285
         Left            =   1815
         TabIndex        =   18
         Top             =   5040
         Width           =   600
      End
      Begin VB.TextBox txtSTAmount 
         Height          =   285
         Left            =   4905
         TabIndex        =   17
         Top             =   4605
         Width           =   1260
      End
      Begin VB.TextBox txtSTRate 
         Height          =   285
         Left            =   1590
         TabIndex        =   16
         Top             =   4605
         Width           =   825
      End
      Begin VB.TextBox txtOtherExpences 
         Height          =   285
         Left            =   4905
         TabIndex        =   15
         Top             =   3870
         Width           =   1245
      End
      Begin VB.TextBox txtServiceCharges 
         Height          =   285
         Left            =   1185
         TabIndex        =   14
         Top             =   3885
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   4890
         TabIndex        =   13
         Top             =   3375
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53870595
         CurrentDate     =   39145
      End
      Begin VB.TextBox txtLCNo 
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3480
         Width           =   1995
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   4890
         TabIndex        =   11
         Top             =   3090
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53870595
         CurrentDate     =   39145
      End
      Begin VB.TextBox txtIGMNo 
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3120
         Width           =   1185
      End
      Begin VB.TextBox txtPerSS 
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   4950
      End
      Begin VB.TextBox txtImportValue 
         Height          =   285
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1950
         Width           =   1020
      End
      Begin VB.TextBox txtPackages 
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2025
         Width           =   1185
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txtParty 
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   975
         Width           =   4920
      End
      Begin VB.TextBox txtSerialNo 
         Height          =   285
         Left            =   1185
         TabIndex        =   2
         Top             =   615
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   4815
         TabIndex        =   1
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53870595
         CurrentDate     =   39145
      End
      Begin VB.TextBox txtInvNo 
         Height          =   285
         Left            =   1185
         TabIndex        =   0
         Top             =   225
         Width           =   1170
      End
      Begin Crystal.CrystalReport r1 
         Left            =   0
         Top             =   0
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
      Begin VB.Label Label26 
         Caption         =   "Date"
         Height          =   195
         Left            =   3390
         TabIndex        =   52
         Top             =   5820
         Width           =   1305
      End
      Begin VB.Label Label25 
         Caption         =   "Index"
         Height          =   225
         Left            =   4125
         TabIndex        =   50
         Top             =   2745
         Width           =   870
      End
      Begin VB.Label Label24 
         Caption         =   "(F4) to Search"
         Height          =   240
         Left            =   2505
         TabIndex        =   49
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Custom B.E. Number"
         Height          =   285
         Left            =   3360
         TabIndex        =   48
         Top             =   5475
         Width           =   1515
      End
      Begin VB.Label Label22 
         Caption         =   "Total Value Incl. S.Tax"
         Height          =   390
         Left            =   240
         TabIndex        =   47
         Top             =   5445
         Width           =   1560
      End
      Begin VB.Label Label21 
         Caption         =   "Add. Sale tax Amount"
         Height          =   225
         Left            =   3360
         TabIndex        =   46
         Top             =   5055
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "%"
         Height          =   255
         Left            =   2535
         TabIndex        =   45
         Top             =   5085
         Width           =   300
      End
      Begin VB.Label Label19 
         Caption         =   "Add. Sales Tax Rate"
         Height          =   255
         Left            =   270
         TabIndex        =   44
         Top             =   5040
         Width           =   1485
      End
      Begin VB.Label Label18 
         Caption         =   "Sale tax Amount"
         Height          =   255
         Left            =   3330
         TabIndex        =   43
         Top             =   4650
         Width           =   1560
      End
      Begin VB.Label Label17 
         Caption         =   "%"
         Height          =   270
         Left            =   2520
         TabIndex        =   42
         Top             =   4605
         Width           =   450
      End
      Begin VB.Label Label16 
         Caption         =   "Rate of Sales Tax"
         Height          =   210
         Left            =   270
         TabIndex        =   41
         Top             =   4605
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Other Un-Receipted Expenses made on Behalf of the receipient (if any)"
         Height          =   795
         Left            =   3345
         TabIndex        =   40
         Top             =   3825
         Width           =   1635
      End
      Begin VB.Label Label14 
         Caption         =   "Service Charges"
         Height          =   420
         Left            =   225
         TabIndex        =   39
         Top             =   3840
         Width           =   1005
      End
      Begin VB.Label Label13 
         Caption         =   "L/C Date"
         Height          =   180
         Left            =   4140
         TabIndex        =   38
         Top             =   3495
         Width           =   780
      End
      Begin VB.Label Label12 
         Caption         =   "L/C No."
         Height          =   255
         Left            =   210
         TabIndex        =   37
         Top             =   3480
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "IGM Date"
         Height          =   240
         Left            =   4140
         TabIndex        =   36
         Top             =   3120
         Width           =   960
      End
      Begin VB.Label Label10 
         Caption         =   "IGM No."
         Height          =   270
         Left            =   225
         TabIndex        =   35
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label Label9 
         Caption         =   "Per S.S."
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   2775
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "Description"
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   2415
         Width           =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Import Value"
         Height          =   240
         Left            =   4155
         TabIndex        =   32
         Top             =   2010
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Packages"
         Height          =   255
         Left            =   270
         TabIndex        =   31
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Label5 
         Caption         =   "S.T.R. #"
         Height          =   180
         Left            =   255
         TabIndex        =   30
         Top             =   1650
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "Party "
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Serial No."
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Date "
         Height          =   270
         Left            =   4200
         TabIndex        =   27
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice No."
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   225
         Width           =   1050
      End
   End
End
Attribute VB_Name = "STInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Sub GetDatafromInvoice()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

Set db = OpenDatabase(App.Path & "\Bloom.mdb")
ssql = "Select * from GSTInvoice where SerialNo=" & Val(txtSerialNo.Text)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    
    txtCustomBENo.Text = tb.Fields("CustomBENo").Value & ""
    If Not IsNull(tb.Fields("BEDate").Value) Then DTPicker4.Value = tb.Fields("BEDate").Value
    txtServiceCharges.Text = tb.Fields("ServiceCharges").Value & ""
    txtOtherExpences.Text = tb.Fields("OtherCharges").Value & ""
    txtSTRate.Text = tb.Fields("STRate").Value & ""
    txtAddSTRate.Text = tb.Fields("AddSTRate").Value & ""
    txtSTAmount.Text = tb.Fields("STAmount").Value & ""
    txtAddSTAmount.Text = tb.Fields("AddSTAmount").Value & ""
    txtTotal.Text = tb.Fields("Total").Value & ""
End If
tb.Close
db.Close
End Sub
Private Sub Save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

Set db = OpenDatabase(App.Path & "\Bloom.mdb")
ssql = "Delete from GSTInvoice"
db.Execute ssql
Set tb = db.OpenRecordset("GSTInvoice", dbOpenTable)
tb.AddNew
    tb.Fields("InvNo").Value = Val(txtInvNo.Text)
    tb.Fields("InvDate").Value = DTPicker1.Value
    tb.Fields("SerialNo").Value = Val(txtSerialNo.Text)
    tb.Fields("PartyName").Value = txtParty.Text
    tb.Fields("Address").Value = txtAddress.Text
    tb.Fields("STRNo").Value = txtSTRNo.Text
    tb.Fields("Packages").Value = txtPackages.Text
    tb.Fields("ImportValue").Value = Val(txtImportValue.Text)
    tb.Fields("Description").Value = txtDescription.Text
    tb.Fields("PerSS").Value = txtPerSS.Text
    tb.Fields("IGMNO").Value = txtIGMNo.Text
    tb.Fields("IGMDate").Value = DTPicker2.Value
    tb.Fields("LCNo").Value = txtLCNo.Text
    tb.Fields("LCDate").Value = DTPicker3.Value
    tb.Fields("CustomBENo").Value = txtCustomBENo.Text
    tb.Fields("BEDate").Value = DTPicker4.Value
    tb.Fields("ServiceCharges").Value = Val(txtServiceCharges.Text)
    tb.Fields("OtherCharges").Value = Val(txtOtherExpences.Text)
    tb.Fields("STRate").Value = Val(txtSTRate.Text)
    tb.Fields("AddSTRate").Value = Val(txtAddSTRate.Text)
    tb.Fields("STAmount").Value = Val(txtSTAmount.Text)
    tb.Fields("AddSTAmount").Value = Val(txtAddSTAmount.Text)
    tb.Fields("Total").Value = Val(txtTotal.Text)
tb.Update
tb.Close
db.Close

End Sub
Private Function Max1() As Double
Dim db As Database
Dim TB1 As Recordset
Dim ssql As String


Set db = OpenDatabase(App.Path & "\Bloom.mdb")
ssql = "Select Max(InvNo) as B from GSTInvoice"
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
Dim tb As Recordset
Dim ssql As String

Set db = OpenDatabase(App.Path & "\Bloom.mdb")

ssql = "Select * from Docs where SrNo=" & Val(txtSerialNo.Text)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    EditMode = True
    txtLCNo.Text = tb.Fields("LCNo").Value & ""
    DTPicker3.Value = tb.Fields("LCDate").Value
    txtPackages.Text = tb.Fields("Packages").Value & ""
    txtDescription.Text = tb.Fields("Goods").Value & ""
    txtPerSS.Text = tb.Fields("PerSS").Value & ""
    txtParty.Text = Blm1.party1(tb.Fields("partyCode").Value)
    txtAddress.Text = Blm1.Address1(tb.Fields("partyCode").Value)
    txtSTRNo.Text = Blm1.GST1(tb.Fields("partyCode").Value)
    txtIGMNo.Text = tb.Fields("IGMNo").Value & ""
    DTPicker2.Value = tb.Fields("IGMDate").Value
    txtImportValue.Text = tb.Fields("ImportValue").Value & ""
    txtIndex.Text = tb.Fields("IndexNo").Value & ""
End If
tb.Close
db.Close
End Sub

Private Sub Command1_Click()
Dim F As String
Screen.MousePointer = vbHourglass
Save
        F = "{GSTInvoice.SerialNo}=" & Val(txtSerialNo.Text)
        r1.ReportFileName = App.Path & "\GSTInvoice.rpt"
        r1.DataFiles(0) = App.Path & "\BLOOM.MDB"
        r1.SelectionFormula = F
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command3_Click()
Dim C As Control
For Each C In Me.Controls
    If TypeOf C Is TextBox Then C.Text = ""
    If TypeOf C Is DTPicker Then C.Value = Now
Next
End Sub

Private Sub Form_Load()
txtInvNo.Text = Max1
DTPicker1.Value = Date

End Sub

Private Sub Timer1_Timer()
txtSTAmount.Text = ((Val(txtSTRate.Text) * Val(txtServiceCharges.Text))) / 100
txtAddSTAmount.Text = ((Val(txtAddSTRate.Text) * Val(txtServiceCharges.Text))) / 100
txtTotal.Text = Val(txtServiceCharges.Text) + Val(txtOtherExpences.Text) + Val(txtSTAmount.Text) + Val(txtAddSTAmount.Text)

End Sub

Private Sub txtSerialNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    SelPartyCode = Val(txtSerialNo.Text)
    DocsList.Show vbModal
    txtSerialNo.Text = SelSerialNo
End If
End Sub

Private Sub txtSerialNo_Validate(Cancel As Boolean)
If Len(txtSerialNo.Text) > 0 Then
    ShowRecord
    GetDatafromInvoice
End If
End Sub
