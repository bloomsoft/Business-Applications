VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form In1 
   Caption         =   "Purchase / Inwards"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   54
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   9255
      Begin VB.TextBox Text21 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   6480
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H8000000E&
         Caption         =   "&DELETE THIS INWARD"
         Height          =   735
         Left            =   6960
         Picture         =   "In1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3240
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   120
         Width           =   975
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   36214
      End
      Begin VB.Label Label27 
         Caption         =   "Brokery"
         Height          =   255
         Left            =   7320
         TabIndex        =   66
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Bro Rate"
         Height          =   255
         Left            =   5760
         TabIndex        =   65
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Broker Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   63
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Broker Code"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Job Balance"
         Height          =   255
         Left            =   2160
         TabIndex        =   60
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Job No."
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "INV_TYPE=1 FOR BUNYAN PURCHASE 3 for Towel, 2 for Socks"
         Height          =   375
         Left            =   6840
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Balance in Ledger"
         Height          =   255
         Left            =   5400
         TabIndex        =   29
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "InWard #"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8520
      TabIndex        =   21
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      TabIndex        =   22
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7440
      Top             =   4320
   End
   Begin VB.Frame Frame6 
      Caption         =   "Options"
      Height          =   855
      Left            =   9480
      TabIndex        =   47
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   960
         TabIndex        =   49
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   960
         TabIndex        =   48
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Actions"
      Height          =   1215
      Left            =   7800
      TabIndex        =   44
      Top             =   6000
      Width           =   3975
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000E&
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2640
         Picture         =   "In1.frx":08B4
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000E&
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1800
         Picture         =   "In1.frx":12D0
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "&Save"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   960
         Picture         =   "In1.frx":394D
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lists"
      Height          =   1575
      Left            =   120
      TabIndex        =   39
      Top             =   5640
      Width           =   7575
      Begin VB.ComboBox Combo3 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   43
         Text            =   "Combo3"
         Top             =   960
         Width           =   6015
      End
      Begin VB.ComboBox Combo2 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   41
         Text            =   "Combo2"
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label Label20 
         Caption         =   "Items List"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "A/c List"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      CausesValidation=   0   'False
      Height          =   2415
      Left            =   120
      TabIndex        =   38
      Top             =   3240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4260
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   105
      TabIndex        =   32
      Top             =   1680
      Width           =   11655
      Begin VB.TextBox Text22 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10875
         TabIndex        =   67
         Top             =   600
         Width           =   690
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   10215
         TabIndex        =   19
         Top             =   600
         Width           =   660
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   4590
         TabIndex        =   14
         Top             =   585
         Width           =   765
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   8760
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7275
         TabIndex        =   17
         Top             =   600
         Width           =   1320
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   10560
         TabIndex        =   20
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Entr&y"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   9120
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6315
         TabIndex        =   16
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5355
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3855
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   2550
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "L-Kami"
         Height          =   255
         Left            =   10440
         TabIndex        =   61
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Bales"
         Height          =   210
         Left            =   4635
         TabIndex        =   56
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "WareHouse"
         Height          =   195
         Left            =   8880
         TabIndex        =   55
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label21 
         Caption         =   "Amount"
         Height          =   255
         Left            =   7680
         TabIndex        =   50
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Stock"
         Height          =   255
         Left            =   6675
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Rate"
         Height          =   255
         Left            =   3885
         TabIndex        =   36
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   5475
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Truck / Lorry#"
      Height          =   735
      Left            =   9480
      TabIndex        =   26
      Top             =   960
      Width           =   2295
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label14 
      Height          =   285
      Left            =   8625
      TabIndex        =   58
      Top             =   7305
      Width           =   585
   End
   Begin VB.Label Label9 
      Height          =   240
      Left            =   7800
      TabIndex        =   57
      Top             =   7335
      Width           =   660
   End
   Begin VB.Label Label23 
      Caption         =   "Balance"
      Height          =   255
      Left            =   7800
      TabIndex        =   52
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Total"
      Height          =   255
      Left            =   9960
      TabIndex        =   51
      Top             =   5760
      Width           =   375
   End
End
Attribute VB_Name = "In1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim crow As Long
Private Sub JobBalance()
Dim db As Database
Dim tb As Recordset
Dim tbP As Recordset
Dim Ssql As String
Set db = OpenDatabase(blm.pathMain)

Ssql = "Select Sum(Quantity-LKamiValue) as Bal from In_DTL where JobNo=" & Val(Text4.Text)
Set tbP = db.OpenRecordset(Ssql)

Ssql = "select * from PContract where Cont_no = " & Val(Text4.Text) & " and SellerCode=" & Val(Text5.Text)
Set tb = db.OpenRecordset(Ssql)
If Not tb.EOF Then
    If Not IsNull(tbP.Fields("Bal")) Then
        Text7.Text = tb.Fields("Quantity") - tbP.Fields("Bal").Value
    Else
        Text7.Text = tb.Fields("Quantity")
    End If
Else
    MsgBox "Invalid Job No. or Don't Belong to the Selected party"
End If
tbP.Close
tb.Close
db.Close
End Sub
Private Sub ledgerbal()
Dim b As Currency
Dim s As String
b = blm.Balance(Val(Text5.Text))
If b < 0 Then
    s = Format(b * -1, "#.000") & " CR"
End If
If b > 0 Then
    s = Format(b, "#.000") & " DR"
End If
Label3.Caption = s
End Sub
Private Function edit1() As Boolean
Dim db As Database
Dim tb As Recordset
Dim Ssql As String
Set db = OpenDatabase(blm.pathMain)
Ssql = "select * from in_mst where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
Set tb = db.OpenRecordset(Ssql)
If Not tb.EOF Then
    Date1.Value = tb.Fields("v_date").Value
    Text5.Text = tb.Fields("Party").Value
    Text6.Text = blm.party1(tb.Fields("Party").Value)
    Text14.Text = tb.Fields("BrokerCode").Value
    Text15.Text = blm.broker1(tb.Fields("BrokerCode").Value)
    Text8.Text = tb.Fields("Note").Value
    edit1 = False
Else
    MsgBox "No Record For This Inward No."
    edit1 = True
    Exit Function
End If
tb.Close
Dim i As Long
Ssql = "select * from in_dtl where p_no = " & Val(Text1.Text)
Ssql = Ssql & " and inv_type=" & Val(Text2.Text)
Set tb = db.OpenRecordset(Ssql)
If Not tb.EOF Then
Do While Not tb.EOF
Grid1.Rows = Grid1.Rows + 1
i = Grid1.Rows - 1
    With Grid1
        .TextMatrix(i, 0) = i
        .TextMatrix(i, 1) = tb.Fields("Item").Value
        .TextMatrix(i, 2) = blm.Item1(tb.Fields("Item").Value)
        .TextMatrix(i, 3) = tb.Fields("Scheme").Value
        .TextMatrix(i, 4) = Format(tb.Fields("Quantity").Value, "#.000")
        .TextMatrix(i, 5) = Format(tb.Fields("Rate").Value, "#.000")
        .TextMatrix(i, 6) = Format(tb.Fields("Quantity").Value * tb.Fields("Rate").Value, "#.000")
        .TextMatrix(i, 7) = tb.Fields("WareHouse").Value
        .TextMatrix(i, 8) = blm.WareHouse(tb.Fields("WareHouse").Value)
    End With
    tb.MoveNext
Loop
End If
tb.Close
db.Close
End Function
Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim Ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
Ssql = "delete from voucher where inv_type=" & Val(Text2.Text) & " and e_type=2 and ent_no = " & Val(Text1.Text)
db.Execute Ssql
Ssql = "delete from in_mst where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
db.Execute Ssql
Ssql = "delete from in_DTL where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
db.Execute Ssql

End If

Set tb = db.OpenRecordset("voucher", dbOpenTable)
tb.AddNew
    tb.Fields("ent_no").Value = Val(Text1.Text)
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("e_type").Value = 2
    tb.Fields("party").Value = Val(Text5.Text)
    tb.Fields("BrokerCode").Value = Val(Text14.Text)
    tb.Fields("debit").Value = 0
    tb.Fields("credit").Value = Round(Val(Text19.Text))
    tb.Fields("Remarks").Value = "Inward # " & Val(Text1.Text) & " Total Wt: " & Label9.Caption & " Total Bales: " & Label14.Caption
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text1.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("e_type").Value = 2
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("party").Value = 6000 + Val(Text2.Text) 'Stock A/c Code Should be This
    tb.Fields("BrokerCode").Value = Val(Text14.Text) 'Stock A/c Code Should be This
    tb.Fields("debit").Value = Round(Val(Text19.Text))
    tb.Fields("credit").Value = 0
    tb.Fields("Remarks").Value = Text6.Text & ", Inward # " & Val(Text1.Text) & " Total Wt: " & Label9.Caption & " Total Bales: " & Label14.Caption
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text1.Text)
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("e_type").Value = 2
    tb.Fields("party").Value = Val(Text14.Text)
    tb.Fields("BrokerCode").Value = -1
    tb.Fields("debit").Value = Round(Val(Text21.Text))
    tb.Fields("credit").Value = 0
    tb.Fields("Remarks").Value = "Brokerage Inward # " & Val(Text1.Text) & " Total Wt: " & Label9.Caption & " Total Bales: " & Label14.Caption
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text1.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("e_type").Value = 2
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("party").Value = Val(Text14.Text) 'bROKERY Credit A/c
    tb.Fields("BrokerCode").Value = -1
    tb.Fields("debit").Value = 0
    tb.Fields("credit").Value = Round(Val(Text21.Text))
    tb.Fields("Remarks").Value = "Brokerage Inward # " & Val(Text1.Text) & " Total Wt: " & Label9.Caption & " Total Bales: " & Label14.Caption
tb.Update

tb.Close
    
Set tb = db.OpenRecordset("in_mst", dbOpenTable)
tb.AddNew
    tb.Fields("P_no").Value = Val(Text1.Text)
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("V_Date").Value = Date1.Value
    tb.Fields("Party").Value = Val(Text5.Text)
    tb.Fields("BrokerCode").Value = Val(Text14.Text)
    tb.Fields("Note").Value = Text8.Text & " "
tb.Update
tb.Close
Dim i As Long
Set tb = db.OpenRecordset("in_dtl", dbOpenTable)
For i = 1 To Grid1.Rows - 1
    With Grid1
tb.AddNew
    tb.Fields("P_no").Value = Val(Text1.Text)
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("V_Date").Value = Date1.Value
    tb.Fields("Item").Value = Val(.TextMatrix(i, 1))
    tb.Fields("Scheme").Value = Val(.TextMatrix(i, 3))
    tb.Fields("Quantity").Value = Val(.TextMatrix(i, 4))
    tb.Fields("Rate").Value = Val(.TextMatrix(i, 5))
    tb.Fields("WareHouse").Value = Val(.TextMatrix(i, 7))
    
tb.Update
    End With
    
Next i
tb.Close
db.Close
End Sub
Private Sub Clear1()
Text9.Text = vbNullString
Text10.Text = vbNullString
Text11.Text = vbNullString
Text12.Text = vbNullString

Text16.Text = vbNullString
Text17.Text = vbNullString
'Text14.Text = vbNullString
End Sub
Private Sub Transfer1()
Dim i As Long
With Grid1
    .Rows = .Rows + 1
    i = .Rows - 1
    .TextMatrix(i, 0) = i
    .TextMatrix(i, 1) = Text9.Text
    .TextMatrix(i, 2) = Text10.Text
    .TextMatrix(i, 3) = Val(Text3.Text)
    .TextMatrix(i, 4) = Format(Val(Text12.Text), "#.000")
    .TextMatrix(i, 5) = Format(Val(Text11.Text), "#.000")
    .TextMatrix(i, 6) = Format(Val(Text17.Text), "#.000")
    .TextMatrix(i, 7) = Combo1.ItemData(Combo1.ListIndex)
    .TextMatrix(i, 8) = Combo1.Text
End With
'GRID1.TopRow = GRID1.Rows - 1
Text9.SetFocus
End Sub
Private Sub ClearFull()

Dim cntl As Control
Dim i As Integer

i = Val(Text2.Text)
For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
Combs
Flex1
If Option1 = True Then
    Text1.Text = max1
End If
Text2.Text = i
End Sub

Private Function max1() As Long
Dim db As Database
Dim tb As Recordset
Dim Ssql As String
Ssql = "select MAX(P_no) AS C FROM In_Mst where inv_type=" & Val(Text2.Text)
'Clipboard.SetText ssql
'MsgBox ssql
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(Ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
db.Close
End Function

Private Sub Combs()
Dim Ssql As String

'Accounts
Ssql = "select * from Parties order by Name"
blm.fill_comb Ssql, Combo2, "Name", "Code"
'Items
Ssql = "select * from Item order by Name"
blm.fill_comb_Item Ssql, Combo3, "Name", "Code"
'WareHouse
Ssql = "select * from WareHouse order by Name"
blm.fill_comb Ssql, Combo1, "Name", "Code"

End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Combo2_Click()
Text5.Text = Combo2.ItemData(Combo2.ListIndex)
Text6.Text = Combo2.Text
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text5.SetFocus
End Sub

Private Sub Combo2_LostFocus()
Text5.Text = Combo2.ItemData(Combo2.ListIndex)
Text6.Text = Combo2.Text

End Sub

Private Sub Combo3_Click()
Text9.Text = Combo3.ItemData(Combo3.ListIndex)
Text10.Text = Combo3.Text
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text9.SetFocus
End Sub

Private Sub Combo3_LostFocus()
Text9.Text = Combo3.ItemData(Combo3.ListIndex)
Text10.Text = Combo3.Text

End Sub

Private Sub Flex1()
With Grid1
    .Rows = 1
    .Cols = 9
    .ColWidth(0) = 1000
    .TextMatrix(0, 0) = "Sr #"
    .ColWidth(1) = 1500
    .TextMatrix(0, 1) = "Item Code"
    .ColWidth(2) = 4000
    .TextMatrix(0, 2) = "Item Name"
    .ColWidth(3) = 1000
    .TextMatrix(0, 3) = "Bales"
    .ColWidth(4) = 1500
    .TextMatrix(0, 4) = "Quantity"
    .ColWidth(5) = 1500
    .TextMatrix(0, 5) = "Rate"
    .ColWidth(6) = 2000
    .TextMatrix(0, 6) = "Amount"
    .ColWidth(7) = 15
    .TextMatrix(0, 7) = "Scheme"
    .ColWidth(8) = 1000
    .TextMatrix(0, 8) = "WareHouse"
    
    
    
End With
End Sub

Private Sub Command1_Click()
Call save
Command2_Click

End Sub

Private Sub Command2_Click()
Call ClearFull
If Option1 = True Then
    Text5.SetFocus
Else
    Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me

End Sub

Private Sub Command4_Click()
Call Clear1

Combo3_Click
Text9.SetFocus
End Sub

Private Sub Command5_Click()
'If Val(Text13.Text) = 0 Then
'    MsgBox "Please Give Bobins..."
'    Exit Sub
'End If
If Val(Text11.Text) = 0 Then
    MsgBox "Please Give Quantity..."
    Exit Sub
End If
If Val(Text12.Text) = 0 Then
    MsgBox "Please Give Rate..."
    Exit Sub
End If

Call Transfer1
Clear1
Text9.SetFocus
End Sub

Private Sub Command6_Click()
Load vour
vour.Text2.Text = 11
vour.Label1.Caption = "Inward #"
vour.Caption = "Inward Note Print"
vour.Show

End Sub

Private Sub Command7_Click()
Dim db As Database

Dim Ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
Ssql = "delete from voucher where inv_type=" & Val(Text2.Text) & " and e_type=2 and ent_no = " & Val(Text1.Text)
db.Execute Ssql
Ssql = "delete from in_mst where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
db.Execute Ssql
Ssql = "delete from in_DTL where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
db.Execute Ssql
End If
db.Close
Command2_Click

End Sub

Private Sub Command8_Click()
Dim i As Long

For i = 1 To Grid1.Rows - 1
    Grid1.TextMatrix(i, 4) = Format(Val(Text20.Text), "#.000")
    Grid1.TextMatrix(i, 5) = Format(Val(Text20.Text) * Grid1.TextMatrix(i, 3), "#.000")
Next i
End Sub

Private Sub Form_Activate()
Text1.Text = max1
End Sub

Private Sub Form_Load()
Combs
Text1.Text = max1
Date1.Value = Date
Flex1
End Sub

Private Sub Grid1_DblClick()
Dim thisrow As Long
Dim R As Integer
If crow > 0 Then
    thisrow = crow
Else
    thisrow = Grid1.Row
End If
If Val(Text12.Text) > 0 Then
    MsgBox "You Already Have Entry There ...."
Else
With Grid1
Text9.Text = .TextMatrix(thisrow, 1)
Text10.Text = .TextMatrix(thisrow, 2)
Text3.Text = .TextMatrix(thisrow, 3)
Text12.Text = .TextMatrix(thisrow, 4)
Text11.Text = .TextMatrix(thisrow, 5)

For R = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(R) = Val(.TextMatrix(thisrow, 7)) Then
        Combo1.ListIndex = R
        Exit For
    End If
Next R
End With
'MsgBox thisrow
If Grid1.Rows = 2 Then
    Grid1.Rows = 1
Else
    Grid1.RemoveItem (thisrow)
End If
Dim i As Long

For i = 1 To Grid1.Rows - 1
    Grid1.TextMatrix(i, 0) = i
Next i
End If
Text9.SetFocus
End Sub

Private Sub Option1_Click()
ClearFull
Command7.Visible = False
Text1.Enabled = False
Date1.SetFocus
End Sub

Private Sub Option2_Click()
ClearFull
Text1.Enabled = True
Command7.Visible = True
Text1.SetFocus


End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    Cancel = edit1
'Else
 '   MsgBox "Please Give Any Purchase No."
End If
End Sub

Private Sub Text11_GotFocus()
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12.Text)

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub




Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF1 Then
            Load Search3
            Search3.Text3.Text = 2
            Search3.Show vbModal
        End If

End Sub

Private Sub Text14_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text15.Text = blm.broker1(Val(Text14.Text))
    If Text15.Text = "Wrong" Then
        MsgBox "Invalid Broker Code..."
'        Cancel = True
    End If
'Else
 '   Cancel = True
End If
End Sub

Private Sub Text17_Change()
If Val(Text17.Text) > 0 Then
    Command5.Enabled = True
Else
    Command5.Enabled = False
End If
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_Change()
Text1.Text = max1
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If


End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) > 0 Then
    JobBalance
End If
End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)

End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
If KeyCode = vbKeyF2 Then
    Search2.Text3.Text = 5
    Search2.Show
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text5_Validate(Cancel As Boolean)
Dim b As Boolean

If Val(Text5.Text) > 0 Then
    Text6.Text = blm.party1(Val(Text5.Text))
    If Text6.Text = "NOT FOUND" Then
        MsgBox "Invalid A/c Code...."
        Cancel = True
    Else
        ledgerbal
    End If
Else
    MsgBox "Please Give Some A/c Code...."
    Cancel = True
End If
End Sub


Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo3.SetFocus
If KeyCode = vbKeyF2 Then
    
    Search1.Text3.Text = 5
    Search1.Show
End If

If Shift = vbAltMask Then
    If KeyCode >= 48 And KeyCode <= 57 Then
        crow = Val(Chr(KeyCode))
       ' MsgBox crow
        Grid1_DblClick
    End If
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text9_Validate(Cancel As Boolean)
Dim b As Boolean

If Val(Text9.Text) > 0 Then
    Text10.Text = blm.Item1(Val(Text9.Text))
    If Text10.Text = "NOT FOUND" Then
        MsgBox "Invalid Item Code...."
        Cancel = True
    Else
        Text16.Text = Format(blm.ITEMstocks(Val(Text9.Text), Date1.Value), "#.000")
    End If
Else
    MsgBox "Please Give Some Item Code...."
    'Cancel = True
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim amt As Currency
Dim wt As Currency
Dim Qty As Double


Dim p As Integer
Dim f As Integer, s As Integer

If Len(Text13.Text) > 0 Then
If Trim(Text13.Text) <> "NIL" Then
p = InStr(1, Text13.Text, "/", vbBinaryCompare)
f = Val(Mid(Text13.Text, 1, p - 1))
End If
End If
'MsgBox "First " & f
s = Val(Mid(Text13.Text, p + 1, Len(Text13.Text) - (p)))
'MsgBox "Second " & s
'If Trim(Text10.Text) = "3/5" Then
'    MsgBox Val(Text7.Text) & " " & Val(Text9.Text)
If f > 0 And s > 0 Then
    If s = 5 Then
        Text22.Text = Round((Val(Text12.Text) / 400) * f)
    Else
        Text22.Text = Round((Val(Text12.Text) / 800) * f)
        'MsgBox "Test"
    End If
    
End If


If Grid1.Rows > 1 Then
    Command1.Enabled = True
    
Else
    Command1.Enabled = False
End If
Text17.Text = Val(Text11.Text) * Val(Text12.Text)
If Val(Text17.Text) > 0 Then
    Command5.Enabled = True
Else
    Command5.Enabled = False
End If
For i = 1 To Grid1.Rows - 1
    amt = amt + Val(Grid1.TextMatrix(i, 6))
    wt = wt + Val(Grid1.TextMatrix(i, 3))
    Qty = Qty + Val(Grid1.TextMatrix(i, 4))
    
Next i
Text18.Text = Format(amt, "#.000")
Text19.Text = Format(Val(Text18.Text), "#.000")
Label9.Caption = wt
Label14.Caption = Qty
End Sub
