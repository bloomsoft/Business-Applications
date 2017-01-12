VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form In1 
   Caption         =   "Inward Note for Knitting Contracts"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text19 
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
      Left            =   8640
      TabIndex        =   59
      Top             =   6240
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
      TabIndex        =   57
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7440
      Top             =   4320
   End
   Begin VB.Frame Frame6 
      Caption         =   "Options"
      Height          =   1335
      Left            =   7800
      TabIndex        =   51
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   855
         Left            =   2400
         Picture         =   "In1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   855
         Left            =   600
         Picture         =   "In1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Actions"
      Height          =   855
      Left            =   7800
      TabIndex        =   47
      Top             =   6480
      Width           =   3975
      Begin VB.CommandButton Command6 
         Caption         =   "&Print"
         Height          =   375
         Left            =   2880
         TabIndex        =   60
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   2040
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   1200
         TabIndex        =   49
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   375
         Left            =   360
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lists"
      Height          =   1215
      Left            =   120
      TabIndex        =   42
      Top             =   6120
      Width           =   7575
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1440
         TabIndex        =   46
         Text            =   "Combo3"
         Top             =   720
         Width           =   5655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   44
         Text            =   "Combo2"
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label20 
         Caption         =   "Items List"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "A/c List"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   41
      Top             =   3840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   11655
      Begin VB.CommandButton Command8 
         Caption         =   "Ra&te Set"
         Height          =   375
         Left            =   7920
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   6960
         TabIndex        =   63
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   55
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   10560
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Entr&y"
         Height          =   375
         Left            =   9120
         TabIndex        =   20
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
         Left            =   10560
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7920
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Rate Set"
         Height          =   255
         Left            =   6000
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Amount"
         Height          =   255
         Left            =   2400
         TabIndex        =   54
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Stock"
         Height          =   255
         Left            =   10680
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Bobins"
         Height          =   255
         Left            =   6000
         TabIndex        =   39
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Ware House"
         Height          =   255
         Left            =   9240
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "LOT#"
         Height          =   255
         Left            =   8040
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Bag #"
         Height          =   255
         Left            =   7080
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Rate"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Note"
      Height          =   855
      Left            =   7800
      TabIndex        =   21
      Top             =   1440
      Width           =   3975
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton Command7 
         Caption         =   "&DELETE THIS INWARD"
         Height          =   375
         Left            =   5160
         TabIndex        =   61
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6600
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   3720
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24576003
         CurrentDate     =   36214
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Goods Name"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Due Days"
         Height          =   255
         Left            =   5280
         TabIndex        =   27
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Bilty #"
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Gate Pass #"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Balance in Ledger"
         Height          =   255
         Left            =   5280
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "InWard #"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label23 
      Caption         =   "Total Wt."
      Height          =   255
      Left            =   7800
      TabIndex        =   58
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Total"
      Height          =   255
      Left            =   9960
      TabIndex        =   56
      Top             =   6240
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
Dim ssql As String
Set db = OpenDatabase(blm.pathMain)
ssql = "select * from in_mst where p_no = " & Val(Text1.Text)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("v_date").Value
    Text2.Text = tb.Fields("Gate_no").Value
    Text3.Text = tb.Fields("Bilty_no").Value
    Text4.Text = tb.Fields("Due_Days").Value
    Text5.Text = tb.Fields("Party").Value
    Text6.Text = blm.party1(tb.Fields("Party").Value)
    Text7.Text = tb.Fields("Goods").Value
    Text8.Text = tb.Fields("Note").Value
    edit1 = False
Else
    MsgBox "No Record For This Purchase No."
    edit1 = True
    Exit Function
End If
tb.Close
Dim i As Long
ssql = "select * from in_dtl where p_no = " & Val(Text1.Text)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
Do While Not tb.EOF
Grid1.Rows = Grid1.Rows + 1
i = Grid1.Rows - 1
    With Grid1
        .TextMatrix(i, 0) = i
        .TextMatrix(i, 1) = tb.Fields("Item").Value
        .TextMatrix(i, 2) = blm.Item1(tb.Fields("Item").Value)
        .TextMatrix(i, 3) = Format(tb.Fields("Quantity").Value, "#.000")
        .TextMatrix(i, 4) = Format(tb.Fields("Rate").Value, "#.000")
        .TextMatrix(i, 5) = Format(tb.Fields("Quantity").Value * tb.Fields("Rate").Value, "#.000")
        .TextMatrix(i, 6) = tb.Fields("Bobins").Value
        .TextMatrix(i, 7) = tb.Fields("CTN_No").Value
        .TextMatrix(i, 8) = tb.Fields("LOT_no").Value
        .TextMatrix(i, 9) = tb.Fields("W_H_Code").Value
        .TextMatrix(i, 10) = blm.WareHouse(tb.Fields("W_H_Code").Value)
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
Dim ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
ssql = "delete from voucher where e_type=2 and ent_no = " & Val(Text1.Text)
db.Execute ssql
ssql = "delete from in_mst where p_no = " & Val(Text1.Text)
db.Execute ssql
ssql = "delete from in_DTL where p_no = " & Val(Text1.Text)
db.Execute ssql
End If

Set tb = db.OpenRecordset("voucher", dbOpenTable)
tb.AddNew
    tb.Fields("ent_no").Value = Val(Text1.Text)
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("e_type").Value = 2
    tb.Fields("party").Value = Val(Text5.Text)
    tb.Fields("debit").Value = 0
    tb.Fields("credit").Value = Val(Text18.Text)
    tb.Fields("Remarks").Value = "Inward # " & Val(Text1.Text) & " Bilty = " & CStr(Text3.Text)
tb.Update
tb.Close
    
Set tb = db.OpenRecordset("in_mst", dbOpenTable)
tb.AddNew
    tb.Fields("P_no").Value = Val(Text1.Text)
    tb.Fields("V_Date").Value = date1.Value
    tb.Fields("Gate_no").Value = CStr(Text2.Text)
    tb.Fields("Bilty_no").Value = CStr(Text3.Text)
    tb.Fields("Due_Days").Value = Val(Text4.Text)
    tb.Fields("Party").Value = Val(Text5.Text)
    tb.Fields("Goods").Value = UCase(CStr(Text7.Text))
    'If Len(Text8.Text) = 0 Then
    tb.Fields("Note").Value = Text8.Text & " "
tb.Update
tb.Close
Dim i As Long
Set tb = db.OpenRecordset("in_dtl", dbOpenTable)
For i = 1 To Grid1.Rows - 1
    With Grid1
tb.AddNew
    tb.Fields("P_no").Value = Val(Text1.Text)
    tb.Fields("V_Date").Value = date1.Value
    tb.Fields("Item").Value = Val(.TextMatrix(i, 1))
    tb.Fields("Quantity").Value = Val(.TextMatrix(i, 3))
    tb.Fields("Rate").Value = Val(.TextMatrix(i, 4))
    tb.Fields("Bobins").Value = Val(.TextMatrix(i, 6))
    tb.Fields("CTN_no").Value = Val(.TextMatrix(i, 7))
    tb.Fields("LOT_no").Value = Val(.TextMatrix(i, 8))
'    MsgBox Val(.TextMatrix(i, 9))
    tb.Fields("W_H_Code").Value = Val(.TextMatrix(i, 9))
    
tb.Update
    End With
Next i
tb.Close
db.Close
End Sub
Private Sub clear1()
Text9.Text = vbNullString
Text10.Text = vbNullString
Text11.Text = vbNullString
Text12.Text = vbNullString
Text13.Text = vbNullString
Text14.Text = vbNullString
Text15.Text = vbNullString
Text16.Text = vbNullString
Text17.Text = vbNullString
End Sub
Private Sub transfer1()
Dim i As Long
With Grid1
    .Rows = .Rows + 1
    i = .Rows - 1
    .TextMatrix(i, 0) = i
    .TextMatrix(i, 1) = Text9.Text
    .TextMatrix(i, 2) = Text10.Text
    .TextMatrix(i, 4) = Format(Val(Text11.Text), "#.000")
    .TextMatrix(i, 3) = Format(Val(Text12.Text), "#.000")
    .TextMatrix(i, 5) = Format(Val(Text17.Text), "#.000")
    .TextMatrix(i, 6) = Text13.Text
    .TextMatrix(i, 7) = Text14.Text
    .TextMatrix(i, 8) = Text15.Text
    .TextMatrix(i, 9) = Combo1.ItemData(Combo1.ListIndex)
    .TextMatrix(i, 10) = Combo1.Text
End With
Grid1.TopRow = Grid1.Rows - 1
    Text9.SetFocus
End Sub
Private Sub Clearfull()

Dim CNTL As Control

For Each CNTL In Me.Controls
    If TypeOf CNTL Is TextBox Then CNTL.Text = vbNullString
    If TypeOf CNTL Is DTPicker Then CNTL.Value = Date - 1
Next
Combs
flex1
If Option1 = True Then
    Text1.Text = MAX1
End If
End Sub

Private Function MAX1() As Long
Dim db As Database
Dim tb As Recordset
Dim ssql As String
ssql = "select MAX(P_no) AS C FROM In_Mst"
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    MAX1 = tb.Fields("C").Value + 1
Else
    MAX1 = 1
End If
tb.Close
db.Close
End Function

Private Sub Combs()
Dim ssql As String

'WareHouses
ssql = "select * from WareHouse order by Name"
blm.fill_comb ssql, Combo1, "Name", "Code"
'Accounts
ssql = "select * from Acchart where Code < 5000 order by Name"
blm.fill_comb ssql, Combo2, "Name", "Code"
'Items
ssql = "select * from Item order by Name"
blm.fill_comb ssql, Combo3, "Name", "Code"

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

Private Sub flex1()
With Grid1
    .Rows = 1
    .Cols = 11
    .ColWidth(0) = 500
    .TextMatrix(0, 0) = "Sr #"
    .ColWidth(1) = 1000
    .TextMatrix(0, 1) = "Item Code"
    .ColWidth(2) = 2500
    .TextMatrix(0, 2) = "Item Name"
    .ColWidth(3) = 1000
    .TextMatrix(0, 3) = "Quantity"
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "Rate"
    .ColWidth(5) = 1000
    .TextMatrix(0, 5) = "Amount"
    .ColWidth(6) = 1500
    .TextMatrix(0, 6) = "Bobins"
    .ColWidth(7) = 1000
    .TextMatrix(0, 7) = "CTN#"
    .ColWidth(8) = 1000
    .TextMatrix(0, 8) = "LOT#"
    .ColWidth(9) = 10  'warehouse Code
    .ColWidth(10) = 1700
    .TextMatrix(0, 10) = "WareHouse"
End With
End Sub

Private Sub Command1_Click()
Call save
Command2_Click

End Sub

Private Sub Command2_Click()
Call Clearfull
If Option1 = True Then
    Text2.SetFocus
Else
    Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Command4_Click()
Call clear1
Combo3_Click
Text9.SetFocus
End Sub

Private Sub Command5_Click()
'If Val(Text13.Text) = 0 Then
'    MsgBox "Please Give Bobins..."
'    Exit Sub
'End If
If Val(Text14.Text) = 0 Then
    MsgBox "Please Give CTN no..."
    Exit Sub
End If
If Val(Text15.Text) = 0 Then
    MsgBox "Please Give Lot No..."
    Exit Sub
End If

Call transfer1
Text14.Text = Val(Text14.Text) + 1
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

Dim ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
ssql = "delete from voucher where e_type=2 and ent_no = " & Val(Text1.Text)
db.Execute ssql
ssql = "delete from in_mst where p_no = " & Val(Text1.Text)
db.Execute ssql
ssql = "delete from in_DTL where p_no = " & Val(Text1.Text)
db.Execute ssql
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

Private Sub Form_Load()
Combs
Text1.Text = MAX1
date1.Value = Date - 1
flex1
End Sub

Private Sub grid1_DblClick()
Dim thisrow As Long

If crow > 0 Then
    thisrow = crow
Else
    thisrow = Grid1.Row
End If
With Grid1
Text9.Text = .TextMatrix(thisrow, 1)
Text10.Text = .TextMatrix(thisrow, 2)
Text11.Text = .TextMatrix(thisrow, 4)
Text12.Text = .TextMatrix(thisrow, 3)
Text13.Text = .TextMatrix(thisrow, 6)
Text14.Text = .TextMatrix(thisrow, 7)
Text15.Text = .TextMatrix(thisrow, 8)
Combo1.ListIndex = Val(.TextMatrix(thisrow, 9)) - 1
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
Text9.SetFocus
End Sub

Private Sub Option1_Click()
Clearfull
Command7.Visible = False
Text1.Enabled = False
date1.SetFocus
End Sub

Private Sub Option2_Click()
Clearfull
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

Private Sub Text13_GotFocus()
Text13.SelStart = 0
Text13.SelLength = Len(Text13.Text)

End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
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

Private Sub Text14_GotFocus()
Text14.SelStart = 0
Text14.SelLength = Len(Text14.Text)
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
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

Private Sub Text14_Validate(Cancel As Boolean)
If Val(Text14.Text) > 0 Then
    Cancel = blm.CTNCheck(Val(Text14.Text))
    If Cancel = True Then
        Text14.Text = vbNullString
        MsgBox "Cotton No Already Exist...."
    End If
Else
    Cancel = True
End If

End Sub

Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15.Text)
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
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

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
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

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)

End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
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

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo3.SetFocus
If Shift = vbAltMask Then
    If KeyCode >= 48 And KeyCode <= 57 Then
        crow = Val(Chr(KeyCode))
       ' MsgBox crow
        grid1_DblClick
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
        Text16.Text = Format(blm.Stock(Val(Text9.Text), date1.Value), "#.000")
    End If
Else
    MsgBox "Please Give Some Item Code...."
    Cancel = True
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim amt As Currency
Dim wt As Currency
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
    amt = amt + Val(Grid1.TextMatrix(i, 5))
    wt = wt + Val(Grid1.TextMatrix(i, 3))
Next i
Text18.Text = Format(amt, "#.000")
Text19.Text = Format(wt, "#.000")
End Sub
