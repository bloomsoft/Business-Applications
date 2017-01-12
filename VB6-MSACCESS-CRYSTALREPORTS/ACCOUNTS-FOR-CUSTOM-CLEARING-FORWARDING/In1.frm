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
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   46
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton Command7 
         Caption         =   "&DELETE THIS INWARD"
         Height          =   375
         Left            =   5160
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   3720
         TabIndex        =   44
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   36214
      End
      Begin VB.Label Label4 
         Caption         =   "INV_TYPE=1 FOR BUNYAN PURCHASE 3 for Towel, 2 for Socks"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   1080
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Label Label8 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Balance in Ledger"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "InWard #"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
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
      Left            =   8640
      TabIndex        =   42
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
      TabIndex        =   40
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
      Height          =   735
      Left            =   7800
      TabIndex        =   34
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   36
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Actions"
      Height          =   1215
      Left            =   7800
      TabIndex        =   30
      Top             =   6000
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2640
         Picture         =   "In1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1800
         Picture         =   "In1.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   960
         Picture         =   "In1.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lists"
      Height          =   1575
      Left            =   120
      TabIndex        =   25
      Top             =   5640
      Width           =   7575
      Begin VB.ComboBox Combo3 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   29
         Text            =   "Combo3"
         Top             =   720
         Width           =   5655
      End
      Begin VB.ComboBox Combo2 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Text            =   "Combo2"
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label20 
         Caption         =   "Items List"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "A/c List"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      CausesValidation=   0   'False
      Height          =   2415
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4260
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   11655
      Begin VB.TextBox Text17 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9960
         TabIndex        =   38
         Top             =   600
         Width           =   1470
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   10560
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Entr&y"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   9120
         TabIndex        =   11
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
         Left            =   8145
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9060
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8220
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   6840
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Amount"
         Height          =   255
         Left            =   10680
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Stock"
         Height          =   255
         Left            =   8505
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Rate"
         Height          =   255
         Left            =   8340
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   9180
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Note"
      Height          =   735
      Left            =   7800
      TabIndex        =   12
      Top             =   960
      Width           =   3975
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Label Label23 
      Caption         =   "Balance"
      Height          =   255
      Left            =   7800
      TabIndex        =   41
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Total"
      Height          =   255
      Left            =   9960
      TabIndex        =   39
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
Dim DB As Database
Dim tb As Recordset
Dim ssql As String
Set DB = OpenDatabase(blm.pathMain)
ssql = "select * from in_mst where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
Set tb = DB.OpenRecordset(ssql)
If Not tb.EOF Then
    Date1.Value = tb.Fields("v_date").Value
    Text5.Text = tb.Fields("Party").Value
    Text6.Text = blm.party1(tb.Fields("Party").Value)
    Text8.Text = tb.Fields("Note").Value
    edit1 = False
Else
    MsgBox "No Record For This Inward No."
    edit1 = True
    Exit Function
End If
tb.Close
Dim i As Long
ssql = "select * from in_dtl where p_no = " & Val(Text1.Text)
ssql = ssql & " and inv_type=" & Val(Text2.Text)
Set tb = DB.OpenRecordset(ssql)
If Not tb.EOF Then
Do While Not tb.EOF
GRID1.Rows = GRID1.Rows + 1
i = GRID1.Rows - 1
    With GRID1
        .TextMatrix(i, 0) = i
        .TextMatrix(i, 1) = tb.Fields("Item").Value
        .TextMatrix(i, 2) = blm.Item1(tb.Fields("Item").Value)
        .TextMatrix(i, 3) = Format(tb.Fields("Quantity").Value, "#.000")
        .TextMatrix(i, 4) = Format(tb.Fields("Rate").Value, "#.000")
        .TextMatrix(i, 5) = Format(tb.Fields("Quantity").Value * tb.Fields("Rate").Value, "#.000")
        .TextMatrix(i, 6) = tb.Fields("Scheme").Value
    End With
    tb.MoveNext
Loop
End If
tb.Close
DB.Close
End Function
Private Sub save()
Dim DB As Database
Dim tb As Recordset
Dim ssql As String
Set DB = OpenDatabase(blm.pathMain)
If Option2 = True Then
ssql = "delete from voucher where inv_type=" & Val(Text2.Text) & " and e_type=2 and ent_no = " & Val(Text1.Text)
DB.Execute ssql
ssql = "delete from in_mst where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
DB.Execute ssql
ssql = "delete from in_DTL where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
DB.Execute ssql

End If

Set tb = DB.OpenRecordset("voucher", dbOpenTable)
tb.AddNew
    tb.Fields("ent_no").Value = Val(Text1.Text)
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("e_type").Value = 2
    tb.Fields("party").Value = Val(Text5.Text)
    tb.Fields("debit").Value = 0
    tb.Fields("credit").Value = Round(Val(Text19.Text))
    tb.Fields("Remarks").Value = "Inward # " & Val(Text1.Text)
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text1.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("e_type").Value = 2
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("party").Value = 6000 + Val(Text2.Text) 'Stock A/c Code Should be This
    tb.Fields("debit").Value = Round(Val(Text19.Text))
    tb.Fields("credit").Value = 0
    tb.Fields("Remarks").Value = "Inward # " & Val(Text1.Text)
tb.Update

tb.Close
    
Set tb = DB.OpenRecordset("in_mst", dbOpenTable)
tb.AddNew
    tb.Fields("P_no").Value = Val(Text1.Text)
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("V_Date").Value = Date1.Value
    tb.Fields("Party").Value = Val(Text5.Text)
    tb.Fields("Note").Value = Text8.Text & " "
tb.Update
tb.Close
Dim i As Long
Set tb = DB.OpenRecordset("in_dtl", dbOpenTable)
For i = 1 To GRID1.Rows - 1
    With GRID1
tb.AddNew
    tb.Fields("P_no").Value = Val(Text1.Text)
    tb.Fields("inv_type").Value = Val(Text2.Text)
    tb.Fields("V_Date").Value = Date1.Value
    tb.Fields("Item").Value = Val(.TextMatrix(i, 1))
    tb.Fields("Quantity").Value = Val(.TextMatrix(i, 3))
    tb.Fields("Rate").Value = Val(.TextMatrix(i, 4))
    tb.Fields("Scheme").Value = Val(.TextMatrix(i, 6))
tb.Update
    End With
    
Next i
tb.Close
DB.Close
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
With GRID1
    .Rows = .Rows + 1
    i = .Rows - 1
    .TextMatrix(i, 0) = i
    .TextMatrix(i, 1) = Text9.Text
    .TextMatrix(i, 2) = Text10.Text
    .TextMatrix(i, 4) = Format(Val(Text11.Text), "#.000")
    .TextMatrix(i, 3) = Format(Val(Text12.Text), "#.000")
    .TextMatrix(i, 5) = Format(Val(Text17.Text), "#.000")

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
Dim DB As Database
Dim tb As Recordset
Dim ssql As String
ssql = "select MAX(P_no) AS C FROM In_Mst where inv_type=" & Val(Text2.Text)
'Clipboard.SetText ssql
'MsgBox ssql
Set DB = OpenDatabase(blm.pathMain)
Set tb = DB.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
DB.Close
End Function

Private Sub Combs()
Dim ssql As String

'Accounts
ssql = "select * from Parties order by Name"
blm.fill_comb ssql, Combo2, "Name", "Code"
'Items
ssql = "select * from Item order by Name"
blm.fill_comb_Item ssql, Combo3, "Name", "Code"

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
With GRID1
    .Rows = 1
    .Cols = 7
    .ColWidth(0) = 1000
    .TextMatrix(0, 0) = "Sr #"
    .ColWidth(1) = 1500
    .TextMatrix(0, 1) = "Item Code"
    .ColWidth(2) = 4000
    .TextMatrix(0, 2) = "Item Name"
    .ColWidth(3) = 1500
    .TextMatrix(0, 3) = "Quantity"
    .ColWidth(4) = 1500
    .TextMatrix(0, 4) = "Rate"
    .ColWidth(5) = 2000
    .TextMatrix(0, 5) = "Amount"
    .ColWidth(6) = 15
    .TextMatrix(0, 6) = "Scheme"
    
    
    
    
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
Dim DB As Database

Dim ssql As String
Set DB = OpenDatabase(blm.pathMain)
If Option2 = True Then
ssql = "delete from voucher where inv_type=" & Val(Text2.Text) & " and e_type=2 and ent_no = " & Val(Text1.Text)
DB.Execute ssql
ssql = "delete from in_mst where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
DB.Execute ssql
ssql = "delete from in_DTL where inv_type=" & Val(Text2.Text) & " and p_no = " & Val(Text1.Text)
DB.Execute ssql
End If
DB.Close
Command2_Click

End Sub

Private Sub Command8_Click()
Dim i As Long

For i = 1 To GRID1.Rows - 1
    GRID1.TextMatrix(i, 4) = Format(Val(Text20.Text), "#.000")
    GRID1.TextMatrix(i, 5) = Format(Val(Text20.Text) * GRID1.TextMatrix(i, 3), "#.000")
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

If crow > 0 Then
    thisrow = crow
Else
    thisrow = GRID1.Row
End If
If Val(Text12.Text) > 0 Then
    MsgBox "You Already Have Entry There ...."
Else
With GRID1
Text9.Text = .TextMatrix(thisrow, 1)
Text10.Text = .TextMatrix(thisrow, 2)
Text11.Text = .TextMatrix(thisrow, 4)
Text12.Text = .TextMatrix(thisrow, 3)

End With
'MsgBox thisrow
If GRID1.Rows = 2 Then
    GRID1.Rows = 1
Else
    GRID1.RemoveItem (thisrow)
End If
Dim i As Long

For i = 1 To GRID1.Rows - 1
    GRID1.TextMatrix(i, 0) = i
Next i
End If
Text9.SetFocus
End Sub

Private Sub Label14_Click()
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

Private Sub Text21_KeyPress(KeyAscii As Integer)
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

Private Sub Text2_Change()
Text1.Text = max1
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

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
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
If GRID1.Rows > 1 Then
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
For i = 1 To GRID1.Rows - 1
    amt = amt + Val(GRID1.TextMatrix(i, 5))
    wt = wt + Val(GRID1.TextMatrix(i, 3))
Next i
Text18.Text = Format(amt, "#.000")
Text19.Text = Format(Val(Text18.Text), "#.000")

End Sub
