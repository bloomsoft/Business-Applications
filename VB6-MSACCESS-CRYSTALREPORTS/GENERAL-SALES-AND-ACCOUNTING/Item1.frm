VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Item1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Information"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Supporting Items"
      Height          =   2535
      Left            =   240
      TabIndex        =   26
      Top             =   7170
      Visible         =   0   'False
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   1335
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "(F1 to Clear, F2 to Accept,F3 to Search)"
         Height          =   375
         Left            =   4680
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Per"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Qty"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Group Information"
      Height          =   855
      Left            =   240
      TabIndex        =   22
      Top             =   1440
      Width           =   6255
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Text            =   "Combo2"
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "Select Group"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   270
      TabIndex        =   20
      Top             =   4125
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   975
         Left            =   3720
         Picture         =   "Item1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   975
         Left            =   2520
         Picture         =   "Item1.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   975
         Left            =   1320
         Picture         =   "Item1.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1815
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   6255
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Top             =   975
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Dozen"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label14 
         Caption         =   "Bales"
         Height          =   270
         Left            =   2880
         TabIndex        =   35
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label13 
         Caption         =   "Price"
         Height          =   255
         Left            =   2880
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "(F4) to Search Item to Edit"
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Stock"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Unit"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Item Description"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1455
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   6255
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   915
         Left            =   3720
         Picture         =   "Item1.frx":0F56
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   915
         Left            =   1320
         Picture         =   "Item1.frx":147C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   5760
      TabIndex        =   21
      Top             =   7920
      Width           =   735
   End
End
Attribute VB_Name = "Item1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Dim PDESC As String
Private Sub Clear1()
Text4.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = vbNullString
Text10.Text = ""
End Sub

Private Sub Transfer1()
With GRID1
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = Text4.Text
    .TextMatrix(.Rows - 1, 1) = Text6.Text
    .TextMatrix(.Rows - 1, 2) = Text7.Text
    .TextMatrix(.Rows - 1, 3) = Text8.Text
End With
End Sub
Private Sub Flex1()
With GRID1
    .Rows = 1
    .Cols = 4
    .ColWidth(0) = 1200
    .ColWidth(1) = 2800
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Qty"
    .TextMatrix(0, 3) = "Per"
End With
End Sub

Public Function group(c As Integer) As Integer
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select groupC,name,unit  from item where code = " & c
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
'MsgBox c
If Not tb.EOF Then
    group = tb.Fields("Group").Value
    Text3.Text = tb.Fields("Unit").Value
    
Else
    group = 0
End If
tb.Close
db.Close

End Function
Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim R As Integer
Dim ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
    ssql = "delete from Item where code = " & Val(Text1.Text)
    db.Execute ssql
    
    ssql = "delete from RawItem where code = " & Val(Text1.Text)
    db.Execute ssql
    
End If


Set tb = db.OpenRecordset("Item", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    
    tb.Fields("Unit").Value = UCase(CStr(Text3.Text))
    tb.Fields("Stock").Value = Val(Text5.Text)
    tb.Fields("GroupC").Value = Combo2.ItemData(Combo2.ListIndex)
    tb.Fields("Rate").Value = Val(Text9.Text)
    tb.Fields("bales").Value = Val(Text10.Text)
    
tb.Update
tb.Close

Set tb = db.OpenRecordset("RawItem", dbOpenTable)
For R = 1 To GRID1.Rows - 1
tb.AddNew
    tb.Fields("RawCODE").Value = Val(GRID1.TextMatrix(R, 0))
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("Qty").Value = Val(GRID1.TextMatrix(R, 2))
    tb.Fields("Unit").Value = Val(GRID1.TextMatrix(R, 3))
tb.Update
Next R
tb.Close
db.Close

End Sub
Private Function check(s As String) As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM Item WHERE NAME = '" & s & "'"
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)

If Not tb.EOF Then
    check = True
Else
    check = False
End If
tb.Close
db.Close
End Function
Private Function max1() As Long
Dim db As Database
Dim tb As Recordset
Dim ssql As String
If Combo2.ListCount > 0 Then
ssql = "select MAX(CODE) AS C FROM Item where Code Between " & Combo2.ItemData(Combo2.ListIndex) * 10000
ssql = ssql & " and " & Combo2.ItemData(Combo2.ListIndex) * 10000 + 10000

Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = (Combo2.ItemData(Combo2.ListIndex) * 10000) + 1
End If
tb.Close
db.Close
End If
End Function

Private Sub Combo1_Change()
If Option2 = True Then
Text2.Text = Combo1.Text
End If
End Sub

Private Sub Combo1_Click()
If Option2 = True Then
Text1.Text = Combo1.ItemData(Combo1.ListIndex)
Text1.Enabled = False
Text2.Text = Combo1.Text
'Combo2.ListIndex = group(Combo1.ItemData(Combo1.ListIndex)) - 1
For i = 0 To Combo2.ListCount - 1
       If Combo2.ItemData(i) = Val(Mid(Text1.Text, 1, 2)) Then
                Combo2.ListIndex = i
                Exit For
        End If
Next i
Text3.Text = UnitRet
Text4.Text = RateRet
Text5.Text = StockRet
Text6.Text = OLDCODERet
Text7.Text = PurRateRet
'edit1
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus 'SendKeys ("{TAB}")
End Sub

Private Sub Combo2_Click()
If Option1 = True Then
Text1.Text = max1
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Command1_Click()
Call save

Combs
Command2_Click

If Option1 = True Then
Text2.SetFocus
Else
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
PDESC = Text2.Text

Text2.Text = vbNullString
Text9.Text = ""
Text5.Text = vbNullString
GRID1.Rows = 1
Clear1
If Option2 = True Then
    Text1.Enabled = True
    Text1.Text = vbNullString
Else
    Text1.Enabled = False
    Text1.Text = max1

End If
Command1.Enabled = False
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

Private Sub Combs()
End Sub

Private Sub Form_Activate()
Me.Top = 10
Me.Left = 10
Dim ssql As String
ssql = "select * from Groups Order by Code"
blm.fill_comb ssql, Combo2, "Name", "Code"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
Clear1
Text4.SetFocus
End If
If KeyCode = vbKeyF2 Then
    Transfer1
    Clear1
    Text4.SetFocus
End If
If KeyCode = vbKeyF3 Then
                Screen.MousePointer = vbHourglass
                Search1.Text3.Text = 3
                Search1.Show
                Screen.MousePointer = vbDefault
End If
If Option2 = True Then
If KeyCode = vbKeyF4 Then
                Screen.MousePointer = vbHourglass
                Search1.Text3.Text = 2
                Search1.Show
                Screen.MousePointer = vbDefault
End If
End If
If KeyCode = vbKeyF12 Then


    Text2.Text = PDESC
End If
    
End Sub

Private Sub Form_Load()
Set blm = New bloom1
Dim ssql As String
ssql = "select * from Groups Order by Code"
blm.fill_comb ssql, Combo2, "Name", "Code"
Flex1
Combs
Text1.Text = max1
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub Grid1_DblClick()
If GRID1.Rows > 1 Then
With GRID1
    Text4.Text = .TextMatrix(.Row, 0)
    Text6.Text = .TextMatrix(.Row, 1)
    Text7.Text = .TextMatrix(.Row, 2)
    Text8.Text = .TextMatrix(.Row, 3)
    If GRID1.Rows = 2 Then
        GRID1.Rows = 1
    Else
        GRID1.RemoveItem .Row
    End If
End With
End If
End Sub

Private Sub Option1_Click()
Command2_Click
If Option1 = True Then
    Combo2.Enabled = True
    Text1.Enabled = False
    Label12.Visible = False
    Text1.Text = max1
    Text2.Text = vbNullString
    Text2.SetFocus
    
Else
    Text1.Enabled = True
End If
End Sub

Private Sub Option2_Click()
Command2_Click
If Option2 = True Then

    'Combo2.Enabled = False
    Text1.Enabled = True
    Label12.Visible = True
    Text1.Text = vbNullString
     Text1.SetFocus
    
Else
    Text1.Enabled = False
End If

End Sub

Private Function edit1() As Boolean
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(blm.pathMain)
ssql = "select * from item where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    Text1.Text = tb.Fields("Code").Value & ""
    Text2.Text = tb.Fields("Name").Value & ""
   
    Text5.Text = tb.Fields("Stock").Value & ""
    Text3.Text = tb.Fields("Unit").Value & ""
    Text9.Text = tb.Fields("Rate").Value & ""
    Text10.Text = tb.Fields("Bales").Value & ""
    edit1 = False
Else
    MsgBox "Invalid Item Code....."
    edit1 = True
    Exit Function
End If
tb.Close

ssql = "select * from Rawitem where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
GRID1.Rows = 1
Do While Not tb.EOF
GRID1.Rows = GRID1.Rows + 1
  GRID1.TextMatrix(GRID1.Rows - 1, 0) = tb.Fields("RawCode").Value & ""
  GRID1.TextMatrix(GRID1.Rows - 1, 1) = blm.Item1(tb.Fields("RawCode").Value) & ""
  GRID1.TextMatrix(GRID1.Rows - 1, 2) = tb.Fields("Qty").Value & ""
  GRID1.TextMatrix(GRID1.Rows - 1, 3) = tb.Fields("Unit").Value & ""
tb.MoveNext
Loop
End If
tb.Close
db_m.Close

End Function
Private Function UnitRet() As String
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(blm.pathMain)
ssql = "select unit from item where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    UnitRet = tb.Fields("Unit").Value & ""
    
End If
tb.Close
db_m.Close
End Function
Private Function StockRet() As Currency
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(blm.pathMain)
ssql = "select Stock from item where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    StockRet = tb.Fields("Stock").Value
    
End If
tb.Close
db_m.Close
End Function

Private Function RateRet() As Currency
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(blm.pathMain)
ssql = "select Rate from item where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    RateRet = tb.Fields("Rate").Value
    
End If
tb.Close
db_m.Close
End Function
Private Function OLDCODERet() As Currency
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(blm.pathMain)
ssql = "select OLDCODE from item where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    OLDCODERet = tb.Fields("OLDCODE").Value
    
End If
tb.Close
db_m.Close
End Function

Private Function PurRateRet() As Currency
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(blm.pathMain)
ssql = "select PurRate from item where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    PurRateRet = tb.Fields("PurRate").Value
    
End If
tb.Close
db_m.Close
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
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
Dim i As Long
If Option2 = True Then
If Val(Text1.Text) > 0 Then
        Cancel = edit1
End If
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
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

Private Sub Text2_Change()
If Text2.Text <> vbNullString Or Text2.Text <> "" Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")

End Sub

Private Sub Text2_LostFocus()
If Option1 = True Then
Dim b As Boolean
If Len(Text2.Text) > 0 Then
b = check(UCase(CStr(Text2.Text)))
If b = True Then
    'MsgBox "ITEM ALREADY EXIST,,,,"
    'Text2.SetFocus
End If
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
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
    Text6.Text = blm.Item1(Val(Text4.Text))
    If Text6.Text = "NOT" Then
        MsgBox "Invalid Supporting Item Code...."
        Cancel = True
    End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
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

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys ("{TAB}")
        End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
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
