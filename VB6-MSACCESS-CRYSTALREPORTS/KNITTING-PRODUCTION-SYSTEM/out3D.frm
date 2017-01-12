VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form out3D 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OutWard Gate Pass of Purchase Dyeing Contract"
   ClientHeight    =   7080
   ClientLeft      =   2265
   ClientTop       =   615
   ClientWidth     =   7095
   Icon            =   "out3D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   855
      Left            =   240
      TabIndex        =   48
      Top             =   4560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1508
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sent Items Info"
      Height          =   1455
      Left            =   240
      TabIndex        =   37
      Top             =   5520
      Width           =   6615
      Begin VB.Label Label36 
         Caption         =   "Label36"
         Height          =   255
         Left            =   2040
         TabIndex        =   47
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label35 
         Caption         =   "Fabric"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "Label32"
         Height          =   255
         Left            =   4200
         TabIndex        =   45
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Cloth Rolls Sent"
         Height          =   375
         Left            =   2760
         TabIndex        =   44
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Label30"
         Height          =   255
         Left            =   2040
         TabIndex        =   43
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Cloth Sent Qty."
         Height          =   375
         Left            =   840
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "Label28"
         Height          =   255
         Left            =   4200
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Cloth Rec. Rolls"
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Label26"
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Cloth Rec. Qty."
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1095
      Left            =   3360
      TabIndex        =   31
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   2280
         Picture         =   "out3D.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   735
         Left            =   1200
         Picture         =   "out3D.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   735
         Left            =   120
         Picture         =   "out3D.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   28
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         Height          =   735
         Left            =   1800
         Picture         =   "out3D.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   735
         Left            =   240
         Picture         =   "out3D.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   6615
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   5040
         TabIndex        =   11
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5640
         Top             =   4680
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   2880
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5640
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   36749
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Del. Date"
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Rate"
         Height          =   255
         Left            =   5040
         TabIndex        =   25
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   4920
         TabIndex        =   24
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Contract Date"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Cloth Rolls"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Quality"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Out ward #"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "out3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim org_q As Currency
Dim rej As Currency
Private Type ContInfo
    clothsent As Long
    yarnsent As Long
    clothrec As Long
    yarnrec As Long
    lycrarec As Long
    lbagsrec As Long
    lbagssent As Long
    lycrasent As Long
    sclothrolls As Long
    syarnbags As Long
    rclothrolls As Long
    ryarnbags As Long
    YarnCount As String
    Cloth As String
    
End Type
Private Sub clear()
Text10.Text = vbNullString
Text11.Text = vbNullString
End Sub
Private Sub Transfer1()
    With Grid1
        .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Text2.Text
            .TextMatrix(.Rows - 1, 1) = Text10.Text
            .TextMatrix(.Rows - 1, 2) = Text11.Text
    End With
End Sub
    
Private Sub Flex1()
With Grid1
    .Rows = 1
    .Cols = 3
    .ColWidth(0) = 1000
    .TextMatrix(0, 0) = "Cont. #"
    .ColWidth(1) = 1000
    .TextMatrix(0, 1) = "Rolls"
    .ColWidth(2) = 1000
    .TextMatrix(0, 2) = "Quantity"
    
End With
End Sub
Private Function Getinfo(c As Long, e As Byte) As ContInfo
    Dim db As Database
    Dim tb As Recordset
    Dim ssql As String
    Dim cn As ContInfo
    
    Set db = OpenDatabase(blm.pathMain)
    
    ssql = "select sum(quantity)as q,sum(bags)as b,sum(lycra)as l,sum(l_bags)as lb from inward where cont_no = " & c
    ssql = ssql & " and e_type= " & e
    
    
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("Q").Value) Then
        cn.yarnrec = tb.Fields("Q").Value
    End If
    If Not IsNull(tb.Fields("B").Value) Then
        cn.ryarnbags = tb.Fields("B").Value
    End If
    If Not IsNull(tb.Fields("l").Value) Then
        cn.lycrarec = tb.Fields("l").Value
    End If
    If Not IsNull(tb.Fields("lb").Value) Then
        cn.lbagsrec = tb.Fields("lb").Value
    End If
    
    
    tb.Close
    
    ssql = "select sum(quantity)as q,sum(Bags)as r from inward where cont_no = " & c
    ssql = ssql & " and e_type= " & e
    
    
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("Q").Value) Then
        cn.clothsent = tb.Fields("Q").Value
    End If
    If Not IsNull(tb.Fields("r").Value) Then
        cn.sclothrolls = tb.Fields("r").Value
    End If
    tb.Close
    
    ssql = "select sum(quantity)as q,sum(rolls)as r from outward where cont_no = " & c
    ssql = ssql & " and e_type= " & e
    
    
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("Q").Value) Then
        
        
        cn.clothrec = tb.Fields("Q").Value
    End If
    If Not IsNull(tb.Fields("r").Value) Then
        
        
        cn.rclothrolls = tb.Fields("r").Value
    End If
    tb.Close
    
    db.Close
    
    
    Getinfo = cn

End Function


Private Function Check(c As Long) As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String
    
ssql = "select * from inward where in_no = " & c
ssql = ssql & " and E_type=3"
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    MsgBox "Outward No already Exist..."
    Check = True
Else
    Check = False
End If
tb.Close
db.Close
End Function


Private Function edit1() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim u As ContInfo

ssql = "select * from cont_1 where cont_no = " & Val(Text2.Text)
ssql = ssql & " and e_type = 3"
org_q = 0
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date3.Value = tb.Fields("v_date").Value
    Text3.Text = tb.Fields("party").Value
    Text4.Text = blm.party1(tb.Fields("party").Value)
    Label21.Caption = Format(tb.Fields("del_date").Value, "dd/MM/yyyy")
    Label23.Caption = Format(tb.Fields("Rate").Value, "#.00")
    org_q = tb.Fields("Cquantity").Value
    Label13.Caption = Format(tb.Fields("CQuantity").Value, "#.00")

'    Label34.Caption = blm.Yarn(tb.Fields("YarnCount").Value)
    Label36.Caption = blm.Cloth(tb.Fields("Item").Value)
    Text9.Text = tb.Fields("item").Value
    Text8.Text = blm.Cloth(tb.Fields("item").Value)
    If Not IsNull(tb.Fields("complete").Value) Then
    If tb.Fields("Complete").Value = 1 Then
        MsgBox "You Have Marked this Contract as Completed....."
        
    End If
    End If
    edit1 = True
    u = Getinfo(Val(Text2.Text), 3)
    'Label16.Caption = Format(u.yarnrec, "#.00")
    'Label20.Caption = Format(u.ryarnbags, "#.00")
    Label26.Caption = Format(u.clothrec, "#.00")
    Label28.Caption = Format(u.rclothrolls, "#.00")
    Label30.Caption = Format(u.clothsent, "#.00")
    Label32.Caption = Format(u.sclothrolls, "#.00")
    
Else
    MsgBox "Not Found ...!"
    edit1 = False
End If
tb.Close
db.Close

    
End Function
Private Function max1() As Double
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
    
    ssql = "select max(in_no)as c from inward where e_type=3"
    Set db = OpenDatabase(blm.pathMain)
    Set tb = db.OpenRecordset(ssql)
    If IsNull(tb.Fields("c").Value) = False Then
        max1 = tb.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    tb.Close
    db.Close
End Function
Private Function edit_kachi() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim i As Integer

ssql = "select * from inward where in_no = " & Val(Text1.Text)
ssql = ssql & " and e_type=3"

Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("in_date").Value
    Text17.Text = tb.Fields("remarks").Value & ""
'    If Not IsNull(tb.Fields("cancel").Value) Then
'        Check1.Value = tb.Fields("cancel").Value
'    End If
'    If Not IsNull(tb.Fields("c_date").Value) Then
'        date4.Value = tb.Fields("c_date").Value
'    End If
Do While Not tb.EOF
    Grid1.Rows = Grid1.Rows + 1
     Grid1.TextMatrix(Grid1.Rows - 1, 0) = tb.Fields("cont_no").Value
    Grid1.TextMatrix(Grid1.Rows - 1, 1) = tb.Fields("Bags").Value
    Grid1.TextMatrix(Grid1.Rows - 1, 2) = tb.Fields("Quantity").Value
    tb.MoveNext
Loop
    
    edit_kachi = True
Else
    MsgBox "Not Found ...!"
    edit_kachi = False
End If
tb.Close
db.Close

    
End Function

Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim i As Integer
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
    ssql = "delete from inward where in_no = " & Val(Text1.Text)
    ssql = ssql & " and e_type = 3"
    db.Execute ssql
End If
db.Close
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset("inward", dbOpenTable)
For i = 1 To Grid1.Rows - 1
tb.AddNew
    tb.Fields("in_no").Value = Val(Text1.Text)
    tb.Fields("in_date").Value = date1.Value
    tb.Fields("E_Type").Value = 3
    
    
    With Grid1
    tb.Fields("cont_no").Value = Val(.TextMatrix(i, 0))
    tb.Fields("Bags").Value = Val(.TextMatrix(i, 1))
    tb.Fields("quantity").Value = Val(.TextMatrix(i, 2))
    
    End With
    tb.Fields("remarks").Value = CStr(Text17.Text)
    If Option2 = True Then
        tb.Fields("cancel").Value = Check1.Value
        tb.Fields("c_date").Value = date4.Value
    End If
tb.Update
Next i
tb.Close
db.Close

    
End Sub

Private Sub Clearfull()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
Label23.Caption = vbNullString
Label21.Caption = vbNullString
End Sub

Private Sub Command1_Click()
Call save
Command2_Click

End Sub

Private Sub Command2_Click()
Call Clearfull
Flex1
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide

End Sub

Private Sub Command4_Click()
Transfer1
clear
Text10.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = ((Screen.Height - Me.Height) / 2) - 1000
Me.Left = (Screen.Width - Me.Width) / 2
date1.Value = Date

date3.Value = Date
Flex1
End Sub

Private Sub Grid1_DblClick()
If Grid1.Rows > 1 Then
    With Grid1
    Text2.Text = .TextMatrix(.Row, 0)
    Text10.Text = .TextMatrix(.Row, 1)
    Text11.Text = .TextMatrix(.Row, 2)
    If .Rows > 2 Then
        .RemoveItem .Row
    Else
        .Rows = 1
    End If
    End With
    Text2.SetFocus
End If
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Option1_Click()
'Text1.Enabled = False
Check1.Visible = False
date4.Visible = False
Text2.SetFocus
End Sub

Private Sub Option2_Click()
'Text1.Enabled = True


Text1.SetFocus

End Sub

Private Sub Text1_GotFocus()
If Option1 = True Then
    Text1.Text = max1
End If
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
Dim b As Boolean
If Val(Text1.Text) > 0 Then
    If Option1 = True Then
        b = Check(Val(Text1.Text))
        Cancel = b
    End If
    If Option2 = True Then
        b = edit_kachi
        If b = False Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
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

Private Sub Text10_Validate(Cancel As Boolean)
If Val(Text10.Text) > 0 Then
    Exit Sub
Else
'    Cancel = True
End If
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


Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Text2_Change()
Dim b As Boolean
If Option2 = True Then
If Val(Text2.Text) > 0 Then
    b = edit1
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
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

Private Sub Text2_Validate(Cancel As Boolean)
Dim b As Boolean

If Val(Text2.Text) > 0 Then
    b = edit1
    If b = False Then
        Cancel = True
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim f As Integer, s As Integer

'Text16.Text = Val(Text11.Text) - Val(Text13.Text) - Val(Text18.Text)
'Text15.Text = Val(Text13.Text) + Val(Text18.Text)
End Sub
