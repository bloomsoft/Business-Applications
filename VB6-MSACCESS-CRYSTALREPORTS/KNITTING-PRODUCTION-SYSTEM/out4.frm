VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form out4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "InWard Gate Pass of Purchase Dyeing Contract"
   ClientHeight    =   7545
   ClientLeft      =   2745
   ClientTop       =   1065
   ClientWidth     =   7095
   Icon            =   "out4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "A/C List"
      Height          =   690
      Left            =   270
      TabIndex        =   53
      Top             =   5580
      Width           =   6585
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1305
         TabIndex        =   54
         Text            =   "Combo2"
         Top             =   225
         Width           =   5160
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sent Items Info"
      Height          =   1215
      Left            =   270
      TabIndex        =   44
      Top             =   6300
      Width           =   6615
      Begin VB.Label Label32 
         Caption         =   "Label32"
         Height          =   255
         Left            =   5160
         TabIndex        =   52
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Cloth Rolls Recvd."
         Height          =   375
         Left            =   3720
         TabIndex        =   51
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Label30"
         Height          =   255
         Left            =   1560
         TabIndex        =   50
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Cloth Recvd. Qty."
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   5160
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Cloth Rolls Sent"
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   255
         Left            =   1560
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cloth Sent Qty."
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1095
      Left            =   3360
      TabIndex        =   35
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2280
         Picture         =   "out4.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1200
         Picture         =   "out4.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   735
         Left            =   120
         Picture         =   "out4.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   32
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1800
         Picture         =   "out4.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   240
         Picture         =   "out4.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   6615
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   1320
         TabIndex        =   58
         Top             =   2040
         Visible         =   0   'False
         Width           =   3105
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   825
         Left            =   2745
         TabIndex        =   57
         Top             =   2070
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1455
         _Version        =   393216
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1305
         TabIndex        =   8
         Top             =   2070
         Width           =   1365
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1305
         TabIndex        =   6
         Top             =   1710
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date4 
         Height          =   285
         Left            =   5160
         TabIndex        =   27
         Top             =   3735
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   36921
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cancel this Inward GatePass"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   3780
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5640
         Top             =   4680
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   4050
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   11
         Top             =   3330
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   3330
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5625
         TabIndex        =   9
         Top             =   2025
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3375
         TabIndex        =   7
         Top             =   1710
         Width           =   3105
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -270
         TabIndex        =   23
         Top             =   2115
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1320
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
         Format          =   20643843
         CurrentDate     =   36749
      End
      Begin VB.Label Label26 
         Caption         =   "Color"
         Height          =   240
         Left            =   315
         TabIndex        =   56
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label Label25 
         Caption         =   "Qulity Code"
         Height          =   255
         Left            =   315
         TabIndex        =   55
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label15 
         Height          =   255
         Left            =   5400
         TabIndex        =   43
         Top             =   2940
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label14 
         Caption         =   "Yarn Quantity"
         Height          =   255
         Left            =   3840
         TabIndex        =   42
         Top             =   2940
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   2925
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   315
         TabIndex        =   40
         Top             =   2925
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Lycra Quantity"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   3765
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Del. Date"
         Height          =   255
         Left            =   4815
         TabIndex        =   31
         Top             =   2025
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
         TabIndex        =   30
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Rate"
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   4920
         TabIndex        =   28
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   315
         TabIndex        =   25
         Top             =   4050
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Contract Date"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   3330
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Cloth Rolls"
         Height          =   255
         Left            =   315
         TabIndex        =   21
         Top             =   3330
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Width"
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   3750
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         Height          =   255
         Left            =   2745
         TabIndex        =   19
         Top             =   1710
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   2745
         TabIndex        =   17
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Inward #"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "out4"
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
End Type
    
    
Private Function Getinfo(c As Long, e As Byte) As ContInfo
    Dim tb As ADODB.Recordset
    Dim ssql As String
    Dim CN2 As ContInfo
    
    
    ssql = "select sum(quantity)as q,sum(bags)as b,sum(lycra)as l,sum(l_bags)as lb from inward where cont_no = " & c
    ssql = ssql & " and e_type= " & e
    
    
    Set tb = CN.Execute(ssql)
    If Not IsNull(tb.Fields("Q").Value) Then
        CN2.clothsent = tb.Fields("Q").Value
    End If
    If Not IsNull(tb.Fields("B").Value) Then
        CN2.sclothrolls = tb.Fields("B").Value
    End If
    If Not IsNull(tb.Fields("l").Value) Then
        CN2.lycrasent = tb.Fields("l").Value
    End If
    If Not IsNull(tb.Fields("lb").Value) Then
        CN2.lbagssent = tb.Fields("lb").Value
    End If
    
    
    tb.Close
    
    ssql = "select sum(quantity)as q,sum(rolls)as r from Outward where cont_no = " & c
    ssql = ssql & " and e_type= " & e
    
    
    Set tb = CN.Execute(ssql)
    If Not IsNull(tb.Fields("Q").Value) Then
        CN2.clothrec = tb.Fields("Q").Value
    End If
    If Not IsNull(tb.Fields("r").Value) Then
        CN2.rclothrolls = tb.Fields("r").Value
    End If
    tb.Close
    Getinfo = CN2

End Function


Private Function Check(c As Long) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String
    
ssql = "select * from outward where out_no = " & c
ssql = ssql & " and E_type=3"
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    MsgBox "Inward No already Exist..."
    Check = True
Else
    Check = False
End If
tb.Close

End Function


Private Function edit1() As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String
Dim u As ContInfo

ssql = "select * from cont_1 where cont_no = " & Val(Text2.Text)
ssql = ssql & " and e_type = 3"
org_q = 0
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    date3.Value = tb.Fields("v_date").Value
    Text3.Text = tb.Fields("party").Value
    Text4.Text = blm.party1(tb.Fields("party").Value)
    Label21.Caption = Format(tb.Fields("del_date").Value, "dd/MM/yyyy")
    Label23.Caption = Format(tb.Fields("Rate").Value, "#.00")
    org_q = tb.Fields("Cquantity").Value
    Label13.Caption = Format(tb.Fields("CQuantity").Value, "#.00")
    Label15.Caption = Format(tb.Fields("YQuantity").Value, "#.00")
    
    Text9.Text = tb.Fields("item").Value
    Text8.Text = blm.Cloth(tb.Fields("item").Value)
    If Not IsNull(tb.Fields("complete").Value) Then
    If tb.Fields("Complete").Value = 1 Then
        MsgBox "You Have Marked this Contract as Completed....."
        
    End If
    End If
    edit1 = True
    u = Getinfo(Val(Text2.Text), 3)
    Label16.Caption = Format(u.clothsent, "#.00")
    Label20.Caption = Format(u.sclothrolls, "#.00")
    Label30.Caption = Format(u.clothrec, "#.00")
    Label32.Caption = Format(u.rclothrolls, "#.00")
    
Else
    MsgBox "Not Found ...!"
    edit1 = False
End If
tb.Close
    
End Function
Private Function max1() As Double
    Dim ssql As String
    Dim tb As ADODB.Recordset
    
    ssql = "select max(out_no)as c from outward where e_type=3"
    Set tb = CN.Execute(ssql)
    If IsNull(tb.Fields("c").Value) = False Then
        max1 = tb.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    tb.Close

End Function
Private Function edit_kachi() As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from outward where out_no = " & Val(Text1.Text)
ssql = ssql & " and e_type=3"

Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("v_date").Value
    Text3.Text = tb.Fields("accode").Value & ""
    Text4.Text = tb.Fields("acname").Value & ""
    Text8.Text = tb.Fields("quality").Value & ""
    Text6.Text = tb.Fields("item").Value
    If tb.Fields("cont_no").Value > 0 Then
    Text2.Text = tb.Fields("cont_no").Value
    End If
    Text10.Text = tb.Fields("rolls").Value
    Text11.Text = tb.Fields("Quantity").Value
    Text17.Text = tb.Fields("remarks").Value & ""
    If Not IsNull(tb.Fields("cancel").Value) Then
        Check1.Value = tb.Fields("cancel").Value
    End If
    If Not IsNull(tb.Fields("c_date").Value) Then
        date4.Value = tb.Fields("c_date").Value
    End If
   
   Do While Not tb.EOF
   With Grid1
   .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = .Rows - 1
    .TextMatrix(.Rows - 1, 1) = tb.Fields("color").Value & ""
    End With
    tb.MoveNext
   Loop
    edit_kachi = True
Else
    MsgBox "Not Found ...!"
    edit_kachi = False
End If
tb.Close
    
End Function

Private Sub save()
Dim tb As New ADODB.Recordset
Dim ssql As String
If Option2 = True Then
    ssql = "delete from Outward where out_no = " & Val(Text1.Text)
    ssql = ssql & " and e_type = 3"
    CN.Execute ssql
End If
tb.Open "Outward", CN, 0, 3, 0
For i = 1 To Grid1.Rows - 1
tb.AddNew
    tb.Fields("out_no").Value = Val(Text1.Text)
    tb.Fields("accode").Value = Text3.Text
    tb.Fields("acname").Value = Text4.Text
    tb.Fields("quality").Value = Text8.Text
    tb.Fields("item").Value = Text6.Text
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("E_Type").Value = 3
    tb.Fields("cont_no").Value = Val(Text2.Text)
    tb.Fields("rolls").Value = Val(Text10.Text)
    tb.Fields("quantity").Value = Val(Text11.Text)
    tb.Fields("remarks").Value = CStr(Text17.Text)
    If Option2 = True Then
        tb.Fields("cancel").Value = Check1.Value
        tb.Fields("c_date").Value = date4.Value
    End If
    tb.Fields("color").Value = Grid1.TextMatrix(i, 1)
tb.Update
Next i
tb.Close
    
End Sub

Private Sub clear()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
Label23.Caption = vbNullString
Label21.Caption = vbNullString
Grid1.Rows = 1
End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
    Text3.Text = Combo2.ItemData(Combo2.ListIndex)
    Text4.Text = Combo2.Text
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If

End Sub

Private Sub Command1_Click()
If Len(Text3.Text) <= 0 Then
MsgBox "Please Select Party", , BLOOMSOFT
Exit Sub
End If

Call save
Command2_Click

End Sub

Private Sub Command2_Click()
Call clear
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide

End Sub

Private Sub Form_Load()
Dim ssql As String
Me.Top = ((Screen.Height - Me.Height) / 2) - 1000
Me.Left = (Screen.Width - Me.Width) / 2
date1.Value = Date

date3.Value = Date

ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo2, "name", "code"
Grid1.Rows = 1
Grid1.Cols = 2
    Grid1.ColWidth(0) = 300
    Grid1.TextMatrix(0, 0) = "Sr.#"
    Grid1.ColWidth(1) = 1700
    Grid1.TextMatrix(0, 1) = "Color"
    Text1.Text = max1

End Sub

Private Sub Grid1_DblClick()
If Grid1.Rows > 1 Then
    With Grid1
        Text5.Text = .TextMatrix(.Row, 1)
    
    If .Rows > 2 Then
        .RemoveItem .Row
    Else
        .Rows = 1
    End If
    End With

End If

End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
Text6.Text = List1.ItemData(List1.ListIndex)
Text8.Text = List1.Text
End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
List1.Visible = False
End If

End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub Option1_Click()
'Text1.Enabled = False
Check1.Visible = False
date4.Visible = False
Text2.SetFocus
End Sub

Private Sub Option2_Click()
'Text1.Enabled = True
Check1.Visible = True
date4.Visible = True
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
    Cancel = True
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

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus

End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Val(Text3.Text) > 0 Then
    Text4.Text = blm.party1(Val(Text3.Text))
        If Text4.Text = "NOT FOUND" Then
            Cancel = True
            
        End If
Else
        Cancel = True
End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 96 And KeyAscii <= 122 Then
KeyAscii = KeyAscii - 32
End If

If Len(Text5.Text) > 0 Then
If KeyAscii = 13 Then
    Grid1.Rows = Grid1.Rows + 1
    Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.Rows - 1
    Grid1.TextMatrix(Grid1.Rows - 1, 1) = Text5.Text

Text5.Text = ""
Text5.SetFocus
End If
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim tb As New ADODB.Recordset
Dim i As Integer
tb.Open "cloth", CN, 0, 3, 0
List1.Visible = True
'Dim S As String
If Not tb.EOF Then
List1.clear
Do While Not tb.EOF
List1.AddItem tb.Fields("name").Value
List1.ItemData(List1.NewIndex) = tb.Fields("code").Value
tb.MoveNext
Loop
List1.ListIndex = 0
End If
tb.Close
Set tb = Nothing
List1.SetFocus
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If

End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Val(Text6.Text) > 0 Then
    Text8.Text = blm.Cloth(Val(Text6.Text))
End If

End Sub

Private Sub Timer1_Timer()
Dim f As Integer, S As Integer

'Text16.Text = Val(Text11.Text) - Val(Text13.Text) - Val(Text18.Text)
'Text15.Text = Val(Text13.Text) + Val(Text18.Text)
End Sub
