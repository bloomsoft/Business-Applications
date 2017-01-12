VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form kachi3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OutWard Gate Pass of Purchase Dyeing Contract"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7035
   Icon            =   "KACHI3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "A/C List"
      Height          =   780
      Left            =   225
      TabIndex        =   47
      Top             =   5310
      Width           =   6630
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1305
         TabIndex        =   48
         Text            =   "Combo2"
         Top             =   270
         Width           =   5205
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1095
      Left            =   3360
      TabIndex        =   36
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2280
         Picture         =   "KACHI3.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1200
         Picture         =   "KACHI3.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   735
         Left            =   120
         Picture         =   "KACHI3.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   33
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1800
         Picture         =   "KACHI3.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   240
         Picture         =   "KACHI3.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   6615
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   5640
         MaxLength       =   30
         TabIndex        =   13
         Top             =   3720
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   1320
         TabIndex        =   50
         Top             =   2040
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1305
         TabIndex        =   6
         Top             =   1665
         Width           =   1320
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5640
         TabIndex        =   42
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   3360
         Visible         =   0   'False
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
         Height          =   375
         Left            =   5160
         TabIndex        =   28
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   49545219
         CurrentDate     =   36921
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cancel this OutWard"
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   3360
         Visible         =   0   'False
         Width           =   2175
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
         Top             =   3720
         Width           =   3615
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
         Format          =   49545219
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   10
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5640
         TabIndex        =   8
         Top             =   2025
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3285
         TabIndex        =   7
         Top             =   1665
         Width           =   3195
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -450
         TabIndex        =   24
         Top             =   2070
         Visible         =   0   'False
         Width           =   615
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
         Format          =   49545219
         CurrentDate     =   36749
      End
      Begin VB.Label Label17 
         Caption         =   "Van No"
         Height          =   255
         Left            =   5040
         TabIndex        =   51
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Name"
         Height          =   285
         Left            =   2745
         TabIndex        =   49
         Top             =   1710
         Width           =   510
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   255
         Left            =   5640
         TabIndex        =   46
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Yarn Quantity"
         Height          =   255
         Left            =   3840
         TabIndex        =   45
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   1560
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Lycra Quantity"
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Lycra Bags"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Del. Date"
         Height          =   255
         Left            =   4050
         TabIndex        =   32
         Top             =   2070
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
         TabIndex        =   31
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Rate"
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   4920
         TabIndex        =   29
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Contract Date"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Cloth Rolls"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Width"
         Height          =   255
         Left            =   5040
         TabIndex        =   21
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Quality Code"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1665
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   2745
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Outward #"
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "kachi3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim org_q As Currency
Dim rej As Currency


Private Function Check(c As Long) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String
    
ssql = "select * from inward where in_no = " & c
ssql = ssql & " and E_type=3"
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    MsgBox "Outward No already Exist..."
    Check = True
Else
    Check = False
End If
tb.Close
End Function


Private Function edit1() As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

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
Else
    MsgBox "Not Found ...!"
    edit1 = False
End If
tb.Close
End Function
Private Function max1() As Double
    Dim ssql As String
    Dim tb As ADODB.Recordset
    
    ssql = "select max(in_no)as c from inward where e_type=3"
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

ssql = "select * from inward where in_no = " & Val(Text1.Text)
ssql = ssql & " and e_type=3"

Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("in_date").Value
    Text3.Text = tb.Fields("accode").Value & ""
    Text4.Text = tb.Fields("acname").Value & ""
    Text8.Text = tb.Fields("quality").Value & ""
    Text12.Text = tb.Fields("item").Value
    Text13.Text = tb.Fields("vanno").Value & ""
    If tb.Fields("cont_no").Value > 0 Then
    Text2.Text = tb.Fields("cont_no").Value
    End If
    Text10.Text = tb.Fields("Bags").Value
    Text11.Text = tb.Fields("Quantity").Value
    Text17.Text = tb.Fields("remarks").Value & ""
    Text5.Text = tb.Fields("Lycra").Value
    Text6.Text = tb.Fields("L_Bags").Value
    If Not IsNull(tb.Fields("cancel").Value) Then
        Check1.Value = tb.Fields("cancel").Value
    End If
    If Not IsNull(tb.Fields("c_date").Value) Then
        date4.Value = tb.Fields("c_date").Value
    End If

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
    ssql = "delete from Inward where in_no = " & Val(Text1.Text)
    ssql = ssql & " and e_type = 3"
    CN.Execute ssql
End If
tb.Open "inward", CN, 0, 3, 0
tb.AddNew
    tb.Fields("in_no").Value = Val(Text1.Text)
    tb.Fields("in_date").Value = date1.Value
    tb.Fields("accode").Value = Text3.Text
    tb.Fields("acname").Value = Text4.Text
    tb.Fields("quality").Value = Text8.Text
    tb.Fields("item").Value = Text12.Text
    tb.Fields("vanno").Value = Text13.Text
    
    tb.Fields("E_Type").Value = 3
    tb.Fields("cont_no").Value = Val(Text2.Text)
    tb.Fields("bags").Value = Val(Text10.Text)
    tb.Fields("quantity").Value = Val(Text11.Text)
    tb.Fields("remarks").Value = CStr(Text17.Text)
    tb.Fields("Lycra").Value = Val(Text5.Text)
    tb.Fields("L_Bags").Value = Val(Text6.Text)
    If Option2 = True Then
        tb.Fields("cancel").Value = Check1.Value
        tb.Fields("c_date").Value = date4.Value
    End If
tb.Update
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If

End Sub

Private Sub Form_Load()
Dim ssql As String
Me.Top = ((Screen.Height - Me.Height) / 2) - 1000
Me.Left = (Screen.Width - Me.Width) / 2
date1.Value = Date

date3.Value = Date

ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo2, "name", "code"
Text1.Text = max1
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
Text12.Text = List1.ItemData(List1.ListIndex)
Text8.Text = List1.Text
End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text12.SetFocus
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
        Beep
        KeyAscii = 0
    
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
        Beep
        KeyAscii = 0
    
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
        Beep
        KeyAscii = 0
    
End If
End Sub


Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Text12_Validate(Cancel As Boolean)
If Val(Text12.Text) > 0 Then
    Text8.Text = blm.Cloth(Val(Text12.Text))
End If

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
        Beep
        KeyAscii = 0
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

Private Sub Timer1_Timer()
Dim f As Integer, S As Integer

End Sub
