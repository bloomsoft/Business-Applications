VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Items 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items Information"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "Items.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   210
      TabIndex        =   33
      Top             =   3930
      Width           =   5415
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   3360
         TabIndex        =   9
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   55640067
         CurrentDate     =   39206
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Opening Stock (BALES)"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Opening Stock (WT)"
         Height          =   255
         Left            =   2880
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Amount"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Item Sub Group"
      Height          =   600
      Left            =   225
      TabIndex        =   30
      Top             =   1635
      Width           =   4200
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   225
         Width           =   3870
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Item Group"
      Height          =   600
      Left            =   225
      TabIndex        =   29
      Top             =   1020
      Width           =   4200
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Text            =   "Combo2"
         Top             =   195
         Width           =   3870
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2235
      Left            =   4440
      TabIndex        =   21
      Top             =   0
      Width           =   1320
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   660
         Left            =   120
         Picture         =   "Items.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   660
         Left            =   120
         Picture         =   "Items.frx":2D03
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   780
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   630
         Left            =   120
         Picture         =   "Items.frx":320A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1635
      Left            =   225
      TabIndex        =   18
      Top             =   2280
      Width           =   5550
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1785
         TabIndex        =   11
         Text            =   "Combo4"
         Top             =   1620
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Finished Item"
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1290
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3270
         TabIndex        =   13
         Top             =   2310
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   2310
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Com1 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   840
         Width           =   3645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Item Head Name"
         Height          =   210
         Left            =   360
         TabIndex        =   32
         Top             =   1635
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Bales"
         Height          =   255
         Left            =   2820
         TabIndex        =   28
         Top             =   2310
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label6 
         Caption         =   "Kgs"
         Height          =   225
         Left            =   360
         TabIndex        =   27
         Top             =   2355
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Opening Stock"
         Height          =   255
         Left            =   495
         TabIndex        =   26
         Top             =   2325
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Unit For Measurements"
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Item's Description"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Item's Code"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   975
      Left            =   225
      TabIndex        =   16
      Top             =   0
      Width           =   4200
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   720
         Left            =   2265
         Picture         =   "Items.frx":36F8
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   165
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   720
         Left            =   705
         Picture         =   "Items.frx":3C1E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   165
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   1740
      TabIndex        =   31
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   210
      TabIndex        =   22
      Top             =   4920
      Width           =   4200
   End
End
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As bloom1
Private Sub saveAccount()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim opendate As Date
Set DB = OpenDatabase(Blm.patHmain)
'ssql = "delete from voumst where v_type = 6 and v_no = 1"
'db.Execute ssql
If Option2 = True Then
    Ssql = "delete from acchart where code = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where v_type = 10 and v_no = 1 and party = " & Val(Text1.Text)
    DB.Execute Ssql
End If


Set TB = DB.OpenRecordset("acchart", dbOpenTable)
TB.AddNew

    TB.Fields("CODE").Value = Val(Text1.Text)
    TB.Fields("NAME").Value = CStr(Text2.Text)
    TB.Fields("opdate").Value = DTPicker1.Value
    TB.Fields("debit").Value = Val(Text6.Text) * Val(Text8.Text)
    TB.Fields("Credit").Value = 0
TB.Update
TB.Close
Dim Tb2 As Recordset
Ssql = "Select * from voumst where v_type=10 and v_no = 1"
Set TB = DB.OpenRecordset(Ssql)
If TB.EOF Then
Set Tb2 = DB.OpenRecordset("voumst", dbOpenTable)
Tb2.AddNew
    Tb2.Fields("v_date").Value = DTPicker1.Value
    Tb2.Fields("v_type").Value = 10
    Tb2.Fields("v_no").Value = 1
    Tb2.Fields("narration").Value = "Open Balance"
Tb2.Update
Tb2.Close
opendate = DTPicker1.Value
Else
    opendate = TB.Fields("v_date").Value
End If

Set Tb2 = DB.OpenRecordset("voudtl", dbOpenTable)
Tb2.AddNew
    Tb2.Fields("v_date").Value = DTPicker1.Value
    Tb2.Fields("v_type").Value = 10
    Tb2.Fields("v_no").Value = 1
    Tb2.Fields("party").Value = Val(Text1.Text)
'    MsgBox "Test"
    Tb2.Fields("debit").Value = Val(Text8.Text)
    Tb2.Fields("credit").Value = 0
Tb2.Update
Tb2.Close

DB.Close

End Sub

Private Sub FillAcComb()
Dim Ssql As String
Ssql = "Select * from Acchart Order by Name"
comb_fill Combo4, Ssql

End Sub
Private Sub combs2()
Dim Ssql As String
If Combo2.ListIndex > -1 Then
Ssql = "select * from Subgroups where GroupCode = " & Combo2.ItemData(Combo2.ListIndex) & " order by name"
comb_fill Combo3, Ssql
End If

End Sub
Private Sub Combs()
Dim Ssql As String

Ssql = "Select * from Groups Order by Name"
comb_fill Combo2, Ssql

End Sub
Private Sub comb_fill(CNTL As Control, Ssql As String)
Dim DB As Database
Dim TB As Recordset
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
CNTL.clear
If Not TB.EOF Then

    Do While Not TB.EOF
        CNTL.AddItem TB.Fields("name").Value
        CNTL.ItemData(CNTL.NewIndex) = TB.Fields("code").Value
        TB.MoveNext
    Loop
CNTL.ListIndex = 0
End If
TB.Close
DB.Close
End Sub
Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)
If Option2 = True Then
    Ssql = "delete from Items where code = " & Val(Text1.Text)
    DB.Execute Ssql
End If


Set TB = DB.OpenRecordset("Items", dbOpenTable)
TB.AddNew
    TB.Fields("CODE").Value = Val(Text1.Text)
    TB.Fields("NAME").Value = CStr(Text2.Text)
    TB.Fields("Unit").Value = CStr(Text3.Text)
    TB.Fields("Stock").Value = Val(Text4.Text)
    TB.Fields("Bales").Value = Val(Text5.Text)
    TB.Fields("IType").Value = Check1.Value
    TB.Fields("OpWT").Value = Val(Text6.Text)
    TB.Fields("OpBales").Value = Val(Text7.Text)
    TB.Fields("Rate").Value = Val(Text8.Text)
    TB.Fields("OpDate").Value = DTPicker1.Value
    TB.Fields("GroupCode").Value = Combo2.ItemData(Combo2.ListIndex)
    TB.Fields("SubGroupCode").Value = Combo3.ItemData(Combo3.ListIndex)
    TB.Fields("AcCode").Value = Val(Text1.Text)
TB.Update
TB.Close
DB.Close
DoEvents
saveAccount
End Sub
Private Function UnitRet(S As Double) As String
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Items WHERE Code = " & S
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)

If Not TB.EOF Then
    UnitRet = TB.Fields("Unit").Value
    Text4.Text = TB.Fields("Stock").Value & ""
    Text5.Text = TB.Fields("Bales").Value & ""
    
Else
    UnitRet = ""
End If
TB.Close
DB.Close

End Function

Private Function check(S As String) As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Items WHERE NAME = '" & S & "'"
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)

If Not TB.EOF Then
    check = True
Else
    check = False
End If
TB.Close
DB.Close
End Function
Private Function max1() As Long
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
If Combo3.ListIndex > -1 Then
Ssql = "select MAX(CODE) AS C FROM Items where Code Between " & Combo3.ItemData(Combo3.ListIndex) * 10000 & " and " & Combo3.ItemData(Combo3.ListIndex) * 10000 + 10000
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("C").Value) Then
    max1 = TB.Fields("C").Value + 1
Else
    max1 = Combo3.ItemData(Combo3.ListIndex) * 10000 + 1
End If
TB.Close
DB.Close
End If
End Function

Private Sub Com1_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
If Result = vbNo Then Exit Sub

Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)
If Option2 = True Then
    Ssql = "Select * from VouDtl where V_type<>10 And Party = " & Val(Text1.Text)
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then
        MsgBox "You Already Have Transactions in this Item"
        Exit Sub
    Else

        Ssql = "delete from Items where code = " & Val(Text1.Text)
        DB.Execute Ssql
        Ssql = "delete from Acchart where code = " & Val(Text1.Text)
        DB.Execute Ssql
        Ssql = "delete from voudtl where v_type = 10 and v_no = 1 and party = " & Val(Text1.Text)
        DB.Execute Ssql
    End If
End If
DB.Close
Command2_Click
End Sub

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
Call edit1
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Combo2_Click()
If Combo2.ListIndex > -1 Then
    combs2
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
'If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Combo3_Click()
If Option1 = True Then
If Combo3.ListIndex > -1 Then
    Text1.Text = max1
End If
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command1_Click()
Dim p As Boolean

'p = Option2.Value
Call save
'MSAVE Val(Text1.Text), UCase(Text2.Text), p
'Combs
Check1.Value = 0
Command2_Click
If Option1 = True Then
Text2.SetFocus
Else
Text1.Enabled = True
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
FillAcComb
Check1.Value = 0
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text7.Text = vbNullString
Text6.Text = vbNullString
Text8.Text = vbNullString
If Option2 = True Then

    Text1.Text = vbNullString
    Text1.Enabled = True
    Text1.SetFocus
Else
    Text1.Enabled = False
    Text1.Text = max1

End If
Command1.Enabled = False
If Option1 = True Then
Text2.SetFocus
Else
Text1.Enabled = True
Text1.SetFocus
End If
DTPicker1.Value = FStartDate - 1
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
Set Blm = New bloom1
DTPicker1.Value = FStartDate - 1

Text1.Text = max1
Combs
FillAcComb
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Blm = Nothing
End Sub

Private Sub Option1_Click()
Command2_Click
Com1.Visible = False
If Option1 = True Then
    Text1.Enabled = False
    Text1.Text = max1
    Combo2.SetFocus
    'Text2.SetFocus
    
Else

End If
End Sub

Private Sub Option2_Click()

Command2_Click
Com1.Visible = True
If Option2 = True Then
    Combs
   
    Text1.Text = vbNullString
    Text1.Enabled = True
    Combo2.SetFocus
    'Text1.SetFocus
Else
    Text1.Enabled = False
End If

End Sub

Private Sub edit1()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer

Ssql = "SELECT * FROM Items WHERE code = " & Val(Text1.Text)
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
'MsgBox Ssql
If Not TB.EOF Then
    'Text2.Text = tb.Fields("name").Value
    Text3.Text = UnitRet(Val(Text1.Text))
    Check1.Value = TB.Fields("Itype").Value
    
    For R = 0 To Combo2.ListCount - 1
        If Combo2.ItemData(R) = TB.Fields("GroupCode").Value Then
            Combo2.ListIndex = R
            Exit For
        End If
    Next R
    
    For R = 0 To Combo3.ListCount - 1
        If Combo3.ItemData(R) = TB.Fields("SubGroupCode").Value Then
            Combo3.ListIndex = R
            Exit For
        End If
    Next R
    If Not IsNull(TB.Fields("Accode").Value) Then
    For R = 0 To Combo4.ListCount - 1
        If Combo4.ItemData(R) = TB.Fields("AcCode").Value Then
            Combo4.ListIndex = R
            Exit For
        End If
    Next R
    End If
    
    Text6.Text = TB.Fields("OpWT").Value & ""
    Text7.Text = TB.Fields("OpBales").Value & ""
    Text8.Text = TB.Fields("Rate").Value & ""
    'DTPicker1.Value = TB.Fields("OpDate").Value
    
'    MsgBox "Test"
Else
    MsgBox "Invalid Item's Code"
    Text1.Enabled = True
    
End If
TB.Close
DB.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search1
    Search1.Show vbModal
    Text1.Text = SelectedItemCode
    Text2.Text = SelectedItemName
'    MsgBox Text1.Text
    edit1
End If
End Sub

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
If Val(Text1.Text) > 0 Then
    Text2.Text = Blm.item1(Val(Text1.Text))
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
Dim B As Boolean

B = check(UCase(CStr(Text2.Text)))
If B = True Then
    MsgBox "ITEM ALREADY EXIST,,,,"
    'Text2.SetFocus
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
