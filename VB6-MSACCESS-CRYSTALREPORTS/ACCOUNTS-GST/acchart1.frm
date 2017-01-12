VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form acchart1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Information"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "acchart1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Heads Information"
      Height          =   735
      Left            =   240
      TabIndex        =   27
      Top             =   960
      Width           =   5775
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Heads Info"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sub Head Information"
      Height          =   855
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   5775
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Select A Sub Head"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   19
      Top             =   5760
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   4200
         Picture         =   "acchart1.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   735
         Left            =   2400
         Picture         =   "acchart1.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   735
         Left            =   600
         Picture         =   "acchart1.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   3135
      Left            =   240
      TabIndex        =   16
      Top             =   2640
      Width           =   5775
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   10
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   615
         Left            =   3000
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   50331651
         CurrentDate     =   36747
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1815
         TabIndex        =   4
         Top             =   690
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Phone"
         Height          =   255
         Left            =   3600
         TabIndex        =   32
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Address"
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "S.T. Reg. #"
         Height          =   255
         Left            =   840
         TabIndex        =   30
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Credit"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Debit"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "A/c Name"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   975
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   5775
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   555
         Left            =   3720
         Picture         =   "acchart1.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   555
         Left            =   840
         Picture         =   "acchart1.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6960
      Width           =   1335
   End
End
Attribute VB_Name = "acchart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Private Function opdate() As Date
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String

Set DB = OpenDatabase(blm.patHmain)
Ssql = "select v_date from voumst where v_type=6"
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    opdate = tb.Fields("v_date").Value
Else
    opdate = Date
End If
tb.Close
DB.Close
End Function

Private Sub save()
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Dim opendate As Date
Set DB = OpenDatabase(blm.patHmain)
'ssql = "delete from voumst where v_type = 6 and v_no = 1"
'db.Execute ssql
If Option2 = True Then
    Ssql = "delete from acchart where code = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where v_type = 10 and v_no = 1 and party = " & Val(Text1.Text)
    DB.Execute Ssql
End If


Set tb = DB.OpenRecordset("acchart", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("opdate").Value = date1.Value
    tb.Fields("debit").Value = Val(Text3.Text)
    tb.Fields("Credit").Value = Val(Text4.Text)
    tb.Fields("STNo").Value = Text5.Text
    tb.Fields("Phone").Value = Text7.Text
    tb.Fields("Address").Value = Text6.Text
tb.Update
tb.Close
Dim tb2 As Recordset
Ssql = "Select * from voumst where v_type=10 and v_no = 1"
Set tb = DB.OpenRecordset(Ssql)
If tb.EOF Then
Set tb2 = DB.OpenRecordset("voumst", dbOpenTable)
tb2.AddNew
    tb2.Fields("v_date").Value = date1.Value
    tb2.Fields("v_type").Value = 10
    tb2.Fields("v_no").Value = 1
    tb2.Fields("narration").Value = "Open Balance"
tb2.Update
tb2.Close
opendate = date1.Value
Else
    opendate = tb.Fields("v_date").Value
End If

Set tb2 = DB.OpenRecordset("voudtl", dbOpenTable)
tb2.AddNew
    tb2.Fields("v_date").Value = opendate
    tb2.Fields("v_type").Value = 10
    tb2.Fields("v_no").Value = 1
    tb2.Fields("party").Value = Val(Text1.Text)
    tb2.Fields("debit").Value = Val(Text3.Text)
    tb2.Fields("credit").Value = Val(Text4.Text)
tb2.Update
tb2.Close

DB.Close

End Sub
Private Function check(s As String) As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM acchart WHERE NAME = '" & s & "'"
Set DB = OpenDatabase(blm.patHmain)
Set tb = DB.OpenRecordset(Ssql)

If Not tb.EOF Then
    check = True
Else
    check = False
End If
tb.Close
DB.Close
End Function
Private Function max1() As Long
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
If Combo2.ListIndex > -1 Then
Ssql = "select MAX(CODE) AS C FROM acchart WHERE CODE between " & Combo2.ItemData(Combo2.ListIndex) * 10000 & " and " & Combo2.ItemData(Combo2.ListIndex) * 10000 + 10000
Set DB = OpenDatabase(blm.patHmain)
Set tb = DB.OpenRecordset(Ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = Combo2.ItemData(Combo2.ListIndex) * 10000 + 1
End If
tb.Close
DB.Close
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
Call edit1

End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
Text1.Text = max1
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'SendKeys ("{TAB}")
If Option1 = True Then
    Text2.SetFocus
End If
End If
End Sub

Private Sub Combo3_Click()
If Combo3.ListCount > 0 Then
Dim Ssql As String
Ssql = "select * from heads where code > 100 and Mid(Cstr(Code),1,2) = '" & Combo3.ItemData(Combo3.ListIndex) & "' order by name"
Combo2.clear
comb_fill Combo2, Ssql
If Combo2.ListCount < 1 Then
    Command1.Enabled = False
    
End If
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command1_Click()
If Option2 = True Then
    If Val(Text1.Text) <= 0 Then
        MsgBox "Select New Option then Enter the New Account...."
        Exit Sub
    End If
End If
Call save

combs
Command2_Click

If Option1 = True Then
Text2.SetFocus
Else
Combo1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text5.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
If Option2 = True Then
    
    Text1.Text = vbNullString
    Combo1.Visible = True
    Text2.Visible = False
Else
    Text1.Enabled = False
    Text1.Text = max1
    Combo1.Visible = False
    Text2.Visible = True
End If
Text3.Text = vbNullString
Text4.Text = vbNullString

Command1.Enabled = False
'date1.Value = opdate
If Option1 = True Then
Text2.SetFocus
Else
Combo1.SetFocus
End If

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub combs()
Dim Ssql As String


Ssql = "select * from acchart where code > 99999 order by name"
comb_fill Combo1, Ssql
End Sub

Private Sub Command4_Click()
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.patHmain)
If Option2 = True Then
    Ssql = "Select * from VouDtl where Party = " & Val(Text1.Text)
    Set tb = DB.OpenRecordset(Ssql)
    If Not tb.EOF Then
        MsgBox "You Already Have Transactions in this Detail A/c"
        Exit Sub
    Else
        Ssql = "delete from Acchart where code = " & Val(Text1.Text)
        DB.Execute Ssql
    End If
    tb.Close
End If
DB.Close
Command2_Click

End Sub

Private Sub Form_Activate()
Dim Ssql As String
Ssql = "select * from heads where code < 100 order by name"
comb_fill Combo3, Ssql
End Sub

Private Sub Form_Load()
Set blm = New bloom1
Me.Top = 10
Me.Left = 10
date1.Value = Date
combs
Text1.Text = max1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Combo2.Enabled = True
    Text1.Enabled = False
    Text1.Text = max1
    Text2.Visible = True
    Text2.Text = vbNullString
    Command4.Visible = False
    Text2.SetFocus
    Combo1.Visible = False
    
Else

End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
combs
    Combo2.Enabled = False
    'Text1.Enabled = True
    Text1.Text = vbNullString
    Command4.Visible = True
    
    Combo1.Visible = True
    Text2.Visible = False
    
    
Combo1.SetFocus
    
Else
    Text1.Enabled = False
End If

End Sub

Private Sub edit1()
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM acchart WHERE code = " & Val(Text1.Text)
Set DB = OpenDatabase(blm.patHmain)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    Text2.Text = tb.Fields("name").Value
    Text1.Enabled = False
    Text3.Text = tb.Fields("debit").Value
    Text4.Text = tb.Fields("credit").Value
    date1.Value = tb.Fields("opdate").Value
    Text5.Text = tb.Fields("STNo").Value & ""
    Text7.Text = tb.Fields("Phone").Value & ""
    Text6.Text = tb.Fields("Address").Value & ""
Else
    MsgBox "Invalid Accounts's Code"
    Combo1.SetFocus
End If
tb.Close
DB.Close
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

Private Sub Text1_LostFocus()
If Val(Text1.Text) > 0 Then
Call edit1
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
If KeyAscii = 13 Then date1.SetFocus

End Sub

Private Sub Text2_LostFocus()
If Option1 = True Then
Dim b As Boolean

b = check(UCase(CStr(Text2.Text)))
If b = True Then
    MsgBox "A/c ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text6_Change()
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text7_Change()
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
