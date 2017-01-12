VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form acchart1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Information"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "acchart1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Sub Head Information"
      Height          =   855
      Left            =   240
      TabIndex        =   20
      Top             =   960
      Width           =   6615
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Text            =   "Combo2"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "Select A Sub Head"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   5520
      Width           =   6615
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   4200
         Picture         =   "acchart1.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   735
         Left            =   2880
         Picture         =   "acchart1.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1560
         Picture         =   "acchart1.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   3615
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   6615
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5040
         Top             =   2520
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   2760
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19529731
         CurrentDate     =   36747
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Fax"
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Phone"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Address"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Credit"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Debit"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "A/c Name"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   975
      Left            =   240
      TabIndex        =   13
      Top             =   0
      Width           =   6615
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   555
         Left            =   4560
         Picture         =   "acchart1.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   25
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   6600
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
Dim db As Database
Dim tb As Recordset
Dim ssql As String

'Set db = OpenDatabase(blm.pathMain)
'ssql = "select v_date from voumst where v_type=6"
'Set tb = db.OpenRecordset(ssql)
'If Not tb.EOF Then
'    opdate = tb.Fields("v_date").Value
'Else
'    opdate = Date
'End If
'tb.Close
'db.Close
End Function

Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim opendate As Date
Set db = OpenDatabase(blm.pathMain)
'ssql = "delete from voumst where v_type = 6 and v_no = 1"
'db.Execute ssql
If Option2 = True Then
    ssql = "delete from acchart where code = " & Val(Text1.Text)
    db.Execute ssql
'    ssql = "delete from voudtl where v_type = 6 and v_no = 1 and party = " & Val(Text1.Text)
'    db.Execute ssql
End If


Set tb = db.OpenRecordset("acchart", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("opdate").Value = Date1.Value
    tb.Fields("debit").Value = Val(Text3.Text)
    tb.Fields("Credit").Value = Val(Text4.Text)
    tb.Fields("Add").Value = Text5.Text
    tb.Fields("Phone").Value = Text6.Text
    tb.Fields("Fax").Value = Text7.Text
tb.Update
tb.Close
'Dim tb2 As Recordset
'ssql = "Select * from voumst where v_type = 6 and v_no = 1"
'Set tb = db.OpenRecordset(ssql)
'If tb.EOF Then
'Set tb2 = db.OpenRecordset("voumst", dbOpenTable)
'tb2.AddNew
'    tb2.Fields("v_date").Value = date1.Value
'    tb2.Fields("v_type").Value = 6
'    tb2.Fields("v_no").Value = 1
'    tb2.Fields("narration").Value = "Open Balance"
'tb2.Update
'tb2.Close
'opendate = date1.Value
'Else
 '   opendate = tb.Fields("v_date").Value
'End If

'Set tb2 = db.OpenRecordset("voudtl", dbOpenTable)
'tb2.AddNew
'    tb2.Fields("v_date").Value = opendate
'    tb2.Fields("v_type").Value = 6
'    tb2.Fields("v_no").Value = 1
'    tb2.Fields("party").Value = Val(Text1.Text)
'    tb2.Fields("debit").Value = Val(Text3.Text)
'    tb2.Fields("credit").Value = Val(Text4.Text)
'tb2.Update
'tb2.Close

db.Close

End Sub
Private Function check(s As String) As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM acchart WHERE NAME = '" & s & "'"
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
ssql = "select MAX(CODE) AS C FROM acchart WHERE CODE between " & Combo2.ItemData(Combo2.ListIndex) * 10000 & " and " & Combo2.ItemData(Combo2.ListIndex) * 10000 + 10000
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = Combo2.ItemData(Combo2.ListIndex) * 10000 + 1
End If
tb.Close
db.Close
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
Call Edit1

End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Combo2_Click()
Text1.Text = max1
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'SendKeys ("{TAB}")
If Option1 = True Then
    Text2.SetFocus
End If
End If
End Sub

Private Sub Command1_Click()
If Option2 = True Then
    If Val(Text1.Text) <= 0 Then
        MsgBox "Select New Option then Enter the New Account...."
        Exit Sub
    End If
End If
Call save

Combs
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
Date1.Value = opdate
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

Private Sub Combs()
Dim ssql As String


ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo1, "Name", "Code"
End Sub
Private Sub Form_Load()
Set blm = New bloom1
Dim ssql As String
ssql = "select * from heads where code > 100 order by name"
blm.fill_comb ssql, Combo2, "Name", "Code"

Combs
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
    Text2.SetFocus
    Combo1.Visible = False
    
Else

End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
Combs
    Combo2.Enabled = False
    'Text1.Enabled = True
    Text1.Text = vbNullString
    
    
    Combo1.Visible = True
    Text2.Visible = False
    
    
Combo1.SetFocus
    
Else
    Text1.Enabled = False
End If

End Sub

Private Sub Edit1()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM acchart WHERE code = " & Val(Text1.Text)
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    Text2.Text = tb.Fields("name").Value
    Text1.Enabled = False
    Text3.Text = tb.Fields("debit").Value
    Text4.Text = tb.Fields("credit").Value
    Date1.Value = tb.Fields("opdate").Value
    Text5.Text = tb.Fields("Add").Value
    Text6.Text = tb.Fields("Phone").Value
    Text7.Text = tb.Fields("Fax").Value
    
Else
    MsgBox "Invalid Account Code"
    Combo1.SetFocus
End If
tb.Close
db.Close
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
Call Edit1
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
If KeyAscii = 13 Then SendKeys ("{TAB}")

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

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Timer1_Timer()
If Val(Text1.Text) = 0 Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If
End Sub
