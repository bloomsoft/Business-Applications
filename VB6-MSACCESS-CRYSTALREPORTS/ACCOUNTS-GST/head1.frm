VERSION 5.00
Begin VB.Form head1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Heads Information"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "head1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   615
         Left            =   2760
         Picture         =   "head1.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   615
         Left            =   1560
         Picture         =   "head1.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   615
         Left            =   360
         Picture         =   "head1.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1935
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   4215
      Begin VB.CommandButton Com1 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Head's Type"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Head's Description"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Head's Code"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   555
         Left            =   2400
         Picture         =   "head1.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   555
         Left            =   840
         Picture         =   "head1.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "head1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Private Sub comb2()
Combo2.clear

    Combo2.AddItem "Assets & Liabilities"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    Combo2.AddItem "Income & Expances"
    Combo2.ItemData(Combo2.NewIndex) = 2
    
Combo2.ListIndex = 0
    
End Sub
Private Sub combs()
Dim ssql As String

ssql = "select * from heads where code < 100 order by name"
comb_fill Combo1, ssql
comb2
End Sub
Private Sub comb_fill(cntl As Control, ssql As String)
Dim db As Database
Dim tb As Recordset
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
cntl.clear
    Do While Not tb.EOF
        cntl.AddItem tb.Fields("name").Value
        cntl.ItemData(cntl.NewIndex) = tb.Fields("code").Value
        tb.MoveNext
    Loop
cntl.ListIndex = 0
End If
tb.Close
db.Close
End Sub
Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Set db = OpenDatabase(blm.patHmain)
If Option2 = True Then
    ssql = "delete from heads where code = " & Val(Text1.Text)
    db.Execute ssql
End If


Set tb = db.OpenRecordset("heads", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("h_type").Value = Combo2.ItemData(Combo2.ListIndex)
tb.Update
tb.Close
db.Close

End Sub
Private Function check(s As String) As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM heads WHERE NAME = '" & s & "'"
Set db = OpenDatabase(blm.patHmain)
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

ssql = "select MAX(CODE) AS C FROM heads WHERE CODE BETWEEN 11 AND 99"
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 11
End If
tb.Close
db.Close
End Function

Private Sub Com1_Click()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Set db = OpenDatabase(blm.patHmain)
If Option2 = True Then
    ssql = "Select * from heads where Code between " & Val(Text1.Text) * 1000 & " and " & Val(Text1.Text) * 1000 + 1000
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
        MsgBox "You Already Have Sub heads in this Head"
        Exit Sub
    Else
        ssql = "delete from heads where code = " & Val(Text1.Text)
        db.Execute ssql
    End If
    tb.Close
End If
db.Close
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
If KeyAscii = 13 Then Combo2.SetFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Command1_Click()
Dim p As Boolean

p = Option2.Value
Call save
'MSAVE Val(Text1.Text), UCase(Text2.Text), p
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
combs
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
Command1.Enabled = False
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

Private Sub Form_Load()
Set blm = New bloom1
Text1.Text = max1
combs
End Sub

Private Sub Form_Paint()
Option1 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Text1.Enabled = False
    Text1.Text = max1
    Text2.Visible = True
    Text2.SetFocus
    Combo1.Visible = False
Else

End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
    combs
   
    Text1.Text = vbNullString

    Combo1.Visible = True
    Text2.Visible = False
    Combo1.SetFocus
Else
    Text1.Enabled = False
End If

End Sub

Private Sub edit1()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM heads WHERE code = " & Val(Text1.Text)
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    'Text2.Text = tb.Fields("name").Value
    Combo2.ListIndex = tb.Fields("h_type").Value - 1
    Text1.Enabled = False
    'Combo1.Visible = False
Else
    MsgBox "Invalid Head's Code"
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
If KeyAscii = 13 Then SendKeys ("{tab}")

End Sub

Private Sub Text2_LostFocus()
If Option1 = True Then
Dim b As Boolean

b = check(UCase(CStr(Text2.Text)))
If b = True Then
    MsgBox "HEAD ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
End Sub
