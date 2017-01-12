VERSION 5.00
Begin VB.Form Items 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items Information"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Items.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   615
         Left            =   2760
         Picture         =   "Items.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   615
         Left            =   1560
         Picture         =   "Items.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   615
         Left            =   360
         Picture         =   "Items.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   2295
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   4215
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Com1 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   13
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
      Begin VB.Label Label5 
         Caption         =   "Opening Stock"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Unit For Measurements"
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Item's Description"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Item's Code"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   555
         Left            =   2400
         Picture         =   "Items.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   555
         Left            =   840
         Picture         =   "Items.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Private Sub combs()
Dim Ssql As String

Ssql = "select * from Items order by name"
comb_fill Combo1, Ssql

End Sub
Private Sub comb_fill(cntl As Control, Ssql As String)
Dim db As Database
Dim tb As Recordset
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(Ssql)
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
Dim Ssql As String
Set db = OpenDatabase(blm.patHmain)
If Option2 = True Then
    Ssql = "delete from Items where code = " & Val(Text1.Text)
    db.Execute Ssql
End If


Set tb = db.OpenRecordset("Items", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("Unit").Value = UCase(CStr(Text3.Text))
    tb.Fields("Stock").Value = Val(Text4.Text)
tb.Update
tb.Close
db.Close

End Sub
Private Function UnitRet(s As Integer) As String
Dim db As Database
Dim tb As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Items WHERE Code = " & s
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(Ssql)

If Not tb.EOF Then
    UnitRet = tb.Fields("Unit").Value
    Text4.Text = tb.Fields("Stock").Value
Else
    UnitRet = ""
End If
tb.Close
db.Close

End Function

Private Function check(s As String) As Boolean
Dim db As Database
Dim tb As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Items WHERE NAME = '" & s & "'"
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(Ssql)

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
Dim Ssql As String

Ssql = "select MAX(CODE) AS C FROM Items"
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(Ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
db.Close
End Function

Private Sub Com1_Click()
Dim db As Database
Dim tb As Recordset
Dim Ssql As String
Set db = OpenDatabase(blm.patHmain)
If Option2 = True Then
        Ssql = "delete from Items where code = " & Val(Text1.Text)
        db.Execute Ssql
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
If KeyAscii = 13 Then Text3.SetFocus
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
Text3.Text = vbNullString
Text4.Text = vbNullString
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
Command2_Click
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
Command2_Click
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
Dim Ssql As String

Ssql = "SELECT * FROM Items WHERE code = " & Val(Text1.Text)
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(Ssql)
If Not tb.EOF Then
    'Text2.Text = tb.Fields("name").Value
    Text3.Text = UnitRet(Val(Text1.Text))
    
    Text1.Enabled = False
    'Combo1.Visible = False
Else
    MsgBox "Invalid Item's Code"
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
    MsgBox "ITEM ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
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
