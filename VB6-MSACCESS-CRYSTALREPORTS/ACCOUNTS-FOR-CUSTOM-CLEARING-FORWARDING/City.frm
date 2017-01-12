VERSION 5.00
Begin VB.Form City 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "City Coding"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   855
         Left            =   2760
         Picture         =   "City.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   1560
         Picture         =   "City.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   855
         Left            =   360
         Picture         =   "City.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "Combo2"
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
      Begin VB.Label Label2 
         Caption         =   "City Name"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "City Code"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "City"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1

Private Sub Combs()
Dim ssql As String

ssql = "select * from City order by name"
blm.fill_comb ssql, Combo1, "Name", "Code"
End Sub

Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
    ssql = "delete from City where code = " & Val(Text1.Text)
    db.Execute ssql
End If


Set tb = db.OpenRecordset("City", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
tb.Update
tb.Close
db.Close

End Sub
Private Function check(s As String) As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM City WHERE NAME = '" & s & "'"
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

ssql = "select MAX(CODE) AS C FROM City"
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 11
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
Call edit1
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Command1_Click()
Dim p As Boolean

p = Option2.Value
Call save
'MSAVE Val(Text1.Text), UCase(Text2.Text), p
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
Combs
Me.Top = 10
Me.Left = 10
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
    Text1.Enabled = True
End If
End Sub

Private Sub Option2_Click()
Command2_Click
If Option2 = True Then
    Combs
   
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

ssql = "SELECT * FROM City WHERE code = " & Val(Text1.Text)
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    'Text2.Text = tb.Fields("name").Value
    Text1.Enabled = False
    'Combo1.Visible = False
Else
    MsgBox "Invalid City Code"
    
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
    MsgBox "CITY ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
End Sub
