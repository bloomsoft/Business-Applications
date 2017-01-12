VERSION 5.00
Begin VB.Form Party 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Party Information"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   975
         Left            =   3720
         Picture         =   "Party.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   975
         Left            =   2520
         Picture         =   "Party.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   975
         Left            =   1320
         Picture         =   "Party.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   2055
      Left            =   225
      TabIndex        =   12
      Top             =   1440
      Width           =   6255
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4140
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1650
         Width           =   1995
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1635
         Width           =   1800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "NTN"
         Height          =   255
         Left            =   3660
         TabIndex        =   21
         Top             =   1665
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "GST No."
         Height          =   240
         Left            =   225
         TabIndex        =   20
         Top             =   1650
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "(F1) To Search Party to Edit"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "Phone"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Party Name"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Party Code"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   6255
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   915
         Left            =   3720
         Picture         =   "Party.frx":0F56
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   915
         Left            =   1320
         Picture         =   "Party.frx":147C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   5745
      TabIndex        =   16
      Top             =   4800
      Width           =   735
   End
End
Attribute VB_Name = "Party"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim R As Integer
Dim ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
    ssql = "delete from parties where code = " & Val(Text1.Text)
    db.Execute ssql
    
    
End If


Set tb = db.OpenRecordset("Parties", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("Address").Value = UCase(CStr(Text3.Text))
    tb.Fields("Phone").Value = Text5.Text
    tb.Fields("GSTNo").Value = Text4.Text
    tb.Fields("NTN").Value = Text6.Text
tb.Update
tb.Close
db.Close

End Sub
Private Function check(s As String) As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM Parties WHERE NAME = '" & s & "'"
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

ssql = "select MAX(CODE) AS C FROM Parties"

Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
db.Close

End Function


Private Sub Command1_Click()
Call save


Command2_Click

If Option1 = True Then
Text2.SetFocus
Else
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text5.Text = vbNullString
Text4.Text = vbNullString
Text6.Text = vbNullString
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Screen.MousePointer = vbHourglass
    Search2.Text3.Text = 7
    Search2.Show
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub Form_Load()
Set blm = New bloom1
Dim ssql As String

Text1.Text = max1
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub


Private Sub Option1_Click()
Command2_Click
If Option1 = True Then

    Text1.Enabled = False
    Text1.Text = max1
    Label6.Visible = False
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
    Label6.Visible = True
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
ssql = "select * from Parties where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    Text1.Text = tb.Fields("Code").Value & ""
    Text2.Text = tb.Fields("Name").Value & ""
    Text5.Text = tb.Fields("Phone").Value & ""
    Text3.Text = tb.Fields("Address").Value & ""
    Text4.Text = tb.Fields("GSTNO").Value & ""
    Text6.Text = tb.Fields("NTN").Value & ""
    edit1 = False
Else
    MsgBox "Invalid Party Code....."
    edit1 = True
    Exit Function
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
If KeyAscii = 13 Then SendKeys ("{tab}")
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
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub
