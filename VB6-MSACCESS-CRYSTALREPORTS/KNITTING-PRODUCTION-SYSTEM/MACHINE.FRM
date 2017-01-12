VERSION 5.00
Begin VB.Form machine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Machines Coding"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "MACHINE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2760
         Picture         =   "MACHINE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1560
         Picture         =   "MACHINE.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   855
         Left            =   360
         Picture         =   "MACHINE.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   4215
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "MACHINE.frx":0C28
         Left            =   1800
         List            =   "MACHINE.frx":0C32
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
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
      Begin VB.Label Label6 
         Caption         =   "Status"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Needles"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Sinker"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Machine Name"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Machine Code"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   2280
         Picture         =   "MACHINE.frx":0C45
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   600
         Picture         =   "MACHINE.frx":0F4F
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "machine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Private Sub combs()
Dim ssql As String

ssql = "select * from Machine order by name"
comb_fill Combo1, ssql

End Sub
Private Sub comb_fill(cntl As Control, ssql As String)
Dim tb As ADODB.Recordset
Set tb = CN.Execute(ssql)
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

End Sub
Private Sub save()
Dim tb As New ADODB.Recordset
Dim ssql As String
If Option2 = True Then
    ssql = "delete from Machine where code = " & Val(Text1.Text)
    CN.Execute ssql
End If


tb.Open "Machine", CN, 0, 3, 0
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("Sinker").Value = Text3.Text
    tb.Fields("Needles").Value = Text4.Text
    tb.Fields("Status").Value = Combo2.ItemData(Combo2.ListIndex)
tb.Update
tb.Close
End Sub
Private Function Check(S As String) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM Machine WHERE NAME = '" & S & "'"
Set tb = CN.Execute(ssql)

If Not tb.EOF Then
    Check = True
Else
    Check = False
End If
tb.Close

End Function
Private Function max1() As Long
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select MAX(CODE) AS C FROM Machine"
Set tb = CN.Execute(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close

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
'If KeyAscii = 13 Then Combo2.SetFocus
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
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = ""
Text4.Text = ""
If Option2 = True Then
    Text1.Enabled = True
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
Text1.SetFocus
End If

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
    Text1.Enabled = True
End If
End Sub

Private Sub Option2_Click()
Command2_Click
If Option2 = True Then
    combs
    'Text1.Enabled = True
    Text1.Text = vbNullString
    Combo1.SetFocus
    Combo1.Visible = True
    Text2.Visible = False
Else
    Text1.Enabled = False
End If

End Sub

Private Sub edit1()
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM Machine WHERE code = " & Val(Text1.Text)
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    'Text2.Text = tb.Fields("name").Value
'    Combo2.ListIndex = tb.Fields("h_type").Value - 1
    Text1.Enabled = False
    Text3.Text = tb.Fields("Needles").Value & ""
    Text4.Text = tb.Fields("Sinker").Value & ""
    Combo2.ListIndex = tb.Fields("Status").Value
    
    'Combo1.Visible = False
Else
    MsgBox "Invalid Machine's Code"
    'Combo1.SetFocus
End If
tb.Close

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
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


Private Sub Text2_LostFocus()
If Option1 = True Then
Dim b As Boolean

b = Check(UCase(CStr(Text2.Text)))
If b = True Then
    MsgBox "MACHINE ALREADY EXIST,,,,"
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
    Text2.SetFocus
End If
End If
End Sub
