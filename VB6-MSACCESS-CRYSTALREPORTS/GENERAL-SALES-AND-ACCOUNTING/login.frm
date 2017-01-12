VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create A New User"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      CausesValidation=   0   'False
      Height          =   975
      Left            =   2880
      Picture         =   "login.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      CausesValidation=   0   'False
      Height          =   975
      Left            =   1800
      Picture         =   "login.frx":0561
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   975
      Left            =   720
      Picture         =   "login.frx":0A68
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete This User"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm Password"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "User Name"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clear()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text1.SetFocus
End Sub
Private Sub Command1_Click()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Set db = OpenDatabase(App.Path & "\user.mdb")
If Check3.Value = 1 Then
    ssql = "delete from list where username = '" & UCase(Trim(Text1.Text)) & "'"
    db.Execute ssql
End If

Set tb = db.OpenRecordset("list", dbOpenTable)
    tb.AddNew
        tb.Fields("username").Value = UCase(Trim(Text1.Text))
        tb.Fields("password").Value = Trim(Text2.Text)
    tb.Update
tb.Close
db.Close
Command2_Click
End Sub
Private Sub edit1()
Dim ssql As String
Dim db As Database
Dim tb As Recordset

ssql = "select * from list where username='" & UCase(Trim(Text1.Text)) & "'"
Set db = OpenDatabase(App.Path & "\user.mdb")
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    Check3.Value = 1
    Text1.Text = tb.Fields("username").Value
    Text2.Text = tb.Fields("password").Value
    Text3.Text = tb.Fields("password").Value
    
End If
tb.Close
db.Close
End Sub

Private Sub Command2_Click()
Call clear
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Command4_Click()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Set db = OpenDatabase(App.Path & "\user.mdb")
If Check3.Value = 1 Then
    ssql = "delete from list where username = '" & UCase(Trim(Text1.Text)) & "'"
    db.Execute ssql
End If
db.Close
Command2_Click
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Text1.Text <> vbNullString Or Text1.Text <> "" Then
    Call edit1
Else
    Cancel = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Timer1_Timer()
If Text1.Text <> vbNullString Or Text1.Text <> "" Then
    Command4.Enabled = True
Else
    Command4.Enabled = False
End If
End Sub
