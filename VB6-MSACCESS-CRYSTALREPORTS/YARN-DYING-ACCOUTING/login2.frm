VERSION 5.00
Begin VB.Form login2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   2400
      Picture         =   "login2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   855
      Left            =   915
      Picture         =   "login2.frx":0561
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "User Name"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "login2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub edit1()
Dim ssql As String
Dim db As Database
Dim tb As Recordset

ssql = "select * from list where username='" & UCase(Trim(Text1.Text)) & "'"
ssql = ssql & " and password = '" & Trim(Text2.Text) & "'"
Set db = OpenDatabase(App.Path & "\user.mdb")
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    User = UCase(Text1.Text)
    Me.Hide
    Unload Me
    frmSplash.Show
    
Else
    MsgBox "Invalid User Information....."
    Text1.SetFocus
End If
tb.Close
db.Close
End Sub

Private Sub Command1_Click()
Call edit1

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
