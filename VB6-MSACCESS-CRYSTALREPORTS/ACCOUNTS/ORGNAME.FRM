VERSION 5.00
Begin VB.Form orgname 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Or Change the Organization Name"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   3240
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   360
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      MaxLength       =   150
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      MaxLength       =   100
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "Cash A/c Code"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Phone"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "G.S.T. Registration No."
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Organization Name"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "orgname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm1 As New bloom1
Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

Set db = OpenDatabase(blm1.SettingsPath)
ssql = "delete from options"
db.Execute ssql
Set tb = db.OpenRecordset("Options", dbOpenTable)
tb.AddNew
    tb.Fields("OrgInfo").Value = Text1.Text
    tb.Fields("GSTNo").Value = Text2.Text
    tb.Fields("Address").Value = Text3.Text
    tb.Fields("Phone").Value = Text4.Text
    tb.Fields("CashAc").Value = Combo1.ItemData(Combo1.ListIndex)
tb.Update
tb.Close
db.Close

End Sub
Private Sub edit1()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim R As Long

Set db = OpenDatabase(blm1.SettingsPath)
ssql = "Select * from options"
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    Text1.Text = tb.Fields("OrgInfo").Value & ""
    Text2.Text = tb.Fields("GSTNo").Value & ""
    Text3.Text = tb.Fields("Address").Value & ""
    Text4.Text = tb.Fields("Phone").Value & ""
    For R = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(R) = tb.Fields("CashAc").Value Then
            Combo1.ListIndex = R
            Exit For
        End If
    Next R
Else
    MsgBox "No Organization Name Has Been Set..."
    Text1.Text = "Enter Orgnization Name Here"
    
End If
tb.Close
db.Close
End Sub
Private Sub Command1_Click()
save
Command2_Click
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me

End Sub

Private Sub Form_Load()

Dim ssql As String
ssql = "Select * from Acchart Order by Name"
blm1.fill_comb ssql, Combo1, "Name", "Code"

edit1

End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
