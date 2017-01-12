VERSION 5.00
Begin VB.Form Dyingcoding 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dying Coding"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   8655
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6360
         Picture         =   "DyingCoding.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3360
         Picture         =   "DyingCoding.frx":09A6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   600
         Picture         =   "DyingCoding.frx":1290
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "BloomSoft"
         Height          =   255
         Left            =   7680
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   8655
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3240
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   1320
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3240
         TabIndex        =   1
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3225
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Dying Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Dying Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   3120
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   8775
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   5400
         Picture         =   "DyingCoding.frx":1A87
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1800
         Picture         =   "DyingCoding.frx":24CC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   345
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Dyingcoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Private Function CheckCode() As Boolean
Dim DB As Database
Set DB = OpenDatabase(blm.pathMain)
Dim RST As Recordset
Dim Ssql As String
Dim R As Long
Ssql = "Select * from DyingChart where Code = " & Val(Text1.Text)
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    MsgBox "Dying Code Already Exist"
    CheckCode = True
Else
    
    CheckCode = False
End If
RST.Close
DB.Close
End Function

Private Sub Combs()
Dim Ssql As String

Ssql = "select * from DyingChart order by name"
blm.fill_comb Ssql, Combo1, "Name", "Code"
End Sub

Private Sub Save()
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
If Option2 = True Then
    Ssql = "delete from DyingChart where code = " & Val(Text1.Text)
    DB.Execute Ssql
End If


Set tb = DB.OpenRecordset("Dyingchart", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
tb.Update
tb.Close
DB.Close

End Sub
Private Function check(s As String) As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM DyingChart WHERE NAME = '" & s & "'"
Set DB = OpenDatabase(blm.pathMain)
Set tb = DB.OpenRecordset(Ssql)

If Not tb.EOF Then
    check = True
Else
    check = False
End If
tb.Close
DB.Close
End Function
Private Function Max1() As Long
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String

Ssql = "select MAX(CODE) AS C FROM Dyingchart"
Set DB = OpenDatabase(blm.pathMain)
Set tb = DB.OpenRecordset(Ssql)
If Not IsNull(tb.Fields("C").Value) Then
    Max1 = tb.Fields("C").Value + 1
Else
    Max1 = 11
End If
tb.Close
DB.Close
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


Private Sub Command1_Click()
Dim p As Boolean

p = Option2.Value
Call Save
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
    Combo1.Visible = False
    Text2.Visible = True
End If
Command1.Enabled = False
If Option1 = True Then
Text1.Enabled = True
Text1.SetFocus
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
Text1.Text = Max1
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
    Text2.Visible = True
    Text1.SetFocus
    Combo1.Visible = False
Else

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
End If

End Sub

Private Sub edit1()
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Dyingchart WHERE code = " & Val(Text1.Text)
Set DB = OpenDatabase(blm.pathMain)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    'Text2.Text = tb.Fields("name").Value
    Text1.Enabled = False
    'Combo1.Visible = False
Else
    MsgBox "Invalid Dying Code"
    
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
If Option2 = True Then
If Val(Text1.Text) > 0 Then
Call edit1
End If
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Option1 = True Then
    Cancel = CheckCode
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
    MsgBox "Dying ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
End Sub
