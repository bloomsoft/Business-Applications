VERSION 5.00
Begin VB.Form Groups1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items Groups"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
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
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2760
         Picture         =   "Groups1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1560
         Picture         =   "Groups1.frx":0561
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
         Picture         =   "Groups1.frx":0A68
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
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   390
         Left            =   2940
         TabIndex        =   14
         Top             =   390
         Visible         =   0   'False
         Width           =   1035
      End
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
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Group Name"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Group Code"
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
         CausesValidation=   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
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
Attribute VB_Name = "Groups1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As bloom1

Private Sub Combs()
Dim Ssql As String

Ssql = "select * from Groups order by name"
Blm.fill_comb Ssql, Combo1, "Name", "Code"
End Sub

Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)
If Option2 = True Then
    Ssql = "delete from Groups where code = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from Heads where code = " & Val(Text1.Text)
    DB.Execute Ssql
End If


Set TB = DB.OpenRecordset("Groups", dbOpenTable)
TB.AddNew
    TB.Fields("CODE").Value = Val(Text1.Text)
    TB.Fields("NAME").Value = CStr(Text2.Text)
TB.Update
TB.Close

Set TB = DB.OpenRecordset("Heads", dbOpenTable)
TB.AddNew
    TB.Fields("CODE").Value = Val(Text1.Text)
    TB.Fields("NAME").Value = CStr(Text2.Text)
    TB.Fields("H_type").Value = 1
TB.Update
TB.Close


DB.Close

End Sub
Private Function check(s As String) As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Groups WHERE NAME = '" & s & "'"
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)

If Not TB.EOF Then
    check = True
Else
    check = False
End If
TB.Close
DB.Close
End Function
Private Function max1() As Long
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "select MAX(CODE) AS C FROM Groups"
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("C").Value) Then
    max1 = TB.Fields("C").Value + 1
Else
    max1 = 11
End If
TB.Close
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
Text2.Text = Combo1.Text

End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
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
Text1.SetFocus
Else
Combo1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
Text2.Text = vbNullString
Combs
If Option2 = True Then
    
    Text1.Text = vbNullString
    Combo1.Visible = True

Else


    Combo1.Visible = False
 
End If
Command1.Enabled = False
If Option1 = True Then
Text1.SetFocus
Else
Combo1.SetFocus
End If

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Command4_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
If Result = vbNo Then Exit Sub


Dim TB As Recordset
Dim Ssql As String
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "Select * from SubGroups where GroupCode=" & Val(Text1.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    MsgBox "You Cannot Delete this Group as It Has Sub Groups"
Else
    Ssql = "Delete from Groups where Code=" & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "Delete from Heads where Code=" & Val(Text1.Text)
    DB.Execute Ssql
End If
TB.Close
DB.Close
Command2_Click

End Sub

Private Sub Form_Load()
Set Blm = New bloom1
'Text1.Text = max1
Combs
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Form_Paint()
Option1 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Blm = Nothing
End Sub

Private Sub Option1_Click()

If Option1 = True Then
    
    Command4.Visible = False
    Combo1.Visible = False
Else
End If
Command2_Click
End Sub

Private Sub Option2_Click()

If Option2 = True Then
    Combs
    Command4.Visible = True
    Text1.Text = vbNullString
    
    Combo1.Visible = True
    
    Combo1.SetFocus
Else
End If
Command2_Click
End Sub

Private Sub edit1()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Groups WHERE code = " & Val(Text1.Text)
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    'Text2.Text = tb.Fields("name").Value
    
    'Combo1.Visible = False
Else
    MsgBox "Invalid Group Code"
    
End If
TB.Close
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

Private Sub Text1_Validate(Cancel As Boolean)
If Option1 = True Then
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM Groups WHERE CODE = " & Val(Text1.Text)
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)

If Not TB.EOF Then
   MsgBox "Group Code Already Exist"
   Cancel = True
Else
    
End If
TB.Close
DB.Close
    
ElseIf Option2 = True Then
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
Dim B As Boolean

B = check(UCase(CStr(Text2.Text)))
If B = True Then
    MsgBox "GROUP ALREADY EXIST,,,,"
End If
End If
End Sub
