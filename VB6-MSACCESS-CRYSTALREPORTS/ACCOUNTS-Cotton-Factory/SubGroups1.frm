VERSION 5.00
Begin VB.Form SubGroups1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items Groups"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Item Group"
      Height          =   600
      Left            =   225
      TabIndex        =   15
      Top             =   855
      Width           =   4200
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Text            =   "Combo2"
         Top             =   180
         Width           =   3870
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   255
      TabIndex        =   13
      Top             =   2910
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2760
         Picture         =   "SubGroups1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1560
         Picture         =   "SubGroups1.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   855
         Left            =   360
         Picture         =   "SubGroups1.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1335
      Left            =   255
      TabIndex        =   10
      Top             =   1470
      Width           =   4215
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2910
         TabIndex        =   16
         Top             =   375
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Text            =   "Combo2"
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1785
         MaxLength       =   5
         TabIndex        =   1
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Sub Group Name"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Sub Group Code"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   810
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   3615
      TabIndex        =   14
      Top             =   4110
      Width           =   855
   End
End
Attribute VB_Name = "SubGroups1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As bloom1

Private Sub Combs()
Dim Ssql As String

Ssql = "select * from SubGroups order by name"
Blm.fill_comb Ssql, Combo1, "Name", "Code"
End Sub

Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)
If Option2 = True Then
    Ssql = "delete from SubGroups where code = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from Heads where code = " & Val(Text1.Text)
    DB.Execute Ssql
End If


Set TB = DB.OpenRecordset("SubGroups", dbOpenTable)
TB.AddNew
    TB.Fields("CODE").Value = Val(Text1.Text)
    TB.Fields("NAME").Value = CStr(Text2.Text)
    TB.Fields("GroupCode").Value = Combo2.ItemData(Combo2.ListIndex)
TB.Update
TB.Close

Set TB = DB.OpenRecordset("Heads", dbOpenTable)
TB.AddNew
    TB.Fields("CODE").Value = Val(Text1.Text)
    TB.Fields("NAME").Value = CStr(Text2.Text)
TB.Update
TB.Close

DB.Close

End Sub
Private Function check(s As String) As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM SubGroups WHERE NAME = '" & s & "'"
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

Ssql = "select MAX(CODE) AS C FROM SubGroups where Code Between " & Combo2.ItemData(Combo2.ListIndex) * 1000 & " and " & Combo2.ItemData(Combo2.ListIndex) * 1000 + 1000
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("C").Value) Then
    max1 = TB.Fields("C").Value + 1
Else
    max1 = Combo2.ItemData(Combo2.ListIndex) * 1000 + 1
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



Private Sub Combo2_Click()
If Option1 = True Then
If Combo2.ListIndex > -1 Then
    Text1.Text = max1
End If
End If
If Option2 = True Then
    Dim Ssql As String
    
    Ssql = "Select * from SubGroups where Code Between " & Combo2.ItemData(Combo2.ListIndex) * 1000 & " and " & Combo2.ItemData(Combo2.ListIndex) * 1000 + 1000
    comb_fill Combo1, Ssql
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command1_Click()
Dim p As Boolean

p = Option2.Value
Call save
'MSAVE Val(Text1.Text), UCase(Text2.Text), p
Combs
Command2_Click
If Option1 = True Then
Combo2.SetFocus
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
Text1.Text = max1
Combo2.SetFocus
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

Ssql = "Select * from Items where SubGroupCode=" & Val(Text1.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    MsgBox "You Cannot Delete this Sub Group as It Has Items"
Else
    Ssql = "Delete from SubGroups where Code=" & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from Heads where code = " & Val(Text1.Text)
    DB.Execute Ssql
End If
TB.Close
DB.Close
Command2_Click

End Sub

Private Sub Form_Activate()
Dim Ssql As String
Ssql = "Select * from Groups Order by Name"
comb_fill Combo2, Ssql
End Sub

Private Sub Form_Load()
Set Blm = New bloom1

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
    Text1.Enabled = False
    Text1.Text = max1
    Text2.Visible = True
    Combo1.Visible = False
Else
End If
Command2_Click
End Sub

Private Sub Option2_Click()
Command2_Click
If Option2 = True Then
    Combs
    Command4.Visible = True
    Text1.Text = vbNullString
    Text1.Enabled = True
    Text2.Visible = False
    Combo1.Visible = True
    
    Combo2.SetFocus
Else
End If

End Sub

Private Sub edit1()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Ssql = "SELECT * FROM SubGroups WHERE code = " & Val(Text1.Text)
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    For R = 0 To Combo2.ListCount - 1
        If Combo2.ItemData(R) = TB.Fields("GroupCode").Value Then
            Combo2.ListIndex = R
            Exit For
        End If
    Next R
Else
    MsgBox "Invalid Sub Group Code"
    
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

Ssql = "SELECT * FROM SubGroups WHERE CODE = " & Val(Text1.Text)
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)

If Not TB.EOF Then
   MsgBox "Sub Group Code Already Exist"
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
    MsgBox "SUB GROUP ALREADY EXIST,,,,"
    
End If
End If
End Sub
