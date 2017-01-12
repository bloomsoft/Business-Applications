VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form acchart1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Information"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "acchart1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Heads Information"
      Height          =   735
      Left            =   240
      TabIndex        =   29
      Top             =   960
      Width           =   5775
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Heads Info"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sub Head Information"
      Height          =   855
      Left            =   240
      TabIndex        =   23
      Top             =   1680
      Width           =   5775
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Select A Sub Head"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   21
      Top             =   5760
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   735
         Left            =   4200
         Picture         =   "acchart1.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   735
         Left            =   2400
         Picture         =   "acchart1.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   735
         Left            =   600
         Picture         =   "acchart1.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   3135
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   5775
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5070
         TabIndex        =   9
         Top             =   1650
         Width           =   555
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check if Employee Left"
         Height          =   285
         Left            =   2985
         TabIndex        =   36
         Top             =   2025
         Width           =   2190
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   1650
         Width           =   1200
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   12
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   420
         Left            =   4065
         TabIndex        =   31
         Top             =   270
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   44498947
         CurrentDate     =   36747
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Timing"
         Height          =   285
         Left            =   5100
         TabIndex        =   37
         Top             =   1290
         Width           =   585
      End
      Begin VB.Label Label12 
         Caption         =   "Sal. Rate"
         Height          =   270
         Left            =   2985
         TabIndex        =   35
         Top             =   1665
         Width           =   870
      End
      Begin VB.Label Label11 
         Caption         =   "Phone"
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Address"
         Height          =   255
         Left            =   840
         TabIndex        =   33
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "S.T. Reg. #"
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Credit"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Debit"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "A/c Name"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   5775
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   555
         Left            =   3720
         Picture         =   "acchart1.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   555
         Left            =   825
         Picture         =   "acchart1.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6960
      Width           =   1335
   End
End
Attribute VB_Name = "acchart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As bloom1
Private Function opdate() As Date
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(Blm.patHmain)
Ssql = "select v_date from voumst where v_type=6"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    opdate = TB.Fields("v_date").Value
Else
    opdate = Date
End If
TB.Close
DB.Close
End Function

Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim opendate As Date
Set DB = OpenDatabase(Blm.patHmain)
'ssql = "delete from voumst where v_type = 6 and v_no = 1"
'db.Execute ssql
If Option2 = True Then
    Ssql = "delete from acchart where code = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where v_type = 10 and v_no = 1 and party = " & Val(Text1.Text)
    DB.Execute Ssql
End If


Set TB = DB.OpenRecordset("acchart", dbOpenTable)
TB.AddNew
    TB.Fields("CODE").Value = Val(Text1.Text)
    TB.Fields("NAME").Value = CStr(Text2.Text)
    TB.Fields("opdate").Value = date1.Value
    TB.Fields("debit").Value = Val(Text3.Text)
    TB.Fields("Credit").Value = Val(Text4.Text)
    TB.Fields("STNo").Value = Text5.Text
    TB.Fields("Phone").Value = Text7.Text
    TB.Fields("Address").Value = Text6.Text
    TB.Fields("SalRate").Value = Val(Text8.Text)
    TB.Fields("Hours").Value = Val(Text9.Text)
    TB.Fields("Status").Value = Check1.Value
TB.Update
TB.Close
Dim Tb2 As Recordset
If Val(Text3.Text) > 0 Or Val(Text4.Text) > 0 Then
Ssql = "Select * from voumst where v_type=10 and v_no = 1"
Set TB = DB.OpenRecordset(Ssql)
If TB.EOF Then
Set Tb2 = DB.OpenRecordset("voumst", dbOpenTable)
Tb2.AddNew
    Tb2.Fields("v_date").Value = date1.Value
    Tb2.Fields("v_type").Value = 10
    Tb2.Fields("v_no").Value = 1
    Tb2.Fields("narration").Value = "Open Balance"
Tb2.Update
Tb2.Close
opendate = date1.Value
Else
    opendate = TB.Fields("v_date").Value
End If

Set Tb2 = DB.OpenRecordset("voudtl", dbOpenTable)
Tb2.AddNew
    Tb2.Fields("v_date").Value = date1.Value
    Tb2.Fields("v_type").Value = 10
    Tb2.Fields("v_no").Value = 1
    Tb2.Fields("party").Value = Val(Text1.Text)
    Tb2.Fields("debit").Value = Val(Text3.Text)
    Tb2.Fields("credit").Value = Val(Text4.Text)
Tb2.Update
Tb2.Close
TB.Close
End If
DB.Close
If Left(Text1.Text, 5) = "32001" Then
    'save2
End If
End Sub
Private Sub save2()
'Dim DB As Database
'Dim TB As Recordset
'Dim Ssql As String
'Dim opendate As Date
'Set DB = OpenDatabase(Blm.patHmain)
''ssql = "delete from voumst where v_type = 6 and v_no = 1"
''db.Execute ssql
'If Option2 = True Then
'    Ssql = "delete from acchart where code = " & Val(Text1.Text)
'    DB.Execute Ssql
'
'End If
'
'
'Set TB = DB.OpenRecordset("acchart", dbOpenTable)
'TB.AddNew
'    TB.Fields("CODE").Value = Val(Replace(Val(Text1.Text), "32001", "58003"))
'    MsgBox TB.Fields("Code").Value
'    TB.Fields("NAME").Value = UCase(CStr(Text2.Text))
'    TB.Fields("opdate").Value = date1.Value
'    TB.Fields("debit").Value = Val(Text3.Text)
'    TB.Fields("Credit").Value = Val(Text4.Text)
'    TB.Fields("STNo").Value = Text5.Text
'    TB.Fields("Phone").Value = Text7.Text
'    TB.Fields("Address").Value = Text6.Text
'    TB.Fields("SalRate").Value = Val(Text8.Text)
'    TB.Fields("Status").Value = Check1.Value
'TB.Update
'TB.Close
'
'DB.Close

End Sub

Private Function check(S As String) As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM acchart WHERE NAME = '" & S & "'"
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
If Combo2.ListIndex > -1 Then
Ssql = "select MAX(CODE) AS C FROM acchart WHERE CODE between " & Combo2.ItemData(Combo2.ListIndex) * 10000 & " and " & Combo2.ItemData(Combo2.ListIndex) * 10000 + 10000
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("C").Value) Then
    max1 = TB.Fields("C").Value + 1
Else
    max1 = Combo2.ItemData(Combo2.ListIndex) * 10000 + 1
End If
TB.Close
DB.Close
End If
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
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Combo2_Click()
If Option1 = True Then
If Combo2.ListCount > 0 Then
Text1.Text = max1
End If
End If
If Option2 = True Then
If Combo2.ListIndex > -1 Then
Dim Ssql As String
Ssql = "select * from Acchart where code Between " & Combo2.ItemData(Combo2.ListIndex) * 10000 & " and " & Combo2.ItemData(Combo2.ListIndex) * 10000 + 10000 & " Order By Name"
'MsgBox ssql
Combo1.clear
comb_fill Combo1, Ssql
End If
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
'If Option1 = True Then
'    Text2.SetFocus
'End If
'End If
End Sub

Private Sub Combo3_Click()
If Combo3.ListCount > 0 Then
Dim Ssql As String
Ssql = "select * from heads where code > 100 and Mid(Cstr(Code),1,2) = '" & Combo3.ItemData(Combo3.ListIndex) & "' order by name"
Combo2.clear
comb_fill Combo2, Ssql
If Combo2.ListCount < 1 Then
    Command1.Enabled = False
    
End If
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command1_Click()
If Option2 = True Then
    If Val(Text1.Text) <= 0 Then
        MsgBox "Select New Option then Enter the New Account...."
        Exit Sub
    End If
End If
Call save

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
Text5.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = ""
Text9.Text = ""
Check1.Value = 0
If Option2 = True Then
    Combs
    Text1.Text = vbNullString
    Combo1.Visible = True
    Text2.Visible = False
Else
    Text1.Enabled = False
    Text1.Text = max1
    Combo1.Visible = False
    Text2.Visible = True
End If
Text3.Text = vbNullString
Text4.Text = vbNullString

Command1.Enabled = False
date1.Value = FStartDate - 1
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

Private Sub Combs()
Dim Ssql As String


Ssql = "select * from acchart where code > 99999 order by name"
comb_fill Combo1, Ssql
End Sub

Private Sub Command4_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete This Account", vbYesNo)
If Result = vbNo Then Exit Sub

Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)
If Option2 = True Then
    Ssql = "Select * from VouDtl where V_Type<>10 and Party = " & Val(Text1.Text)
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then
        MsgBox "You Already Have Transactions in this Detail A/c"
        Exit Sub
    Else
        Ssql = "delete from Acchart where code = " & Val(Text1.Text)
        DB.Execute Ssql
    End If
    TB.Close
End If
DB.Close
Command2_Click

End Sub

Private Sub Form_Activate()
Dim Ssql As String
Ssql = "select * from heads where code < 100 order by name"
comb_fill Combo3, Ssql
End Sub

Private Sub Form_Load()
Set Blm = New bloom1

date1.Value = FStartDate - 1
Combs
Text1.Text = max1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Blm = Nothing
End Sub

Private Sub Option1_Click()
If Option1 = True Then
    Text1.Enabled = False
    Text1.Text = max1
    Text2.Visible = True
    Text2.Text = vbNullString
    Command4.Visible = False
    Text2.SetFocus
    Combo1.Visible = False
    
Else

End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
Combs
    
    'Text1.Enabled = True
    Text1.Text = vbNullString
    Command4.Visible = True
    
    Combo1.Visible = True
    Text2.Visible = False
    
    
Combo3.SetFocus
    
Else
    Text1.Enabled = False
End If

End Sub

Private Sub edit1()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "SELECT * FROM acchart WHERE code = " & Val(Text1.Text)
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    Text2.Text = TB.Fields("name").Value
    Text1.Enabled = False
    Text3.Text = TB.Fields("debit").Value & ""
    Text4.Text = TB.Fields("credit").Value & ""
    'date1.Value = TB.Fields("opdate").Value
    Text5.Text = TB.Fields("STNo").Value & ""
    Text7.Text = TB.Fields("Phone").Value & ""
    Text6.Text = TB.Fields("Address").Value & ""
    Text8.Text = TB.Fields("SalRate").Value & ""
    Text9.Text = TB.Fields("Hours").Value & ""
    Check1.Value = Val(TB.Fields("Status").Value & "")
Else
    MsgBox "Invalid Accounts's Code"
    Combo1.SetFocus
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
'If KeyAscii = 13 Then date1.SetFocus
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_LostFocus()
If Option1 = True Then
Dim B As Boolean

B = check(UCase(CStr(Text2.Text)))
If B = True Then
    MsgBox "A/c ALREADY EXIST,,,,"
'    Text2.SetFocus
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text6_Change()
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text7_Change()
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
