VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EmpOverTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Over Time Management"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Lists"
      Height          =   795
      Left            =   135
      TabIndex        =   20
      Top             =   3000
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   135
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   390
         Width           =   4245
      End
      Begin VB.Label Label7 
         Caption         =   "Employee List"
         Height          =   285
         Left            =   135
         TabIndex        =   21
         Top             =   180
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   555
         Left            =   810
         Picture         =   "EmpOvertime.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   555
         Left            =   2550
         Picture         =   "EmpOvertime.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   255
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   4620
      TabIndex        =   16
      Top             =   0
      Width           =   1320
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   855
         Left            =   240
         Picture         =   "EmpOvertime.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2790
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   855
         Left            =   240
         Picture         =   "EmpOvertime.frx":5386
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   240
         Picture         =   "EmpOvertime.frx":57C8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1050
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   855
         Left            =   240
         Picture         =   "EmpOvertime.frx":5C0A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1350
      Left            =   135
      TabIndex        =   12
      Top             =   1665
      Width           =   4455
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   975
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1470
         TabIndex        =   4
         Top             =   960
         Width           =   450
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1470
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label8 
         Caption         =   "Short Time Hours"
         Height          =   255
         Left            =   2220
         TabIndex        =   23
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Over Time Hours"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Emp. A/c Title"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Emp. A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   945
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3510
         TabIndex        =   1
         Top             =   240
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   780
         TabIndex        =   0
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   73859075
         CurrentDate     =   37710
      End
      Begin VB.Label Label2 
         Caption         =   "Adv.Deduction"
         Height          =   255
         Left            =   2370
         TabIndex        =   24
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "EmpOverTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Private Sub Combs()
Dim Ssql As String
Ssql = "select * from acchart where Mid(Code,1,2)='" & EmpHead & "' order by name"
Blm.fill_comb Ssql, Combo1, "name", "code"
End Sub
Private Sub SHowRecord()
Dim B As Boolean
Dim Rs As Recordset
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)
Dim Ssql As String

Ssql = "Select * from OverTime where Accode=" & Text3.Text
Ssql = Ssql & " and ADate = #" & DTPicker1.Value & "#"
Set Rs = DB.OpenRecordset(Ssql)
If Not Rs.EOF Then
    Option2 = True
    Command4.Enabled = True
    Text3.Text = Rs.Fields("Accode").Value
    Text4.Text = Blm.party1(Rs.Fields("Accode").Value)
    Text5.Text = Rs.Fields("OHours").Value & ""
    Text6.Text = Rs.Fields("SHours").Value & ""
    Text1.Text = Rs.Fields("AdvDeduction").Value & ""
Else
    Command4.Enabled = False
    MsgBox "No Record Found"
    Option1 = True
End If
Rs.Close


DB.Close

End Sub
Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim B As Boolean
If Option2 = True Then
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from OverTime where "
    Ssql = Ssql & " ADate = #" & DTPicker1.Value & "#"
    Ssql = Ssql & " and Accode = " & Text3.Text
    DB.Execute Ssql
    
    DB.Close
End If
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset("OverTime", dbOpenTable)
TB.AddNew
    TB.Fields("Adate").Value = DTPicker1.Value
    TB.Fields("Accode").Value = Text3.Text
    TB.Fields("OHours").Value = Val(Text5.Text)
    TB.Fields("SHours").Value = Val(Text6.Text)
    TB.Fields("AdvDeduction").Value = Val(Text1.Text)
TB.Update
TB.Close

End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex > -1 Then
    Text3.Text = Combo1.ItemData(Combo1.ListIndex)
    Text4.Text = Combo1.Text
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Command1_Click()
If Val(Text3.Text) > 0 Then
Screen.MousePointer = vbHourglass
save
Command2_Click
Screen.MousePointer = vbDefault
Else
    MsgBox "Please Complete This Voucher"
End If
End Sub

Private Sub Command2_Click()
Option1 = True
Text1.Text = ""
Text3.Text = ""
Text5.Text = ""
Text6.Text = ""
Text4.Text = ""
Combs
Command4.Enabled = False
DTPicker1.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command4_Click()
Dim Ssql As String
Dim R As VbMsgBoxResult
R = MsgBox("Do You Want to Delete This OverTime", vbApplicationModal + vbYesNo)
If R = vbYes Then
If Option2 = True Then
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from OverTime where "
    Ssql = Ssql & " Accode=" & Text3.Text
    Ssql = Ssql & " and ADate = #" & DTPicker1.Value & "#"
    DB.Execute Ssql
    
    DB.Close
    Command2_Click
End If
End If
End Sub

Private Sub Command5_Click()
End Sub


Private Sub DTPicker1_LostFocus()
    If DTPicker1.Value >= FStartDate And DTPicker1.Value <= FEndDate Then
      ' Text1.Text = max1
      ' Edit1
    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If
End Sub

Private Sub Form_Activate()
Combs
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
End Sub

Private Sub Option1_Click()
Command2_Click

DTPicker1.SetFocus
End Sub

Private Sub Option2_Click()

DTPicker1.SetFocus
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text3.Text = SelectedAccountCode
    Text4.Text = SelectedAccountName
End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Val(Text3.Text) <> 0 Then
    Text4.Text = Blm.party1(Val(Text3.Text))
    If Text4.Text = "NOT" Then
        Cancel = True
    Else
        If Blm.GetEmpStatus(Val(Text3.Text)) = 1 Then
            MsgBox "This Employee has Left So Please Don't Post Any Data in It"
            Cancel = True
        End If
        'If Option2 = True Then
            SHowRecord
        'End If
    End If
        
End If

End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Combo2.SetFocus
End If

End Sub

Private Sub Text8_Validate(Cancel As Boolean)
If Val(Text8.Text) <> 0 Then
    
    Text9.Text = Blm.party1(Val(Text8.Text))
    If Text9.Text = "NOT" Then
        Cancel = True
        
    End If
        
End If

End Sub
