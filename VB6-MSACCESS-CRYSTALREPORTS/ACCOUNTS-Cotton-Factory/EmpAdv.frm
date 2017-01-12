VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EmpAdv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Advance Management"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Lists"
      Height          =   1155
      Left            =   135
      TabIndex        =   27
      Top             =   4500
      Width           =   5775
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1410
         TabIndex        =   31
         Text            =   "Combo1"
         Top             =   645
         Width           =   4245
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1425
         TabIndex        =   30
         Text            =   "Combo1"
         Top             =   270
         Width           =   4245
      End
      Begin VB.Label Label11 
         Caption         =   "Credit List"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Employee List"
         Height          =   285
         Left            =   225
         TabIndex        =   28
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   555
         Left            =   810
         Picture         =   "EmpAdv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   555
         Left            =   2550
         Picture         =   "EmpAdv.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   255
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   135
      TabIndex        =   21
      Top             =   3525
      Width           =   4455
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Credit A/c Title"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Credit A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4500
      Left            =   4620
      TabIndex        =   20
      Top             =   0
      Width           =   1320
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   855
         Left            =   240
         Picture         =   "EmpAdv.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3465
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   855
         Left            =   240
         Picture         =   "EmpAdv.frx":5386
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2475
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   240
         Picture         =   "EmpAdv.frx":57C8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1470
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   855
         Left            =   240
         Picture         =   "EmpAdv.frx":5C0A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   465
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1350
      Left            =   120
      TabIndex        =   16
      Top             =   2145
      Width           =   4455
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   1125
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "0.00"
         Height          =   255
         Left            =   3060
         TabIndex        =   32
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Advance"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Emp. A/c Title"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Emp. A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   945
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20774915
         CurrentDate     =   37710
      End
      Begin VB.Label Label3 
         Caption         =   "Details"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Voucher No."
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "EmpAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Private Sub Combs()
Dim Ssql As String
Ssql = "select * from acchart where Mid(Code,1,2)='" & EmpHead & "' order by name"
Blm.fill_comb Ssql, Combo1, "name", "code"
Ssql = "select * from acchart where Mid(Code,1,2)<>'" & EmpHead & "' order by name"
Blm.fill_comb Ssql, Combo2, "name", "code"
End Sub
Private Sub SHowRecord()
Dim B As Boolean
Dim Rs As Recordset
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)
Dim Ssql As String

Ssql = "Select * from VouMST where V_Type=11 and V_No=" & Val(Text1.Text)
Set Rs = DB.OpenRecordset(Ssql)
If Not Rs.EOF Then
    DTPicker1.Value = Rs.Fields("V_Date").Value
Else
    MsgBox "No Data Found For This Advance Voucher"
    Exit Sub
End If
Rs.Close

Ssql = "Select * from VouDTL where V_type=11 and V_No=" & Val(Text1.Text) & " and Debit>0"
Set Rs = DB.OpenRecordset(Ssql)
If Not Rs.EOF Then
    Text3.Text = Rs.Fields("Party").Value
    Text4.Text = Blm.party1(Rs.Fields("Party").Value)
    Text2.Text = Rs.Fields("Remarks").Value & ""
    Text5.Text = Rs.Fields("Debit").Value
End If
Rs.Close

Ssql = "Select * from VouDTL where V_type=11 and V_No=" & Val(Text1.Text) & " and Credit>0"
Set Rs = DB.OpenRecordset(Ssql)
If Not Rs.EOF Then
    Text8.Text = Rs.Fields("Party").Value
    Text9.Text = Blm.party1(Rs.Fields("Party").Value)
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
    'b = checkdate(date1.Value)
    If B = True Then
        Exit Sub
    End If
End If
If Option2 = True Then
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from voumst where "
    Ssql = Ssql & " v_type = 11"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where "
    Ssql = Ssql & " v_type = 11"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    DB.Close
End If
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset("voumst", dbOpenTable)
TB.AddNew
    TB.Fields("v_date").Value = DTPicker1.Value
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = 11
    TB.Fields("narration").Value = "Employee Advance Voucher for Emp : " & Text4.Text
TB.Update
TB.Close
Set TB = DB.OpenRecordset("voudtl", dbOpenTable)

TB.AddNew
    TB.Fields("v_date").Value = DTPicker1.Value
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = 11
    TB.Fields("party").Value = Val(Text3.Text)
    TB.Fields("remarks").Value = Text2.Text
    TB.Fields("debit").Value = Val(Text5.Text)
    TB.Fields("credit").Value = 0
TB.Update

TB.AddNew
    TB.Fields("v_date").Value = DTPicker1.Value
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = 11
    TB.Fields("party").Value = Text8.Text
    TB.Fields("remarks").Value = "Advance Voucher of Employee " & Text4.Text
    TB.Fields("debit").Value = 0
    TB.Fields("credit").Value = Val(Text5.Text)
TB.Update


TB.Close
DB.Close

End Sub
Private Function AccountName(C As String, TBox As Control) As Boolean
Dim RST As New ADODB.Recordset
Dim Ssql As String

Ssql = "Select * from Acchart where Ac_Code = '" & C & "'"
Set RST = CN.Execute(Ssql)
If Not RST.EOF Then
TBox = RST.Fields("AC_Name").Value & ""
AccountName = False
Else
TBox = "Invalid Account Code"
AccountName = True
End If
RST.Close


End Function

Private Function max1()
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(v_no)as c from voumst where v_type = 11"
    
    Set DB = OpenDatabase(Blm.patHmain)
    Set TB = DB.OpenRecordset(Ssql)
    If IsNull(TB.Fields("c").Value) = False Then
        max1 = TB.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    TB.Close
    DB.Close
End Function

Private Sub Check1_Click()
Text1.Enabled = Not Text1.Enabled
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
Private Sub Combo2_Click()
If Combo2.ListIndex > -1 Then
    Text8.Text = Combo2.ItemData(Combo2.ListIndex)
    Text9.Text = Combo2.Text
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text8.SetFocus
End Sub

Private Sub Command1_Click()
If Val(Text1.Text) > 0 And Val(Text3.Text) > 0 And Val(Text8.Text) > 0 Then
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
Text1.Text = max1
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

Combs
DTPicker1.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command4_Click()
Dim Ssql As String
Dim R As VbMsgBoxResult
R = MsgBox("Do You Want to Delete This Voucher", vbApplicationModal + vbYesNo)
If R = vbYes Then
If Option2 = True Then
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from voumst where "
    Ssql = Ssql & " v_type = 11"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where "
    Ssql = Ssql & " v_type = 11"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    DB.Close
    
End If
End If
End Sub

Private Sub Command5_Click()
End Sub

Private Sub Command6_Click()
Rs.MoveFirst
SHowRecord

End Sub

Private Sub Command7_Click()
If Not Rs.BOF Then
    Rs.MovePrevious
    SHowRecord
Else
    MsgBox "You Already At First Record"
End If


End Sub

Private Sub Command8_Click()
If Not Rs.EOF Then
    Rs.MoveNext
    SHowRecord
Else
    MsgBox "You Already At Last Record"
End If

End Sub

Private Sub DTPicker1_LostFocus()
    If DTPicker1.Value >= FStartDate And DTPicker1.Value <= FEndDate Then
       '
        'edit1
    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If
Text1.Text = max1
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
Text1.Enabled = False
DTPicker1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Text = ""
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
    Command4.Enabled = True
Else
    Command4.Enabled = False
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Option2 = True Then
    SHowRecord
End If
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
        Label8.Caption = Blm.CurrentBalance(Val(Text3.Text))
        If Blm.GetEmpStatus(Val(Text3.Text)) = 1 Then
            MsgBox "This Employee has Left So Please Don't Post Any Data in It"
            Cancel = True
        End If
    End If
        
End If

End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text8.Text = SelectedAccountCode
    Text9.Text = SelectedAccountName
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
