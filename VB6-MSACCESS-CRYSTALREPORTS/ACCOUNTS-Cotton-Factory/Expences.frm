VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Expences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Expenses"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "Expences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4530
   Begin VB.Frame Frame7 
      Height          =   945
      Left            =   1860
      TabIndex        =   40
      Top             =   990
      Width           =   915
      Begin VB.OptionButton Option4 
         Caption         =   "&New"
         Height          =   255
         Left            =   30
         TabIndex        =   42
         Top             =   150
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&Update"
         Height          =   255
         Left            =   30
         TabIndex        =   41
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Height          =   945
      Left            =   2820
      TabIndex        =   37
      Top             =   990
      Width           =   1665
      Begin VB.CommandButton Command7 
         Caption         =   "&Save as"
         Height          =   705
         Left            =   90
         TabIndex        =   39
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Delete"
         Height          =   705
         Left            =   840
         TabIndex        =   38
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Voucher Info"
      Height          =   915
      Left            =   30
      TabIndex        =   27
      Top             =   990
      Width           =   1785
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   510
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label16 
         Caption         =   "Ref#"
         Height          =   255
         Left            =   210
         TabIndex        =   29
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Number"
         Height          =   255
         Left            =   210
         TabIndex        =   28
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2940
      Top             =   2280
   End
   Begin VB.Frame Frame4 
      Caption         =   "Actions"
      Height          =   960
      Left            =   2460
      TabIndex        =   18
      Top             =   30
      Width           =   2040
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   690
         Left            =   1320
         Picture         =   "Expences.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   195
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   690
         Left            =   705
         Picture         =   "Expences.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   195
         Width           =   630
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   690
         Left            =   90
         Picture         =   "Expences.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   195
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction"
      Height          =   3540
      Left            =   30
      TabIndex        =   15
      Top             =   1890
      Width           =   4485
      Begin VB.Frame Frame3 
         Caption         =   "Material Cost"
         Height          =   2025
         Left            =   2160
         TabIndex        =   30
         Top             =   780
         Width           =   2205
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1110
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Sub Head Wise Cost"
            Height          =   195
            Left            =   270
            TabIndex        =   35
            Top             =   870
            Width           =   1785
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   540
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Item Wise Cost"
            Height          =   255
            Left            =   270
            TabIndex        =   33
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox txtMaterialCost 
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1665
            Width           =   1740
         End
         Begin VB.Label Label9 
            Caption         =   "Material Cost"
            Height          =   210
            Left            =   150
            TabIndex        =   32
            Top             =   1440
            Width           =   1110
         End
      End
      Begin VB.TextBox txtTotal 
         Height          =   300
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3015
         Width           =   1755
      End
      Begin VB.TextBox txtContractors 
         Height          =   285
         Left            =   255
         MaxLength       =   150
         TabIndex        =   9
         Top             =   3030
         Width           =   1755
      End
      Begin VB.TextBox txtSalaries 
         Height          =   285
         Left            =   255
         MaxLength       =   150
         TabIndex        =   8
         Top             =   2430
         Width           =   1755
      End
      Begin VB.TextBox txtMics 
         Height          =   285
         Left            =   255
         MaxLength       =   150
         TabIndex        =   7
         Top             =   1800
         Width           =   1755
      End
      Begin VB.TextBox txtELectricAmount 
         Height          =   285
         Left            =   2805
         TabIndex        =   5
         Top             =   435
         Width           =   1260
      End
      Begin VB.TextBox txtElectricRate 
         Height          =   285
         Left            =   1365
         TabIndex        =   4
         Top             =   450
         Width           =   1320
      End
      Begin VB.TextBox txtMaintain 
         Height          =   285
         Left            =   240
         MaxLength       =   150
         TabIndex        =   6
         Top             =   1215
         Width           =   1755
      End
      Begin VB.TextBox txtElectricUnits 
         Height          =   285
         Left            =   240
         MaxLength       =   150
         TabIndex        =   3
         Top             =   465
         Width           =   990
      End
      Begin VB.Label Label10 
         Caption         =   "Total"
         Height          =   240
         Left            =   2265
         TabIndex        =   25
         Top             =   2805
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Contractors Amount"
         Height          =   255
         Left            =   270
         TabIndex        =   24
         Top             =   2790
         Width           =   1830
      End
      Begin VB.Label Label7 
         Caption         =   "Salaries Expenses"
         Height          =   255
         Left            =   255
         TabIndex        =   23
         Top             =   2190
         Width           =   1830
      End
      Begin VB.Label Label4 
         Caption         =   "Miscelleneous Expenses"
         Height          =   255
         Left            =   270
         TabIndex        =   22
         Top             =   1560
         Width           =   1830
      End
      Begin VB.Label Label3 
         Caption         =   "Electric Amount"
         Height          =   225
         Left            =   2895
         TabIndex        =   21
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Electric Rate/Unit"
         Height          =   210
         Left            =   1380
         TabIndex        =   20
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label lblAvgBaleWT 
         Caption         =   "..."
         Height          =   255
         Left            =   7365
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Maintenance Expenses"
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   975
         Width           =   1830
      End
      Begin VB.Label Label5 
         Caption         =   "Electric Units"
         Height          =   255
         Left            =   225
         TabIndex        =   16
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voucher Information"
      Height          =   960
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   2430
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   810
         TabIndex        =   0
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74973187
         CurrentDate     =   36757
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Expences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(vno)as c from Expences"
    
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
Private Function GetCost()
Dim Ssql As String
Dim DB As Database
Dim TB As Recordset
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "Select Sum(Amount) as Amt from Issue where V_Date=#" & date1.Value & "#"
'Ssql = Ssql & " and VNo=" & Val(Text1.Text)
Ssql = Ssql & " and RefNo=" & Val(Text7.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Amt").Value) Then
    GetCost = TB.Fields("Amt").Value
Else
    GetCost = 0
End If
TB.Close
DB.Close
End Function

Private Function GetCostSH()
Dim Ssql As String
Dim DB As Database
Dim TB As Recordset
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "Select Sum(Amount) as Amt from IssueSH where V_Date=#" & date1.Value & "#"
'Ssql = Ssql & " and VNo=" & Val(Text1.Text)
Ssql = Ssql & " and RefNo=" & Val(Text7.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Amt").Value) Then
    GetCostSH = TB.Fields("Amt").Value
Else
    GetCostSH = 0
End If
TB.Close
DB.Close
End Function

Private Function edit1() As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Tb2 As Recordset
Dim p As Long
Set DB = OpenDatabase(Blm.patHmain)
Ssql = "select * from Expences where"
Ssql = Ssql & " VNo=" & Val(Text1.Text)
'Ssql = Ssql & " and RefNo=" & Val(Text7.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    Text7.Text = TB.Fields("RefNo").Value
    date1.Value = TB.Fields("EDate").Value
    txtElectricUnits.Text = TB.Fields("ElectricUnits") & ""
    txtElectricRate.Text = TB.Fields("ElectricRate").Value & ""
    txtELectricAmount.Text = TB.Fields("ElectricAmount").Value & ""
    
    txtMaintain.Text = TB.Fields("MainTain").Value & ""
    txtMics.Text = TB.Fields("Misc").Value & ""
    txtSalaries.Text = TB.Fields("Salaries").Value & ""
    txtContractors.Text = TB.Fields("Contractor").Value & ""
    txtMaterialCost.Text = TB.Fields("MatCost").Value & ""
    
    Text2.Text = TB.Fields("ItemCost").Value & ""
    Text3.Text = TB.Fields("SHCost").Value & ""
    If TB.Fields("CostType").Value = 1 Then Option1 = True
    If TB.Fields("CostType").Value = 2 Then Option2 = True
Else
    MsgBox "No Data For This Date..."
    edit1 = False
End If
TB.Close
DB.Close
End Function
Private Sub clearfull()
Dim CNTL As Control

For Each CNTL In Me.Controls
    If TypeOf CNTL Is TextBox Then CNTL.Text = vbNullString
'    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
Option1 = False
Option2 = False
Text1.Text = max1
End Sub

Private Sub Command1_Click()
If Val(Text1.Text) > 0 And Val(Text7.Text) > 0 Then
        Call save
        Command2_Click
Else
        MsgBox "Please Complete This Voucher"
End If
End Sub

Private Sub Command2_Click()
Call clearfull

date1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Itm As String, Qty As Double, Comm As String, NetRate As Double
Dim B As Boolean
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from Expences where"
    Ssql = Ssql & " VNo=" & Val(Text1.Text)
    'Ssql = Ssql & " and RefNo=" & Val(Text7.Text)
    DB.Execute Ssql
    DB.Close
Set DB = OpenDatabase(Blm.patHmain)

Set TB = DB.OpenRecordset("Expences", dbOpenTable)
TB.AddNew
            TB.Fields("Vno").Value = Val(Text1.Text)
            TB.Fields("RefNo").Value = Val(Text7.Text)
            TB.Fields("EDate").Value = date1.Value
            TB.Fields("ElectricUnits").Value = Val(txtElectricUnits.Text)
            TB.Fields("ElectricRate").Value = Val(txtElectricRate.Text)
            TB.Fields("ElectricAmount").Value = Val(txtELectricAmount.Text)
            TB.Fields("MainTain").Value = Val(txtMaintain.Text)
            TB.Fields("Misc").Value = Val(txtMics.Text)
            TB.Fields("Salaries").Value = Val(txtSalaries.Text)
            TB.Fields("Contractor").Value = Val(txtContractors.Text)
            TB.Fields("MatCost").Value = Val(txtMaterialCost.Text)
            TB.Fields("ItemCost").Value = Val(Text2.Text)
            TB.Fields("SHCost").Value = Val(Text3.Text)
            If Option1 = True Then TB.Fields("CostType").Value = 1
            If Option2 = True Then TB.Fields("CostType").Value = 2
            
TB.Update

TB.Close
DB.Close
End Sub


Private Sub Command6_Click()
End Sub

Private Sub Command7_Click()
Text1.Text = max1
save
Command2_Click
End Sub

Private Sub Command8_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
If Result = vbNo Then Exit Sub

Dim DB As Database
Dim Ssql As String

    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from Expences where"
    Ssql = Ssql & " VNo=" & Val(Text1.Text)
    'Ssql = Ssql & " and RefNo=" & Val(Text7.Text)
    DB.Execute Ssql
    DB.Close
    Command2_Click


End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub date1_LostFocus()

    If date1.Value >= FStartDate And date1.Value <= FEndDate Then
    '    Text1.Text = max1

    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If

End Sub

Private Sub Form_Load()
date1.Value = Date
If Screen.Width < 800 And Screen.Height < 600 Then
MsgBox "Please Set your Desktop 800 x 600 Then Try"
Me.Hide
Unload Me
End If
Text1.Text = max1
    

End Sub


Private Sub Option1_Click()
'Text1.Enabled = False
'date1.SetFocus
'Command6.Visible = False
End Sub

Private Sub Option2_Click()

'Command6.Visible = True
'Text1.Enabled = True
'Text1.SetFocus
End Sub


Private Sub Option3_Click()
Command2_Click
Text1.Enabled = True
Text1.SetFocus
Command7.Visible = True
Command8.Visible = True

End Sub

Private Sub Option4_Click()
Command2_Click
Text1.Enabled = False
Text7.SetFocus
Command7.Visible = False
Command8.Visible = False

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Val(Text1.Text) > 0 Then
        edit1
        Text2.Text = GetCost
        Text3.Text = GetCostSH
    End If

End Sub

Private Sub Text7_Validate(Cancel As Boolean)
If Val(Text7.Text) > 0 Then
      Text2.Text = GetCost
      Text3.Text = GetCostSH
End If

End Sub

Private Sub Timer1_Timer()
If Option1 = True Then txtMaterialCost.Text = Val(Text2.Text)
If Option2 = True Then txtMaterialCost.Text = Val(Text3.Text)
txtELectricAmount.Text = Val(txtElectricUnits.Text) * Val(txtElectricRate.Text)
txtTotal.Text = Val(txtELectricAmount.Text) + Val(txtMaintain.Text) + Val(txtMaterialCost.Text) + Val(txtMics.Text) + Val(txtSalaries.Text) + Val(txtContractors.Text)
End Sub

Private Sub txtContractors_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtELectricAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtElectricRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtElectricUnits_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtMaintain_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtMics_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtSalaries_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
