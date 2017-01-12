VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SalVoucher 
   Caption         =   "Employee Sallary Payment Voucher"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   795
      Left            =   120
      TabIndex        =   42
      Top             =   6090
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1980
         TabIndex        =   47
         Top             =   450
         Width           =   1170
      End
      Begin VB.TextBox txtOverTime 
         Height          =   285
         Left            =   60
         TabIndex        =   44
         Top             =   450
         Width           =   930
      End
      Begin VB.TextBox txtShortTime 
         Height          =   285
         Left            =   990
         TabIndex        =   43
         Top             =   450
         Width           =   960
      End
      Begin VB.Label Label13 
         Caption         =   "Adv.Deduction"
         Height          =   180
         Left            =   2010
         TabIndex        =   48
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Over Time"
         Height          =   180
         Left            =   90
         TabIndex        =   46
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Short Time"
         Height          =   180
         Left            =   1020
         TabIndex        =   45
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   10635
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   6195
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5370
      TabIndex        =   40
      Text            =   "Combo1"
      Top             =   6240
      Width           =   4125
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Apply"
      Height          =   390
      Left            =   3420
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Frame Frame4 
      Height          =   1485
      Left            =   5340
      TabIndex        =   35
      Top             =   90
      Width           =   2355
      Begin VB.OptionButton Option4 
         Caption         =   "&Update"
         Height          =   975
         Left            =   1185
         Picture         =   "SalVoucher.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   255
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&New"
         Height          =   975
         Left            =   225
         Picture         =   "SalVoucher.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   255
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   6660
      TabIndex        =   27
      Top             =   7590
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Command9 
         CausesValidation=   0   'False
         Height          =   615
         Left            =   1200
         Picture         =   "SalVoucher.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         CausesValidation=   0   'False
         Height          =   615
         Left            =   2160
         Picture         =   "SalVoucher.frx":5386
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         CausesValidation=   0   'False
         Height          =   615
         Left            =   3120
         Picture         =   "SalVoucher.frx":57C8
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         CausesValidation=   0   'False
         Height          =   615
         Left            =   4080
         Picture         =   "SalVoucher.frx":5C0A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&new"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2640
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   11655
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   10320
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7500
         TabIndex        =   23
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1305
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60162051
         CurrentDate     =   37709
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Month Days"
         Height          =   255
         Left            =   9000
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Total Deductions"
         Height          =   255
         Left            =   9000
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Total Sallaries"
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Debit Title"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Debit Code"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Voucher Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher No."
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid G1 
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1485
      Left            =   7740
      TabIndex        =   12
      Top             =   90
      Width           =   4035
      Begin VB.CommandButton Command5 
         Caption         =   "&Delete"
         Height          =   975
         Left            =   2970
         Picture         =   "SalVoucher.frx":604C
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   270
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Exit"
         Height          =   975
         Left            =   1980
         Picture         =   "SalVoucher.frx":648E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   975
         Left            =   990
         Picture         =   "SalVoucher.frx":68D0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   975
         Left            =   120
         Picture         =   "SalVoucher.frx":6D12
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   120
      TabIndex        =   9
      Top             =   90
      Width           =   5175
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   1050
         Width           =   3735
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3330
         Top             =   120
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Calculate Salaries"
         Height          =   825
         Left            =   3240
         Picture         =   "SalVoucher.frx":7154
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60162051
         CurrentDate     =   37709
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   630
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60162051
         CurrentDate     =   37709
      End
      Begin VB.Label Label14 
         Caption         =   "Sub Head"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Debit A/c List"
      Height          =   240
      Left            =   4350
      TabIndex        =   39
      Top             =   6300
      Width           =   1035
   End
End
Attribute VB_Name = "SalVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Dim CRow As Integer
Private Sub SHowRecord()
Dim B As Boolean
Dim R As Long
Dim DB As Database
Dim RST As Recordset
Dim Ssql As String

flex1

Ssql = "Select * from SalVoucher where V_No = " & Val(Text1.Text)
Set DB = OpenDatabase(Blm1.patHmain)
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
Text7.Text = RST.Fields("MonthDays").Value & ""
DTPicker1.Value = RST.Fields("SDate")
DTPicker2.Value = RST.Fields("EDate")
        DTPicker3.Value = RST.Fields("V_date").Value
        Text1.Text = RST.Fields("V_No").Value
        Text2.Text = RST.Fields("DebitCode").Value
        Text3.Text = Blm1.party1(RST.Fields("DebitCode").Value)
        
        For R = 0 To Combo1.ListCount - 1
            If Mid(Combo1.ItemData(R), 3, 3) = Mid(Text2.Text, 3, 3) Then
                Combo1.ListIndex = R
                Exit For
            End If
        Next R
        
        For R = 0 To Combo2.ListCount - 1
            If Mid(Combo2.ItemData(R), 3, 3) = Mid(Text2.Text, 3, 3) Then
                Combo2.ListIndex = R
                Exit For
            End If
        Next R

     
     Do While Not RST.EOF
        With G1
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = .Rows - 1
        .TextMatrix(.Rows - 1, 1) = RST.Fields("EmpCode").Value
        .TextMatrix(.Rows - 1, 2) = Blm1.party1(RST.Fields("EmpCode").Value)
        .TextMatrix(.Rows - 1, 3) = RST.Fields("Presents").Value
        .TextMatrix(.Rows - 1, 4) = RST.Fields("Absents").Value
        .TextMatrix(.Rows - 1, 5) = RST.Fields("Leaves").Value
        .TextMatrix(.Rows - 1, 6) = RST.Fields("Holidays").Value & ""
        .TextMatrix(.Rows - 1, 7) = RST.Fields("OverTime").Value & ""
        .TextMatrix(.Rows - 1, 8) = RST.Fields("ShortTime").Value & ""
        .TextMatrix(.Rows - 1, 9) = RST.Fields("SalRate").Value & ""
        .TextMatrix(.Rows - 1, 10) = RST.Fields("Advance").Value & ""
        .TextMatrix(.Rows - 1, 11) = RST.Fields("Deduction").Value & ""
        .TextMatrix(.Rows - 1, 12) = Round(RST.Fields("OTAmount").Value, 2) & ""
        .TextMatrix(.Rows - 1, 13) = Round(RST.Fields("STAmount"), 2) & ""
        .TextMatrix(.Rows - 1, 14) = Round(RST.Fields("SalAmount").Value, 2) & ""
        .TextMatrix(.Rows - 1, 15) = Round(RST.Fields("AdvBal").Value, 2) & ""
        .TextMatrix(.Rows - 1, 16) = Round(RST.Fields("TotalDeduction").Value, 2) & ""
        .TextMatrix(.Rows - 1, 17) = Round(RST.Fields("Payment").Value, 2) & ""
        .TextMatrix(.Rows - 1, 18) = RST.Fields("C").Value & ""
        .TextMatrix(.Rows - 1, 19) = RST.Fields("Y").Value & ""
        End With
        RST.MoveNext
    Loop
End If
RST.Close
DB.Close

End Sub
Private Sub save()
Dim Ssql As String
Dim R As Long
Dim A As String
Dim RST As Recordset
Dim DB As Database

Set DB = OpenDatabase(Blm1.patHmain)
If Option4 = True Then
    Ssql = "Delete from VouMSt where V_Type=15 and V_No = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "Delete from VouDTL where V_Type=15 and V_No = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "Delete from SalVoucher where V_No=" & Val(Text1.Text)
    DB.Execute Ssql
End If


Set RST = DB.OpenRecordset("VouMST", dbOpenTable)
RST.AddNew
    RST.Fields("V_Date").Value = DTPicker3.Value
    RST.Fields("V_Type").Value = 15
    RST.Fields("V_No").Value = Val(Text1.Text)
    RST.Fields("Narration").Value = "Salaries Voucher from : " & Format(DTPicker1.Value, "dd-MMM-yyyy") & " To : " & Format(DTPicker2.Value, "dd-MMM-yyyy")
RST.Update
RST.Close

Set RST = DB.OpenRecordset("VouDTL", dbOpenTable)
For R = 1 To G1.Rows - 1
RST.AddNew
    RST.Fields("V_Date").Value = DTPicker3.Value
    RST.Fields("V_Type").Value = 15
    RST.Fields("V_No").Value = Val(Text1.Text)
    RST.Fields("Remarks").Value = "Sal. " & Format(DTPicker1.Value, "dd-MMM-yyyy") & " To : " & Format(DTPicker2.Value, "dd-MMM-yyyy") & " P:" & G1.TextMatrix(R, 3) & " L:" & G1.TextMatrix(R, 5) & " H:" & G1.TextMatrix(R, 6) & " A:" & G1.TextMatrix(R, 4) & " OT:" & G1.TextMatrix(R, 7) & " ST: " & G1.TextMatrix(R, 8) & " C: " & G1.TextMatrix(R, 18) & " Y: " & G1.TextMatrix(R, 19) & " Adv.Ded:" & G1.TextMatrix(R, 11)
    RST.Fields("Party").Value = Val(G1.TextMatrix(R, 1))
    RST.Fields("Debit").Value = 0
    RST.Fields("Credit").Value = Val(G1.TextMatrix(R, 17))
RST.Update
Next R

RST.AddNew
    RST.Fields("V_Date").Value = DTPicker3.Value
    RST.Fields("V_Type").Value = 15
    RST.Fields("V_No").Value = Val(Text1.Text)
    RST.Fields("Remarks").Value = "Salaries Voucher from : " & Format(DTPicker1.Value, "dd-MMM-yyyy") & " To : " & Format(DTPicker2.Value, "dd-MMM-yyyy")
    RST.Fields("Party").Value = Val(Text2.Text)
    RST.Fields("Debit").Value = Val(Text4.Text)
    RST.Fields("Credit").Value = 0
RST.Update

RST.Close

Set RST = DB.OpenRecordset("SalVoucher", dbOpenTable)
For R = 1 To G1.Rows - 1
RST.AddNew
    RST.Fields("MonthDays").Value = Val(Text7.Text)
    RST.Fields("V_no").Value = Val(Text1.Text)
    RST.Fields("V_Date").Value = DTPicker3.Value
    RST.Fields("SDate").Value = DTPicker1.Value
    RST.Fields("EDate").Value = DTPicker2.Value
    RST.Fields("DebitCode").Value = Val(Text2.Text)
    RST.Fields("SHCode").Value = Val(Mid(G1.TextMatrix(R, 1), 1, 5))
    RST.Fields("SHName").Value = Blm1.SubHeadName(Val(Mid(G1.TextMatrix(R, 1), 1, 5)))
    RST.Fields("EmpCode").Value = Val(G1.TextMatrix(R, 1))
    RST.Fields("EmpName").Value = G1.TextMatrix(R, 2)
    RST.Fields("Presents").Value = Val(G1.TextMatrix(R, 3))
    RST.Fields("Absents").Value = Val(G1.TextMatrix(R, 4))
    RST.Fields("Leaves").Value = Val(G1.TextMatrix(R, 5))
    RST.Fields("Holidays").Value = Val(G1.TextMatrix(R, 6))
    RST.Fields("OverTime").Value = Val(G1.TextMatrix(R, 7))
    RST.Fields("ShortTime").Value = Val(G1.TextMatrix(R, 8))
    RST.Fields("SalRate").Value = Val(G1.TextMatrix(R, 9))
    RST.Fields("Advance").Value = Val(G1.TextMatrix(R, 10))
    RST.Fields("Deduction").Value = Val(G1.TextMatrix(R, 11))
    RST.Fields("OTAmount").Value = Val(G1.TextMatrix(R, 12))
    RST.Fields("STAmount").Value = Val(G1.TextMatrix(R, 13))
    RST.Fields("SalAmount").Value = Val(G1.TextMatrix(R, 14))
    RST.Fields("AdvBal").Value = Val(G1.TextMatrix(R, 15))
    RST.Fields("TotalDeduction").Value = Val(G1.TextMatrix(R, 16))
    RST.Fields("Payment").Value = Val(G1.TextMatrix(R, 17))
    RST.Fields("C").Value = Val(G1.TextMatrix(R, 18))
    RST.Fields("Y").Value = Val(G1.TextMatrix(R, 19))
RST.Update
Next R
RST.Close
DB.Close
End Sub
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(v_no)as c from voumst where v_type = 15"
    
    Set DB = OpenDatabase(Blm1.patHmain)
    Set TB = DB.OpenRecordset(Ssql)
    If IsNull(TB.Fields("c").Value) = False Then
        max1 = TB.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    TB.Close
    DB.Close
End Function

Private Sub flex1()
With G1
    .Rows = 1
    .Cols = 20
    
    .ColWidth(0) = 500
    .TextMatrix(0, 0) = "Sr#"
    
    .ColWidth(1) = 1800
    .TextMatrix(0, 1) = "Emp A/c Code"
    
    .ColWidth(2) = 2800
    .TextMatrix(0, 2) = "Emp A/c Title"
    
    .ColWidth(3) = 800
    .TextMatrix(0, 3) = "Presents"
    
    .ColWidth(4) = 800
    .TextMatrix(0, 4) = "Absents"
    
    .ColWidth(5) = 800
    .TextMatrix(0, 5) = "Leaves"
    
    .ColWidth(6) = 800
    .TextMatrix(0, 6) = "Holidays"
    
    .ColWidth(7) = 800
    .TextMatrix(0, 7) = "OverTime"
    
    .ColWidth(8) = 800
    .TextMatrix(0, 8) = "Short Time"
    
    .ColWidth(9) = 1200
    .TextMatrix(0, 9) = "Sal. Rate"
    
    
    .ColWidth(10) = 1200
    .TextMatrix(0, 10) = "T.Advance"
    
    .ColWidth(11) = 1200
    .TextMatrix(0, 11) = "Adv.Deduction"
    
    
    .ColWidth(12) = 1000
    .TextMatrix(0, 12) = "OT Amount"
    
    .ColWidth(13) = 1000
    .TextMatrix(0, 13) = "ST Amount"
    
    .ColWidth(14) = 1000
    .TextMatrix(0, 14) = "Sal. Amt"
    
    .ColWidth(15) = 1200
    .TextMatrix(0, 15) = "Adv.Balance."
    
    .ColWidth(16) = 1200
    .TextMatrix(0, 16) = "Total Deduction"
    
    .ColWidth(17) = 1000
    .TextMatrix(0, 17) = "Net.Total"
    
    .ColWidth(18) = 1000
    .TextMatrix(0, 18) = "C"
    
    .ColWidth(19) = 1000
    .TextMatrix(0, 19) = "Y"
    
    
End With
End Sub
Private Sub CalcSallries()
Dim Ssql As String
Dim R As Integer
Dim PACCode As String
Dim DB As Database
Dim RST As Recordset
Dim CBal As Double
Dim DRate As Double
Dim SRate As Double, SRateHour As Double
Dim GSal As Double, OTimeSal As Double
Dim NSal As Double
Dim RSEmp As Recordset
Dim RSF As Recordset
Set DB = OpenDatabase(Blm1.patHmain)
Ssql = "Select * from Acchart where Status=0 and Mid(Code,1,5)='" & Val(Combo2.ItemData(Combo2.ListIndex)) & "' Order by Code"
Set RSEmp = DB.OpenRecordset(Ssql)
If Not RSEmp.EOF Then
    RSEmp.MoveLast
    RSEmp.MoveFirst
End If

If Val(Text7.Text) <= 0 Then
    MsgBox "Please Enter The Days of the Month to Calculate Sallaries"
    Exit Sub
End If
flex1
Ssql = "Select Ac_Code,Status,Count(*) as C from EmpATD where A_Date Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# Group by Ac_Code,Status"
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    RST.MoveLast
    RST.MoveFirst
End If
If Not RSEmp.EOF Then
    Do While Not RSEmp.EOF
        With G1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = RSEmp.Fields("Code").Value
            .TextMatrix(.Rows - 1, 2) = RSEmp.Fields("Name").Value & ""
            If RST.RecordCount > 0 Then
            RST.Filter = "Ac_Code=" & RSEmp.Fields("Code").Value
            Set RSF = RST.OpenRecordset
            If Not RSF.EOF Then
                Do While Not RSF.EOF
                Select Case RSF.Fields("Status")
                    Case "P"
                        .TextMatrix(.Rows - 1, 3) = RSF.Fields("C").Value
                    Case "A"
                        .TextMatrix(.Rows - 1, 4) = RSF.Fields("C").Value
                    Case "L"
                        .TextMatrix(.Rows - 1, 5) = RSF.Fields("C").Value
                    Case "*"
                    
                        .TextMatrix(.Rows - 1, 6) = RSF.Fields("C").Value
                    Case "C"
                    
                        .TextMatrix(.Rows - 1, 18) = RSF.Fields("C").Value
                    Case "Y"
                    
                        .TextMatrix(.Rows - 1, 19) = RSF.Fields("C").Value
                      
                End Select
                RSF.MoveNext
                Loop
            End If
            RSF.Close
            RST.Filter = ""
            End If
            .TextMatrix(.Rows - 1, 7) = Blm1.GetEmpOverTime(RSEmp.Fields("Code").Value, DTPicker1.Value, DTPicker2.Value)
            .TextMatrix(.Rows - 1, 8) = Blm1.GetEmpShorTime(RSEmp.Fields("Code").Value, DTPicker1.Value, DTPicker2.Value)
            .TextMatrix(.Rows - 1, 9) = Blm1.GetEmpSalRate(RSEmp.Fields("Code").Value)
            .TextMatrix(.Rows - 1, 10) = Blm1.GetEmpAdvance(RSEmp.Fields("Code").Value, DTPicker1.Value, DTPicker2.Value)
            .TextMatrix(.Rows - 1, 11) = Blm1.GetEmpAdvDeduction(RSEmp.Fields("Code").Value, DTPicker1.Value, DTPicker2.Value)
            'Per Day Salary Rate
            SRate = Val(.TextMatrix(.Rows - 1, 9)) / Val(Text7.Text)
            'Per HourSalary Rate
            SRateHour = SRate / Val(RSEmp.Fields("Hours").Value & "")
            GSal = SRate * ((Val(.TextMatrix(.Rows - 1, 3)) + Val(.TextMatrix(.Rows - 1, 5))) + Val(.TextMatrix(.Rows - 1, 6)) + Val(.TextMatrix(.Rows - 1, 18)) + Val(.TextMatrix(.Rows - 1, 19)))
            
            OTimeSal = (Val(.TextMatrix(.Rows - 1, 7))) * SRateHour
            STimeSal = Val(.TextMatrix(.Rows - 1, 8)) * SRateHour
'            MsgBox GSal
            '==============================
            .TextMatrix(.Rows - 1, 12) = Round((Val(.TextMatrix(.Rows - 1, 7)) * SRateHour), 2)
            .TextMatrix(.Rows - 1, 13) = Round((Val(.TextMatrix(.Rows - 1, 8)) * SRateHour), 2)
            .TextMatrix(.Rows - 1, 14) = Round(GSal, 2)
            .TextMatrix(.Rows - 1, 15) = Round(Val(.TextMatrix(.Rows - 1, 10)) - Val(.TextMatrix(.Rows - 1, 11)), 2)
            .TextMatrix(.Rows - 1, 16) = Format(OTimeSal + Val(.TextMatrix(.Rows - 1, 11)), "#.00")
            NSal = ((GSal + OTimeSal) - STimeSal) '- Val(.TextMatrix(.Rows - 1, 11))
          '  MsgBox NSal
            .TextMatrix(.Rows - 1, 17) = Format(NSal, "#.00")
            
            
   
            
            
        End With
        PACCode = RSEmp.Fields("Code")
    RSEmp.MoveNext
    Loop

End If
RSEmp.Close

Ssql = "Select Accode,Sum(OHours) as O,Sum(SHours) as S from OverTime where ADate Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# Group by Accode"
Set RSEmp = DB.OpenRecordset(Ssql)
Do While Not RSEmp.EOF
    For R = 1 To G1.Rows - 1
        If G1.TextMatrix(R, 1) = RSEmp.Fields("Accode").Value Then
            G1.TextMatrix(R, 7) = RSEmp.Fields("O").Value & ""
            G1.TextMatrix(R, 8) = RSEmp.Fields("S").Value & ""
            Exit For
        End If
    Next R
RSEmp.MoveNext
Loop
RSEmp.Close
RST.Close
DB.Close
End Sub


Private Sub Combo1_Click()
If Combo1.ListIndex > -1 Then
  Text2.Text = Combo1.ItemData(Combo1.ListIndex)
  Text3.Text = Combo1.Text
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Combo2_Click()
If Combo2.ListIndex > -1 Then
'MsgBox "Test"
Dim R As Integer
For R = 0 To Combo1.ListCount - 1
    If Mid(Combo1.ItemData(R), 3, 3) = Mid(Combo2.ItemData(Combo2.ListIndex), 3, 3) Then
        Combo1.ListIndex = R
        Exit For
    End If
Next R
End If
End Sub

Private Sub Command1_Click()
CalcSallries
End Sub



Private Sub Command2_Click()
Screen.MousePointer = vbHourglass
If G1.Rows > 1 And Val(Text1.Text) > 0 And Val(Text2.Text) > 0 Then
    save
    Command3_Click
Else
    MsgBox "Please Complete This Voucher"
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text6.Visible = False
DTPicker3.Value = Date
G1.Rows = 1
Text1.Text = max1
DTPicker3.SetFocus
Option1 = True
End Sub

Private Sub Command4_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command5_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
If Result = vbNo Then Exit Sub


    Ssql = "Delete from VouMSt where V_Type=15 and V_No = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "Delete from VouDTL where V_Type=15 and V_No = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "Delete from SalVoucher where V_No=" & Val(Text1.Text)
    DB.Execute Ssql
    
    Command3_Click
End Sub

Private Sub Command6_Click()
If CRow > 0 Then
G1.TextMatrix(CRow, 7) = txtOverTime.Text
G1.TextMatrix(CRow, 8) = txtShortTime.Text
G1.TextMatrix(CRow, 11) = Text8.Text
G1.TextMatrix(CRow, 15) = Val(G1.TextMatrix(CRow, 10)) - Val(G1.TextMatrix(CRow, 11))
G1.TextMatrix(CRow, 16) = Val(G1.TextMatrix(CRow, 11)) + Val(G1.TextMatrix(CRow, 13))
G1.TextMatrix(CRow, 17) = Val(G1.TextMatrix(CRow, 14)) - Val(G1.TextMatrix(CRow, 16))
End If
End Sub

Private Sub DTPicker3_LostFocus()
    If DTPicker3.Value >= FStartDate And DTPicker3.Value <= FEndDate Then
    '    Text1.Text = max1
    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If
End Sub

Private Sub Form_Activate()
Dim Ssql As String
Dim RS As Recordset
Dim DB As Database
Ssql = "Select * from Acchart where Mid(Code,1,2)='59' Order by Name"
Blm1.fill_comb Ssql, Combo1, "Name", "Code"

Ssql = "Select * from Heads where Code >= 32000 and Code <=32999  Order by Code"
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set RS = DB.OpenRecordset(Ssql)
Combo2.clear
If Not RS.EOF Then
    Do While Not RS.EOF
        Combo2.AddItem RS.Fields("Code").Value & " - " & RS.Fields("Name").Value
        Combo2.ItemData(Combo2.NewIndex) = RS.Fields("Code").Value
    RS.MoveNext
    Loop
End If
RS.Close
DB.Close
'Blm1.fill_comb Ssql, Combo2, "Name", "Code"
End Sub

Private Sub Form_Load()


Text1.Text = max1
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
End Sub

Private Sub G1_Click()
If G1.Rows > 1 And G1.Cols > 2 Then
    CRow = G1.Row
    txtOverTime.Text = G1.TextMatrix(G1.Row, 7)
    txtShortTime.Text = G1.TextMatrix(G1.Row, 8)
End If
End Sub

Private Sub Option3_Click()
Command3_Click
Command5.Visible = False
Text1.Enabled = False
Text1.Text = max1
DTPicker1.SetFocus

End Sub

Private Sub Option4_Click()
Command3_Click
Command5.Visible = True
Text1.Enabled = True
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    SHowRecord
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text2.Text = SelectedAccountCode
    Text3.Text = SelectedAccountName
End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) <> 0 Then
    
    Text3.Text = Blm1.party1(Val(Text2.Text))
    If Text3.Text = "NOT" Then
        Cancel = True
    
        
    End If
        
End If


End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)

End Sub

Private Sub Timer1_Timer()
Dim R As Long
Dim TSAl As Double
Dim TDed As Double
Dim AMT As Double
Dim SRate As Double, GSal As Double
If G1.Cols >= 9 Then
            


For R = 1 To G1.Rows - 1

'SRate = Val(G1.TextMatrix(G1.Rows - 1, 6)) / Val(Text7.Text)
'GSal = Val(G1.TextMatrix(G1.Rows - 1, 6)) - (Val(G1.TextMatrix(G1.Rows - 1, 4)) * SRate)
'NSal = GSal - Val(G1.TextMatrix(G1.Rows - 1, 8))
'G1.TextMatrix(G1.Rows - 1, 9) = Format(NSal, "#.00")

TSAl = TSAl + Val(G1.TextMatrix(R, 17))
TDed = TDed + Val(G1.TextMatrix(R, 16))
Next R
Text4.Text = Format(TSAl, "#.00")
Text5.Text = Format(TDed, "#.00")
End If

If Len(Text2.Text) <= 0 Then
    Command2.Enabled = False
    Exit Sub
Else
    Command2.Enabled = True
End If
If G1.Rows <= 1 Then
    Command2.Enabled = False
    Exit Sub
Else
    Command2.Enabled = True
End If
    
End Sub

