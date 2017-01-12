VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vou1P 
   Caption         =   "Purchase Voucher Entry"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "vou1P.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Save This Voucher as"
      Height          =   375
      Left            =   9870
      TabIndex        =   55
      Top             =   1230
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      Height          =   675
      Left            =   6780
      TabIndex        =   56
      Top             =   1560
      Width           =   4935
      Begin VB.Label lblBalance 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2550
         TabIndex        =   60
         Top             =   450
         Width           =   2355
      End
      Begin VB.Label lblrecWT 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   30
         TabIndex        =   59
         Top             =   450
         Width           =   2355
      End
      Begin VB.Label lblJobInfo 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      Height          =   1155
      Left            =   6750
      TabIndex        =   33
      Top             =   30
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   840
         Left            =   120
         Picture         =   "vou1P.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   840
         Left            =   1080
         Picture         =   "vou1P.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   4680
   End
   Begin VB.Frame Frame4 
      Caption         =   "Actions"
      Height          =   1155
      Left            =   9000
      TabIndex        =   30
      Top             =   30
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   840
         Left            =   1800
         Picture         =   "vou1P.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   840
         Left            =   960
         Picture         =   "vou1P.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   840
         Left            =   120
         Picture         =   "vou1P.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3315
      Left            =   240
      TabIndex        =   29
      Top             =   3900
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction"
      Height          =   1575
      Left            =   240
      TabIndex        =   25
      Top             =   2250
      Width           =   11535
      Begin VB.CheckBox Check1 
         Caption         =   "Check if GST Invoice Required"
         Height          =   255
         Left            =   90
         TabIndex        =   54
         Top             =   1020
         Width           =   2745
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   10260
         TabIndex        =   16
         Top             =   615
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox Text9 
         Height          =   300
         Left            =   4845
         TabIndex        =   12
         Top             =   645
         Width           =   630
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   7785
         MaxLength       =   255
         TabIndex        =   15
         Top             =   630
         Width           =   3600
      End
      Begin VB.TextBox Text12 
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
         Left            =   6405
         TabIndex        =   14
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5490
         TabIndex        =   13
         Top             =   645
         Width           =   915
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   8880
         Picture         =   "vou1P.frx":EDCC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   10080
         Picture         =   "vou1P.frx":1156E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   11
         Top             =   645
         Width           =   840
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
         Left            =   1095
         TabIndex        =   10
         Top             =   645
         Width           =   2925
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label lblStockAmt 
         Caption         =   "0.00"
         Height          =   255
         Left            =   6420
         TabIndex        =   65
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label lblBaleStock 
         Caption         =   "0"
         Height          =   255
         Left            =   4860
         TabIndex        =   64
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lblStock 
         Caption         =   "0.00"
         Height          =   255
         Left            =   5490
         TabIndex        =   63
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   6510
         TabIndex        =   32
         Top             =   285
         Width           =   1185
      End
      Begin VB.Label Label29 
         Caption         =   "Freight"
         Height          =   210
         Left            =   10290
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label28 
         Caption         =   "Bales"
         Height          =   240
         Left            =   4905
         TabIndex        =   48
         Top             =   285
         Width           =   525
      End
      Begin VB.Label Label25 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   7815
         TabIndex        =   45
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Weight (KGS)"
         Height          =   255
         Left            =   5490
         TabIndex        =   41
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "Rate"
         Height          =   255
         Left            =   4035
         TabIndex        =   28
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1335
         TabIndex        =   27
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   270
         TabIndex        =   26
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voucher Information"
      Height          =   2220
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   6450
      Begin VB.CommandButton Command6 
         Caption         =   "&Delete "
         Height          =   375
         Left            =   5430
         TabIndex        =   61
         Top             =   1530
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   1560
         Width           =   945
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   5250
         TabIndex        =   2
         Top             =   360
         Width           =   780
      End
      Begin VB.TextBox Text14 
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
         Left            =   3840
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1530
         Width           =   735
      End
      Begin VB.TextBox Text10 
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
         Left            =   3840
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   3300
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   36757
      End
      Begin VB.Label Label31 
         Caption         =   "Vehicle No."
         Height          =   285
         Left            =   2880
         TabIndex        =   53
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label30 
         Caption         =   "Job No."
         Height          =   240
         Left            =   4680
         TabIndex        =   50
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "...."
         Height          =   255
         Left            =   1440
         TabIndex        =   47
         Top             =   1875
         Width           =   1185
      End
      Begin VB.Label Label26 
         Caption         =   "Party Balance :"
         Height          =   255
         Left            =   330
         TabIndex        =   46
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Debit Name"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Debit Code"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "==="
         Height          =   255
         Left            =   1305
         TabIndex        =   40
         Top             =   1890
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "S.T. Reg. #"
         Height          =   255
         Left            =   345
         TabIndex        =   39
         Top             =   1890
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Invoice No."
         Height          =   255
         Left            =   330
         TabIndex        =   38
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Seller Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Seller Code"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher #"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   2910
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label16 
      Caption         =   "When in Seller Code Press (F1) to Select Accounts from List && (F1) in Item Code to Select Item From List"
      Height          =   255
      Left            =   300
      TabIndex        =   62
      Top             =   7590
      Width           =   7935
   End
   Begin VB.Label lblJobType 
      Height          =   255
      Left            =   6810
      TabIndex        =   57
      Top             =   1260
      Width           =   3075
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5760
      TabIndex        =   52
      Top             =   7215
      Width           =   750
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4980
      TabIndex        =   51
      Top             =   7215
      Width           =   750
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6525
      TabIndex        =   42
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10560
      TabIndex        =   31
      Top             =   7200
      Width           =   1215
   End
End
Attribute VB_Name = "vou1P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Private Function VerifyJob(j As Long) As Boolean
Dim RS As Recordset
Dim RS2 As Recordset
Dim DB As Database
Dim Ssql As String
Dim B As Boolean
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Ssql = "Select * from PurJob where V_no=" & j
'Ssql = Ssql & " and Buyer=" & Val(Text2.Text)
Set RS = DB.OpenRecordset(Ssql)
If Not RS.EOF Then
    Text2.Text = RS.Fields("Buyer").Value
    Text10.Text = Blm.party1(RS.Fields("Buyer"))
    Text3.Text = RS.Fields("Item").Value
    Text4.Text = Blm.item1(RS.Fields("Item"))
    Select Case RS.Fields("CTerms").Value
    
         Case 1 'WT Based
            lblJobType.Caption = "Job Information : " & "Weight Based"
            lblJobInfo.Caption = "Quantity : " & RS.Fields("Qty").Value & " ,  Rate : " & RS.Fields("Rate").Value
            lblrecWT.Caption = ""
            lblBalance.Caption = ""
        Case 2
            lblJobType.Caption = "Job Information : " & "Consignments"
            lblJobInfo.Caption = "Contracted Consignments : " & RS.Fields("Consignments").Value & " ,  Rate : " & RS.Fields("Rate").Value
            lblrecWT.Caption = ""
            lblBalance.Caption = ""
        Case 3
            lblJobType.Caption = "Job Information : " & "Periodic"
            lblJobInfo.Caption = "From : " & Format(RS.Fields("SDate").Value, "dd-MMM-yyyy") & " To : " & Format(RS.Fields("EDate").Value, "dd-MMM-yyyy") & " ,  Rate : " & RS.Fields("Rate").Value
            lblrecWT.Caption = "Days Passed : " & Abs(DateDiff("d", date1.Value, RS.Fields("SDate").Value))
            lblBalance.Caption = "Days Left : " & Abs(DateDiff("d", date1.Value, RS.Fields("EDate").Value))
            'MsgBox "Test"
            If date1.Value >= RS.Fields("SDate").Value And date1.Value <= RS.Fields("EDate").Value Then
            
            Else
                MsgBox "Please Check The Voucher Date, As It is Not in the Contracted Period of This Contract"
            End If
        Case 4
            lblJobType.Caption = "Job Information : " & "Others"
            lblJobInfo.Caption = RS.Fields("OtherTerms").Value & "" & " ,  Rate : " & RS.Fields("Rate").Value
        Case 5
            lblJobType.Caption = "Job Information : " & "Rate Based"
            lblJobInfo.Caption = "Contracted Rate : " & RS.Fields("Rate").Value & ""
    End Select
Else
    MsgBox "Either This No. Contract Don't Exist or Don't Belong to This Party"
    VerifyJob = True
End If

If B = False Then
Ssql = "Select Sum(Qty) as Q,Count(Bales) as C from Purchase where JobNo=" & j
Set RS2 = DB.OpenRecordset(Ssql)
If Not IsNull(RS2.Fields("Q").Value) Then
'    lblrecWT.Caption = "Purchased Qty. " & RS2.Fields("Q").Value & ""
'    lblBalance.Caption = "Balance : " & Rs.Fields("Qty").Value - RS2.Fields("Q").Value
    
    Select Case RS.Fields("CTerms")
        Case 1 'WT Based
            If RS.Fields("Qty").Value <= RS2.Fields("Q").Value Then
                MsgBox "The Contracted Weight Has Been Received"
            End If
            lblrecWT.Caption = "Purchased Quantity. " & RS2.Fields("Q").Value & ""
            lblBalance.Caption = "Balance : " & RS.Fields("Qty").Value - RS2.Fields("Q").Value
        Case 2
            If RS2.Fields("C").Value >= RS.Fields("Consignments").Value Then
                MsgBox "Contracted Consignments Has Been Received"
            End If
            lblrecWT.Caption = "Purchased Consignments " & RS2.Fields("C").Value & ""
            lblBalance.Caption = "Balance : " & RS.Fields("Consignments").Value - RS2.Fields("C").Value
        Case 3
            
            If date1.Value < RS.Fields("SDate").Value And date1.Value > RS.Fields("EDate").Value Then
                MsgBox "Please Check The Voucher Date, As It is Not in the Contracted Period of This Contract"
                
            End If
    End Select

    
    
Else
    'lblrecWT.Caption = ""
End If
RS2.Close
End If
RS.Close
DB.Close
End Function
Private Sub ShowStockAccount(C As Double)
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)
Ssql = "Select * from Items where Code=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    If Not IsNull(TB.Fields("AccodE").Value) Then
        Text13.Text = TB.Fields("Accode").Value
        Text14.Text = Blm.party1(Val(Text13.Text))
    End If
End If
TB.Close
DB.Close
End Sub

Private Function checkdate(V_Date As Date) As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "Select * from voudtl where v_date > #" & V_Date & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    MsgBox "You Cannot Enter or Update the Voucher in Back Dates...."
    checkdate = True
Else
    checkdate = False
End If
TB.Close
DB.Close
End Function
Private Function edit1() As Boolean
flex1
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Tb2 As Recordset
Dim p As Long
Set DB = OpenDatabase(Blm.patHmain)
Ssql = "select * from Purchase where v_type = 4"
Ssql = Ssql & " and v_no = " & Val(Text1.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    If Not IsNull(TB.Fields("GSTInv").Value) Then Check1.Value = TB.Fields("GSTINV").Value
    date1.Value = TB.Fields("v_date").Value
    Text2.Text = TB.Fields("Seller").Value
    Text10.Text = Blm.party1(TB.Fields("Seller").Value)
    Text13.Text = TB.Fields("DebitCode").Value
    Text14.Text = Blm.party1(TB.Fields("DebitCode").Value)
    Text11.Text = TB.Fields("Inv_no").Value & ""
    Text6.Text = TB.Fields("Lorryno").Value & ""
    Text17.Text = TB.Fields("JobNo").Value & ""
            grid1.Rows = 1
            Do While Not TB.EOF
                grid1.Rows = grid1.Rows + 1
                With grid1
                    .TextMatrix(.Rows - 1, 0) = .Rows - 1
                    
                    .TextMatrix(.Rows - 1, 1) = TB.Fields("Item").Value
                    .TextMatrix(.Rows - 1, 2) = Blm.item1(TB.Fields("item").Value)
                    .TextMatrix(.Rows - 1, 3) = Format(TB.Fields("Rate").Value, "#.00")
                    .TextMatrix(.Rows - 1, 4) = TB.Fields("Bales").Value
                    .TextMatrix(.Rows - 1, 5) = Format(TB.Fields("Qty").Value, "#.00")
                    .TextMatrix(.Rows - 1, 6) = Format(TB.Fields("Amount").Value, "#.00")
                    .TextMatrix(.Rows - 1, 7) = TB.Fields("Remarks").Value
                    .TextMatrix(.Rows - 1, 8) = TB.Fields("Freight").Value & ""
                    
                End With
                TB.MoveNext
            Loop
Else
    MsgBox "No Voucher With this No. in This Type..."
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
flex1
Combs
lblJobType.Caption = ""
lblJobInfo.Caption = ""
lblrecWT.Caption = ""
Label11.Caption = vbNullString
Label12.Caption = vbNullString
End Sub

Private Sub transfer1()
grid1.Rows = grid1.Rows + 1
With grid1
    .TextMatrix(.Rows - 1, 0) = .Rows - 1
    .TextMatrix(.Rows - 1, 1) = Val(Text3.Text)
    .TextMatrix(.Rows - 1, 2) = Text4.Text
    .TextMatrix(.Rows - 1, 3) = Format(Text5.Text, "#.00")
    .TextMatrix(.Rows - 1, 4) = Format(Text9.Text, "#.00")
    .TextMatrix(.Rows - 1, 5) = Format(Val(Text8.Text), "#.00")
    .TextMatrix(.Rows - 1, 6) = Text12.Text
    .TextMatrix(.Rows - 1, 7) = Text15.Text
    .TextMatrix(.Rows - 1, 8) = Text16.Text
    
End With
End Sub
Private Sub flex1()
grid1.Rows = 1
grid1.Cols = 9
With grid1
    .ColWidth(0) = 500
    .TextMatrix(0, 0) = "Sr#"
    .ColWidth(1) = 1200
    .TextMatrix(0, 1) = "Item Code"
    .ColWidth(2) = 2000
    .TextMatrix(0, 2) = "Item Name"
    .ColWidth(3) = 1000
    .TextMatrix(0, 3) = "Rate"
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "Bales"
    .ColWidth(5) = 1000
    .TextMatrix(0, 5) = "Weight"
    .ColWidth(6) = 1200
    .TextMatrix(0, 6) = "Net Amount"
    .ColWidth(7) = 3000
    .TextMatrix(0, 7) = "Remarks"
    .ColWidth(8) = 0
    .TextMatrix(0, 8) = "Freight"
End With
End Sub
Private Sub Combs()

End Sub
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(v_no)as c from voumst where v_type = 4"
    
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

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus

End Sub

Private Sub clear1()
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString


Text8.Text = vbNullString
Text9.Text = vbNullString
Text16.Text = vbNullString
'Text15.Text = ""
End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
Text2.Text = Combo2.ItemData(Combo2.ListIndex)
Text10.Text = Combo2.Text
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus

End Sub
Private Sub Combo3_Click()
If Combo3.ListCount > 0 Then
Text13.Text = Combo3.ItemData(Combo3.ListIndex)
Text14.Text = Combo3.Text
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text13.SetFocus

End Sub

Private Sub Combo1_Click()
If Combo1.ListCount > 0 Then
Text3.Text = Combo1.ItemData(Combo1.ListIndex)
Text4.Text = Combo1.Text
End If
End Sub

Private Sub Combo4_Click()
If Combo4.ListCount > 0 Then
Text6.Text = Combo4.ItemData(Combo4.ListIndex)
Text7.Text = Combo4.Text
End If

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text6.SetFocus
End Sub

Private Sub Command1_Click()
If grid1.Rows > 1 And Val(Text2.Text) > 0 And Val(Text1.Text) > 0 Then
        Call save
        Command2_Click
Else
    MsgBox "Please Complete the Voucher"
End If
End Sub

Private Sub Command2_Click()
Call clearfull
Check1.Value = 0
date1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Command4_Click()
If Val(Text3.Text) > 0 Then

Call transfer1
Call clear1
Text3.SetFocus

End If
End Sub

Private Function GetRemarks(VNo As Double) As String
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Ssql = "Select * from VouDTL where V_Type=4 and V_No=" & VNo
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    GetRemarks = TB.Fields("Remarks").Value & ""
End If
TB.Close
DB.Close
End Function

Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Itm As String, Qty As Double, Comm As String, NetRate As Double
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
    Ssql = Ssql & " v_type = 4"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where "
    Ssql = Ssql & " v_type = 4"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    
    Ssql = "delete from Purchase where "
    Ssql = Ssql & " v_type = 4"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    
    DB.Close
End If
Set DB = OpenDatabase(Blm.patHmain)

Set TB = DB.OpenRecordset("Purchase", dbOpenTable)
For I% = 1 To grid1.Rows - 1
TB.AddNew
            TB.Fields("GSTInv").Value = Check1.Value
            TB.Fields("V_no").Value = Val(Text1.Text)
            TB.Fields("V_Type").Value = 4
            TB.Fields("V_Date").Value = date1.Value
            TB.Fields("Seller").Value = Val(Text2.Text)
            TB.Fields("DebitCode").Value = Val(Text13.Text)
            TB.Fields("Inv_no").Value = Text11.Text
            TB.Fields("Lorryno").Value = Text6.Text
            TB.Fields("Jobno").Value = Val(Text17.Text)
    With grid1
            
            TB.Fields("Item").Value = Val(.TextMatrix(I%, 1))
            TB.Fields("Rate").Value = Val(.TextMatrix(I%, 3))
            TB.Fields("Bales").Value = Val(.TextMatrix(I%, 4))
            TB.Fields("QTY").Value = Val(.TextMatrix(I%, 5))
            TB.Fields("Amount").Value = Val(.TextMatrix(I%, 6))
            TB.Fields("Remarks").Value = .TextMatrix(I%, 7)
            TB.Fields("Freight").Value = Val(.TextMatrix(I%, 8))
            
    End With
TB.Update
Next I%
TB.Close
Set TB = DB.OpenRecordset("voumst", dbOpenTable)

TB.AddNew
    TB.Fields("v_date").Value = date1.Value
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = 4
    TB.Fields("narration").Value = "Purchase JV # " & Text1.Text
    TB.Fields("RefNo").Value = Val(Text17.Text)
TB.Update
TB.Close
Set TB = DB.OpenRecordset("VouDtl", dbOpenTable)
For I% = 1 To grid1.Rows - 1
TB.AddNew
    TB.Fields("v_date").Value = date1.Value
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = 4
    TB.Fields("party").Value = Val(Text2.Text)
    TB.Fields("remarks").Value = "( " & grid1.TextMatrix(I%, 2) & " B:" & grid1.TextMatrix(I%, 4) & " W:" & grid1.TextMatrix(I%, 5) & "@" & grid1.TextMatrix(I%, 3) & " ) " & grid1.TextMatrix(I%, 7) & " V#" & Text6.Text
    TB.Fields("debit").Value = 0
    TB.Fields("credit").Value = Val(grid1.TextMatrix(I%, 6))
    
TB.Update
TB.AddNew
    TB.Fields("v_date").Value = date1.Value
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = 4
    TB.Fields("party").Value = Val(grid1.TextMatrix(I%, 1))
    TB.Fields("remarks").Value = "( " & grid1.TextMatrix(I%, 2) & " B:" & grid1.TextMatrix(I%, 4) & " W:" & grid1.TextMatrix(I%, 5) & "@" & grid1.TextMatrix(I%, 3) & " ) " & grid1.TextMatrix(I%, 7) & " V#" & Text6.Text
    TB.Fields("debit").Value = Val(grid1.TextMatrix(I%, 6))
    TB.Fields("credit").Value = 0
    
TB.Update
Next I%
TB.Close
DB.Close
End Sub

Private Sub Command5_Click()
Call clear1
Text3.SetFocus
End Sub

Private Sub Command6_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
If Result = vbNo Then Exit Sub

Dim DB As Database
Dim Ssql As String
If Option2 = True Then
Set DB = OpenDatabase(Blm.patHmain)
Ssql = "delete from voumst where "
    Ssql = Ssql & " v_type = 4"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where "
    Ssql = Ssql & " v_type = 4"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
DB.Execute Ssql
    Ssql = "delete from Purchase where "
    Ssql = Ssql & " v_type = 4"
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
DB.Execute Ssql
DB.Close
Command2_Click
End If
End Sub

Private Sub Command7_Click()
Text1.Text = max1
save
Command2_Click
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub date1_LostFocus()
If Option1 = True Then
    If date1.Value >= FStartDate And date1.Value <= FEndDate Then
        Text1.Text = max1
    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If
End If
End Sub

Private Sub Form_Load()
date1.Value = Date
Combs

Call flex1
If Screen.Width < 800 And Screen.Height < 600 Then
MsgBox "Please Set your Desktop 800 x 600 Then Try"
Me.Hide
Unload Me
End If

    Text1.Text = max1

End Sub

Private Sub grid1_Click()
If grid1.Row > 0 Then
    Text5.Text = grid1.TextMatrix(grid1.Row, 2)
End If
End Sub

Private Sub grid1_DblClick()
    With grid1
    
    Text3.Text = .TextMatrix(.Row, 1)
    Text4.Text = .TextMatrix(.Row, 2)
    Text5.Text = .TextMatrix(.Row, 3)
    Text9.Text = .TextMatrix(.Row, 4)
    Text8.Text = .TextMatrix(.Row, 5)
    Text12.Text = .TextMatrix(.Row, 6)
    Text15.Text = .TextMatrix(.Row, 7)
    Text16.Text = .TextMatrix(.Row, 8)
    End With
If grid1.Rows = 2 Then
    grid1.Rows = 1
Else
    grid1.RemoveItem (grid1.Row)
End If
End Sub

Private Sub lblTo_Click()

End Sub

Private Sub Option1_Click()
Text1.Enabled = False
date1.SetFocus
Command7.Visible = False
Command6.Visible = False
End Sub

Private Sub Option2_Click()

Command6.Visible = True
Command7.Visible = True
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text1.Text) <> 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim B As Boolean
If Val(Text1.Text) > 0 Then
    B = edit1
        If B = True Then
            Cancel = True
            Text1.Text = vbNullString
        End If
End If

End Sub

Private Sub Text11_GotFocus()
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Text13_GotFocus()
Text13.SelStart = 0
Text13.SelLength = Len(Text13.Text)

End Sub

Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text13.Text = SelectedAccountCode
    Text14.Text = SelectedAccountName
End If


End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text13.Text) <> 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text13_Validate(Cancel As Boolean)
If Val(Text13.Text) <> 0 Then
    Text14.Text = Blm.party1(Val(Text13.Text))
    If Text14.Text = "NOT" Then
        Cancel = True
    Else
        Label27.Caption = Blm.CurrentBalance(Val(Text13.Text))
    End If
        
End If

End Sub

Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15.Text)

End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys ("{TAB}")
End If

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
If Val(Text17.Text) > 0 Then
    Cancel = VerifyJob(Val(Text17.Text))
End If
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)


End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text2.Text = SelectedAccountCode
    Text10.Text = SelectedAccountName
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text1.Text) <> 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) <> 0 Then
    Text10.Text = Blm.party1(Val(Text2.Text))
    If Text10.Text = "NOT" Then
        Cancel = True
    Else
        Label19.Caption = Blm.SalesTaxNo(Val(Text2.Text))
        Label27.Caption = Blm.CurrentBalance(Val(Text2.Text))
    End If
        
End If
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    'Combo1.SetFocus
    Load Search1
    Search1.Show vbModal
    Text3.Text = SelectedItemCode
    Text4.Text = SelectedItemName
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Val(Text3.Text) <> 0 Then
    Text4.Text = Blm.item1(Val(Text3.Text))
    If Text4.Text = "NOT" Then
        Cancel = True
    Else
        lblStock.Caption = Blm.ClosingStock(Val(Text3.Text))
        lblBaleStock.Caption = Blm.ClosingStockBales(Val(Text3.Text))
        lblStockAmt.Caption = Blm.CurrentBalance(Val(Text3.Text))
        ShowStockAccount Val(Text3.Text)
    End If
        
End If

End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text8_GotFocus()
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim I As Long
Dim deb As Currency
Dim cred As Currency
Dim B As Double
Dim W As Double
Text12.Text = Val(Text5.Text) * Val(Text8.Text)
If grid1.Rows > 1 Then
    For I = 1 To grid1.Rows - 1
        deb = deb + Val(grid1.TextMatrix(I, 6))
        B = B + Val(grid1.TextMatrix(I, 4))
        W = W + Val(grid1.TextMatrix(I, 5))
        cred = cred + Val(grid1.TextMatrix(I, 8))
    Next I
    Label11.Caption = deb
    Label12.Caption = Format(cred, "#.00")
    Label8.Caption = B
    Label9.Caption = Format(W, "#.000")
End If
End Sub
