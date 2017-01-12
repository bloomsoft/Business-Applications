VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Issue 
   Caption         =   "Daily Issue Voucher Entry"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ISSUE.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   945
      Left            =   7110
      TabIndex        =   45
      Top             =   30
      Width           =   2565
      Begin VB.CommandButton Command8 
         Caption         =   "&Delete"
         Height          =   705
         Left            =   960
         TabIndex        =   47
         Top             =   150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Save as"
         Height          =   705
         Left            =   90
         TabIndex        =   46
         Top             =   150
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   945
      Left            =   5910
      TabIndex        =   42
      Top             =   30
      Width           =   1155
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   255
         Left            =   180
         TabIndex        =   44
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   240
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Voucher Info"
      Height          =   975
      Left            =   2760
      TabIndex        =   37
      Top             =   0
      Width           =   3045
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Ref#"
         Height          =   255
         Left            =   210
         TabIndex        =   41
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Number"
         Height          =   255
         Left            =   210
         TabIndex        =   38
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2145
      Top             =   4680
   End
   Begin VB.Frame Frame4 
      Caption         =   "Actions"
      Height          =   1020
      Left            =   9750
      TabIndex        =   24
      Top             =   -15
      Width           =   2040
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   690
         Left            =   1320
         Picture         =   "ISSUE.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   195
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   690
         Left            =   705
         Picture         =   "ISSUE.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   195
         Width           =   630
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   690
         Left            =   90
         Picture         =   "ISSUE.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   195
         Width           =   630
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4065
      Left            =   255
      TabIndex        =   23
      Top             =   2625
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7170
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction"
      Height          =   1575
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   11535
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   7980
         TabIndex        =   10
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9240
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3885
         TabIndex        =   6
         Top             =   600
         Width           =   630
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   6480
         MaxLength       =   255
         TabIndex        =   9
         Top             =   600
         Width           =   1470
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
         Left            =   5445
         TabIndex        =   8
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   600
         Width           =   915
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   8880
         Picture         =   "ISSUE.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   10080
         Picture         =   "ISSUE.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
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
         TabIndex        =   4
         Top             =   600
         Width           =   1965
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblStockAmt 
         Caption         =   "0.00"
         Height          =   255
         Left            =   5490
         TabIndex        =   48
         Top             =   990
         Width           =   1545
      End
      Begin VB.Label Label10 
         Caption         =   "Debit Code"
         Height          =   255
         Left            =   8400
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Debit Name"
         Height          =   255
         Left            =   9480
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   6480
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   5400
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblBaleStock 
         Caption         =   "0.00"
         Height          =   195
         Left            =   3930
         TabIndex        =   36
         Top             =   990
         Width           =   615
      End
      Begin VB.Label lblAvgBaleWT 
         Caption         =   "..."
         Height          =   255
         Left            =   6780
         TabIndex        =   33
         Top             =   990
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblStock 
         Caption         =   "0.00"
         Height          =   285
         Left            =   4620
         TabIndex        =   32
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "Stock :"
         Height          =   210
         Left            =   3210
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Bales"
         Height          =   240
         Left            =   3945
         TabIndex        =   30
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label20 
         Caption         =   "Weight (KGS)"
         Height          =   255
         Left            =   4530
         TabIndex        =   27
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Rate"
         Height          =   255
         Left            =   3075
         TabIndex        =   22
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1095
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   270
         TabIndex        =   20
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voucher Information"
      Height          =   975
      Left            =   240
      TabIndex        =   17
      Top             =   -15
      Width           =   2475
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   75
         TabIndex        =   0
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   72548355
         CurrentDate     =   36757
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   75
         TabIndex        =   18
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6660
      TabIndex        =   35
      Top             =   6750
      Width           =   960
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7650
      TabIndex        =   34
      Top             =   6750
      Width           =   960
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5670
      TabIndex        =   28
      Top             =   6750
      Width           =   960
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4410
      TabIndex        =   25
      Top             =   6750
      Width           =   1215
   End
End
Attribute VB_Name = "Issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Dim EditMode As Boolean
Private Function CurrentBalance(AcCode As Long) As String
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim S As String
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "select Sum(Debit - Credit) as Bal from Voudtl where Party = " & AcCode
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    If TB.Fields("Bal").Value > 0 Then
        S = Format(TB.Fields("Bal").Value, "#.00") & " DR"
    ElseIf TB.Fields("Bal").Value < 0 Then
        S = Format(TB.Fields("Bal").Value, "#.00") & " CR"
    End If
Else
    S = "...."
End If
TB.Close
DB.Close
CurrentBalance = S
End Function
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(vno)as c from Issue"
    
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

Private Sub saveAccount()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim opendate As Date
Dim R As Integer
Set DB = OpenDatabase(Blm.patHmain)

    Ssql = "delete from vouMST where v_type = 20 and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where v_type = 20 and v_no = " & Val(Text1.Text)
    DB.Execute Ssql

Dim Tb2 As Recordset
Set Tb2 = DB.OpenRecordset("voumst", dbOpenTable)
Tb2.AddNew
    Tb2.Fields("v_date").Value = date1.Value
    Tb2.Fields("v_type").Value = 20
    Tb2.Fields("v_no").Value = Val(Text1.Text)
    Tb2.Fields("narration").Value = "Issue Voucher"
    Tb2.Fields("RefNo").Value = Val(Text7.Text)
Tb2.Update
Tb2.Close


Set Tb2 = DB.OpenRecordset("voudtl", dbOpenTable)
With grid1
For R = 1 To .Rows - 1
    Tb2.AddNew
        Tb2.Fields("v_date").Value = date1.Value
        Tb2.Fields("v_type").Value = 20
        Tb2.Fields("v_no").Value = Val(Text1.Text)
        Tb2.Fields("party").Value = Val(.TextMatrix(R, 1))
        Tb2.Fields("debit").Value = 0
        Tb2.Fields("credit").Value = Val(.TextMatrix(R, 6))
        Tb2.Fields("Remarks").Value = "@ " & .TextMatrix(R, 3) & " + " & .TextMatrix(R, 4) & " + " & .TextMatrix(R, 5)
    Tb2.Update
    Tb2.AddNew
        Tb2.Fields("v_date").Value = date1.Value
        Tb2.Fields("v_type").Value = 20
        Tb2.Fields("v_no").Value = Val(Text1.Text)
        Tb2.Fields("party").Value = Val(.TextMatrix(R, 12))
        Tb2.Fields("debit").Value = Val(.TextMatrix(R, 6))
        Tb2.Fields("credit").Value = 0
        Tb2.Fields("Remarks").Value = "@ " & .TextMatrix(R, 3) & " + " & .TextMatrix(R, 4) & " + " & .TextMatrix(R, 5)
    Tb2.Update
Next R
End With
Tb2.Close

DB.Close

End Sub

Private Sub AdjustPercentage()
Dim R As Integer
Dim p As Single
For R = 1 To grid1.Rows - 1
'    MsgBox "Test"
    If Val(Label11.Caption) <> 0 Then
        p = (Val(grid1.TextMatrix(R, 5)) * 100) / Val(Label11.Caption)
    Else
        p = 100
    End If
    grid1.TextMatrix(R, 9) = Round(p, 2) & "%"
Next R
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
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Tb2 As Recordset
Dim p As Long
Set DB = OpenDatabase(Blm.patHmain)
Ssql = "select * from Issue where "
Ssql = Ssql & " VNo=" & Text1.Text
'Ssql = Ssql & " and RefNo=" & Text7.Text
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    EditMode = True
    date1.Value = TB.Fields("v_date").Value
    Text1.Text = TB.Fields("VNo").Value & ""
    Text7.Text = TB.Fields("RefNo").Value & ""
            grid1.Rows = 1
            Do While Not TB.EOF
                grid1.Rows = grid1.Rows + 1
                With grid1
                    .TextMatrix(.Rows - 1, 0) = .Rows - 1
                    .TextMatrix(.Rows - 1, 1) = TB.Fields("ItemCode").Value
                    .TextMatrix(.Rows - 1, 2) = Blm.item1(TB.Fields("itemCode").Value)
                    .TextMatrix(.Rows - 1, 3) = Format(TB.Fields("Rate").Value, "#.00")
                    .TextMatrix(.Rows - 1, 4) = TB.Fields("Bales").Value
                    .TextMatrix(.Rows - 1, 5) = Format(TB.Fields("Qty").Value, "#.00")
                    .TextMatrix(.Rows - 1, 6) = Format(TB.Fields("Amount").Value, "#.00")
                    .TextMatrix(.Rows - 1, 7) = TB.Fields("Remarks").Value
                    .TextMatrix(.Rows - 1, 8) = TB.Fields("AvgBaleWT").Value & ""
                    .TextMatrix(.Rows - 1, 9) = TB.Fields("Percent").Value & ""
                    .TextMatrix(.Rows - 1, 10) = TB.Fields("BaleStock").Value & ""
                    .TextMatrix(.Rows - 1, 11) = TB.Fields("QtyStock").Value & ""
                    .TextMatrix(.Rows - 1, 12) = TB.Fields("DrCode").Value & ""
                    .TextMatrix(.Rows - 1, 13) = Blm.party1(TB.Fields("DrCode").Value)
                    
                End With
                TB.MoveNext
            Loop
Else
    MsgBox "No Data For This Date..."
    edit1 = False
'    Text1.Text = max1
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
Text1.Text = max1
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
    .TextMatrix(.Rows - 1, 8) = Val(lblAvgBaleWT.Caption)
    .TextMatrix(.Rows - 1, 10) = Val(lblBaleStock.Caption)
    .TextMatrix(.Rows - 1, 11) = Val(lblStock.Caption)
    .TextMatrix(.Rows - 1, 12) = Text6.Text
    .TextMatrix(.Rows - 1, 13) = Text2.Text
    
End With

End Sub
Private Sub flex1()
grid1.Rows = 1
grid1.Cols = 14
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
    .ColWidth(7) = 1800
    .TextMatrix(0, 7) = "Remarks"
    .ColWidth(8) = 1200
    .TextMatrix(0, 8) = "Avg Bale Wt"
    .ColWidth(9) = 1200
    .TextMatrix(0, 9) = "Percent"
    .ColWidth(10) = 0 'bales
    .ColWidth(11) = 0 'KGS STock
    .ColWidth(12) = 1200
    .TextMatrix(0, 12) = "Debit Code"
    .ColWidth(13) = 1800
    .TextMatrix(0, 13) = "Debit Name"
End With
End Sub
Private Sub Combs()

End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub clear1()
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text8.Text = vbNullString
Text9.Text = vbNullString
'Text15.Text = ""
End Sub

Private Sub Combo1_Click()
If Combo1.ListCount > 0 Then
Text3.Text = Combo1.ItemData(Combo1.ListIndex)
Text4.Text = Combo1.Text
End If
End Sub

Private Sub Combo4_Click()
If Combo4.ListIndex > -1 Then
    Text6.Text = Combo4.ItemData(Combo4.ListIndex)
    Text2.Text = Combo4.Text
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text6.SetFocus
End Sub

Private Sub Command1_Click()
If grid1.Rows > 1 And Val(Text1.Text) > 0 And Val(Text7.Text) > 0 Then
        Call save
        Command2_Click
Else
    MsgBox "Please Complete This Voucher"
End If
End Sub

Private Sub Command2_Click()
Call clearfull
EditMode = False
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
End Function

Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Itm As String, Qty As Double, Comm As String, NetRate As Double
Dim B As Boolean
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from Issue where "
    Ssql = Ssql & "VNo=" & Text1.Text
    'Ssql = Ssql & " and RefNo=" & Text7.Text
    DB.Execute Ssql
    DB.Close
Set DB = OpenDatabase(Blm.patHmain)

Set TB = DB.OpenRecordset("Issue", dbOpenTable)
For I% = 1 To grid1.Rows - 1
TB.AddNew
            TB.Fields("V_Date").Value = date1.Value
            TB.Fields("VNo").Value = Val(Text1.Text)
            TB.Fields("RefNo").Value = Val(Text7.Text)
    With grid1
            TB.Fields("ItemCode").Value = Val(.TextMatrix(I%, 1))
            TB.Fields("Rate").Value = Val(.TextMatrix(I%, 3))
            TB.Fields("Bales").Value = Val(.TextMatrix(I%, 4))
            TB.Fields("QTY").Value = Val(.TextMatrix(I%, 5))
            TB.Fields("Amount").Value = Val(.TextMatrix(I%, 6))
            TB.Fields("Remarks").Value = .TextMatrix(I%, 7)
            TB.Fields("AvgBaleWT").Value = Val(.TextMatrix(I%, 8))
            TB.Fields("Percent").Value = .TextMatrix(I%, 9)
            TB.Fields("BaleStock") = Val(.TextMatrix(I%, 10))
            TB.Fields("QtyStock") = Val(.TextMatrix(I%, 11))
            TB.Fields("DrCode") = Val(.TextMatrix(I%, 12))
    End With
TB.Update
Next I%
TB.Close
DB.Close
DoEvents
saveAccount
End Sub

Private Sub Command5_Click()
Call clear1
Text3.SetFocus
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()
Text1.Text = max1
Call save
Command2_Click

End Sub

Private Sub Command8_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
If Result = vbNo Then Exit Sub

Dim DB As Database
Dim Ssql As String
If Option2 = True Then
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from Issue where"
    Ssql = Ssql & " VNo=" & Text1.Text
    'Ssql = Ssql & " and RefNo=" & Text7.Text
      DB.Execute Ssql
      
    Ssql = "delete from vouMST where v_type = 20 and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where v_type = 20 and v_no = " & Val(Text1.Text)
    DB.Execute Ssql


  
    DB.Close
    Command2_Click
End If
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

Private Sub Form_Activate()
Combs
End Sub

Private Sub Form_Load()
date1.Value = Date
Text1.Text = max1
Call flex1
If Screen.Width < 800 And Screen.Height < 600 Then
MsgBox "Please Set your Desktop 800 x 600 Then Try"
Me.Hide
Unload Me
End If

    

End Sub

Private Sub grid1_Click()
If grid1.Row > 0 Then
    Text5.Text = grid1.TextMatrix(grid1.Row, 2)
End If
End Sub

Private Sub grid1_DblClick()
    With grid1
    Text3.Text = .TextMatrix(grid1.Row, 1)
    Text4.Text = .TextMatrix(grid1.Row, 2)
    Text5.Text = .TextMatrix(grid1.Row, 3)
    Text9.Text = .TextMatrix(grid1.Row, 4)
    Text8.Text = .TextMatrix(grid1.Row, 5)
    Text12.Text = .TextMatrix(grid1.Row, 6)
    Text15.Text = .TextMatrix(grid1.Row, 7)
    lblAvgBaleWT.Caption = .TextMatrix(grid1.Row, 8) & ""
    lblBaleStock.Caption = .TextMatrix(grid1.Row, 10) & ""
    lblStock.Caption = .TextMatrix(grid1.Row, 11) & ""
    Text6.Text = .TextMatrix(grid1.Row, 12)
    Text2.Text = .TextMatrix(grid1.Row, 13)
    End With
If grid1.Rows = 2 Then
    grid1.Rows = 1
Else
    grid1.RemoveItem (grid1.Row)
End If
End Sub

Private Sub Label21_Click()

End Sub

Private Sub Option1_Click()
Text1.Enabled = False
date1.SetFocus
Command7.Visible = False
Command8.Visible = False
End Sub

Private Sub Option2_Click()
Command7.Visible = True
Command8.Visible = True
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
If Option2 = True Then
If Val(Text1.Text) > 0 Then
    edit1
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
If KeyCode = vbKeyF1 Then Combo3.SetFocus

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

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)


End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
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
Dim R As CottonIssue
If Val(Text3.Text) <> 0 Then
    Text4.Text = Blm.item1(Val(Text3.Text))
    If Text4.Text = "NOT" Then
        Cancel = True
    Else
        'If EditMode = False Then
        lblStock.Caption = Blm.ClosingStock(Val(Text3.Text))
        lblBaleStock.Caption = Blm.ClosingStockBales(Val(Text3.Text))
        
        'R = AvgRateAndWeight(date1.Value, Val(Text3.Text))
         'MsgBox "Test"
        lblStockAmt.Caption = CurrentBalance(Val(Text3.Text))
'        MsgBox Val(lblStockAmt.Caption)
        If Val(lblStock.Caption) <> 0 Then
            Text5.Text = Round(Val(lblStockAmt.Caption) / Val(lblStock.Caption), 2)
        Else
            Text5.Text = Round(Val(lblStockAmt.Caption), 0)
        End If
        'lblAvgBaleWT.Caption = Round(R.AvgBaleWeight, 3)
        If Val(lblBaleStock.Caption) <> 0 Then
            lblAvgBaleWT.Caption = Round(Val(lblStock.Caption) / Val(lblBaleStock.Caption), 2)
        Else
            lblAvgBaleWT.Caption = Round(Val(lblStock.Caption), 2)
        End If
        'End If
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

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text6.Text = SelectedAccountCode
    Text2.Text = SelectedAccountName
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
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

Private Sub Text6_Validate(Cancel As Boolean)
If Val(Text6.Text) <> 0 Then
    
    Text2.Text = Blm.party1(Val(Text6.Text))
    If Text2.Text = "NOT" Then
        Cancel = True
        
    End If
        
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

Private Sub Text8_Validate(Cancel As Boolean)
If (Val(lblStock.Caption) - Val(Text8.Text)) <= 0 Then
            MsgBox "This Item is Out of Stock"
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

Private Sub Text9_LostFocus()
Text8.Text = Val(Text9.Text) * Val(lblAvgBaleWT.Caption)
End Sub

Private Sub Timer1_Timer()
Dim I As Long
Dim TAmount As Double
Dim TBales As Double
Dim TQty As Double
If grid1.Rows > 1 Then
    For I = 1 To grid1.Rows - 1
        TBales = TBales + Val(grid1.TextMatrix(I, 6))
        TQty = TQty + Val(grid1.TextMatrix(I, 5))
        TAmount = TAmount + Val(grid1.TextMatrix(I, 4))
    Next I
    Label4.Caption = "100"
Else
    Label4.Caption = ""
End If

Label3.Caption = TBales
Label11.Caption = TQty
Label12.Caption = TAmount
'Text8.Text = Round(Val(Text9.Text) * Val(lblAvgBaleWT.Caption), 3)
Text12.Text = Round(Val(Text5.Text) * Val(Text8.Text), 2)
DoEvents
'MsgBox "Test"
AdjustPercentage
End Sub
