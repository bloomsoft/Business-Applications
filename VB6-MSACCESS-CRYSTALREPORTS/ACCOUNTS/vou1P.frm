VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form vou1P 
   Caption         =   "Purchase Voucher Entry"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "vou1P.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      Height          =   975
      Left            =   6720
      TabIndex        =   43
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   615
         Left            =   120
         Picture         =   "vou1P.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   615
         Left            =   1080
         Picture         =   "vou1P.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2190
      Top             =   4680
   End
   Begin VB.Frame Frame4 
      Caption         =   "Actions"
      Height          =   975
      Left            =   9000
      TabIndex        =   37
      Top             =   1320
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   615
         Left            =   1800
         Picture         =   "vou1P.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   615
         Left            =   960
         Picture         =   "vou1P.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   615
         Left            =   120
         Picture         =   "vou1P.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   34
      Top             =   3960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   6720
      TabIndex        =   28
      Top             =   0
      Width           =   5055
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1200
         TabIndex        =   58
         Text            =   "Combo3"
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   53
         Text            =   "Combo1"
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   36
         Text            =   "Combo2"
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label24 
         Caption         =   "Debit List"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Items List"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "A/c List"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction"
      Height          =   1575
      Left            =   270
      TabIndex        =   27
      Top             =   2280
      Width           =   11535
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   5070
         TabIndex        =   18
         Top             =   1185
         Width           =   975
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   4080
         MaxLength       =   25
         TabIndex        =   9
         Top             =   585
         Width           =   960
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   5070
         TabIndex        =   10
         Top             =   585
         Width           =   975
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   240
         MaxLength       =   255
         TabIndex        =   17
         Top             =   1200
         Width           =   4815
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
         Left            =   9720
         TabIndex        =   16
         Top             =   570
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         TabIndex        =   12
         Top             =   570
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8160
         TabIndex        =   14
         Top             =   570
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   8880
         Picture         =   "vou1P.frx":EDCC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   10080
         Picture         =   "vou1P.frx":1156E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
         Left            =   7440
         TabIndex        =   13
         Top             =   570
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   9000
         TabIndex        =   15
         Top             =   570
         Width           =   705
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   6060
         MaxLength       =   50
         TabIndex        =   11
         Top             =   570
         Width           =   795
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
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   2625
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Discount"
         Height          =   285
         Left            =   5160
         TabIndex        =   66
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Label28 
         Caption         =   "Size (n1*n2)"
         Height          =   255
         Left            =   4110
         TabIndex        =   65
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label30 
         Caption         =   "Pcs"
         Height          =   285
         Left            =   5280
         TabIndex        =   64
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label25 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   8160
         TabIndex        =   51
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label15 
         Caption         =   "Comm."
         Height          =   255
         Left            =   6840
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   9870
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Comm %"
         Height          =   255
         Left            =   7440
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Net Rate"
         Height          =   255
         Left            =   9000
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Rate"
         Height          =   255
         Left            =   6150
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1530
         TabIndex        =   30
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voucher Information"
      Height          =   2295
      Left            =   240
      TabIndex        =   24
      Top             =   0
      Width           =   6375
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
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   5520
         TabIndex        =   2
         Top             =   360
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
         TabIndex        =   4
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
      Begin VB.CommandButton Command6 
         Caption         =   "&Delete "
         Height          =   375
         Left            =   5400
         TabIndex        =   39
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59506691
         CurrentDate     =   36757
      End
      Begin Crystal.CrystalReport r1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   3
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label27 
         Caption         =   "...."
         Height          =   255
         Left            =   3600
         TabIndex        =   63
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Party Balance :"
         Height          =   255
         Left            =   2400
         TabIndex        =   62
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Stock :"
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblStock 
         Height          =   255
         Left            =   1080
         TabIndex        =   59
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Debit Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   56
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Debit Code"
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "==="
         Height          =   255
         Left            =   1320
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "S.T. Reg. #"
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Invoice No."
         Height          =   255
         Left            =   4680
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Seller Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Seller Code"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher #"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   54
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "When in Seller Code Press (F1) to Select Accounts from List && (F1) in Item Code to Select Item From List"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   7200
      Width           =   7935
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10560
      TabIndex        =   38
      Top             =   7200
      Width           =   1215
   End
End
Attribute VB_Name = "vou1P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private blmr As New bloom_r
Private Sub printReport()
Dim Result As VbMsgBoxResult
Result = MsgBox("Want to Print This Voucher", vbYesNo)
If Result = vbYes Then

r1.ReportFileName = blmr.report_path & "pjv.rpt"
r1.ReportTitle = blm.orgname
f = "{Purchaseview.v_no} = " & Val(Text1.Text)

r1.DataFiles(0) = blm.patHmain
'r1.SelectionFormula = f
r1.WindowTop = 0
r1.WindowLeft = 0
r1.WindowState = crptMaximized
r1.printReport

End If
End Sub

Private Function checkdate(v_date As Date) As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.patHmain)

Ssql = "Select * from voudtl where v_date > #" & v_date & "#"
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    MsgBox "You Cannot Enter or Update the Voucher in Back Dates...."
    checkdate = True
Else
    checkdate = False
End If
tb.Close
DB.Close
End Function
Private Function edit1() As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Dim tb2 As Recordset
Dim p As Long
Set DB = OpenDatabase(blm.patHmain)
Ssql = "select * from Purchase where v_type = 4"
Ssql = Ssql & " and v_no = " & Val(Text1.Text)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("v_date").Value
    Text2.Text = tb.Fields("Seller").Value
    Text10.Text = blm.party1(tb.Fields("Seller").Value)
    Text13.Text = tb.Fields("DebitCode").Value
    Text14.Text = blm.party1(tb.Fields("DebitCode").Value)
    Text11.Text = tb.Fields("Inv_no").Value & ""
            grid1.Rows = 1
            Do While Not tb.EOF
                grid1.Rows = grid1.Rows + 1
                With grid1
                    .TextMatrix(.Rows - 1, 0) = .Rows - 1
                    .TextMatrix(.Rows - 1, 1) = tb.Fields("Item").Value
                    .TextMatrix(.Rows - 1, 2) = blm.item1(tb.Fields("item").Value)
                    .TextMatrix(.Rows - 1, 3) = Format(tb.Fields("Rate").Value, "#.00")
                    .TextMatrix(.Rows - 1, 4) = Format(tb.Fields("Qty").Value, "#.00")
                    .TextMatrix(.Rows - 1, 5) = Format(tb.Fields("STRate").Value, "#.000")
                    .TextMatrix(.Rows - 1, 6) = Format(tb.Fields("Rate").Value + tb.Fields("STRate").Value, "#.00")
                    .TextMatrix(.Rows - 1, 7) = Format(tb.Fields("CommRate").Value, "#.000")
                    .TextMatrix(.Rows - 1, 8) = Format((Val(.TextMatrix(.Rows - 1, 6)) * tb.Fields("Qty").Value) - tb.Fields("Discount").Value, "#.00")
                    .TextMatrix(.Rows - 1, 9) = tb.Fields("Remarks").Value & ""
                    .TextMatrix(.Rows - 1, 10) = tb.Fields("Size").Value & ""
                    .TextMatrix(.Rows - 1, 11) = tb.Fields("Pcs").Value & ""
                    .TextMatrix(.Rows - 1, 12) = tb.Fields("Discount").Value & ""
                    
                    
                End With
                tb.MoveNext
            Loop
Else
    MsgBox "No Voucher With this No. in This Type..."
    edit1 = False
End If
tb.Close
DB.Close
End Function
Private Sub clearfull()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
'    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
flex1
combs
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
    .TextMatrix(.Rows - 1, 4) = Format(Val(Text8.Text), "#.00")
    .TextMatrix(.Rows - 1, 5) = Format(Text9.Text, "#.000")
    .TextMatrix(.Rows - 1, 6) = Format((Val(Text5.Text) + Val(Text9.Text)), "#.00")
    .TextMatrix(.Rows - 1, 7) = Text7.Text
    .TextMatrix(.Rows - 1, 8) = Format(Val(Text6.Text) * Val(Text8.Text), "#.00")
    .TextMatrix(.Rows - 1, 9) = Text15.Text
    .TextMatrix(.Rows - 1, 10) = Text16.Text
    .TextMatrix(.Rows - 1, 11) = Text18.Text
    .TextMatrix(.Rows - 1, 12) = Text17.Text
    
End With
End Sub
Private Sub flex1()
grid1.Rows = 1
grid1.Cols = 13
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
    .TextMatrix(0, 4) = "Quantity"
    .ColWidth(5) = 800
    .TextMatrix(0, 5) = "Comm."
    .ColWidth(6) = 1000
    .TextMatrix(0, 6) = "Net Rate"
    .ColWidth(7) = 800
    .TextMatrix(0, 7) = "Comm%"
    .ColWidth(8) = 1200
    .TextMatrix(0, 8) = "Net Amount"
    .ColWidth(9) = 1800
    .TextMatrix(0, 9) = "Remarks"
    .ColWidth(10) = 1000
    .TextMatrix(0, 10) = "Size"
    .ColWidth(11) = 1000
    .TextMatrix(0, 11) = "Pcs"
    .ColWidth(12) = 1000
    .TextMatrix(0, 12) = "Discount"
End With
End Sub
Private Sub combs()
Dim Ssql As String
Ssql = "select * from acchart order by name"
blm.fill_comb Ssql, Combo2, "name", "code"
blm.fill_comb Ssql, Combo3, "name", "code"
Ssql = "select * from Items order by name"
blm.fill_comb Ssql, Combo1, "name", "code"

End Sub
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim tb As Recordset
    
    Ssql = "select max(v_no)as c from voumst where v_type = 4"
    
    Set DB = OpenDatabase(blm.patHmain)
    Set tb = DB.OpenRecordset(Ssql)
    If IsNull(tb.Fields("c").Value) = False Then
        max1 = tb.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    tb.Close
    DB.Close
End Function

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus

End Sub

Private Sub clear1()
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = vbNullString
Text9.Text = vbNullString
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

Private Sub Command1_Click()
If grid1.Rows > 1 Then
        Call save
        printReport
        Command2_Click
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

Private Sub Command4_Click()
If Val(Text3.Text) > 0 Then
If Val(Text6.Text) > 0 Or Val(Text7.Text) > 0 Then
Call transfer1
Call clear1
Text3.SetFocus
End If
End If
End Sub

Private Function GetRemarks(VNo As Double) As String
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Ssql = "Select * from VouDTL where V_Type=4 and V_No=" & VNo
Set DB = OpenDatabase(blm.patHmain)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    GetRemarks = tb.Fields("Remarks").Value & ""
End If
tb.Close
DB.Close
End Function

Private Sub save()
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Dim Itm As String, Qty As Double, Comm As String, NetRate As Double
Dim b As Boolean
If Option2 = True Then
    'b = checkdate(date1.Value)
    If b = True Then
        Exit Sub
    End If
End If
If Option2 = True Then
    Set DB = OpenDatabase(blm.patHmain)
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
Set DB = OpenDatabase(blm.patHmain)

Set tb = DB.OpenRecordset("Purchase", dbOpenTable)
For i% = 1 To grid1.Rows - 1
tb.AddNew
            tb.Fields("V_no").Value = Val(Text1.Text)
            tb.Fields("V_Type").Value = 4
            tb.Fields("V_Date").Value = date1.Value
            tb.Fields("Seller").Value = Val(Text2.Text)
            tb.Fields("DebitCode").Value = Val(Text13.Text)
            tb.Fields("Inv_no").Value = Text11.Text
        
    With grid1
            tb.Fields("Item").Value = Val(.TextMatrix(i%, 1))
            tb.Fields("Rate").Value = Val(.TextMatrix(i%, 3))
            tb.Fields("QTY").Value = Val(.TextMatrix(i%, 4))
            tb.Fields("STRate").Value = Val(.TextMatrix(i%, 5))
            tb.Fields("CommRate").Value = Val(.TextMatrix(i%, 7))
            tb.Fields("Remarks").Value = .TextMatrix(i%, 9)
            tb.Fields("Size").Value = .TextMatrix(i%, 10)
            tb.Fields("Pcs").Value = Val(.TextMatrix(i%, 11))
            tb.Fields("Discount").Value = Val(.TextMatrix(i%, 12))
            
    End With
tb.Update
Next i%
tb.Close
Set tb = DB.OpenRecordset("voumst", dbOpenTable)

tb.AddNew
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("v_no").Value = Val(Text1.Text)
    tb.Fields("v_type").Value = 4
    tb.Fields("narration").Value = "Purchase JV # " & Text1.Text
tb.Update
tb.Close
Set tb = DB.OpenRecordset("VouDtl", dbOpenTable)
For i% = 1 To grid1.Rows - 1
tb.AddNew
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("v_no").Value = Val(Text1.Text)
    tb.Fields("v_type").Value = 4
    tb.Fields("party").Value = Val(Text2.Text)
    tb.Fields("remarks").Value = "( " & grid1.TextMatrix(i%, 2) & " " & grid1.TextMatrix(i%, 4) & "@" & grid1.TextMatrix(i%, 3) & " ) " & grid1.TextMatrix(i%, 9)
    tb.Fields("debit").Value = 0
    tb.Fields("credit").Value = Val(grid1.TextMatrix(i%, 8))
    
tb.Update
tb.AddNew
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("v_no").Value = Val(Text1.Text)
    tb.Fields("v_type").Value = 4
    tb.Fields("party").Value = Text13.Text
    tb.Fields("remarks").Value = "( " & grid1.TextMatrix(i%, 2) & " " & grid1.TextMatrix(i%, 4) & "@" & grid1.TextMatrix(i%, 3) & " ) " & grid1.TextMatrix(i%, 9)
    tb.Fields("debit").Value = Val(grid1.TextMatrix(i%, 8))
    tb.Fields("credit").Value = 0
    
tb.Update
Next i%
tb.Close
DB.Close
End Sub

Private Sub Command5_Click()
Call clear1
Text3.SetFocus
End Sub

Private Sub Command6_Click()
Dim DB As Database
Dim Ssql As String
If Option2 = True Then
Set DB = OpenDatabase(blm.patHmain)
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

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub date1_LostFocus()
If Option1 = True Then
    Text1.Text = max1
End If
End Sub

Private Sub Form_Activate()
combs
End Sub

Private Sub Form_Load()
date1.Value = Date


Call flex1
'If Screen.Width = 800 And Screen.Height = 600 Then Me.WindowState = 2
'If Screen.Width > 800 And Screen.Height > 600 Then
'    Me.Height = 8085
'    Me.Width = 12060
'End If
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
If grid1.Rows > 2 Then
    With grid1
        Text3.Text = .TextMatrix(.Row, 1)
        Text4.Text = .TextMatrix(.Row, 2)
        Text5.Text = .TextMatrix(.Row, 3)
        Text8.Text = .TextMatrix(.Row, 4)
        Text9.Text = .TextMatrix(.Row, 5)
        Text6.Text = .TextMatrix(.Row, 6)
        Text7.Text = .TextMatrix(.Row, 7)
        Text12.Text = .TextMatrix(.Row, 8)
        Text15.Text = .TextMatrix(.Row, 9)
        Text16.Text = .TextMatrix(.Row, 10)
        Text18.Text = .TextMatrix(.Row, 11)
        Text17.Text = .TextMatrix(.Row, 12)
    End With
    grid1.RemoveItem (grid1.Row)
Else
If grid1.Rows = 2 Then
    With grid1
        Text3.Text = .TextMatrix(.Row, 1)
        Text4.Text = .TextMatrix(.Row, 2)
        Text5.Text = .TextMatrix(.Row, 3)
        Text8.Text = .TextMatrix(.Row, 4)
        Text9.Text = .TextMatrix(.Row, 5)
        Text6.Text = .TextMatrix(.Row, 6)
        Text7.Text = .TextMatrix(.Row, 7)
        Text12.Text = .TextMatrix(.Row, 8)
        Text15.Text = .TextMatrix(.Row, 9)
        Text16.Text = .TextMatrix(.Row, 10)
        Text18.Text = .TextMatrix(.Row, 11)
        Text17.Text = .TextMatrix(.Row, 12)
    End With
    grid1.Rows = 1
End If
End If
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
date1.SetFocus
Command6.Visible = False
End Sub

Private Sub Option2_Click()

Command6.Visible = True
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
Dim b As Boolean
If Val(Text1.Text) > 0 Then
    b = edit1
        If b = True Then
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
    Text14.Text = blm.party1(Val(Text13.Text))
    If Text14.Text = "NOT" Then
        Cancel = True
    Else
        Label27.Caption = blm.CurrentBalance(Val(Text13.Text))
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

Private Sub Text16_GotFocus()
Text16.SelStart = 0
Text16.SelLength = Len(Text16.Text)

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If

End Sub

Private Sub Text16_Validate(Cancel As Boolean)
If Len(Text16.Text) > 0 Then
    Text8.Text = (Eval(Text16.Text) * Val(Text18.Text))
End If

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 <> 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
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

Private Sub Text18_Validate(Cancel As Boolean)
If Len(Text18.Text) > 0 Then
    Text8.Text = (Eval(Text16.Text) * Val(Text18.Text))
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
    Text10.Text = blm.party1(Val(Text2.Text))
    If Text10.Text = "NOT" Then
        Cancel = True
    Else
        Label19.Caption = blm.SalesTaxNo(Val(Text2.Text))
    End If
        
End If
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo1.SetFocus
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
    Text4.Text = blm.item1(Val(Text3.Text))
    If Text4.Text = "NOT" Then
        Cancel = True
    Else
        lblStock.Caption = blm.ClosingStock(Val(Text3.Text))
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

Private Sub Text9_Change()
'Text7.Text = ""
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
Dim i As Long
Dim deb As Currency
Dim cred As Currency
If Val(Text7.Text) > 0 Then
    Text9.Text = (Val(Text5.Text) * Val(Text7.Text)) / 100
End If
Text6.Text = Val(Text5.Text) + Val(Text9.Text)
Text12.Text = Round((Val(Text6.Text) * Val(Text8.Text)))
'Text7.Text = Val(Text6.Text) * Val(Text9.Text) / 100
'Text12.Text = Val(Text6.Text) - Val(Text7.Text)
If grid1.Rows > 1 Then
    For i = 1 To grid1.Rows - 1
        deb = deb + Val(grid1.TextMatrix(i, 5))
        cred = cred + Val(grid1.TextMatrix(i, 8))
    Next i
    Label11.Caption = Format(deb, "#.00")
    Label12.Caption = Format(cred - Val(Text17.Text), "#.00")
End If
End Sub
