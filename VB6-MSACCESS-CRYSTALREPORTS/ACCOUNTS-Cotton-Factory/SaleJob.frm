VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SaleJob 
   Caption         =   "Sale Job Definition"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "SaleJob.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Caption         =   "Other Terms"
      Height          =   1155
      Left            =   2790
      TabIndex        =   47
      Top             =   1200
      Width           =   6165
      Begin VB.TextBox Text9 
         Height          =   915
         Left            =   60
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   180
         Width           =   6015
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Terms Type"
      Height          =   1125
      Left            =   210
      TabIndex        =   43
      Top             =   1200
      Width           =   2535
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "SaleJob.frx":27A2
         Left            =   90
         List            =   "SaleJob.frx":27B5
         TabIndex        =   4
         Text            =   "Combo3"
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      Height          =   1125
      Left            =   6630
      TabIndex        =   34
      Top             =   60
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   810
         Left            =   120
         Picture         =   "SaleJob.frx":27ED
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   810
         Left            =   1080
         Picture         =   "SaleJob.frx":4F8F
         Style           =   1  'Graphical
         TabIndex        =   35
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
      Height          =   1125
      Left            =   9000
      TabIndex        =   29
      Top             =   1200
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   810
         Left            =   1800
         Picture         =   "SaleJob.frx":7731
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   810
         Left            =   960
         Picture         =   "SaleJob.frx":9ED3
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   810
         Left            =   120
         Picture         =   "SaleJob.frx":C675
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   28
      Top             =   3960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   240
      TabIndex        =   24
      Top             =   2325
      Width           =   11535
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   6300
         TabIndex        =   10
         Top             =   630
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55967747
         CurrentDate     =   39297
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5040
         TabIndex        =   9
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55967747
         CurrentDate     =   39297
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4230
         TabIndex        =   8
         Top             =   630
         Width           =   795
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
         Left            =   10035
         TabIndex        =   14
         Top             =   630
         Width           =   1245
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   8460
         TabIndex        =   12
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   240
         MaxLength       =   255
         TabIndex        =   15
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9060
         TabIndex        =   13
         Top             =   630
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   8880
         Picture         =   "SaleJob.frx":EE17
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   10050
         Picture         =   "SaleJob.frx":115B9
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   7800
         MaxLength       =   50
         TabIndex        =   11
         Top             =   630
         Width           =   615
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
         Left            =   1560
         TabIndex        =   7
         Top             =   630
         Width           =   2715
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   285
         Left            =   6390
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "From"
         Height          =   255
         Left            =   5130
         TabIndex        =   45
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label9 
         Caption         =   "Consignments"
         Height          =   255
         Left            =   4020
         TabIndex        =   44
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Bales"
         Height          =   240
         Left            =   8430
         TabIndex        =   42
         Top             =   255
         Width           =   555
      End
      Begin VB.Label Label25 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Weight (KGS)"
         Height          =   255
         Left            =   9060
         TabIndex        =   39
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   10080
         TabIndex        =   32
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Rate"
         Height          =   255
         Left            =   7800
         TabIndex        =   27
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   255
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   240
      TabIndex        =   21
      Top             =   0
      Width           =   6375
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
         TabIndex        =   3
         Top             =   750
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   750
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5520
         TabIndex        =   31
         Top             =   330
         Visible         =   0   'False
         Width           =   735
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
         Format          =   55967747
         CurrentDate     =   36757
      End
      Begin VB.Label Label4 
         Caption         =   "Party Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Party Code"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Job #"
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   40
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "When in Seller Code Press (F1) to Select Accounts from List && (F1) in Item Code to Select Item From List"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   7200
      Width           =   7935
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10560
      TabIndex        =   30
      Top             =   7200
      Width           =   1215
   End
End
Attribute VB_Name = "SaleJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Private Function edit1() As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Tb2 As Recordset
Dim p As Long
Set DB = OpenDatabase(Blm.patHmain)
Ssql = "select * from SaleJob where"
Ssql = Ssql & " v_no = " & Val(Text1.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    date1.Value = TB.Fields("v_date").Value
    Text2.Text = TB.Fields("Buyer").Value
    Text10.Text = Blm.party1(TB.Fields("Buyer").Value)
    For R = 0 To Combo3.ListCount - 1
        If Combo3.ItemData(R) = TB.Fields("CTerms").Value Then
            Combo3.ListIndex = R
            Exit For
        End If
    Next R
        Text9.Text = TB.Fields("OtherTerms").Value & ""
    
            grid1.Rows = 1
            Do While Not TB.EOF
                grid1.Rows = grid1.Rows + 1
                With grid1
                    .TextMatrix(.Rows - 1, 0) = .Rows - 1
                    .TextMatrix(.Rows - 1, 1) = TB.Fields("Item").Value
                    .TextMatrix(.Rows - 1, 2) = Blm.item1(TB.Fields("item").Value)
                    .TextMatrix(.Rows - 1, 3) = TB.Fields("Consignments").Value
                    If Not IsNull(TB.Fields("SDate").Value) Then .TextMatrix(.Rows - 1, 4) = Format(TB.Fields("SDate").Value, "dd-MMM-yyyy")
                    If Not IsNull(TB.Fields("EDate").Value) Then .TextMatrix(.Rows - 1, 5) = Format(TB.Fields("EDate").Value, "dd-MMM-yyyy")
                    .TextMatrix(.Rows - 1, 6) = Format(TB.Fields("Rate").Value, "#.00")
                    .TextMatrix(.Rows - 1, 7) = Format(TB.Fields("STRate").Value, "#.00")
                    .TextMatrix(.Rows - 1, 8) = Format(TB.Fields("Qty").Value, "#.00")
                    .TextMatrix(.Rows - 1, 9) = Format(TB.Fields("Rate").Value + TB.Fields("STRate").Value, "#.00")
                    .TextMatrix(.Rows - 1, 10) = TB.Fields("Remarks").Value
                    
                    
                End With
                TB.MoveNext
            Loop
Else
    MsgBox "No Purchase Job With this No. in This Type..."
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
Label11.Caption = vbNullString
Label12.Caption = vbNullString
End Sub

Private Sub transfer1()
grid1.Rows = grid1.Rows + 1
With grid1
    .TextMatrix(.Rows - 1, 0) = .Rows - 1
    .TextMatrix(.Rows - 1, 1) = Val(Text3.Text)
    .TextMatrix(.Rows - 1, 2) = Text4.Text
    .TextMatrix(.Rows - 1, 3) = Text7.Text
    If Combo3.ListIndex = 2 Then
        .TextMatrix(.Rows - 1, 4) = Format(DTPicker1.Value, "dd-MMM-yyyy")
        .TextMatrix(.Rows - 1, 5) = Format(DTPicker2.Value, "dd-MMM-yyyy")
    End If
    .TextMatrix(.Rows - 1, 6) = Format(Text5.Text, "#.00")
    .TextMatrix(.Rows - 1, 7) = Format(Val(Text6.Text), "#.00")
    .TextMatrix(.Rows - 1, 8) = Format(Text8.Text, "#.00")
    .TextMatrix(.Rows - 1, 9) = Format(Val(Text12.Text), "#.00")
    .TextMatrix(.Rows - 1, 10) = Text15.Text
    
End With
End Sub
Private Sub flex1()
grid1.Rows = 1
grid1.Cols = 11
With grid1
    .ColWidth(0) = 500
    .TextMatrix(0, 0) = "Sr#"
    .ColWidth(1) = 1200
    .TextMatrix(0, 1) = "Item Code"
    .ColWidth(2) = 1500
    .TextMatrix(0, 2) = "Item Name"
    .ColWidth(3) = 700
    .TextMatrix(0, 3) = "Consignments"
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "From"
    .ColWidth(5) = 1000
    .TextMatrix(0, 5) = "To"
    .ColWidth(6) = 1000
    .TextMatrix(0, 6) = "Rate"
    .ColWidth(7) = 1000
    .TextMatrix(0, 7) = "Bales"
    .ColWidth(8) = 1000
    .TextMatrix(0, 8) = "Weight"
    .ColWidth(9) = 1200
    .TextMatrix(0, 9) = "Net Amount"
    .ColWidth(10) = 1800
    .TextMatrix(0, 10) = "Remarks"
End With
End Sub
Private Sub Combs()


End Sub
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(v_no)as c from SaleJob"
    
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
Text6.Text = vbNullString
Text8.Text = vbNullString
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
Private Sub Combo1_Click()
If Combo1.ListCount > 0 Then
Text3.Text = Combo1.ItemData(Combo1.ListIndex)
Text4.Text = Combo1.Text
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")

End Sub

Private Sub Command1_Click()
If grid1.Rows > 1 And Val(Text1.Text) > 0 And Val(Text2.Text) > 0 Then
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

Private Sub Command4_Click()
If Val(Text3.Text) > 0 Then
Call transfer1
Call clear1
Text3.SetFocus
End If
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
    Ssql = "delete from SaleJob where "
    Ssql = Ssql & " v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    
    DB.Close
End If
Set DB = OpenDatabase(Blm.patHmain)

Set TB = DB.OpenRecordset("SaleJob", dbOpenTable)
For I% = 1 To grid1.Rows - 1
TB.AddNew
            TB.Fields("V_no").Value = Val(Text1.Text)
            TB.Fields("V_Date").Value = date1.Value
            TB.Fields("Buyer").Value = Val(Text2.Text)
            TB.Fields("CTerms").Value = Combo3.ItemData(Combo3.ListIndex)
            TB.Fields("OtherTerms").Value = Text9.Text
    With grid1
            TB.Fields("Item").Value = Val(.TextMatrix(I%, 1))
            TB.Fields("Consignments").Value = Val(.TextMatrix(I%, 3))
            If Combo3.ListIndex = 2 Then
                TB.Fields("SDate").Value = CDate(.TextMatrix(I%, 4))
                TB.Fields("EDate").Value = CDate(.TextMatrix(I%, 5))
            End If
            TB.Fields("Rate").Value = Val(.TextMatrix(I%, 6))
            TB.Fields("STRate").Value = Val(.TextMatrix(I%, 7)) 'Bales
            TB.Fields("QTY").Value = Val(.TextMatrix(I%, 8))
            TB.Fields("Remarks").Value = .TextMatrix(I%, 10)
    End With
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
    Ssql = "delete from SaleJob where "
    Ssql = Ssql & " v_no = " & Val(Text1.Text)
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
     If date1.Value >= FStartDate And date1.Value <= FEndDate Then
        
    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If
    Text1.Text = max1
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
date1.Value = Date
DTPicker1.Value = Date
DTPicker2.Value = Date + 30
Combs
Combo3.ListIndex = 0
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
        Text7.Text = .TextMatrix(.Row, 3)
        If Len(.TextMatrix(.Row, 4)) > 0 Then DTPicker1.Value = CDate(.TextMatrix(.Row, 4))
        If Len(.TextMatrix(.Row, 5)) > 0 Then DTPicker2.Value = CDate(.TextMatrix(.Row, 5))
        Text5.Text = .TextMatrix(.Row, 6)
        Text6.Text = .TextMatrix(.Row, 7)
        Text8.Text = .TextMatrix(.Row, 8)
        Text12.Text = .TextMatrix(.Row, 9)
        Text15.Text = .TextMatrix(.Row, 10)
    End With
    grid1.RemoveItem (grid1.Row)
Else
If grid1.Rows = 2 Then
    With grid1
        Text3.Text = .TextMatrix(.Row, 1)
        Text4.Text = .TextMatrix(.Row, 2)
        Text7.Text = .TextMatrix(.Row, 3)
        If Len(.TextMatrix(.Row, 4)) > 0 Then DTPicker1.Value = CDate(.TextMatrix(.Row, 4))
        If Len(.TextMatrix(.Row, 5)) > 0 Then DTPicker2.Value = CDate(.TextMatrix(.Row, 5))
        Text5.Text = .TextMatrix(.Row, 6)
        Text6.Text = .TextMatrix(.Row, 7)
        Text8.Text = .TextMatrix(.Row, 8)
        Text12.Text = .TextMatrix(.Row, 9)
        Text15.Text = .TextMatrix(.Row, 10)
    End With
    grid1.Rows = 1
End If
End If
End Sub

Private Sub lblStock_Click()

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
    End If
        
End If
End Sub

Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15.Text)
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If

End Sub

Private Sub Text16_GotFocus()
Text16.SelStart = 0
Text16.SelLength = Len(Text16.Text)

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
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
    End If
        
End If
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
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

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    End If

End Sub

Private Sub Timer1_Timer()
Dim I As Long
Dim deb As Currency
Dim cred As Currency
Text12.Text = Val(Text5.Text) * Val(Text8.Text)
'Text7.Text = Val(Text6.Text) * Val(Text9.Text) / 100
'Text12.Text = Val(Text6.Text) - Val(Text7.Text)
If grid1.Rows > 1 Then
    For I = 1 To grid1.Rows - 1
        deb = deb + Val(grid1.TextMatrix(I, 5))
        cred = cred + Val(grid1.TextMatrix(I, 6))
    Next I
    Label11.Caption = Format(deb, "#.00")
    Label12.Caption = Format(cred, "#.00")
End If
End Sub
