VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmManualBS 
   Caption         =   "Balance Sheet Statement"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   2145
      Left            =   4170
      TabIndex        =   29
      Top             =   6180
      Width           =   3375
      Begin VB.TextBox Text36 
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Top             =   840
         Width           =   555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   285
         Left            =   1380
         TabIndex        =   33
         Top             =   870
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Update"
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Top             =   900
         Width           =   1005
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Save"
         Height          =   435
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Reset"
         Height          =   435
         Left            =   810
         TabIndex        =   16
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Prev"
         Height          =   435
         Left            =   1380
         TabIndex        =   15
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Delete"
         Height          =   435
         Left            =   1950
         TabIndex        =   31
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command11 
         Caption         =   "E&xit"
         Height          =   435
         Left            =   2520
         TabIndex        =   30
         Top             =   1560
         Width           =   585
      End
      Begin VB.TextBox txtRefNo 
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   510
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60620803
         CurrentDate     =   39297
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   780
         TabIndex        =   0
         Top             =   180
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60620803
         CurrentDate     =   39297
      End
      Begin VB.Label Label17 
         Caption         =   "From"
         Height          =   285
         Left            =   210
         TabIndex        =   37
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   285
         Left            =   180
         TabIndex        =   36
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label19 
         Caption         =   "Sheet#"
         Height          =   315
         Left            =   210
         TabIndex        =   35
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Ref #"
         Height          =   285
         Left            =   180
         TabIndex        =   34
         Top             =   1200
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3555
      Left            =   90
      TabIndex        =   26
      Top             =   2700
      Width           =   9105
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Text            =   "Liabilities"
         Top             =   150
         Width           =   2835
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3060
         TabIndex        =   10
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   4260
         TabIndex        =   11
         Top             =   150
         Width           =   2805
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   7080
         TabIndex        =   12
         Top             =   150
         Width           =   1155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8310
         TabIndex        =   13
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8310
         TabIndex        =   18
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7020
         TabIndex        =   27
         Top             =   3180
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2685
         Left            =   180
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   450
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4736
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   90
      TabIndex        =   19
      Top             =   -30
      Width           =   9135
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   9030
         Top             =   1320
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7050
         TabIndex        =   25
         Top             =   2400
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1755
         Left            =   180
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   630
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3096
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8280
         TabIndex        =   17
         Top             =   780
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8280
         TabIndex        =   8
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   7080
         TabIndex        =   7
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4230
         TabIndex        =   6
         Top             =   330
         Width           =   2835
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3060
         TabIndex        =   5
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Text            =   "Assets"
         Top             =   330
         Width           =   2835
      End
      Begin VB.Label Label4 
         Caption         =   "Amount"
         Height          =   225
         Left            =   7110
         TabIndex        =   23
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   225
         Left            =   4260
         TabIndex        =   22
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Code"
         Height          =   225
         Left            =   3090
         TabIndex        =   21
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Title"
         Height          =   225
         Left            =   180
         TabIndex        =   20
         Top             =   150
         Width           =   825
      End
   End
   Begin Crystal.CrystalReport R1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "frmManualBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Sub GetSaleInformation(HCode As Long)
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Bales) as B,Sum(Qty) as Q,Sum(Amount) as A from Sales"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("B")) Then
    Text38.Text = TB.Fields("B").Value
End If
If Not IsNull(TB.Fields("Q")) Then
    Text39.Text = TB.Fields("Q").Value
End If
If Val(Text38.Text) > 0 And Val(Text39.Text) > 0 Then
    Text40.Text = Val(Text39.Text) / Val(Text38.Text)
End If
If Not IsNull(TB.Fields("A")) And Val(Text39.Text) > 0 Then
    Text41.Text = TB.Fields("A").Value / Val(Text39.Text)
End If
TB.Close
DB.Close

End Sub

Private Function GetHeadIssue(HCode As Long) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Qty) as Q from Issue where "
If HCode <= 99 Then
    Ssql = Ssql & " Val(Mid(ItemCode, 1, 2)) = " & HCode
ElseIf HCode > 99 Then
    Ssql = Ssql & " Val(Mid(ItemCode, 1, 5)) = " & HCode
End If
Ssql = Ssql & " and V_Date Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
Ssql = Ssql & " and RefNo=" & txtRefNo.Text
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    GetHeadIssue = TB.Fields("Q").Value
End If
TB.Close
DB.Close
End Function
Private Function GetHeadIssueAmount(HCode As Long) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Amount) as A from Issue where "
If HCode <= 99 Then
    Ssql = Ssql & " Val(Mid(ItemCode, 1, 2)) = " & HCode
ElseIf HCode > 99 Then
    Ssql = Ssql & " Val(Mid(ItemCode, 1, 5)) = " & HCode
End If
Ssql = Ssql & " and V_Date Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
Ssql = Ssql & " and RefNo=" & txtRefNo.Text
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("A").Value) Then
    GetHeadIssueAmount = TB.Fields("A").Value
End If
TB.Close
DB.Close
End Function


Private Function GetHeadProduction(HCode As Long) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Qty) as Q from Production where "
If HCode <= 99 Then
    Ssql = Ssql & " Val(Mid(ItemCode, 1, 2)) = " & HCode
ElseIf HCode > 99 Then
    Ssql = Ssql & " Val(Mid(ItemCode, 1, 5)) = " & HCode
End If
Ssql = Ssql & " and V_Date Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
Ssql = Ssql & " and RefNo=" & txtRefNo.Text
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    GetHeadProduction = TB.Fields("Q").Value
End If
TB.Close
DB.Close
End Function
Private Function GetHeadBal(HCode As Long, Optional Side As Integer) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
If Side = 1 Then
    Ssql = "Select Sum(Debit) as B from VouDTL where "
    If HCode <= 99 Then
        Ssql = Ssql & " Val(Mid(Party, 1, 2)) = " & HCode
    ElseIf HCode > 99 Then
        Ssql = Ssql & " Val(Mid(Party, 1, 5)) = " & HCode
    End If

ElseIf Side = 2 Then
    Ssql = "Select Sum(Credit) as B from VouDTL where "
    If HCode <= 99 Then
        Ssql = Ssql & " Val(Mid(Party, 1, 2)) = " & HCode
    ElseIf HCode > 99 Then
        Ssql = Ssql & " Val(Mid(Party, 1, 5)) = " & HCode
    End If
ElseIf Side = 0 Then
    Ssql = "Select Sum(Debit)-Sum(Credit) as B from VouDTL where "
    If HCode <= 99 Then
        Ssql = Ssql & " Val(Mid(Party, 1, 2)) = " & HCode
    ElseIf HCode > 99 Then
        Ssql = Ssql & " Val(Mid(Party, 1, 5)) = " & HCode
    End If
End If
Ssql = Ssql & " and V_Date <= #" & DTPicker2.Value & "#"
Set TB = DB.OpenRecordset(Ssql)

If Not IsNull(TB.Fields("B").Value) Then
    GetHeadBal = TB.Fields("B").Value
End If
TB.Close
DB.Close
End Function
Private Sub edit1()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select * from BS where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=1"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        txtRefNo.Text = TB.Fields("RefNo").Value & ""
With MSFlexGrid1
    Do While Not TB.EOF
        .TextMatrix(.Rows - 1, 0) = TB.Fields("Title").Value
        .TextMatrix(.Rows - 1, 1) = TB.Fields("Code").Value & ""
        .TextMatrix(.Rows - 1, 2) = TB.Fields("Name").Value & ""
        .TextMatrix(.Rows - 1, 3) = TB.Fields("Amount").Value & ""
        .Rows = .Rows + 1
        TB.MoveNext
    Loop
End With
End If
TB.Close
    
Ssql = "Select * from BS where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=2"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        
With MSFlexGrid2
    Do While Not TB.EOF
        .TextMatrix(.Rows - 1, 0) = TB.Fields("Title").Value
        .TextMatrix(.Rows - 1, 1) = TB.Fields("Code").Value & ""
        .TextMatrix(.Rows - 1, 2) = TB.Fields("Name").Value & ""
        .TextMatrix(.Rows - 1, 3) = TB.Fields("Amount").Value & ""
        .Rows = .Rows + 1
        TB.MoveNext
    Loop
End With
End If
TB.Close

DB.Close
End Sub

Private Function max1() As Long
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "select MAX(SheetNo) AS C FROM BS"
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("C").Value) Then
    max1 = TB.Fields("C").Value + 1
Else
    max1 = 1
End If
TB.Close
DB.Close
End Function
Private Sub save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
If Option2 = True Then
    Ssql = "Delete from BS where SheetNo=" & Val(Text36.Text)
    DB.Execute Ssql
End If

Set TB = DB.OpenRecordset("BS", dbOpenTable)
With MSFlexGrid1
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 1
        TB.Fields("Title").Value = .TextMatrix(R, 0)
        TB.Fields("Code").Value = Val(.TextMatrix(R, 1))
        TB.Fields("Name").Value = .TextMatrix(R, 2)
        TB.Fields("Amount").Value = Val(.TextMatrix(R, 3))
    TB.Update
Next R
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 1
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text5.Text)
    TB.Update

End With

With MSFlexGrid2
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 2
        TB.Fields("Title").Value = .TextMatrix(R, 0)
        TB.Fields("Code").Value = Val(.TextMatrix(R, 1))
        TB.Fields("Name").Value = .TextMatrix(R, 2)
        TB.Fields("Amount").Value = Val(.TextMatrix(R, 3))
    TB.Update
Next R
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 2
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text6.Text)
    TB.Update
'    TB.AddNew
'        TB.Fields("SDate").Value = DTPicker1.Value
'        TB.Fields("EDate").Value = DTPicker2.Value
'        TB.Fields("SheetNo").Value = Val(Text36.Text)
'        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
'        TB.Fields("HCode").Value = 2
'        TB.Fields("Title").Value = "Balance"
'        TB.Fields("Amount").Value = Val(Text5.Text) - Val(Text6.Text)
'    TB.Update
    
End With

    
    TB.Close
    DB.Close
    
    MsgBox "Information Has Been Saved"
End Sub
Private Sub flex1(Obj As MSFlexGrid)
With Obj
    .Rows = 0
    .Rows = 1
    .FixedRows = 0
    .FixedCols = 0
    .Cols = 4
    .ColWidth(0) = 2800
    .ColWidth(1) = 1200
    .ColWidth(2) = 2700
    .ColWidth(3) = 1200
    
End With
End Sub
Private Sub transfer1()
With MSFlexGrid1
    .TextMatrix(.Rows - 1, 0) = Text1.Text
    .TextMatrix(.Rows - 1, 1) = Text2.Text
    .TextMatrix(.Rows - 1, 2) = Text3.Text
    .TextMatrix(.Rows - 1, 3) = Abs(Val(Text4.Text))
    .Rows = .Rows + 1
End With
End Sub

Private Sub Transfer2()
With MSFlexGrid2
    .TextMatrix(.Rows - 1, 0) = Text10.Text
    .TextMatrix(.Rows - 1, 1) = Text9.Text
    .TextMatrix(.Rows - 1, 2) = Text8.Text
    .TextMatrix(.Rows - 1, 3) = Text7.Text
    .Rows = .Rows + 1
End With
End Sub

Private Sub Transfer3()
With MSFlexGrid3
    .TextMatrix(.Rows - 1, 0) = Text15.Text
    .TextMatrix(.Rows - 1, 1) = Text14.Text
    .TextMatrix(.Rows - 1, 2) = Text13.Text
    .TextMatrix(.Rows - 1, 3) = Text12.Text
    .Rows = .Rows + 1
End With
End Sub
Private Sub Transfer4()
With MSFlexGrid4
    .TextMatrix(.Rows - 1, 0) = Text32.Text
    .TextMatrix(.Rows - 1, 1) = Text33.Text
    .TextMatrix(.Rows - 1, 2) = Text34.Text
    .TextMatrix(.Rows - 1, 3) = Text35.Text
    .Rows = .Rows + 1
End With
End Sub

Private Sub Command1_Click()
transfer1
Command2_Click
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command10_Click()

If Option2 = True Then
    Dim Result As VbMsgBoxResult
    Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
    If Result = vbNo Then Exit Sub

    Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
    
    Ssql = "Delete from BS where SheetNo=" & Val(Text36.Text)
    DB.Execute Ssql
    

End If
End Sub

Private Sub Command11_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command12_Click()
Transfer4
Command13_Click
End Sub

Private Sub Command13_Click()
Text32.Text = ""
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text32.SetFocus
End Sub

Private Sub Command2_Click()
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
'Text10.Text = ""
Text9.Text = ""
Text8.Text = ""
Text7.Text = ""
Text10.SetFocus
End Sub

Private Sub Command4_Click()
Transfer2
Command3_Click
End Sub

Private Sub Command5_Click()
Text15.Text = ""
Text14.Text = ""
Text13.Text = ""
Text12.Text = ""
Text15.SetFocus
End Sub

Private Sub Command6_Click()
Transfer3
Command5_Click
End Sub

Private Sub Command4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command7_Click()
If Val(Text36.Text) > 0 Then
    Command7.Enabled = False
    save
    Command8_Click
    Command7.Enabled = True
Else
    MsgBox "Please Complete This Voucher"
End If
End Sub

Private Sub Command7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command8_Click()
Dim CNTL As Control

For Each CNTL In Me.Controls
    If TypeOf CNTL Is TextBox Then CNTL.Text = ""
Next
Text1.Text = "Assets"
Text10.Text = "Liabilities"
flex1 MSFlexGrid1
flex1 MSFlexGrid2
DTPicker1.Value = FStartDate
DTPicker2.Value = FEndDate
If Option1 = True Then Text36.Text = max1
If Option1 = True Then DTPicker1.SetFocus
If Option2 = True Then Text36.SetFocus
End Sub

Private Sub Command8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command9_Click()
R1.ReportFileName = App.path & "\BSManual.rpt"
R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
R1.SelectionFormula = "{PL.SheetNo}=" & Val(Text36.Text)
R1.Action = 1
End Sub

Private Sub Command9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
flex1 MSFlexGrid1
flex1 MSFlexGrid2
DTPicker1.Value = FStartDate
DTPicker2.Value = FEndDate
Text36.Text = max1
End Sub

Private Sub Label11_Click()
End Sub

Private Sub MSFlexGrid1_DblClick()
With MSFlexGrid1
    If .Rows > 1 Then
        Text1.Text = .TextMatrix(.Row, 0)
        Text2.Text = .TextMatrix(.Row, 1)
        Text3.Text = .TextMatrix(.Row, 2)
        Text4.Text = .TextMatrix(.Row, 3)
        If .Rows = 2 Then
            .Rows = 1
            .clear
        Else
            .RemoveItem .Row
        End If
    End If
End With

End Sub

Private Sub MSFlexGrid2_DblClick()
With MSFlexGrid2
    If .Rows > 1 Then
        Text10.Text = .TextMatrix(.Row, 0)
        Text9.Text = .TextMatrix(.Row, 1)
        Text8.Text = .TextMatrix(.Row, 2)
        Text7.Text = .TextMatrix(.Row, 3)
        If .Rows = 2 Then
            .Rows = 1
            .clear
            
        Else
            .RemoveItem .Row
        End If
    End If
End With

End Sub

Private Sub MSFlexGrid3_DblClick()
With MSFlexGrid3
    If .Rows > 1 Then
        Text15.Text = .TextMatrix(.Row, 0)
        Text14.Text = .TextMatrix(.Row, 1)
        Text13.Text = .TextMatrix(.Row, 2)
        Text12.Text = .TextMatrix(.Row, 3)
        If .Rows = 2 Then
            .Rows = 1
            .clear
        Else
            .RemoveItem .Row
        End If
    End If
End With

End Sub

Private Sub MSFlexGrid4_DblClick()
With MSFlexGrid4
    If .Rows > 1 Then
        Text32.Text = .TextMatrix(.Row, 0)
        Text33.Text = .TextMatrix(.Row, 1)
        Text34.Text = .TextMatrix(.Row, 2)
        Text35.Text = .TextMatrix(.Row, 3)
        If .Rows = 2 Then
            .Rows = 1
            .clear
        Else
            .RemoveItem .Row
        End If
    End If
End With

End Sub

Private Sub Option1_Click()
Command8_Click
End Sub

Private Sub Option2_Click()
Command8_Click
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
If Val(Text14.Text) > 0 Then
    Text13.Text = Blm1.heads(Val(Text14.Text))
    Text12.Text = GetHeadBal(Val(Text14.Text))
End If

End Sub

Private Sub Text18_Validate(Cancel As Boolean)
If Val(Text18.Text) > 0 Then
    Text17.Text = Blm1.heads(Val(Text18.Text))
    Text16.Text = GetHeadIssueAmount(Val(Text18.Text))
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text3.Text = Blm1.heads(Val(Text2.Text))
    Text4.Text = GetHeadBal(Val(Text2.Text), 0)
    'GetSaleInformation Val(Text2.Text)
End If
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
If Val(Text22.Text) > 0 Then
    Text21.Text = Blm1.heads(Val(Text22.Text))
    Text20.Text = GetHeadIssue(Val(Text22.Text))
End If

End Sub

Private Sub Text26_Validate(Cancel As Boolean)
If Val(Text26.Text) > 0 Then
    Text25.Text = Blm1.heads(Val(Text26.Text))
    Text24.Text = GetHeadProduction(Val(Text26.Text))
End If

End Sub

Private Sub Text30_Validate(Cancel As Boolean)
If Val(Text30.Text) > 0 Then
    Text29.Text = Blm1.heads(Val(Text30.Text))
    Text28.Text = GetHeadIssue(Val(Text30.Text))
End If

End Sub

Private Sub Text33_Validate(Cancel As Boolean)
If Val(Text33.Text) > 0 Then
    Text34.Text = Blm1.heads(Val(Text33.Text))
    Text35.Text = GetHeadBal(Val(Text33.Text))
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text36_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text36_Validate(Cancel As Boolean)
If Val(Text36.Text) > 0 And Option2 = True Then
    edit1
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
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

Private Sub Text9_Validate(Cancel As Boolean)
If Val(Text9.Text) > 0 Then
    Text8.Text = Blm1.heads(Val(Text9.Text))
    Text7.Text = GetHeadBal(Val(Text9.Text), 0)
End If
End Sub

Private Sub Timer1_Timer()
Dim T As Double
Dim R As Integer

With MSFlexGrid1
    For R = 0 To .Rows - 2
        T = T + Val(.TextMatrix(R, 3))
    Next R
End With
Text5.Text = T
T = 0
With MSFlexGrid2
    For R = 0 To .Rows - 2
        T = T + Val(.TextMatrix(R, 3))
    Next R
End With
Text6.Text = T
T = 0

End Sub

Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
