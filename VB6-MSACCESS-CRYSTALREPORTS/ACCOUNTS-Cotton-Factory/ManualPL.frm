VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmManualPL 
   Caption         =   "Profit / Loss Statement"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Height          =   1605
      Left            =   90
      TabIndex        =   62
      Top             =   6270
      Width           =   9105
      Begin VB.TextBox Text25 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6990
         TabIndex        =   64
         Top             =   1260
         Width           =   1185
      End
      Begin VB.CommandButton Command15 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8310
         TabIndex        =   63
         Top             =   540
         Width           =   735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8310
         TabIndex        =   28
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   7080
         TabIndex        =   27
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   4260
         TabIndex        =   26
         Top             =   150
         Width           =   2805
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   3060
         TabIndex        =   25
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Text            =   "Expenses"
         Top             =   150
         Width           =   2835
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   825
         Left            =   180
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   450
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1455
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1605
      Left            =   90
      TabIndex        =   58
      Top             =   4680
      Width           =   9105
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   180
         TabIndex        =   19
         Text            =   "Closing Stocks"
         Top             =   150
         Width           =   2835
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   3060
         TabIndex        =   20
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   4260
         TabIndex        =   21
         Top             =   150
         Width           =   2805
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   7080
         TabIndex        =   22
         Top             =   150
         Width           =   1155
      End
      Begin VB.CommandButton Command13 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8310
         TabIndex        =   23
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8310
         TabIndex        =   60
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox Text16 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6990
         TabIndex        =   59
         Top             =   1260
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   825
         Left            =   180
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   450
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1455
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1545
      Left            =   90
      TabIndex        =   54
      Top             =   3150
      Width           =   9105
      Begin VB.TextBox Text15 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6990
         TabIndex        =   56
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8310
         TabIndex        =   55
         Top             =   540
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8310
         TabIndex        =   18
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   7080
         TabIndex        =   17
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   4260
         TabIndex        =   16
         Top             =   150
         Width           =   2805
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   3060
         TabIndex        =   15
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Text            =   "Purchases"
         Top             =   150
         Width           =   2835
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   735
         Left            =   180
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   450
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1296
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame5 
      Height          =   3135
      Left            =   9240
      TabIndex        =   47
      Top             =   -30
      Width           =   2355
      Begin VB.TextBox Text26 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   810
         TabIndex        =   29
         Top             =   1560
         Width           =   1395
      End
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
         TabIndex        =   49
         Top             =   870
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Update"
         Height          =   285
         Left            =   1380
         TabIndex        =   48
         Top             =   1140
         Width           =   945
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Save"
         Height          =   435
         Left            =   240
         TabIndex        =   30
         Top             =   2160
         Width           =   585
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Reset"
         Height          =   435
         Left            =   810
         TabIndex        =   31
         Top             =   2160
         Width           =   585
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Prev"
         Height          =   435
         Left            =   1380
         TabIndex        =   32
         Top             =   2160
         Width           =   585
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Delete"
         Height          =   435
         Left            =   540
         TabIndex        =   33
         Top             =   2580
         Width           =   585
      End
      Begin VB.CommandButton Command11 
         Caption         =   "E&xit"
         Height          =   435
         Left            =   1110
         TabIndex        =   34
         Top             =   2580
         Width           =   585
      End
      Begin VB.TextBox txtRefNo 
         Height          =   285
         Left            =   780
         TabIndex        =   3
         Top             =   1200
         Width           =   555
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   510
         Width           =   1365
         _ExtentX        =   2408
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
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60620803
         CurrentDate     =   39297
      End
      Begin VB.Label Label6 
         Caption         =   "P/L B/F"
         Height          =   315
         Left            =   180
         TabIndex        =   66
         Top             =   1590
         Width           =   705
      End
      Begin VB.Label Label17 
         Caption         =   "From"
         Height          =   285
         Left            =   210
         TabIndex        =   53
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   285
         Left            =   180
         TabIndex        =   52
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label19 
         Caption         =   "Sheet#"
         Height          =   315
         Left            =   210
         TabIndex        =   51
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Ref #"
         Height          =   285
         Left            =   180
         TabIndex        =   50
         Top             =   1200
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1545
      Left            =   90
      TabIndex        =   44
      Top             =   1620
      Width           =   9105
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Text            =   "Opening Stocks"
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
         TabIndex        =   36
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6990
         TabIndex        =   45
         Top             =   1200
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   765
         Left            =   180
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   450
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1349
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   90
      TabIndex        =   37
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
         TabIndex        =   43
         Top             =   1320
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   705
         Left            =   180
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   630
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1244
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8280
         TabIndex        =   35
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
         Text            =   "Sales"
         Top             =   330
         Width           =   2835
      End
      Begin VB.Label Label4 
         Caption         =   "Amount"
         Height          =   225
         Left            =   7110
         TabIndex        =   41
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   225
         Left            =   4260
         TabIndex        =   40
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Code"
         Height          =   225
         Left            =   3090
         TabIndex        =   39
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Title"
         Height          =   225
         Left            =   180
         TabIndex        =   38
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
Attribute VB_Name = "frmManualPL"
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
'MsgBox Ssql
Set TB = DB.OpenRecordset(Ssql)

If Not IsNull(TB.Fields("B").Value) Then
    GetHeadBal = TB.Fields("B").Value
End If
TB.Close
DB.Close
End Function
Private Function GetHeadOpBal(HCode As Long, Optional Side As Integer) As Double
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
Ssql = Ssql & " and V_Type=10 and V_Date <= #" & DTPicker2.Value & "#"
'MsgBox Ssql
Set TB = DB.OpenRecordset(Ssql)

If Not IsNull(TB.Fields("B").Value) Then
    GetHeadOpBal = TB.Fields("B").Value
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

Ssql = "Select * from PL where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=1"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        txtRefNo.Text = TB.Fields("RefNo").Value & ""
        Text26.Text = TB.Fields("PLBF").Value & ""
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
    
Ssql = "Select * from PL where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=2"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        txtRefNo.Text = TB.Fields("RefNo").Value & ""
        Text26.Text = TB.Fields("PLBF").Value & ""
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

Ssql = "Select * from PL where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=3"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        Text26.Text = TB.Fields("PLBF").Value
        txtRefNo.Text = TB.Fields("RefNo").Value & ""
With MSFlexGrid3
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

Ssql = "Select * from PL where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=4"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        Text26.Text = TB.Fields("PLBF").Value
        txtRefNo.Text = TB.Fields("RefNo").Value & ""
With MSFlexGrid4
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

Ssql = "Select * from PL where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=5"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        Text26.Text = TB.Fields("PLBF").Value
        txtRefNo.Text = TB.Fields("RefNo").Value & ""
With MSFlexGrid5
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

Ssql = "select MAX(SheetNo) AS C FROM PL"
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
    Ssql = "Delete from PL where SheetNo=" & Val(Text36.Text)
    DB.Execute Ssql
End If

Set TB = DB.OpenRecordset("PL", dbOpenTable)
With MSFlexGrid1
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
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
        TB.Fields("PLBF").Value = Val(Text26.Text)
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
        TB.Fields("PLBF").Value = Val(Text26.Text)
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
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 2
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text6.Text)
    TB.Update
End With

With MSFlexGrid3
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 3
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
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 3
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text15.Text)
    TB.Update
End With
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 3
        TB.Fields("Title").Value = "Grand Total"
        TB.Fields("Amount").Value = Val(Text15.Text) + Val(Text6.Text)
    TB.Update
    
With MSFlexGrid4
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 4
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
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 4
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text16.Text)
    TB.Update
End With
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 4
        TB.Fields("Title").Value = "Net Cost of Material Consumed"
        TB.Fields("Amount").Value = (Val(Text15.Text) + Val(Text6.Text)) - Val(Text16.Text)
    TB.Update
    
With MSFlexGrid5
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 5
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
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 5
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text25.Text)
    TB.Update
End With
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = "Grand Total"
        TB.Fields("Amount").Value = ((Val(Text15.Text) + Val(Text6.Text)) - Val(Text16.Text)) + Val(Text25.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = "Net Profit/Loss For The Year"
        TB.Fields("Amount").Value = Val(Text5.Text) - (((Val(Text15.Text) + Val(Text6.Text)) - Val(Text16.Text)) + Val(Text25.Text))
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = "Profit/Loss Brought Forward"
        TB.Fields("Amount").Value = Val(Text26.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("PLBF").Value = Val(Text26.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = "Profit/Loss Carried Forward"
        TB.Fields("Amount").Value = Val(Text26.Text) + (Val(Text5.Text) - (((Val(Text15.Text) + Val(Text6.Text)) - Val(Text16.Text)) + Val(Text25.Text)))
    TB.Update


    
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
    .TextMatrix(.Rows - 1, 0) = Text11.Text
    .TextMatrix(.Rows - 1, 1) = Text12.Text
    .TextMatrix(.Rows - 1, 2) = Text13.Text
    .TextMatrix(.Rows - 1, 3) = Text14.Text
    .Rows = .Rows + 1
End With
End Sub
Private Sub Transfer4()
With MSFlexGrid4
    .TextMatrix(.Rows - 1, 0) = Text20.Text
    .TextMatrix(.Rows - 1, 1) = Text19.Text
    .TextMatrix(.Rows - 1, 2) = Text18.Text
    .TextMatrix(.Rows - 1, 3) = Text17.Text
    .Rows = .Rows + 1
End With
End Sub
Private Sub Transfer5()
With MSFlexGrid5
    .TextMatrix(.Rows - 1, 0) = Text21.Text
    .TextMatrix(.Rows - 1, 1) = Text22.Text
    .TextMatrix(.Rows - 1, 2) = Text23.Text
    .TextMatrix(.Rows - 1, 3) = Text24.Text
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
    
    Ssql = "Delete from PL where SheetNo=" & Val(Text36.Text)
    DB.Execute Ssql
    

End If
End Sub

Private Sub Command11_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command12_Click()
'Text20.Text = ""
Text19.Text = ""
Text18.Text = ""
Text17.Text = ""
Text20.SetFocus

End Sub

Private Sub Command13_Click()
Transfer4
Command12_Click

End Sub

Private Sub Command13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command14_Click()
Transfer5
Command15_Click
End Sub

Private Sub Command14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command15_Click()
'Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text21.SetFocus
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

Private Sub Command4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command5_Click()
Transfer3
Command6_Click

End Sub

Private Sub Command5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command6_Click()
'Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text11.SetFocus

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
Text1.Text = "Sales"
Text10.Text = "Opening Stocks"
Text11.Text = "Purchases"
Text20.Text = "Closing Stocks"
Text21.Text = "Expenses."


flex1 MSFlexGrid1
flex1 MSFlexGrid2
flex1 MSFlexGrid3
flex1 MSFlexGrid4
flex1 MSFlexGrid5
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
R1.ReportFileName = App.path & "\PLManual.rpt"
R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
R1.SelectionFormula = "{PL.SheetNo}=" & Val(Text36.Text)
R1.Action = 1
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
flex1 MSFlexGrid3
flex1 MSFlexGrid4
flex1 MSFlexGrid5
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
        Text11.Text = .TextMatrix(.Row, 0)
        Text12.Text = .TextMatrix(.Row, 1)
        Text13.Text = .TextMatrix(.Row, 2)
        Text14.Text = .TextMatrix(.Row, 3)
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
        Text20.Text = .TextMatrix(.Row, 0)
        Text19.Text = .TextMatrix(.Row, 1)
        Text18.Text = .TextMatrix(.Row, 2)
        Text17.Text = .TextMatrix(.Row, 3)
        If .Rows = 2 Then
            .Rows = 1
            .clear
        Else
            .RemoveItem .Row
        End If
    End If
End With

End Sub

Private Sub MSFlexGrid5_DblClick()
With MSFlexGrid5
    If .Rows > 1 Then
        Text21.Text = .TextMatrix(.Row, 0)
        Text22.Text = .TextMatrix(.Row, 1)
        Text23.Text = .TextMatrix(.Row, 2)
        Text24.Text = .TextMatrix(.Row, 3)
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
If Val(Text12.Text) > 0 Then
    Text13.Text = Blm1.heads(Val(Text12.Text))
    Text14.Text = GetHeadBal(Val(Text12.Text), 1)
    'GetSaleInformation Val(Text2.Text)
End If

End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
'If Val(Text14.Text) > 0 Then
'    Text13.Text = Blm1.heads(Val(Text14.Text))
'    Text12.Text = GetHeadBal(Val(Text14.Text))
'End If

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text18_Validate(Cancel As Boolean)
If Val(Text18.Text) > 0 Then
    Text17.Text = Blm1.heads(Val(Text18.Text))
    Text16.Text = GetHeadIssueAmount(Val(Text18.Text))
End If

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
If Val(Text19.Text) > 0 Then
    Text18.Text = Blm1.heads(Val(Text19.Text))
    Text17.Text = GetHeadBal(Val(Text19.Text))
    'GetSaleInformation Val(Text2.Text)
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text3.Text = Blm1.heads(Val(Text2.Text))
    Text4.Text = GetHeadBal(Val(Text2.Text), 2)
    'GetSaleInformation Val(Text2.Text)
End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
If Val(Text22.Text) > 0 Then
    Text23.Text = Blm1.heads(Val(Text22.Text))
    Text24.Text = GetHeadBal(Val(Text22.Text))
End If

End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
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
    Text7.Text = GetHeadOpBal(Val(Text9.Text), 1)
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
With MSFlexGrid3
    For R = 0 To .Rows - 2
        T = T + Val(.TextMatrix(R, 3))
    Next R
End With
Text15.Text = T
T = 0
With MSFlexGrid4
    For R = 0 To .Rows - 2
        T = T + Val(.TextMatrix(R, 3))
    Next R
End With
Text16.Text = T
T = 0
With MSFlexGrid5
    For R = 0 To .Rows - 2
        T = T + Val(.TextMatrix(R, 3))
    Next R
End With
Text25.Text = T

End Sub

Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
