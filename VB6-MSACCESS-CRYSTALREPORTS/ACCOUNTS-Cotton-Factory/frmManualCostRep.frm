VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmManualCostRep 
   Caption         =   "Cost Report - 2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command15 
      Caption         =   "Refresh Bags Data"
      Height          =   495
      Left            =   9360
      TabIndex        =   99
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Refresh Bales Data"
      Height          =   495
      Left            =   9360
      TabIndex        =   98
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Frame Frame9 
      Height          =   1155
      Left            =   9360
      TabIndex        =   94
      Top             =   4320
      Width           =   2475
      Begin VB.TextBox Text41 
         Height          =   285
         Left            =   1290
         TabIndex        =   86
         Top             =   630
         Width           =   975
      End
      Begin VB.TextBox Text46 
         Height          =   285
         Left            =   1290
         TabIndex        =   85
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Avg. Sale @"
         Height          =   315
         Left            =   120
         TabIndex        =   96
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label22 
         Caption         =   "T.Sold Weight"
         Height          =   285
         Left            =   180
         TabIndex        =   95
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1665
      Left            =   9360
      TabIndex        =   90
      Top             =   2160
      Width           =   2475
      Begin VB.TextBox Text45 
         Height          =   285
         Left            =   1290
         TabIndex        =   82
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox Text44 
         Height          =   285
         Left            =   1290
         TabIndex        =   83
         Top             =   750
         Width           =   975
      End
      Begin VB.TextBox Text43 
         Height          =   285
         Left            =   1290
         TabIndex        =   84
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Bags"
         Height          =   315
         Left            =   180
         TabIndex        =   93
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label20 
         Caption         =   "Sold Weight"
         Height          =   285
         Left            =   180
         TabIndex        =   92
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Avg WT./Bag"
         Height          =   255
         Left            =   180
         TabIndex        =   91
         Top             =   1230
         Width           =   1035
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1665
      Left            =   9360
      TabIndex        =   77
      Top             =   -30
      Width           =   2475
      Begin VB.TextBox Text40 
         Height          =   285
         Left            =   1290
         TabIndex        =   81
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text39 
         Height          =   285
         Left            =   1290
         TabIndex        =   80
         Top             =   750
         Width           =   975
      End
      Begin VB.TextBox Text38 
         Height          =   285
         Left            =   1290
         TabIndex        =   79
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Avg WT./Bale"
         Height          =   255
         Left            =   180
         TabIndex        =   88
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Sold Weight"
         Height          =   285
         Left            =   180
         TabIndex        =   87
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Bales"
         Height          =   315
         Left            =   180
         TabIndex        =   78
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1605
      Left            =   90
      TabIndex        =   71
      Top             =   3300
      Width           =   9225
      Begin VB.CommandButton Command13 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8220
         TabIndex        =   75
         Top             =   570
         Width           =   795
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8220
         TabIndex        =   74
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox Text37 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6960
         TabIndex        =   73
         Top             =   1230
         Width           =   1185
      End
      Begin VB.TextBox Text35 
         Height          =   285
         Left            =   7020
         TabIndex        =   13
         Top             =   120
         Width           =   1155
      End
      Begin VB.TextBox Text34 
         Height          =   285
         Left            =   4170
         TabIndex        =   12
         Top             =   120
         Width           =   2805
      End
      Begin VB.TextBox Text33 
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   120
         Width           =   1155
      End
      Begin VB.TextBox Text32 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Expenses"
         Top             =   120
         Width           =   2835
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   825
         Left            =   90
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   420
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   1455
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2145
      Left            =   8400
      TabIndex        =   65
      Top             =   6270
      Width           =   3375
      Begin VB.TextBox txtRefNo 
         Height          =   315
         Left            =   780
         TabIndex        =   41
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Caption         =   "E&xit"
         Height          =   435
         Left            =   2520
         TabIndex        =   46
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Delete"
         Height          =   435
         Left            =   1950
         TabIndex        =   45
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Prev"
         Height          =   435
         Left            =   1380
         TabIndex        =   44
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Reset"
         Height          =   435
         Left            =   810
         TabIndex        =   43
         Top             =   1560
         Width           =   585
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Save"
         Height          =   435
         Left            =   240
         TabIndex        =   42
         Top             =   1560
         Width           =   585
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Update"
         Height          =   285
         Left            =   2280
         TabIndex        =   70
         Top             =   900
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   285
         Left            =   1380
         TabIndex        =   69
         Top             =   870
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox Text36 
         Height          =   285
         Left            =   780
         TabIndex        =   40
         Top             =   840
         Width           =   555
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   780
         TabIndex        =   39
         Top             =   510
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55967747
         CurrentDate     =   39297
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   780
         TabIndex        =   38
         Top             =   180
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55967747
         CurrentDate     =   39297
      End
      Begin VB.Label Label5 
         Caption         =   "Ref #"
         Height          =   285
         Left            =   180
         TabIndex        =   76
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label19 
         Caption         =   "Sheet#"
         Height          =   315
         Left            =   210
         TabIndex        =   68
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   285
         Left            =   180
         TabIndex        =   67
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label17 
         Caption         =   "From"
         Height          =   285
         Left            =   180
         TabIndex        =   66
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1725
      Left            =   120
      TabIndex        =   60
      Top             =   6540
      Width           =   8235
      Begin VB.TextBox Text31 
         Height          =   285
         Left            =   60
         TabIndex        =   34
         Text            =   "Expenses"
         Top             =   1350
         Width           =   2835
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   2940
         TabIndex        =   35
         Top             =   1350
         Width           =   1155
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   4110
         TabIndex        =   36
         Top             =   1350
         Width           =   2835
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   6960
         TabIndex        =   37
         Top             =   1350
         Width           =   1155
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   60
         TabIndex        =   30
         Text            =   "Production"
         Top             =   1050
         Width           =   2835
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   2940
         TabIndex        =   31
         Top             =   1050
         Width           =   1155
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   4110
         TabIndex        =   32
         Top             =   1050
         Width           =   2835
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   6960
         TabIndex        =   33
         Top             =   1050
         Width           =   1155
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   60
         TabIndex        =   26
         Text            =   "Material Issued"
         Top             =   750
         Width           =   2835
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   2940
         TabIndex        =   27
         Top             =   750
         Width           =   1155
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   4110
         TabIndex        =   28
         Top             =   750
         Width           =   2835
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   6960
         TabIndex        =   29
         Top             =   750
         Width           =   1155
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   60
         TabIndex        =   22
         Text            =   "Material Cost"
         Top             =   450
         Width           =   2835
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   2940
         TabIndex        =   23
         Top             =   450
         Width           =   1155
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   4110
         TabIndex        =   24
         Top             =   450
         Width           =   2835
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   6960
         TabIndex        =   25
         Top             =   450
         Width           =   1155
      End
      Begin VB.Label Label16 
         Caption         =   "Title"
         Height          =   225
         Left            =   60
         TabIndex        =   64
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label15 
         Caption         =   "Code"
         Height          =   225
         Left            =   2970
         TabIndex        =   63
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label14 
         Caption         =   "Description"
         Height          =   225
         Left            =   4140
         TabIndex        =   62
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label13 
         Caption         =   "Amount"
         Height          =   225
         Left            =   6990
         TabIndex        =   61
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1725
      Left            =   120
      TabIndex        =   57
      Top             =   4830
      Width           =   9195
      Begin Crystal.CrystalReport R1 
         Left            =   1860
         Top             =   180
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
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Text            =   "Other Assets"
         Top             =   150
         Width           =   2835
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   3060
         TabIndex        =   17
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   4230
         TabIndex        =   18
         Top             =   150
         Width           =   2835
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   7080
         TabIndex        =   19
         Top             =   150
         Width           =   1155
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8280
         TabIndex        =   20
         Top             =   150
         Width           =   765
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8280
         TabIndex        =   21
         Top             =   540
         Width           =   765
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6990
         TabIndex        =   58
         Top             =   1350
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   825
         Left            =   180
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   510
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1455
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   90
      TabIndex        =   54
      Top             =   1620
      Width           =   9225
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Text            =   "Purchases"
         Top             =   150
         Width           =   2835
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3060
         TabIndex        =   7
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   4260
         TabIndex        =   8
         Top             =   150
         Width           =   2805
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   7080
         TabIndex        =   9
         Top             =   150
         Width           =   1155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8310
         TabIndex        =   14
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   8310
         TabIndex        =   15
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7050
         TabIndex        =   55
         Top             =   1320
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   885
         Left            =   180
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   450
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1561
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   90
      TabIndex        =   47
      Top             =   -30
      Width           =   9225
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   8700
         Top             =   1290
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7050
         TabIndex        =   53
         Top             =   1320
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   705
         Left            =   180
         TabIndex        =   52
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
         TabIndex        =   5
         Top             =   780
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   405
         Left            =   8310
         TabIndex        =   4
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   7080
         TabIndex        =   3
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4230
         TabIndex        =   2
         Top             =   330
         Width           =   2835
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3060
         TabIndex        =   1
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "Sales"
         Top             =   330
         Width           =   2835
      End
      Begin VB.Label Label4 
         Caption         =   "Amount"
         Height          =   225
         Left            =   7110
         TabIndex        =   51
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   225
         Left            =   4260
         TabIndex        =   50
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Code"
         Height          =   225
         Left            =   3090
         TabIndex        =   49
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Title"
         Height          =   225
         Left            =   180
         TabIndex        =   48
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   285
      Left            =   10830
      TabIndex        =   97
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   285
      Left            =   10800
      TabIndex        =   89
      Top             =   5640
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "frmManualCostRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Sub GetSaleInformation()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
If Len(MSFlexGrid1.TextMatrix(0, 1)) >= 5 Then
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Bales) as B,Sum(Qty) as Q,Sum(Amount) as A from Sales"
Ssql = Ssql & " where Mid(Item,1,5)='" & MSFlexGrid1.TextMatrix(0, 1) & "'"
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
If Not IsNull(TB.Fields("A")) Then
    Label10.Caption = TB.Fields("A").Value
End If
TB.Close
DB.Close
End If

End Sub

Private Sub GetSaleInformation2()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
If MSFlexGrid1.Rows > 1 Then
If Len(MSFlexGrid1.TextMatrix(1, 1)) >= 5 Then
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Bales) as B,Sum(Qty) as Q,Sum(Amount) as A from Sales"
Ssql = Ssql & " where Mid(ITem,1,5)='" & MSFlexGrid1.TextMatrix(1, 1) & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("B")) Then
    Text45.Text = TB.Fields("B").Value
End If
If Not IsNull(TB.Fields("Q")) Then
    Text44.Text = TB.Fields("Q").Value
End If
If Val(Text45.Text) > 0 And Val(Text44.Text) > 0 Then
    Text43.Text = Val(Text44.Text) / Val(Text45.Text)
End If
If Not IsNull(TB.Fields("A")) Then
    Label11.Caption = TB.Fields("A").Value
End If

TB.Close
DB.Close
End If
End If
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
Private Function GetIssue() As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Qty) as Q from IssueSH where "
Ssql = Ssql & "V_Date Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
Ssql = Ssql & " and RefNo=" & txtRefNo.Text
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    GetIssue = TB.Fields("Q").Value
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

Ssql = "Select Sum(Amount) as A from IssueSH where "
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
Private Function GetIssueAmount() As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Amount) as A from Issue where "
Ssql = Ssql & "V_Date Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
Ssql = Ssql & " and RefNo=" & txtRefNo.Text
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("A").Value) Then
    GetIssueAmount = TB.Fields("A").Value
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

Private Function GetProduction() As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim R As Integer
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select Sum(Qty) as Q from Production where "
Ssql = Ssql & " V_Date Between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
Ssql = Ssql & " and RefNo=" & txtRefNo.Text
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    GetProduction = TB.Fields("Q").Value
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
'MsgBox Ssql
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

Ssql = "Select * from CostSheet where SheetNo=" & Val(Text36.Text)
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
    
Ssql = "Select * from CostSheet where SheetNo=" & Val(Text36.Text)
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

Ssql = "Select * from CostSheet where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=3"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        
With MSFlexGrid4
    Do While Not TB.EOF
        .TextMatrix(.Rows - 1, 0) = TB.Fields("Title").Value & ""
        .TextMatrix(.Rows - 1, 1) = TB.Fields("Code").Value & ""
        .TextMatrix(.Rows - 1, 2) = TB.Fields("Name").Value & ""
        .TextMatrix(.Rows - 1, 3) = TB.Fields("Amount").Value & ""
        .Rows = .Rows + 1
        TB.MoveNext
    Loop
End With
End If
TB.Close


Ssql = "Select * from CostSheet where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=4"
Ssql = Ssql & " and Code <> Null"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        DTPicker1.Value = TB.Fields("SDate").Value
        DTPicker2.Value = TB.Fields("EDate").Value
        Text36.Text = TB.Fields("SheetNo").Value
        
With MSFlexGrid3
    Do While Not TB.EOF
        .TextMatrix(.Rows - 1, 0) = TB.Fields("Title").Value & ""
        .TextMatrix(.Rows - 1, 1) = TB.Fields("Code").Value & ""
        .TextMatrix(.Rows - 1, 2) = TB.Fields("Name").Value & ""
        .TextMatrix(.Rows - 1, 3) = TB.Fields("Amount").Value & ""
        .Rows = .Rows + 1
        TB.MoveNext
    Loop
End With
End If
TB.Close
Ssql = "Select * from CostSheet where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=6"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        Text19.Text = TB.Fields("Title").Value & ""
        Text18.Text = TB.Fields("Code").Value & ""
        Text17.Text = TB.Fields("Name").Value & ""
        Text16.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text23.Text = TB.Fields("Title").Value & ""
        Text22.Text = TB.Fields("Code").Value & ""
        Text21.Text = TB.Fields("Name").Value & ""
        Text20.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text27.Text = TB.Fields("Title").Value & ""
        Text26.Text = TB.Fields("Code").Value & ""
        Text25.Text = TB.Fields("Name").Value & ""
        Text24.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text31.Text = TB.Fields("Title").Value & ""
        Text30.Text = TB.Fields("Code").Value & ""
        Text29.Text = TB.Fields("Name").Value & ""
        Text28.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
TB.Close

Ssql = "Select * from CostSheet where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=8"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        Text38.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text39.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text40.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text41.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
TB.Close

Ssql = "Select * from CostSheet where SheetNo=" & Val(Text36.Text)
Ssql = Ssql & " and HCode=9"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
        Text45.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text44.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
    
If Not TB.EOF Then
        Text43.Text = TB.Fields("Amount").Value & ""
        TB.MoveNext
End If
TB.Close
DB.Close
End Sub

Private Function max1() As Long
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Ssql = "select MAX(SheetNo) AS C FROM CostSheet"
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
    Ssql = "Delete from CostSheet where SheetNo=" & Val(Text36.Text)
    DB.Execute Ssql
End If

Set TB = DB.OpenRecordset("CostSheet", dbOpenTable)
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
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 2
        TB.Fields("Title").Value = "Balance"
        TB.Fields("Amount").Value = Val(Text5.Text) - Val(Text6.Text)
    TB.Update
    
End With

With MSFlexGrid3
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 4
        TB.Fields("Title").Value = .TextMatrix(R, 0)
        TB.Fields("Code").Value = Val(.TextMatrix(R, 1))
        TB.Fields("Name").Value = .TextMatrix(R, 2)
        TB.Fields("Amount").Value = Val(.TextMatrix(R, 3))
    TB.Update
Next R
End With
TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 4
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text11.Text)
TB.Update

TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 4
        TB.Fields("Title").Value = "Net Amount ###"
        TB.Fields("Amount").Value = (Val(Text5.Text) - (Val(Text6.Text) + Val(Text37.Text))) + Val(Text11.Text)
'        MsgBox "Test"
    TB.Update

With MSFlexGrid4
For R = 0 To .Rows - 2
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
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
        TB.Fields("HCode").Value = 3
        TB.Fields("Title").Value = "Total"
        TB.Fields("Amount").Value = Val(Text37.Text)
    TB.Update
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 3
        TB.Fields("Title").Value = "Grand Total ###"
        TB.Fields("Amount").Value = (Val(Text5.Text) - (Val(Text6.Text) + Val(Text37.Text)))
    TB.Update
    
     
    'MsgBox "Test"
End With

    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = Text19.Text
        TB.Fields("Code").Value = Val(Text18.Text)
        TB.Fields("Name").Value = Text17.Text
        TB.Fields("Amount").Value = Val(Text16.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = Text23.Text
        TB.Fields("Code").Value = Val(Text22.Text)
        TB.Fields("Name").Value = Text21.Text
        TB.Fields("Amount").Value = Val(Text20.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = Text27.Text
        TB.Fields("Code").Value = Val(Text26.Text)
        TB.Fields("Name").Value = Text25.Text
        TB.Fields("Amount").Value = Val(Text24.Text)
    TB.Update

    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 6
        TB.Fields("Title").Value = Text31.Text
        TB.Fields("Code").Value = Val(Text30.Text)
        TB.Fields("Name").Value = Text29.Text
        TB.Fields("Amount").Value = Val(Text28.Text)
    TB.Update

        
   

    
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 7
        TB.Fields("Title").Value = "INPUT (PER KG)"
        TB.Fields("Amount").Value = Val(Text16.Text) / Val(Text20.Text)
    TB.Update
    
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 7
        TB.Fields("Title").Value = "OUTPUT (% YIELD)"
        TB.Fields("Amount").Value = (Val(Text24.Text) / Val(Text20.Text)) * 100
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 7
        TB.Fields("Title").Value = "EXPENCES (PER/KG)"
        TB.Fields("Amount").Value = (Val(Text28.Text) / Val(Text24.Text))
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 7
        TB.Fields("Title").Value = "COST (PER/KG)"
        TB.Fields("Amount").Value = (Val(Text16.Text) / Val(Text24.Text))
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 7
        TB.Fields("Title").Value = "TOTAL COST (PER/KG) ###"
        TB.Fields("Amount").Value = (Val(Text16.Text) / Val(Text24.Text)) + (Val(Text28.Text) / Val(Text24.Text))
    TB.Update
    
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 8
        TB.Fields("Title").Value = "BALES"
        TB.Fields("Amount").Value = Val(Text38.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 8
        TB.Fields("Title").Value = "WEIGHT"
        TB.Fields("Amount").Value = Val(Text39.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 8
        TB.Fields("Title").Value = "AVG. WT./BALE"
        TB.Fields("Amount").Value = Val(Text40.Text)
    TB.Update
    
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 9
        TB.Fields("Title").Value = "BAGS"
        TB.Fields("Amount").Value = Val(Text45.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 9
        TB.Fields("Title").Value = "WEIGHT"
        TB.Fields("Amount").Value = Val(Text44.Text)
    TB.Update
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 9
        TB.Fields("Title").Value = "AVG. WT./BAG"
        TB.Fields("Amount").Value = Val(Text43.Text)
    TB.Update
    
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 10
        TB.Fields("Title").Value = "TOTAL SOLD WEIGHT"
        TB.Fields("Amount").Value = Val(Text46.Text)
    TB.Update
    
    
    TB.AddNew
        TB.Fields("SDate").Value = DTPicker1.Value
        TB.Fields("EDate").Value = DTPicker2.Value
        TB.Fields("SheetNo").Value = Val(Text36.Text)
        TB.Fields("RefNo").Value = Val(txtRefNo.Text)
        TB.Fields("HCode").Value = 10
        TB.Fields("Title").Value = "AVG. SALE RATE"
        TB.Fields("Amount").Value = Val(Text41.Text)
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

Private Sub Command10_Click()

If Option2 = True Then
    Dim Result As VbMsgBoxResult
    Result = MsgBox("Do You Realy Want to Delete This Voucher", vbYesNo)
    If Result = vbNo Then Exit Sub

    Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
    
    Ssql = "Delete from CostSheet where SheetNo=" & Val(Text36.Text)
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
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text32.SetFocus
End Sub

Private Sub Command14_Click()
GetSaleInformation

End Sub

Private Sub Command15_Click()
GetSaleInformation2
End Sub

Private Sub Command2_Click()

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()

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
Text14.Text = ""
Text13.Text = ""
Text12.Text = ""
Text15.SetFocus
End Sub

Private Sub Command6_Click()
Transfer3
Command5_Click
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

Private Sub Command8_Click()
Dim CNTL As Control

For Each CNTL In Me.Controls
    If TypeOf CNTL Is TextBox Then CNTL.Text = ""
Next
Text1.Text = "Sales"
Text10.Text = "Purchases"
Text32.Text = "Expenses"
Text15.Text = "Other Assets"
flex1 MSFlexGrid1
flex1 MSFlexGrid2
flex1 MSFlexGrid3
flex1 MSFlexGrid4
DTPicker1.Value = FStartDate
DTPicker2.Value = Date
Text19.Text = "Material Cost"
Text23.Text = "Material Issued"
Text27.Text = "Production"
Text31.Text = "Expenses"
If Option1 = True Then Text36.Text = max1
End Sub

Private Sub Command9_Click()
R1.ReportFileName = App.path & "\ManualCostSheet.rpt"
R1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
R1.SelectionFormula = "{CostSheet.SheetNo}=" & Val(Text36.Text)
R1.Action = 1
End Sub

Private Sub Form_Load()
flex1 MSFlexGrid1
flex1 MSFlexGrid2
flex1 MSFlexGrid4
flex1 MSFlexGrid3
DTPicker1.Value = FStartDate
DTPicker2.Value = Date
Text36.Text = max1
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
    'Text16.Text = GetIssueAmount()
End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text3.Text = Blm1.heads(Val(Text2.Text))
    Text4.Text = GetHeadBal(Val(Text2.Text), 2)

End If
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
If Val(Text22.Text) > 0 Then
    Text21.Text = Blm1.heads(Val(Text22.Text))
    'Text20.Text = GetIssue()
End If

End Sub

Private Sub Text26_Validate(Cancel As Boolean)
If Val(Text26.Text) > 0 Then
    Text25.Text = Blm1.heads(Val(Text26.Text))
    'Text24.Text = GetProduction()
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

Private Sub Text36_Validate(Cancel As Boolean)
If Val(Text36.Text) > 0 And Option2 = True Then
    edit1
End If
End Sub

Private Sub Text38_Change()
If Val(Text38.Text) > 0 And Val(Text39.Text) > 0 Then
    Text40.Text = Val(Text39.Text) / Val(Text38.Text)
    Text41.Text = Val(Label10.Caption) / Val(Text39.Text)
End If

End Sub

Private Sub Text39_Change()
If Val(Text38.Text) > 0 And Val(Text39.Text) > 0 Then
    Text40.Text = Val(Text39.Text) / Val(Text38.Text)
    Text41.Text = Val(Label10.Caption) / Val(Text39.Text)
End If

End Sub

Private Sub Text9_Validate(Cancel As Boolean)
If Val(Text9.Text) > 0 Then
    Text8.Text = Blm1.heads(Val(Text9.Text))
    Text7.Text = GetHeadBal(Val(Text9.Text), 1)
End If
End Sub

Private Sub Timer1_Timer()
Dim T As Double
Dim R As Integer

With MSFlexGrid1
    For R = 0 To .Rows - 2
        T = T + Val(.TextMatrix(R, 3))
    Next R
    Label10.Caption = .TextMatrix(0, 3)
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
Text11.Text = T
T = 0
With MSFlexGrid4
    For R = 0 To .Rows - 2
        T = T + Val(.TextMatrix(R, 3))
    Next R
End With
Text37.Text = T
Text28.Text = Text37.Text



Text46.Text = Val(Text39.Text) + Val(Text44.Text)
If (Val(Label10.Caption) + Val(Label11.Caption)) > 0 And Val(Text46.Text) > 0 Then
Text41.Text = (Val(Label10.Caption) + Val(Label11.Caption)) / Val(Text46.Text)
End If

End Sub
