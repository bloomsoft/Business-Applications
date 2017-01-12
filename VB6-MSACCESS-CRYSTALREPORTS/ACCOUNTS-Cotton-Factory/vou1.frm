VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vou1 
   Caption         =   "Voucher Entry"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "vou1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      Height          =   1335
      Left            =   6720
      TabIndex        =   38
      Top             =   390
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   975
         Left            =   120
         Picture         =   "vou1.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   975
         Left            =   1080
         Picture         =   "vou1.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   39
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
      Height          =   1335
      Left            =   8970
      TabIndex        =   29
      Top             =   390
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   975
         Left            =   1800
         Picture         =   "vou1.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   975
         Left            =   960
         Picture         =   "vou1.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   975
         Left            =   120
         Picture         =   "vou1.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   28
      Top             =   3630
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction"
      Height          =   1815
      Left            =   240
      TabIndex        =   22
      Top             =   1740
      Width           =   11535
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Format          =   75038721
         CurrentDate     =   37015
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Clear"
         Height          =   735
         Left            =   8880
         Picture         =   "vou1.frx":EDCC
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   735
         Left            =   10080
         Picture         =   "vou1.frx":1156E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   9360
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5160
         MaxLength       =   255
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   ".."
         Height          =   255
         Left            =   6600
         TabIndex        =   43
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Current Balance"
         Height          =   255
         Left            =   6600
         TabIndex        =   42
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Payment Days"
         Height          =   255
         Left            =   5280
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Kachi #"
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Kachi Date"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Credit"
         Height          =   255
         Left            =   10200
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Debit"
         Height          =   255
         Left            =   8880
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Narration"
         Height          =   255
         Left            =   5400
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voucher Information"
      Height          =   1755
      Left            =   240
      TabIndex        =   17
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton Command7 
         Caption         =   "Save This Voucher As"
         Height          =   375
         Left            =   3360
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Delete This Voucher"
         Height          =   375
         Left            =   3360
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "vou1.frx":13D10
         Left            =   3720
         List            =   "vou1.frx":13D12
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1350
         TabIndex        =   2
         Top             =   810
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   75038723
         CurrentDate     =   36757
      End
      Begin VB.Label Label17 
         Caption         =   "Ref #"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Narration"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher #"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   330
         TabIndex        =   18
         Top             =   810
         Width           =   855
      End
   End
   Begin VB.Label Label16 
      Caption         =   "When in A/c Code Press (F1) to Select Accounts from List"
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   7200
      Width           =   3735
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9840
      TabIndex        =   31
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8400
      TabIndex        =   30
      Top             =   7200
      Width           =   1335
   End
End
Attribute VB_Name = "vou1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Dim CurRemarks As String
Private Sub CurrentBalance(AcCode As Long)
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "select Sum(Debit - Credit) as Bal from Voudtl where Party = " & AcCode
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    If TB.Fields("Bal").Value > 0 Then
        Label19.Caption = Format(TB.Fields("Bal").Value, "#.00") & " DR"
    ElseIf TB.Fields("Bal").Value < 0 Then
        Label19.Caption = Format(TB.Fields("Bal").Value, "#.00") & " CR"
    End If
Else
    Label19.Caption = "...."
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

Set DB = OpenDatabase(Blm.patHmain)
Ssql = "select * from voumst where v_type = " & Combo1.ItemData(Combo1.ListIndex)
Ssql = Ssql & " and v_no = " & Val(Text1.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    date1.Value = TB.Fields("v_date").Value
    Text2.Text = TB.Fields("narration").Value
        Ssql = "select * from voudtl where v_type = " & Combo1.ItemData(Combo1.ListIndex)
        Ssql = Ssql & " and v_no = " & Val(Text1.Text)
        Set Tb2 = DB.OpenRecordset(Ssql)
        grid1.Rows = 1
        If Not Tb2.EOF Then
            Do While Not Tb2.EOF
                grid1.Rows = grid1.Rows + 1
                Text10.Text = Tb2.Fields("ICT").Value & ""
                With grid1
                    .TextMatrix(.Rows - 1, 0) = Tb2.Fields("party").Value
                    .TextMatrix(.Rows - 1, 1) = Blm.party1(Tb2.Fields("party").Value)
                    .TextMatrix(.Rows - 1, 2) = Tb2.Fields("remarks").Value & ""
                    .TextMatrix(.Rows - 1, 3) = Tb2.Fields("debit").Value
                    .TextMatrix(.Rows - 1, 4) = Tb2.Fields("credit").Value
                    .TextMatrix(.Rows - 1, 5) = Tb2.Fields("Pakki_no").Value
                    If Not IsNull(Tb2.Fields("p_Date").Value) Then
                    .TextMatrix(.Rows - 1, 6) = Tb2.Fields("p_Date").Value
                    .TextMatrix(.Rows - 1, 7) = Tb2.Fields("Days").Value
                    End If
                End With
                Tb2.MoveNext
            Loop
        End If
        Tb2.Close
        edit1 = True
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

Label11.Caption = vbNullString
Label12.Caption = vbNullString
End Sub

Private Sub transfer1()
grid1.Rows = grid1.Rows + 1
With grid1
    .TextMatrix(.Rows - 1, 0) = Val(Text3.Text)
    .TextMatrix(.Rows - 1, 1) = Text4.Text
    .TextMatrix(.Rows - 1, 2) = Text5.Text
    CurRemarks = Text5.Text

    .TextMatrix(.Rows - 1, 3) = Val(Text6.Text)
    .TextMatrix(.Rows - 1, 4) = Val(Text7.Text)
    .TextMatrix(.Rows - 1, 5) = Val(Text8.Text)
    .TextMatrix(.Rows - 1, 6) = DTPicker1.Value
    .TextMatrix(.Rows - 1, 7) = Val(Text9.Text)
End With
End Sub
Private Sub flex1()
grid1.Rows = 1
grid1.Cols = 8
With grid1
    .ColWidth(0) = 1500
    .TextMatrix(0, 0) = "A/c Code"
    .ColWidth(1) = 3000
    .TextMatrix(0, 1) = "A/c Title"
    .ColWidth(2) = 3500
    .TextMatrix(0, 2) = "Remarks"
    .ColWidth(3) = 1700
    .TextMatrix(0, 3) = "Debit"
    .ColWidth(4) = 1700
    .TextMatrix(0, 4) = "Credit"
    .ColWidth(5) = 0
    .TextMatrix(0, 5) = "Kachi #"
    .ColWidth(6) = 0
    .TextMatrix(0, 6) = "Kachi Date"
    .ColWidth(7) = 0
    .TextMatrix(0, 7) = "Payment"
End With
End Sub
Private Sub Combs()

End Sub
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(v_no)as c from voumst where v_type = " & Combo1.ItemData(Combo1.ListIndex)
    
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
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Combo1_LostFocus()
If Option1 = True Then
    Text1.Text = max1
End If
End Sub

Private Sub clear1()
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = vbNullString
Text9.Text = vbNullString
End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
Text3.Text = Combo2.ItemData(Combo2.ListIndex)
Text4.Text = Combo2.Text
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus

End Sub

Private Sub Command1_Click()
Dim Diff As Currency
If Val(Text1.Text) <= 0 Then
    MsgBox "Please Complete This Voucher"
    Exit Sub
End If
'MsgBox "Test"
If Val(Label11.Caption) > 0 And Val(Label12.Caption) > 0 Then
    If Val(Label11.Caption) = Val(Label12.Caption) Then
        Call save
        Command2_Click
    Else
        If Combo1.ListIndex = 2 Then
        If grid1.Rows > 1 Then
            Diff = Val(Label11.Caption) - Val(Label12.Caption)
            'MsgBox Diff
            If Diff <> 0 Then
                
                With grid1
                    
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = 160010001
                        .TextMatrix(.Rows - 1, 1) = Blm.party1(160010001)
                        .TextMatrix(.Rows - 1, 2) = CurRemarks
                        
                        If Diff > 0 Then
                        .TextMatrix(.Rows - 1, 3) = Format(0, "#.00")
                        .TextMatrix(.Rows - 1, 4) = Format(Diff, "#.00")
                        Else
                        .TextMatrix(.Rows - 1, 3) = Format(Diff * -1, "#.00")
                        .TextMatrix(.Rows - 1, 4) = Format(0, "#.00")
                        End If
                End With
                     Call save
                     Command2_Click
             End If
        End If
        Else
            MsgBox "Please Make Sure Debit & Credit Totals are Equal"
        End If
    End If
Else
    If Combo1.ListIndex = 2 Then
        If grid1.Rows > 1 Then
            Diff = Val(Label11.Caption) - Val(Label12.Caption)
            'MsgBox Diff
            If Diff <> 0 Then
                
                With grid1
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = 160010001
                        .TextMatrix(.Rows - 1, 1) = Blm.party1(160010001)
                        .TextMatrix(.Rows - 1, 2) = CurRemarks
                        If Diff > 0 Then
                        .TextMatrix(.Rows - 1, 3) = Format(0, "#.00")
                        .TextMatrix(.Rows - 1, 4) = Format(Diff, "#.00")
                        Else
                        .TextMatrix(.Rows - 1, 3) = Format(Diff * -1, "#.00")
                        .TextMatrix(.Rows - 1, 4) = Format(0, "#.00")
                        End If
                End With
                     Call save
                     Command2_Click
             End If
        End If
     End If
End If
End Sub

Private Sub Command2_Click()
Call clearfull
Combo1.SetFocus
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
    Ssql = Ssql & " v_type = " & Combo1.ItemData(Combo1.ListIndex)
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where "
    Ssql = Ssql & " v_type = " & Combo1.ItemData(Combo1.ListIndex)
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    DB.Close
End If
Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset("voumst", dbOpenTable)
TB.AddNew
    TB.Fields("v_date").Value = date1.Value
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = Combo1.ItemData(Combo1.ListIndex)
    TB.Fields("narration").Value = Text2.Text
TB.Update
TB.Close
Set TB = DB.OpenRecordset("voudtl", dbOpenTable)
For I% = 1 To grid1.Rows - 1
With grid1
TB.AddNew
    TB.Fields("v_date").Value = date1.Value
    TB.Fields("ICT").Value = Text10.Text
    TB.Fields("v_no").Value = Val(Text1.Text)
    TB.Fields("v_type").Value = Combo1.ItemData(Combo1.ListIndex)
    TB.Fields("party").Value = Val(.TextMatrix(I%, 0))
    TB.Fields("remarks").Value = .TextMatrix(I%, 2)
    TB.Fields("debit").Value = Val(.TextMatrix(I%, 3))
    'MsgBox tb.Fields("Debit").Value
    TB.Fields("credit").Value = Val(.TextMatrix(I%, 4))
    TB.Fields("Pakki_no").Value = Val(.TextMatrix(I%, 5))
    If Len(.TextMatrix(I%, 6)) > 0 Then
    TB.Fields("P_Date").Value = CDate(.TextMatrix(I%, 6))
    TB.Fields("Days").Value = Val(.TextMatrix(I%, 7))
    End If
TB.Update
End With
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
    Ssql = Ssql & " v_type = " & Combo1.ItemData(Combo1.ListIndex)
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
    DB.Execute Ssql
    Ssql = "delete from voudtl where "
    Ssql = Ssql & " v_type = " & Combo1.ItemData(Combo1.ListIndex)
    Ssql = Ssql & " and v_no = " & Val(Text1.Text)
DB.Execute Ssql
DB.Close
Command2_Click
End If
End Sub

Private Sub Command7_Click()
Text1.Text = max1
Call save
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

Blm.vouchercomb Combo1
Call flex1
'If Screen.Width = 800 And Screen.Height = 600 Then Me.WindowState = 2
'If Screen.Width > 800 And Screen.Height > 600 Then
'    Me.Height = 8085
'    Me.Width = 12060
'End If
'If Screen.Width < 800 And Screen.Height < 600 Then
'MsgBox "Please Set your Desktop 800 x 600 Then Try"
'Me.Hide
'Unload Me
'End If
End Sub

Private Sub grid1_Click()
If grid1.Row > 0 Then
    Text5.Text = grid1.TextMatrix(grid1.Row, 2)
End If
End Sub

Private Sub grid1_DblClick()
If grid1.Rows > 2 Then
    With grid1
        Text3.Text = .TextMatrix(.Row, 0)
        Text4.Text = .TextMatrix(.Row, 1)
        Text5.Text = .TextMatrix(.Row, 2)
        Text6.Text = .TextMatrix(.Row, 3)
        Text7.Text = .TextMatrix(.Row, 4)
        Text8.Text = .TextMatrix(.Row, 5)
        If IsDate(.TextMatrix(.Row, 6)) Then
        DTPicker1.Value = .TextMatrix(.Row, 6)
        End If
        Text9.Text = .TextMatrix(.Row, 7)
    End With
    grid1.RemoveItem (grid1.Row)
Else
If grid1.Rows = 2 Then
    With grid1
        Text3.Text = .TextMatrix(.Row, 0)
        Text4.Text = .TextMatrix(.Row, 1)
        Text5.Text = .TextMatrix(.Row, 2)
        Text6.Text = .TextMatrix(.Row, 3)
        Text7.Text = .TextMatrix(.Row, 4)
        Text8.Text = .TextMatrix(.Row, 5)
        DTPicker1.Value = .TextMatrix(.Row, 6)
        Text9.Text = .TextMatrix(.Row, 7)
    End With
    grid1.Rows = 1
End If
End If
End Sub

Private Sub Option1_Click()
Command2_Click
Text1.Enabled = False
Combo1.SetFocus
Command7.Visible = False
Command6.Visible = False
End Sub

Private Sub Option2_Click()
Command2_Click
Combo1.SetFocus
Command6.Visible = True
Command7.Visible = True
Text1.Enabled = True
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
        If B = False Then
            Cancel = True
            Text1.Text = vbNullString
        End If
End If

End Sub

Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Show vbModal
    Text3.Text = SelectedAccountCode
    Text4.Text = SelectedAccountName
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
    Label19.Caption = "...."
    Text4.Text = Blm.party1(Val(Text3.Text))
    If Text4.Text = "NOT" Then
        Cancel = True
    Else
        CurrentBalance Val(Text3.Text)
        
    End If
        
End If

End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text6_Change()
If Val(Text6.Text) > 0 Then
    Text7.Enabled = False
Else
    Text7.Enabled = True
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

Private Sub Text7_Change()
If Val(Text7.Text) > 0 Then
    If Val(Text6.Text) > 0 Then
        Text7.Text = 0
        Text7.Enabled = False
        Text6.SetFocus
    Else
        Text6.Enabled = False
    End If
Else
    Text6.Enabled = True
End If
    
    
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

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

Private Sub Timer1_Timer()
Dim I As Long
Dim deb As Currency
Dim cred As Currency

If grid1.Rows > 1 Then
    For I = 1 To grid1.Rows - 1
        deb = deb + Val(grid1.TextMatrix(I, 3))
        cred = cred + Val(grid1.TextMatrix(I, 4))
    Next I
    Label11.Caption = deb
    Label12.Caption = cred
Command6.Enabled = True
Else
Command6.Enabled = False
End If
End Sub
