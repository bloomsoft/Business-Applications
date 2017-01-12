VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Arrivals 
   Caption         =   "Daily Arrivals"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Arrivals.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Height          =   945
      Left            =   5910
      TabIndex        =   33
      Top             =   0
      Width           =   915
      Begin VB.OptionButton Option3 
         Caption         =   "&Update"
         Height          =   255
         Left            =   30
         TabIndex        =   35
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "&New"
         Height          =   255
         Left            =   30
         TabIndex        =   34
         Top             =   150
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin VB.Frame Frame6 
      Height          =   945
      Left            =   6870
      TabIndex        =   30
      Top             =   0
      Width           =   1665
      Begin VB.CommandButton Command8 
         Caption         =   "&Delete"
         Height          =   705
         Left            =   810
         TabIndex        =   32
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Save as"
         Height          =   705
         Left            =   90
         TabIndex        =   31
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Voucher Info"
      Height          =   1035
      Left            =   2790
      TabIndex        =   27
      Top             =   -30
      Width           =   3075
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Number"
         Height          =   255
         Left            =   210
         TabIndex        =   29
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Ref#"
         Height          =   255
         Left            =   210
         TabIndex        =   28
         Top             =   510
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
      Left            =   9720
      TabIndex        =   20
      Top             =   -30
      Width           =   2040
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   690
         Left            =   1320
         Picture         =   "Arrivals.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   195
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   690
         Left            =   705
         Picture         =   "Arrivals.frx":4F44
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   195
         Width           =   630
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   690
         Left            =   90
         Picture         =   "Arrivals.frx":76E6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   195
         Width           =   630
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4065
      Left            =   240
      TabIndex        =   19
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
      TabIndex        =   16
      Top             =   1080
      Width           =   11535
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   8400
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   9510
         TabIndex        =   7
         Top             =   585
         Width           =   1080
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   5310
         MaxLength       =   150
         TabIndex        =   5
         Top             =   600
         Width           =   3060
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10590
         TabIndex        =   8
         Top             =   570
         Width           =   690
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   8880
         Picture         =   "Arrivals.frx":9E88
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   495
         Left            =   10080
         Picture         =   "Arrivals.frx":C62A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2475
         MaxLength       =   150
         TabIndex        =   4
         Top             =   600
         Width           =   2835
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         MaxLength       =   150
         TabIndex        =   3
         Top             =   600
         Width           =   2250
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Bales"
         Height          =   255
         Left            =   8520
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblAvgBaleWT 
         Caption         =   "..."
         Height          =   255
         Left            =   7365
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Weight"
         Height          =   240
         Left            =   9495
         TabIndex        =   24
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label Label25 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   5295
         TabIndex        =   23
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Freight"
         Height          =   255
         Left            =   10575
         TabIndex        =   21
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "Mill Name"
         Height          =   255
         Left            =   2475
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "PartyName"
         Height          =   255
         Left            =   225
         TabIndex        =   17
         Top             =   225
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voucher Information"
      Height          =   1020
      Left            =   240
      TabIndex        =   14
      Top             =   -15
      Width           =   2475
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67502083
         CurrentDate     =   36757
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Weight"
      Height          =   240
      Left            =   5400
      TabIndex        =   36
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9840
      TabIndex        =   26
      Top             =   6675
      Width           =   960
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10800
      TabIndex        =   22
      Top             =   6720
      Width           =   960
   End
End
Attribute VB_Name = "Arrivals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Private Function max1() As Double
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Ssql = "select max(vno)as c from Arrivals"
    
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

Private Function edit1() As Boolean
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim Tb2 As Recordset
Dim p As Long
Set DB = OpenDatabase(Blm.patHmain)
Ssql = "select * from Arrivals where"
Ssql = Ssql & " VNo=" & Text1.Text
'Ssql = Ssql & " and RefNo=" & Text7.Text
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    date1.Value = TB.Fields("adate").Value
    Text7.Text = TB.Fields("RefNo").Value & ""
            grid1.Rows = 1
            Do While Not TB.EOF
                grid1.Rows = grid1.Rows + 1
                With grid1
                    .TextMatrix(.Rows - 1, 0) = .Rows - 1
                    .TextMatrix(.Rows - 1, 1) = TB.Fields("PartyName").Value & ""
                    .TextMatrix(.Rows - 1, 2) = TB.Fields("MillsName").Value & ""
                    .TextMatrix(.Rows - 1, 3) = TB.Fields("ItemName").Value & ""
                    .TextMatrix(.Rows - 1, 4) = TB.Fields("Bales").Value & ""
                    .TextMatrix(.Rows - 1, 5) = Format(TB.Fields("Weight").Value, "#.00")
                    .TextMatrix(.Rows - 1, 6) = TB.Fields("Freight").Value & ""
                    
                    
                End With
                TB.MoveNext
            Loop
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
flex1
Combs
Text1.Text = max1
Label11.Caption = vbNullString
End Sub

Private Sub transfer1()
grid1.Rows = grid1.Rows + 1
With grid1
    .TextMatrix(.Rows - 1, 0) = .Rows - 1
    .TextMatrix(.Rows - 1, 1) = Text3.Text
    .TextMatrix(.Rows - 1, 2) = Text4.Text
    .TextMatrix(.Rows - 1, 3) = Text15.Text
    .TextMatrix(.Rows - 1, 4) = Val(Text2.Text)
    .TextMatrix(.Rows - 1, 5) = Format(Val(Text9.Text), "#.000")
    .TextMatrix(.Rows - 1, 6) = Format(Text8.Text, "#.00")
    
End With

End Sub
Private Sub flex1()
grid1.Rows = 1
grid1.Cols = 7
With grid1
    .ColWidth(0) = 500
    .TextMatrix(0, 0) = "Sr#"
    .ColWidth(1) = 2000
    .TextMatrix(0, 1) = "Party Name"
    .ColWidth(2) = 2000
    .TextMatrix(0, 2) = "Mill Name"
    .ColWidth(3) = 2000
    .TextMatrix(0, 3) = "Item Name"
    .ColWidth(4) = 1500
    .TextMatrix(0, 4) = "Bales"
    .ColWidth(5) = 1500
    .TextMatrix(0, 5) = "Weight"
    .ColWidth(6) = 1000
    .TextMatrix(0, 6) = "Freight"
    
End With
End Sub
Private Sub Combs()

End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub clear1()
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text8.Text = vbNullString
Text9.Text = vbNullString
Text15.Text = ""
End Sub


Private Sub Command1_Click()
If grid1.Rows > 1 Then
        Call save
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
If Len(Text3.Text) > 0 Then
If Len(Text8.Text) > 0 Or Len(Text9.Text) > 0 Then
Call transfer1
Call clear1
Text3.SetFocus
End If
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
    Ssql = "delete from Arrivals where "
    Ssql = Ssql & " VNo=" & Text1.Text
    
    DB.Execute Ssql
    DB.Close
Set DB = OpenDatabase(Blm.patHmain)

Set TB = DB.OpenRecordset("Arrivals", dbOpenTable)
For I% = 1 To grid1.Rows - 1
TB.AddNew
            TB.Fields("ADate").Value = date1.Value
            TB.Fields("Vno").Value = Val(Text1.Text)
            TB.Fields("RefNo").Value = Val(Text7.Text)
    With grid1
            TB.Fields("PartyName").Value = .TextMatrix(I%, 1)
            TB.Fields("MillsName").Value = .TextMatrix(I%, 2)
            TB.Fields("ItemName").Value = .TextMatrix(I%, 3)
            TB.Fields("Bales").Value = Val(.TextMatrix(I%, 4))
            TB.Fields("Weight").Value = Val(.TextMatrix(I%, 5))
            TB.Fields("Freight").Value = Val(.TextMatrix(I%, 6))
            
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
'If Option2 = True Then
    Set DB = OpenDatabase(Blm.patHmain)
    Ssql = "delete from Arrivals where"
    Ssql = Ssql & " VNo=" & Text1.Text
    'Ssql = Ssql & " and RefNo=" & Text7.Text
    DB.Execute Ssql
    DB.Close
    Command2_Click
'End If

End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub date1_LostFocus()
    If date1.Value >= FStartDate And date1.Value <= FEndDate Then
    '    Text1.Text = max1
'        edit1
    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If
    Text1.Text = max1
End Sub

Private Sub Form_Load()
date1.Value = Date
Combs
Call flex1
Text1.Text = max1
If Screen.Width < 800 And Screen.Height < 600 Then
MsgBox "Please Set your Desktop 800 x 600 Then Try"
Me.Hide
Unload Me
End If

    

End Sub

Private Sub grid1_DblClick()
    With grid1
    Text3.Text = .TextMatrix(.Row, 1)
    Text4.Text = .TextMatrix(.Row, 2)
    Text15.Text = .TextMatrix(.Row, 3)
    Text2.Text = .TextMatrix(.Row, 4)
    Text9.Text = .TextMatrix(.Row, 5)
    Text8.Text = .TextMatrix(.Row, 6)
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
Command6.Visible = False
End Sub

Private Sub Option2_Click()

Command6.Visible = True
Text1.Enabled = True
Text1.SetFocus
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
If Val(Text1.Text) > 0 Then edit1
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


Private Sub Text2_KeyPress(KeyAscii As Integer)
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

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
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
If KeyCode = vbKeyF1 Then Combo4.SetFocus
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
    Text7.Text = Blm.Mill1(Val(Text6.Text))
    If Text4.Text = "NOT" Then
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

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
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
Dim TAmount As Double
Dim TBales As Double
Dim TQty As Double
If grid1.Rows > 1 Then
    For I = 1 To grid1.Rows - 1
        TBales = TBales + Val(grid1.TextMatrix(I, 4))
        TQty = TQty + Val(grid1.TextMatrix(I, 5))
    Next I
    
    
End If

Label3.Caption = TBales
Label11.Caption = TQty
End Sub
