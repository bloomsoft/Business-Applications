VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form vou_1 
   Caption         =   "Daiy Cash Payments and Reciepts"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "&Undo Changes you Made to this Date Before Saving"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   6480
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save and Exit"
      Height          =   375
      Left            =   7920
      TabIndex        =   27
      Top             =   6480
      Width           =   3855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cash Information"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   9135
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   6720
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Cash in Hand"
         Height          =   255
         Left            =   5400
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Cash B/F"
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Info."
      Height          =   2415
      Left            =   9360
      TabIndex        =   20
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "Start &Transactions"
         Height          =   1095
         Left            =   120
         Picture         =   "vou_1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   69337091
         CurrentDate     =   37401
      End
      Begin VB.Label Label1 
         Caption         =   "Select Date"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text7 
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
      Left            =   6240
      TabIndex        =   18
      Top             =   6480
      Width           =   1575
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
      Left            =   4680
      TabIndex        =   17
      Top             =   6480
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid GRID1 
      Height          =   3540
      Left            =   105
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2820
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   6244
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transactions"
      Height          =   1695
      Left            =   105
      TabIndex        =   10
      Top             =   720
      Width           =   9132
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   4830
         TabIndex        =   9
         Top             =   1305
         Width           =   4170
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1065
         TabIndex        =   8
         Top             =   1290
         Width           =   3750
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   225
         TabIndex        =   7
         Top             =   1290
         Width           =   795
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel &Entry"
         Height          =   375
         Left            =   7920
         TabIndex        =   19
         Top             =   240
         Width           =   1092
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   3720
         Top             =   120
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5880
         TabIndex        =   6
         Top             =   720
         Width           =   3132
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblExpCode 
         Height          =   180
         Left            =   8370
         TabIndex        =   33
         Top             =   1050
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label13 
         Caption         =   "Expence Remarks for Invoice"
         Height          =   225
         Left            =   5010
         TabIndex        =   32
         Top             =   1065
         Width           =   3105
      End
      Begin VB.Label Label12 
         Caption         =   "Serial No"
         Height          =   210
         Left            =   225
         TabIndex        =   31
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label11 
         Caption         =   "Clearing Expence Will Be Shown on Invoice"
         Height          =   195
         Left            =   1260
         TabIndex        =   30
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Ledger Remarks"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Credit"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Debit"
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "A/c Code"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   690
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   135
      TabIndex        =   29
      Top             =   2415
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "(F2) to Search A/c"
            TextSave        =   "(F2) to Search A/c"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "(F3) to Open Entries of A Date"
            TextSave        =   "(F3) to Open Entries of A Date"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "(F4) to Select Clearing Case"
            TextSave        =   "(F4) to Select Clearing Case"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4762
            MinWidth        =   4762
            Text            =   "(F5) to Search Clearing Expence"
            TextSave        =   "(F5) to Search Clearing Expence"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "vou_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Sub ClearFull()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Flex1

End Sub
Private Sub edit1()
Dim tb As Recordset
Dim db As Database
Dim I As Integer
Dim ssql As String
Dim CashBF As Currency
Dim PCashBF As Currency

Set db = OpenDatabase(Blm1.pathMain)
ssql = "Select * from Voucher where V_date = #" & Date1.Value & "#"
ssql = ssql & " and e_type=9"
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
GRID1.Rows = 1
    Do While Not tb.EOF
        With GRID1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = tb.Fields("Party").Value
            .TextMatrix(.Rows - 1, 1) = Blm1.party1(tb.Fields("Party").Value)
            .TextMatrix(.Rows - 1, 2) = Format(tb.Fields("Debit").Value, "#.00")
            .TextMatrix(.Rows - 1, 3) = Format(tb.Fields("Credit").Value, "#.00")
            .TextMatrix(.Rows - 1, 4) = tb.Fields("Remarks").Value & ""
            .TextMatrix(.Rows - 1, 5) = tb.Fields("SerialNo").Value & ""
            .TextMatrix(.Rows - 1, 6) = tb.Fields("Expence").Value & ""
            .TextMatrix(.Rows - 1, 7) = tb.Fields("ExpRemarks").Value & ""
            .TextMatrix(.Rows - 1, 8) = tb.Fields("ExpCode").Value & ""
            
        End With
        tb.MoveNext
    Loop
    Blm1.Cash Date1.Value, PCashBF, CashBF
    Label8.Caption = Format(CashBF, "#.00")
    Label10.Caption = Format(CashBF, "#.00")
Else
    MsgBox "No Cash Payments and Reciepts in This Date..."
    GRID1.Rows = 1
    Blm1.Cash Date1.Value, PCashBF, CashBF
    Label8.Caption = Format(CashBF, "#.00")
    Label10.Caption = Format(CashBF, "#.00")
End If
tb.Close
db.Close
Text1.SetFocus
End Sub
Private Sub Save()
Dim tb As Recordset
Dim db As Database
Dim I As Integer
Dim ssql As String

Set db = OpenDatabase(Blm1.pathMain)
ssql = "Delete from Voucher Where v_Date = #" & Date1.Value & "#"
ssql = ssql & " and e_Type=9"
db.Execute ssql

Set tb = db.OpenRecordset("Voucher", dbOpenTable)
For I = 1 To GRID1.Rows - 1
With GRID1
tb.AddNew
        tb.Fields("V_Date").Value = Date1.Value
        tb.Fields("e_type").Value = 9
        tb.Fields("Party").Value = Val(.TextMatrix(I, 0))
        tb.Fields("Debit").Value = Val(.TextMatrix(I, 2))
        tb.Fields("Credit").Value = Val(.TextMatrix(I, 3))
        tb.Fields("Remarks").Value = .TextMatrix(I, 4)
        tb.Fields("SerialNo").Value = Val(.TextMatrix(I, 5))
        tb.Fields("Expence").Value = .TextMatrix(I, 6)
        tb.Fields("ExpRemarks").Value = .TextMatrix(I, 7)
        tb.Fields("ExpCode").Value = Val(.TextMatrix(I, 8))
tb.Update
End With
Next I
tb.Close
ssql = "Delete from Pre_cash where v_date = #" & Date1.Value & "#"
db.Execute ssql
Set tb = db.OpenRecordset("Pre_Cash", dbOpenTable)
tb.AddNew
    tb.Fields("V_date").Value = Date1.Value
    tb.Fields("Opening").Value = Val(Label8.Caption)
    tb.Fields("Closing").Value = Val(Label10.Caption)
tb.Update
tb.Close
db.Close
End Sub
Private Sub Transfer1()
With GRID1
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = Text1.Text
    .TextMatrix(.Rows - 1, 1) = Text2.Text
    .TextMatrix(.Rows - 1, 2) = Text3.Text
    .TextMatrix(.Rows - 1, 3) = Text4.Text
    .TextMatrix(.Rows - 1, 4) = Text5.Text
    .TextMatrix(.Rows - 1, 5) = Text8.Text
    .TextMatrix(.Rows - 1, 6) = Text9.Text
    .TextMatrix(.Rows - 1, 7) = Text10.Text
    .TextMatrix(.Rows - 1, 8) = lblExpCode.Caption
End With
Text1.SetFocus
End Sub
Private Sub Flex1()
With GRID1
    .Rows = 1
    .Cols = 9
    .ColWidth(0) = 1400
    .ColWidth(1) = 1800
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1500
    .ColWidth(5) = 800
    .ColWidth(6) = 1800
    .ColWidth(7) = 1800
    .ColWidth(8) = 0
    .TextMatrix(0, 0) = "Account Code"
    .TextMatrix(0, 1) = "Account Title"
    .TextMatrix(0, 2) = "Debit"
    .TextMatrix(0, 3) = "Credit"
    .TextMatrix(0, 4) = "Ledger Remarks"
    .TextMatrix(0, 5) = "Serial No"
    .TextMatrix(0, 6) = "Clearing Expence"
    .TextMatrix(0, 7) = "Expence Remarks"
    .TextMatrix(0, 8) = "Expence Code"
End With
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
edit1
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Dim I As VbMsgBoxResult

Screen.MousePointer = vbHourglass
Save
ClearFull
Screen.MousePointer = vbDefault
I = MsgBox("Do you Want to Exit...", vbYesNo + vbCritical, "Warning")
If I = vbYes Then
    Me.Hide
    Unload Me
End If
End Sub

Private Sub Command3_Click()
Screen.MousePointer = vbHourglass
ClearFull
edit1
Screen.MousePointer = vbDefault
End Sub

Private Sub Command6_Click()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString

Text1.SetFocus
End Sub

Private Sub Date1_Change()
Command1.Enabled = True
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbHourglass
edit1
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then Date1.SetFocus
If KeyCode = vbKeyF2 Then
    Search2.Text3.Text = 2
    Search2.Show
End If
If KeyCode = vbKeyF4 Then
    SelPartyCode = Val(Text1.Text)
    DocsList.Show vbModal
    Text8.Text = SelSerialNo
End If
If KeyCode = vbKeyF5 Then
    Search1.Text3.Text = 1
    Search1.Show
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'MsgBox Me.ActiveControl.Name
If Me.ActiveControl.Name <> "Text10" Then
SendKeys ("{TAB}")
End If
End If
End Sub

Private Sub Form_Load()
Flex1
Date1.Value = Date
End Sub

Private Sub Grid1_DblClick()
If GRID1.Rows > 1 Then
    With GRID1
        Text1.Text = .TextMatrix(.Row, 0)
        Text2.Text = .TextMatrix(.Row, 1)
        Text3.Text = .TextMatrix(.Row, 2)
        Text4.Text = .TextMatrix(.Row, 3)
        Text5.Text = .TextMatrix(.Row, 4)
        Text8.Text = .TextMatrix(.Row, 5)
        Text9.Text = .TextMatrix(.Row, 6)
        Text10.Text = .TextMatrix(.Row, 7)
        lblExpCode.Caption = .TextMatrix(.Row, 8)
        If .Rows = 2 Then .Rows = 1
        If .Rows > 2 Then .RemoveItem .Row
    End With
End If
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    Text2.Text = Blm1.party1(Val(Text1.Text))
    If Text2.Text = "NOT FOUND" Then
        MsgBox "Invalid A/c Code......"
        Cancel = True
    End If
End If
    
End Sub

Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Transfer1
    Command6_Click
    Text1.SetFocus
End If

End Sub

Private Sub Text3_Change()
If Val(Text3.Text) > 0 Then
    Text4.Enabled = False
Else
    Text4.Enabled = True
End If

End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    Beep
    KeyAscii = 0
End If

End Sub

Private Sub Text4_Change()
If Val(Text4.Text) > 0 Then
    Text3.Enabled = False
Else
    Text3.Enabled = True
End If

End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    Beep
    KeyAscii = 0
End If

End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Text8_GotFocus()
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    Beep
    KeyAscii = 0
End If

End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub

Private Sub Timer1_Timer()
Dim R As Long
Dim TDEB As Currency
Dim TCRED As Currency
'If Grid1.Rows > 1 Then
'    Command2.Enabled = True
'Else
'    Command2.Enabled = False
'End If
Text6.Text = ""
Text7.Text = ""
Label10.Caption = Label8.Caption
For R = 1 To GRID1.Rows - 1
    TDEB = (TDEB + Val(GRID1.TextMatrix(R, 2)))
    TCRED = (TCRED + Val(GRID1.TextMatrix(R, 3)))
Next R
Text6.Text = Format(TDEB, "#.00")
Text7.Text = Format(TCRED, "#.00")
Label10.Caption = Format(Val(Text7.Text) + Val(Label10.Caption) - Val(Text6.Text), "#.00")
End Sub
