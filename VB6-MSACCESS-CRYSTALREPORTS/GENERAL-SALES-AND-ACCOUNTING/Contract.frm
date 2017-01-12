VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form s 
   Caption         =   "Cloth  Purchase Contract"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   7215
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Caption         =   "Note"
      Height          =   2055
      Left            =   6795
      TabIndex        =   56
      Top             =   4920
      Width           =   4980
      Begin VB.TextBox txtNote 
         Height          =   1740
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   4830
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2055
      Left            =   7080
      TabIndex        =   51
      Top             =   8130
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command9 
         Caption         =   "> |"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   28
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   ">"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   27
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "<"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   26
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "| <"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1095
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2055
      Left            =   120
      TabIndex        =   50
      Top             =   4920
      Width           =   6615
      Begin VB.CommandButton Command5 
         Caption         =   "&Print"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   4560
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   735
         Left            =   600
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   4560
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   855
         Left            =   600
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   9360
      TabIndex        =   47
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   49
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   48
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   36
      Top             =   2040
      Width           =   11655
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4920
         TabIndex        =   59
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Contract.frx":0000
         Left            =   10920
         List            =   "Contract.frx":000A
         TabIndex        =   13
         Text            =   "Yes"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   8280
         TabIndex        =   24
         Top             =   2280
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Contract.frx":0017
         Left            =   10560
         List            =   "Contract.frx":0024
         TabIndex        =   20
         Text            =   "Power"
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Text            =   "Pick"
         Top             =   840
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   10080
         Top             =   1680
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   22
         Text            =   "80% 80 Mtrs and Up, 20% 37 Mtrs to 80 Mtrs."
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   8280
         TabIndex        =   18
         Top             =   1800
         Width           =   615
      End
      Begin MSComCtl2.DTPicker Date2 
         Height          =   375
         Left            =   5040
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   37502
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   14
         Text            =   "70% Kachi Purchi, 30% At the End of the Contract"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text19 
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
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   9
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Rate of Warp"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Invoice Mode"
         Height          =   255
         Left            =   9240
         TabIndex        =   55
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Delievery Plan"
         Height          =   255
         Left            =   6720
         TabIndex        =   53
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Loom Type"
         Height          =   255
         Left            =   9240
         TabIndex        =   52
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Piece Length"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Brokery Rate"
         Height          =   255
         Left            =   6720
         TabIndex        =   45
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Del. Date"
         Height          =   255
         Left            =   3720
         TabIndex        =   44
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Cloth  Rate/ Mtr."
         Height          =   255
         Left            =   6720
         TabIndex        =   42
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Rate of Weft"
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Labour / "
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Gazana"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Cloth Quality"
         Height          =   255
         Left            =   3720
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Cloth Code"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   90
      Width           =   9135
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   1440
         TabIndex        =   60
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39394
      End
      Begin MSComCtl2.DTPicker Date3 
         Height          =   375
         Left            =   7680
         TabIndex        =   1
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   37528
      End
      Begin VB.TextBox Text5 
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
         Left            =   5040
         TabIndex        =   5
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text3 
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
         Left            =   5040
         TabIndex        =   3
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5040
         TabIndex        =   0
         Top             =   375
         Width           =   1425
      End
      Begin VB.Label Label30 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   6720
         TabIndex        =   54
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "BrokerName"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Broker Code"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Party Name"
         Height          =   255
         Left            =   3720
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Party Code"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Contract Date"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Dim db As Database
Dim Rs As Recordset
Private Function edit1() As Boolean
Dim ssql As String
'Dim P As ClothInfo
Dim R As Integer
Dim RST As Recordset
ssql = "Select * from PContract where Cont_no = " & Val(Text1.Text)
'ssql = ssql & " and CType=" & Combo4.ItemData(Combo4.ListIndex)
Set RST = db.OpenRecordset(ssql)
If Not RST.EOF Then
    txtNote.Text = RST.Fields("Note").Value & ""
    Date1.Value = RST.Fields("Cont_Date").Value
    Text2.Text = RST.Fields("SellerCode").Value
    'Text3.Text = Blm1.SellerName(Val(Text2.Text))
    Text3.Text = Blm1.party1(Val(Text2.Text))
    Text4.Text = RST.Fields("BrokerCode").Value
    Text5.Text = Blm1.broker1(Val(Text4.Text))
    Text6.Text = RST.Fields("ClothCode").Value
    Text7.Text = Blm1.Item1(Val(Text6.Text))
    
    Text13.Text = RST.Fields("Quantity").Value
    Text14.Text = RST.Fields("Labour").Value
    Text17.Text = RST.Fields("WarpRate").Value
    Text18.Text = RST.Fields("WeftRate").Value
    Text19.Text = RST.Fields("ClothRate").Value
    Text21.Text = RST.Fields("CRate").Value
    Text20.Text = RST.Fields("Payment").Value & ""
    Text22.Text = RST.Fields("Terms").Value
    Date2.Value = RST.Fields("DelDate").Value
'    Text25.Text = RST.Fields("WarpBag").Value
'    Text26.Text = RST.Fields("WeftBag").Value
    
    Date3.Value = RST.Fields("ExpDate").Value
    Combo2.Text = RST.Fields("Loom").Value
    For R = 0 To Combo3.ListCount - 1
    If Combo3.ItemData(R) = RST.Fields("InvMode").Value Then
     Combo3.ListIndex = R
     Exit For
    End If
    Next R
    Text27.Text = RST.Fields("DelPlan").Value & ""
    For R = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(R) = RST.Fields("LPer").Value Then
     Combo1.ListIndex = R
     Exit For
    End If
    Next R
    
    edit1 = False
    
Else
    MsgBox "Invalid  Contract No."
    edit1 = True
End If
RST.Close

End Function
Private Sub save()
Dim ssql As String
If Option2 = True Then
    ssql = "Delete from PContract Where Cont_no = " & Val(Text1.Text)
'    ssql = ssql & " and CType=" & Combo4.ItemData(Combo4.ListIndex)
    db.Execute ssql
End If
Rs.AddNew
    Rs.Fields("Note").Value = txtNote.Text
'    Rs.Fields("Ctype").Value = Combo4.ItemData(Combo4.ListIndex)
    Rs.Fields("Cont_Date").Value = Date1.Value
    Rs.Fields("Cont_no").Value = Val(Text1.Text)
    Rs.Fields("SellerCode").Value = Val(Text2.Text)
    Rs.Fields("BrokerCode").Value = Val(Text4.Text)
    Rs.Fields("ClothCode").Value = Val(Text6.Text)
    Rs.Fields("Quantity").Value = Val(Text13.Text)
    Rs.Fields("Labour").Value = Val(Text14.Text)
    Rs.Fields("WarpRate").Value = Val(Text17.Text)
    Rs.Fields("WeftRate").Value = Val(Text18.Text)
    Rs.Fields("ClothRate").Value = Val(Text19.Text)
    Rs.Fields("CRate").Value = Val(Text21.Text)
    Rs.Fields("Payment").Value = Text20.Text
    Rs.Fields("Terms").Value = Text22.Text
'    Rs.Fields("WarpBags").Value = Val(Text23.Text)
'    Rs.Fields("WeftBags").Value = Val(Text24.Text)
    Rs.Fields("DelDate").Value = Date2.Value
'    Rs.Fields("WarpBag").Value = Val(Text25.Text)
'    Rs.Fields("WeftBag").Value = Val(Text26.Text)
'    Rs.Fields("WarpWt").Value = Val(Text15.Text) / 40
'    Rs.Fields("WeftWt").Value = Val(Text16.Text) / 40
    Rs.Fields("ExpDate").Value = Date3.Value
    Rs.Fields("Loom").Value = Combo2.Text
    Rs.Fields("InvMode").Value = Combo3.ItemData(Combo3.ListIndex)
    Rs.Fields("DelPlan").Value = Text27.Text
'    Rs.Fields("LPer").Value = Combo1.ItemData(Combo1.ListIndex)
    
Rs.Update
Rs.Close
Set Rs = db.OpenRecordset("PContract", dbOpenDynaset)

End Sub
Private Function max1() As Long
Dim RST As Recordset
Dim ssql As String

ssql = "Select Max(Cont_no) as C from PContract"
Set RST = db.OpenRecordset(ssql)
If Not IsNull(RST.Fields("C").Value) Then
    max1 = RST.Fields("C").Value + 1
Else
    max1 = 1
End If
RST.Close

End Function

Private Sub Combo4_LostFocus()
Text1.Text = max1
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
save
Command2_Click
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Dim Cntl As Control
For Each Cntl In Me.Controls
    If TypeOf Cntl Is TextBox Then Cntl.Text = ""
    If TypeOf Cntl Is DTPicker Then Cntl.Value = Date
    
Next
Text20.Text = "70% Kachi Purchi, 30% At the End of the Contract"
Text22.Text = "80% 80 Mtrs and Up, 20% 37 Mtrs to 80 Mtrs."
Text1.Text = max1
Date1.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me

End Sub

Private Sub Command4_Click()
If Option2 = True Then
    ssql = "Delete from PContract Where Cont_no = " & Val(Text1.Text)
    ssql = ssql & " and Ctype=" & Combo4.ItemData(Combo4.ListIndex)
    db.Execute ssql
End If
Command2_Click
End Sub

Private Sub Command5_Click()
Load notes
notes.Caption = "Cloth  Contract"
notes.Text3.Text = 1
notes.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
'If TypeOf Me.ActiveControl Is DTPicker Then MsgBox KeyAscii
 Do While i < 100000

        i = i + 1
Loop
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(Blm1.pathMain)
Set Rs = db.OpenRecordset("PContract", dbOpenDynaset)
'Combo4.ListIndex = 0
Text1.Text = max1
Date1.Value = Date
Date2.Value = Date
Combo3.ListIndex = 0
'Combo1.ListIndex = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Rs.Close
db.Close
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
    Text2.Text = List1.ItemData(List1.ListIndex)
    Text3.Text = List1.Text
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub List2_Click()
If List2.ListIndex > -1 Then
    Text4.Text = List2.ItemData(List2.ListIndex)
    Text5.Text = List2.Text
End If
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
End Sub

Private Sub List2_LostFocus()
List2.Visible = False
End Sub

Private Sub List3_Click()
'Dim P As ClothInfo
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text6.SetFocus
End Sub

Private Sub List3_LostFocus()
List3.Visible = False
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Date1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Text1.SetFocus
'Text1.Enabled = True
'Combo4.SetFocus
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
'Text1.SetFocus
If Option2 = True Then
    If Val(Text1.Text) > 0 Then
        Cancel = edit1
    Else
        MsgBox "Please Give  Contract No."
        Cancel = True
    End If
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF1 Then
            Load Search2
            Search2.Text3.Text = 9
            Search2.Show vbModal
        End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text3.Text = Blm1.party1(Val(Text2.Text))
    If Text3.Text = "Wrong" Then
        MsgBox "Invalid Party Code..."
        Cancel = True
    End If
Else
    Cancel = True
End If
End Sub

Private Sub Text20_GotFocus()
Text20.SelStart = 0
Text20.SelLength = Len(Text20.Text)
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text22_GotFocus()
Text22.SelStart = 0
Text22.SelLength = Len(Text22.Text)

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF1 Then
            Load Search3
            Search3.Text3.Text = 1
            Search3.Show vbModal
        End If

End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text5.Text = Blm1.broker1(Val(Text4.Text))
    If Text5.Text = "Wrong" Then
        MsgBox "Invalid Broker Code..."
'        Cancel = True
    End If
'Else
 '   Cancel = True
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
Load Search1
Search1.Text3.Text = 10
Search1.Show vbModal
End If

End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Val(Text6.Text) > 0 Then
    Text7.Text = Blm1.Item1(Val(Text6.Text))
    If Text7.Text = "Wrong" Then
        MsgBox "Invalid Item Code..."
'        Cancel = True
    End If
'Else
 '   Cancel = True
End If
End Sub

