VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form vou_Dying 
   Caption         =   "Dying Payment Voucher"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   10095
      Begin VB.CommandButton Command3 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7200
         Picture         =   "vou_Dying.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4200
         Picture         =   "vou_Dying.frx":09A6
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1320
         Picture         =   "vou_Dying.frx":1290
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Voucher Entries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   10095
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   2520
         Width           =   8415
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Top             =   1080
         Width           =   5175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Discription"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Quality Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Quality Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Factory Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Factory Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Voucher Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   67239937
         CurrentDate     =   39498
      End
      Begin VB.Label Label2 
         Caption         =   "Reciept No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2160
         Picture         =   "vou_Dying.frx":1B93
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         Picture         =   "vou_Dying.frx":25AA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "vou_Dying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As Recordset
Dim db As Database
Private blm As New bloom1

Private Sub Clear1()
    
Text2.Text = vbNullString
Text4.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString

End Sub


Private Sub Combs()
Dim ssql As String

'Factory
ssql = "select * from FactoryChart order by Name"
blm.Factory ssql, Combo2, "Name", "Code"
'cloth Quality
ssql = "select * from Cloths order by Name"
blm.FillCloth1 ssql, Combo3, "Name", "Code"
'Dying
ssql = "select * from DyingChart order by Name"
blm.Dying ssql, Combo1, "Name", "Code"

End Sub

Private Function Edit1() As Boolean
Dim ssql As String
Dim tb As Recordset
Dim R As Long
Dim B As Boolean


'MsgBox Tb.Fields("YType").Value
'For R = 0 To Combo1.ListCount - 1
'    If Combo1.ItemData(R) = tb.Fields("Ytype").Value Then
''        MsgBox Tb.Fields("YType").Value
'        Combo1.ListIndex = R
'        Exit For
'End If
'Next R
'Combo1_Validate False
'Text14.Text = tb.Fields("DoNo").Value & ""
'Text15.Text = tb.Fields("PurchaseParty").Value & ""
'Edit1 = False
'Else
'MsgBox "Invalid Issue No."
'Edit1 = True
'End If
'tb.Close
'ssql = "Select * from YarnIssue where Issue_no = " & Val(Text1.Text)
'Set tb = db.OpenRecordset(ssql)
'If Not tb.EOF Then
'DTPicker1.Value = tb.Fields("Issue_Date").Value
''Text1.Text = Tb.Fields("Cont_no").Value
''For R = 0 To Combo4.ListCount - 1
''    If Combo4.ItemData(R) = Tb.Fields("Ctype").Value Then
''        Combo4.ListIndex = R
''        Exit For
'End If
'Next R
'B = ContractInfo()
Set db = OpenDatabase(blm.pathMain)
ssql = "select * from CLOTHREC where REC_NO = " & Val(Text1.Text)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    DTPicker1.Value = tb.Fields("Date").Value
    Text2.Text = tb.Fields("FAC_CODE").Value
    Text3.Text = blm.Factory(tb.Fields("Code").Value)
    Text4.Text = tb.Fields("Cloth_CODE").Value
    Text5.Text = blm.Dying(tb.Fields("Code").Value)
    Text6.Text = tb.Fields("AMOUNT").Value
    Text7.Text = tb.Fields("REMARKS").Value
    Edit1 = False
Else
    MsgBox "No Record For This VOUCHER NO."
    Edit1 = True
    Exit Function
End If
tb.Close

End Function

Private Sub Save()
Dim ssql As String
If Option2 = True Then
    ssql = "Delete from PaymentLoom WHere Vou_NO = " & Val(Text1.Text)
    db.Execute (ssql)

End If
    Rs.AddNew
        Rs.Fields("Date").Value = DTPicker1.Value
        Rs.Fields("Vou_No").Value = Val(Text1.Text)
        Rs.Fields("FAC_CODE").Value = Val(Text2.Text)
        Rs.Fields("CLOTH_CODE").Value = Val(Text4.Text) 'Combo4.ItemData(Combo4.ListIndex)
        Rs.Fields("AMOUNT").Value = Text6.Text
        Rs.Fields("REMARKS").Value = Text7.Text
    Rs.Update
    Rs.Close
Set Rs = db.OpenRecordset("PaymentLoom", dbOpenDynaset)

End Sub

Private Function Max1() As Double
Dim ssql As String
Dim tb As Recordset

ssql = "Select Max(Vou_No) as C from PaymentLoom"
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    Max1 = tb.Fields("C").Value + 1
Else
    Max1 = 1
End If
tb.Close
End Function

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Save
Command2_Click
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Clear1
DTPicker1.Value = Date
If Option1 = True Then
Text1.Text = Max1
Text2.SetFocus
Else
Text1.Enabled = True
Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(blm.pathMain)
Set Rs = db.OpenRecordset("ClothRec", dbOpenDynaset)
'Combo4.ListIndex = 0
Text1.Text = Max1
'Combo1.ListIndex = 0
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Text1.Text = Max1
Text2.SetFocus
End Sub

Private Sub Option2_Click()
Command2_Click
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Option2 = True Then
If Val(Text1.Text) > 0 Then
    Edit1
End If
End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search3.Text3.Text = 3
        Search3.Show
    End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search1.Text3.Text = 3
        Search1.Show
    End If

End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF1 Then
'        Search2.Text3.Text = 1
'        Search2.Show
'    End If
'
End Sub
