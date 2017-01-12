VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form voudyingpay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dying Payment Voucher"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10425
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
      TabIndex        =   24
      Top             =   4680
      Width           =   10095
      Begin VB.CommandButton Command3 
         Caption         =   "E&xit"
         CausesValidation=   0   'False
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
         Picture         =   "voudyingpay.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
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
         Picture         =   "voudyingpay.frx":09A6
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Picture         =   "voudyingpay.frx":1290
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   270
      TabIndex        =   19
      Top             =   1770
      Width           =   10095
      Begin VB.TextBox lblDispute 
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3945
         TabIndex        =   38
         Top             =   1725
         Width           =   2940
      End
      Begin VB.TextBox Text3 
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
         Left            =   8280
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3240
         Top             =   2280
      End
      Begin VB.Frame Frame5 
         Caption         =   "Kora Detail"
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
         Left            =   7320
         TabIndex        =   30
         Top             =   840
         Width           =   2655
         Begin VB.TextBox Text11 
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox Text9 
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text8 
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Label14"
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
            Left            =   1920
            TabIndex        =   35
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "%age"
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
            TabIndex        =   33
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Gazana"
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
            TabIndex        =   32
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Total Than"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   735
         End
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
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text14 
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text10 
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
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000F&
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
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   3735
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
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
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
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
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
         Height          =   345
         Left            =   4665
         TabIndex        =   5
         Top             =   2445
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Dispute"
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
         Left            =   2850
         TabIndex        =   39
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         Left            =   3960
         TabIndex        =   29
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblError 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2895
         TabIndex        =   37
         Top             =   1200
         Width           =   3960
      End
      Begin VB.Label Label4 
         Caption         =   "Lot No."
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
         Left            =   7200
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Dying Gazana"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Rate"
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
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Dying Code"
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
         Left            =   240
         TabIndex        =   22
         Top             =   360
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
         Left            =   240
         TabIndex        =   21
         Top             =   840
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
         Left            =   2835
         TabIndex        =   20
         Top             =   2445
         Visible         =   0   'False
         Width           =   1455
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
      Height          =   1695
      Left            =   4200
      TabIndex        =   16
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
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
         Left            =   4080
         Picture         =   "voudyingpay.frx":1B93
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
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
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1440
         TabIndex        =   28
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   20119555
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
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   975
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
         Left            =   240
         TabIndex        =   17
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
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2160
         Picture         =   "voudyingpay.frx":2531
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   315
         Picture         =   "voudyingpay.frx":2F48
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   1215
      End
   End
End
Attribute VB_Name = "voudyingpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private Function CheckLotNo() As Boolean
'Dim Rs As Recordset
Dim DB As Database
Dim tb As Recordset
Dim R As Long
Dim b As Boolean
Dim Ssql As String

Set DB = OpenDatabase(blm.pathMain)
    Ssql = "select * from PaymentDying where Lot_NO = " & Val(Text3.Text)
    Set tb = DB.OpenRecordset(Ssql)
        If Not tb.EOF Then
            MsgBox "This Lot No Already Added"
            CheckLotNo = True
        End If
    tb.Close
DB.Close
End Function
Private Sub Clear1()
lblError.Caption = ""
lblDispute.Text = ""
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text10.Text = vbNullString
Text14.Text = vbNullString
Text8.Text = vbNullString
Text9.Text = vbNullString
Text11.Text = vbNullString
Label14.Caption = vbNullString

If Option1 = True Then
    Text1.Enabled = False
    Text6.SetFocus
Else
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub


Private Sub Combs()
Dim Ssql As String

''Factory
'Ssql = "select * from FactoryChart order by Name"
'Blm.Factory Ssql, Combo2, "Name", "Code"
''cloth Quality
'Ssql = "select * from Cloths order by Name"
'Blm.FillCloth1 Ssql, Combo3, "Name", "Code"
''Dying
'Ssql = "select * from DyingChart order by Name"
'Blm.Dying Ssql, Combo1, "Name", "Code"

End Sub

Private Function edit1() As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from PaymentDying where Vou_NO = " & Val(Text1.Text)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    DTPicker1.Value = tb.Fields("Date").Value
    Text3.Text = tb.Fields("LOT_NO") & ""
    Text4.Text = tb.Fields("Cloth_CODE").Value
    Text5.Text = blm.FillCloth1(tb.Fields("Cloth_Code").Value)
    Text6.Text = tb.Fields("DYING_CODE").Value
    Text7.Text = blm.Dying(tb.Fields("DYING_Code").Value)
    Text2.Text = tb.Fields("GAZANA").Value
    Text10.Text = tb.Fields("RATE").Value
    Text14.Text = tb.Fields("AMOUNT").Value
    Text8.Text = tb.Fields("THANZ").Value
    Text9.Text = tb.Fields("KORA_GAZANA").Value
    Text11.Text = tb.Fields("PERCENT").Value
    Label14.Caption = tb.Fields("PERLabel").Value & ""
    If Val(Text3.Text) > 0 Then
        lblDispute.Text = GetDispute(Val(Text3.Text))
    End If
    
    edit1 = False
Else
    MsgBox "No Record For This VOUCHER No."
    edit1 = True
    Exit Function
End If
tb.Close
DB.Close
End Function

Private Sub Save()
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
If Option2 = True Then
    Ssql = "Delete from PAYMENTDYING WHere VOU_NO = " & Val(Text1.Text)
    DB.Execute (Ssql)
End If
Set RS = DB.OpenRecordset("PAYMENTDYING", dbOpenDynaset)
RS.AddNew
    RS.Fields("Date").Value = DTPicker1.Value
    RS.Fields("VOU_No").Value = Val(Text1.Text)
    RS.Fields("LOT_No").Value = Val(Text3.Text)
    RS.Fields("CLOTH_CODE").Value = Val(Text4.Text) 'Combo4.ItemData(Combo4.ListIndex)
    RS.Fields("DYING_CODE").Value = Val(Text6.Text)
    RS.Fields("GAZANA").Value = Text2.Text
    RS.Fields("RATE").Value = Val(Text10.Text)
    RS.Fields("AMOUNT").Value = Text14.Text
    RS.Fields("THANZ").Value = Text8.Text
    RS.Fields("KORA_GAZANA").Value = Text9.Text
    RS.Fields("PERCENT").Value = Text11.Text
    RS.Fields("PERLabel").Value = Label14.Caption
RS.Update
RS.Close

Ssql = "Update Packing Set Dispute='" & lblDispute.Text & "' where Lot_No=" & Val(Text3.Text)
DB.Execute Ssql
DB.Close
End Sub
Private Function LOTEditKora() As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from ClothRec where LOT_NO =  " & Val(Text3.Text) & " and DYING_CODE =  " & Val(Text6.Text) & ""
Set tb = DB.OpenRecordset(Ssql)
'MsgBox "Test"
If Not tb.EOF Then
    Text8.Text = tb.Fields("THANS").Value
    Text9.Text = tb.Fields("GAZANA").Value
    'LOTEdit1 = False
Else
    Text10.Text = vbNullString
    Text14.Text = vbNullString
    Text8.Text = vbNullString
    Text9.Text = vbNullString
    Text11.Text = vbNullString
    Text2.Text = vbNullString
    MsgBox "This Lot is not Recvd for Selected Party"
    LOTEditLOTEditKora = True
    Exit Function
End If
tb.Close
DB.Close
End Function

Private Function LOTEditDying() As Boolean
Dim DB As Database
Dim tb As Recordset
Dim R As Long
Dim b As Boolean
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from Packing where LOT_NO = " & Val(Text3.Text)
Set tb = DB.OpenRecordset(Ssql)
'MsgBox "Test"
If Not tb.EOF Then
    Text4.Text = tb.Fields("Cloth_CODE").Value
    Text5.Text = blm.FillCloth1(tb.Fields("Cloth_Code").Value)
    Text6.Text = tb.Fields("DYING_CODE").Value
    Text7.Text = blm.Dying(tb.Fields("DYING_Code").Value)
    Text2.Text = tb.Fields("P_GAZANA").Value
    lblDispute.Text = tb.Fields("Dispute").Value & ""
    'LOTEdit1 = False
Else
    'MsgBox "No Record For This VOUCHER No."
    lblError.Caption = "No Dying Gazana Added, Please Add Dying Gazana"
    lblDispute.Text = ""
    LOTEditDying = True
    Exit Function
End If
tb.Close
DB.Close
End Function
Private Function GetDispute(LotNo As Long) As String
Dim DB As Database
Dim tb As Recordset
Dim R As Long
Dim b As Boolean
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from Packing where LOT_NO = " & LotNo
Set tb = DB.OpenRecordset(Ssql)
'MsgBox "Test"
If Not tb.EOF Then
    GetDispute = tb.Fields("Dispute").Value & ""
   
End If
tb.Close
DB.Close
End Function
Private Function Max1() As Double
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "Select Max(VOU_No) as C from PAYMENTDYING"
Set tb = DB.OpenRecordset(Ssql)
If Not IsNull(tb.Fields("C").Value) Then
    Max1 = tb.Fields("C").Value + 1
Else
    Max1 = 1
End If
tb.Close
DB.Close
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
Text6.SetFocus
Else
Text1.Enabled = True
Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command4_Click()
Dim DB As Database
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
If Option2 = True Then
    Ssql = "Delete from PaymentDying WHere Vou_NO = " & Val(Text1.Text)
    DB.Execute (Ssql)
End If
DB.Close
Command2_Click
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
Text1.Text = Max1
'Combo1.ListIndex = 0
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Text1.Text = Max1
Text6.SetFocus
Command4.Visible = False
End Sub

Private Sub Option2_Click()
Command2_Click
Text1.Enabled = True
Text1.SetFocus
Command4.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Option2 = True Then
If Val(Text1.Text) > 0 Then
    edit1
End If
End If

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 Then
'        Search3.Text3.Text = 2
'        Search3.Show
'    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text3_Validate(Cancel As Boolean)

If Val(Text3.Text) > 0 Then
    If Option1 = True Then
        Cancel = CheckLotNo
        If Cancel = True Then Exit Sub
    End If
     LOTEditKora
     LOTEditDying
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search1.Text3.Text = 5
        Search1.Show
    End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
        Search2.Text3.Text = 3
        Search2.Show
    End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
Dim b As Boolean
If Val(Text6.Text) > 0 Then
    Text7.Text = blm.Dying(Val(Text6.Text))
    If Text7.Text = "NOT FOUND" Then
        MsgBox "Invalid Dying Code...."
        Cancel = True
    End If
Else
    MsgBox "Please Give Some Dying Code...."
    Cancel = True
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim amt As Double
If Val(Text2.Text) > 0 And Val(Text10.Text) Then
    amt = Val(Text2.Text) * Val(Text10.Text)
End If
Text14.Text = amt
Dim age As Double
Dim per As Double
If Val(Text2.Text) > 0 And Val(Text9.Text) Then
    age = Val(Text9.Text) - Val(Text2.Text)
End If
'MsgBox age
If age <> 0 Then
    per = (age / Val(Text9.Text)) * 100
    If Val(Text2.Text) < Val(Text9.Text) Then
        Label14.Caption = "Less"
    ElseIf Val(Text2.Text) > Val(Text9.Text) Then
        Label14.Caption = "Xess"
    Else
        Label14.Caption = ""
    End If
End If
Text11.Text = Abs(per)
    
End Sub
