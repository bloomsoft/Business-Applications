VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ClothRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cloth Recieving"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
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
      TabIndex        =   27
      Top             =   4320
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
         Picture         =   "ClothRec.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "ClothRec.frx":09A6
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Picture         =   "ClothRec.frx":1290
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cloth Recieving and Dying Issuence Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   10095
      Begin VB.TextBox txtProgram 
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
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   990
         Width           =   1695
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
         Left            =   8160
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1665
         TabIndex        =   7
         Top             =   1440
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   661
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
         Format          =   20709379
         CurrentDate     =   39501
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
         Left            =   8160
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
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
         Height          =   360
         Left            =   4920
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
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
         Height          =   360
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2040
         Width           =   3615
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
         Left            =   1665
         TabIndex        =   10
         Top             =   2040
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
         Height          =   360
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   960
         Width           =   4215
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
         Left            =   1695
         TabIndex        =   5
         Top             =   930
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   4215
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
         Height          =   345
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Program"
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
         Left            =   8130
         TabIndex        =   32
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Lot No."
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
         Left            =   7080
         TabIndex        =   30
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label13 
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
         Left            =   7080
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Thans"
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
         Left            =   4080
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
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
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Rec. Date"
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
         TabIndex        =   23
         Top             =   1440
         Width           =   1380
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
         TabIndex        =   22
         Top             =   960
         Width           =   1455
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
         TabIndex        =   21
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Issue Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4200
      TabIndex        =   17
      Top             =   0
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
         Height          =   1335
         Left            =   3840
         Picture         =   "ClothRec.frx":1B93
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
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
         Height          =   345
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
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
         Format          =   20709379
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
         Left            =   960
         TabIndex        =   19
         Top             =   480
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
         Left            =   2760
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
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
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   0
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
         Picture         =   "ClothRec.frx":2531
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   360
         Picture         =   "ClothRec.frx":2F48
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "ClothRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1

Private Sub Clear1()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = vbNullString
Text10.Text = vbNullString
Text14.Text = vbNullString
txtProgram.Text = ""
If Option1 = True Then
    Text1.Enabled = False
    Text2.SetFocus
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

Private Function LOTCheck() As Boolean
Dim DB As Database
Dim tb As Recordset
Dim R As Long
Dim B As Boolean
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from ClothRec where LOT_NO =  " & Val(Text8.Text) & " and DYING_CODE =  " & Val(Text6.Text) & ""

Set tb = DB.OpenRecordset(Ssql)
'MsgBox ssql
'MsgBox "Test"
If Not tb.EOF Then
    MsgBox "This Lot No alreasy Exist"
    LOTCheck = True
End If
tb.Close
DB.Close
End Function


Private Function edit1() As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from CLOTHREC where REC_NO = " & Val(Text1.Text)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    DTPicker1.Value = tb.Fields("Date_REC").Value
    DTPicker2.Value = tb.Fields("Date_RECIEVE").Value
    Text2.Text = tb.Fields("FAC_CODE").Value
    Text3.Text = blm.Factory(tb.Fields("FAC_Code").Value)
    Text4.Text = tb.Fields("Cloth_CODE").Value
    Text5.Text = blm.FillCloth1(tb.Fields("Cloth_Code").Value)
    Text6.Text = tb.Fields("DYING_CODE").Value
    Text7.Text = blm.Dying(tb.Fields("DYING_Code").Value)
    Text14.Text = tb.Fields("GAZANA").Value
    Text10.Text = tb.Fields("THANS").Value
    Text8.Text = tb.Fields("LOT_NO").Value & ""
    txtProgram.Text = tb.Fields("Program").Value & ""
    edit1 = False
Else
    MsgBox "No Record For This Reciept No."
    edit1 = True
    Exit Function
End If
tb.Close
DB.Close
End Function

Private Sub Save()
Dim DB As Database
Dim RS As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
If Option2 = True Then
    Ssql = "Delete from ClothRec WHere REC_NO = " & Val(Text1.Text)
    DB.Execute (Ssql)
End If
Set RS = DB.OpenRecordset("ClOTHREC", dbOpenDynaset)
RS.AddNew
    RS.Fields("Date_REC").Value = DTPicker1.Value
    RS.Fields("Date_RECIEVE").Value = DTPicker2.Value
    RS.Fields("REC_No").Value = Val(Text1.Text)
    RS.Fields("FAC_CODE").Value = Val(Text2.Text)
    RS.Fields("CLOTH_CODE").Value = Val(Text4.Text) 'Combo4.ItemData(Combo4.ListIndex)
    RS.Fields("DYING_CODE").Value = Val(Text6.Text)
    RS.Fields("GAZANA").Value = Val(Text14.Text)
    RS.Fields("THANS").Value = Val(Text10.Text)
    RS.Fields("LOT_NO").Value = Val(Text8.Text)
    RS.Fields("Program").Value = txtProgram.Text
RS.Update
RS.Close
DB.Close
End Sub

Private Function Max1() As Double
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "Select Max(Rec_No) as C from ClothRec"
Set tb = DB.OpenRecordset(Ssql)
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
Me.Hide
Unload Me
End Sub

Private Sub Command4_Click()
Dim DB As Database
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
If Option2 = True Then
    Ssql = "Delete from CLOTHREC WHere REC_NO = " & Val(Text1.Text)
    DB.Execute (Ssql)
End If
DB.Close
Command2_Click
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub Form_Load()
'Combo4.ListIndex = 0
DTPicker2.Value = Date
Text1.Text = Max1
'Combo1.ListIndex = 0
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Option1_Click()
Command2_Click
Text1.Enabled = False
Text1.Text = Max1
Text2.SetFocus
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

Private Sub Text14_KeyPress(KeyAscii As Integer)
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
    If KeyCode = vbKeyF1 Then
        Search3.Text3.Text = 2
        Search3.Show
    End If
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

Private Sub Text2_Validate(Cancel As Boolean)
Dim B As Boolean
If Val(Text2.Text) > 0 Then
    Text3.Text = blm.Factory(Val(Text2.Text))
    If Text3.Text = "NOT FOUND" Then
        MsgBox "Invalid Factory Code...."
        Cancel = True
    End If
Else
    MsgBox "Please Give Some Factory Code...."
    Cancel = True
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search1.Text3.Text = 2
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

Private Sub Text4_Validate(Cancel As Boolean)
Dim B As Boolean
If Val(Text4.Text) > 0 Then
    Text5.Text = blm.FillCloth1(Val(Text4.Text))
    If Text5.Text = "NOT FOUND" Then
        MsgBox "Invalid Cloth Quality Code...."
        Cancel = True
    End If
Else
    MsgBox "Please Give Some Cloth Quality Code...."
    Cancel = True
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search2.Text3.Text = 1
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
Dim B As Boolean
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

Private Sub Text8_Validate(Cancel As Boolean)
If Val(Text8.Text) > 0 Then
    'MsgBox "Test"
    Cancel = LOTCheck
Else
Command1.SetFocus
End If
End Sub

Private Sub txtProgram_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys ("{TAB}")
End If
End Sub
