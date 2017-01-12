VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form YarnIssue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YARN ISSUE"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10395
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
      TabIndex        =   40
      Top             =   4800
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
         Picture         =   "YarnISSUE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
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
         Picture         =   "YarnISSUE.frx":09A6
         Style           =   1  'Graphical
         TabIndex        =   41
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
         Picture         =   "YarnISSUE.frx":1290
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ISSUE Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   27
      Top             =   1800
      Width           =   10095
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   6600
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text15 
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
         Left            =   1800
         TabIndex        =   16
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
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
         Left            =   4440
         TabIndex        =   14
         Top             =   2040
         Width           =   1455
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
         Left            =   4440
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   270
         Left            =   8640
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   7920
         TabIndex        =   20
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   255
         Left            =   8640
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   7920
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
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
         Top             =   840
         Width           =   4455
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
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
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
         Width           =   4455
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
      Begin VB.Label Label18 
         Caption         =   "Cons"
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
         Left            =   2040
         TabIndex        =   46
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Cons"
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
         Left            =   2040
         TabIndex        =   45
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Bags"
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
         TabIndex        =   44
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Bags"
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
         TabIndex        =   43
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Carrier Name"
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
         TabIndex        =   39
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "AT Sizing"
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
         Left            =   7680
         TabIndex        =   38
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label12 
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
         Left            =   6000
         TabIndex        =   37
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label11 
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
         Left            =   6000
         TabIndex        =   36
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "BANA Brand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   35
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "TANA Brand"
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
         Left            =   3720
         TabIndex        =   34
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "TANA Name"
         Height          =   255
         Left            =   7800
         TabIndex        =   33
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "BANA Name"
         Height          =   255
         Left            =   7800
         TabIndex        =   32
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
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
         TabIndex        =   31
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
         Left            =   120
         TabIndex        =   30
         Top             =   840
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   360
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
      Height          =   1695
      Left            =   4200
      TabIndex        =   24
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
         Left            =   4200
         Picture         =   "YarnISSUE.frx":1B93
         Style           =   1  'Graphical
         TabIndex        =   47
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
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   480
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
         Format          =   20774915
         CurrentDate     =   39498
      End
      Begin VB.Label Label2 
         Caption         =   "Issue No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
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
         Left            =   480
         TabIndex        =   25
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
         Picture         =   "YarnISSUE.frx":2531
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1215
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
         Height          =   1215
         Left            =   360
         Picture         =   "YarnISSUE.frx":2F48
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
End
Attribute VB_Name = "YarnIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1

Private Sub Clear1()
'Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = vbNullString
Text9.Text = vbNullString
Text10.Text = vbNullString
Text11.Text = vbNullString
Text12.Text = vbNullString
Text13.Text = vbNullString
Text14.Text = vbNullString
Text15.Text = vbNullString
Text16.Text = vbNullString
Text17.Text = vbNullString
Text18.Text = vbNullString
Text19.Text = vbNullString
If Option1 = True Then
    Text1.Enabled = False
    Text2.SetFocus
Else
    Text1.Enabled = True
    Text1.SetFocus
'Text2.Text = "11"
End If
End Sub


Private Sub Combs()
Dim Ssql As String

''Factory
'Ssql = "select * from FactoryChart order by Name"
'Blm.fill_comb Ssql, Combo2, "Name", "Code"
''cloth Quality
'Ssql = "select * from Cloths order by Name"
'Blm.fill_comb_Item Ssql, Combo3, "Name", "Code"
''Yarn
'Ssql = "select * from Yarns order by Name"
'Blm.fill_comb Ssql, Combo1, "Name", "Code"

End Sub

Private Function edit1() As Boolean
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from YARNISSUE where ISSUE_NO = " & Val(Text1.Text)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
    DTPicker1.Value = tb.Fields("Issue_Date").Value
    Text2.Text = tb.Fields("FAC_CODE").Value
    Text3.Text = blm.Factory(tb.Fields("FAC_Code").Value)
    Text4.Text = tb.Fields("Cloth_CODE").Value
    Text5.Text = blm.FillCloth1(tb.Fields("Cloth_Code").Value)
    'Text6.Text = tb.Fields("T_CODE").Value & ""
'    Text7.Text = blm.YarnName(tb.Fields("T_Code").Value) & ""
    'Text8.Text = tb.Fields("B_CODE").Value & ""
'    Text9.Text = blm.YarnName(tb.Fields("B_Code").Value)
    Text10.Text = tb.Fields("T_Brand").Value
    Text11.Text = tb.Fields("B_Brand").Value
    Text12.Text = tb.Fields("T_QTY").Value
    Text13.Text = tb.Fields("B_QTY").Value
    Text16.Text = tb.Fields("T_RATE").Value
    Text17.Text = tb.Fields("B_RATE").Value
    Text18.Text = tb.Fields("T_CONS").Value
    Text19.Text = tb.Fields("B_CONS").Value
    Text14.Text = tb.Fields("SIZING").Value
    Text15.Text = tb.Fields("CARRIER").Value
    edit1 = False
Else
    MsgBox "No Record For This Issue No."
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
    Ssql = "Delete from YARNISSUE WHere Issue_no = " & Val(Text1.Text)
    DB.Execute (Ssql)
End If
Set RS = DB.OpenRecordset("YARNISSUE", dbOpenDynaset)
RS.AddNew
    RS.Fields("ISSUE_DATE").Value = DTPicker1.Value
    RS.Fields("ISSUE_NO").Value = Val(Text1.Text)
    RS.Fields("FAC_CODE").Value = Val(Text2.Text)
    RS.Fields("CLOTH_CODE").Value = Val(Text4.Text) 'Combo4.ItemData(Combo4.ListIndex)
    'RS.Fields("T_CODE").Value = Val(Text6.Text)
    'RS.Fields("B_CODE").Value = Val(Text8.Text)
    RS.Fields("T_Qty").Value = Val(Text12.Text)
    RS.Fields("B_QTY").Value = Val(Text13.Text) 'Combo1.ItemData(Combo1.ListIndex)
    RS.Fields("T_BRAND").Value = Text10.Text
    RS.Fields("B_BRAND").Value = Text11.Text
    RS.Fields("T_RATE").Value = Val(Text16.Text)
    RS.Fields("B_RATE").Value = Val(Text17.Text)
    RS.Fields("T_CONS").Value = Val(Text18.Text)
    RS.Fields("B_CONS").Value = Val(Text19.Text)
    RS.Fields("SIZING").Value = Text14.Text
    RS.Fields("CARRIER").Value = Text15.Text
RS.Update
RS.Close
DB.Close
End Sub

Private Function Max1() As Double
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String
Set DB = OpenDatabase(blm.pathMain)
Ssql = "Select Max(Issue_No) as C from YARNISSUE"
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
If Val(Text2.Text) <= 0 Then
    MsgBox "Please Enter a Weaving Fatory"
    Exit Sub
End If
If Val(Text4.Text) <= 0 Then
    MsgBox "Please Enter a Cloth Quality"
    Exit Sub
End If
Screen.MousePointer = vbHourglass
Save
Command2_Click
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Clear1
Text1.Text = vbNullString
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
    Ssql = "Delete from YARNISSUE WHere ISSUE_NO = " & Val(Text1.Text)
    DB.Execute (Ssql)
End If
DB.Close
Command2_Click
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys ("{TAB}")
End If
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
Text1.Text = Max1
'Combo1.ListIndex = 0
Me.Top = 10
Me.Left = 10
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
'Delete Button
Command4.Visible = True
Text1.Enabled = True
Text1.SetFocus
DTPicker1.Enabled = True
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
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
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

Private Sub Text13_KeyPress(KeyAscii As Integer)
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
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
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

Private Sub Text17_KeyPress(KeyAscii As Integer)
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

Private Sub Text18_KeyPress(KeyAscii As Integer)
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

Private Sub Text19_KeyPress(KeyAscii As Integer)
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
        Search3.Text3.Text = 1
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
Dim b As Boolean
If Val(Text2.Text) > 0 Then
    Text3.Text = blm.Factory(Val(Text2.Text))
    If Text3.Text = "NOT FOUND" Then
        MsgBox "Invalid Factory Code...."
        Cancel = True
    End If
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search1.Text3.Text = 1
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
Dim b As Boolean
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
        Search4.Text3.Text = 1
        Search4.Show
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

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search4.Text3.Text = 2
        Search4.Show
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
