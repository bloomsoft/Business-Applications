VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form vour5 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   5685
   Begin VB.Frame Frame2 
      Caption         =   "Actions"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   5415
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   3240
         Picture         =   "Vour5.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   720
         Picture         =   "Vour5.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Only Packing Pending Lots List"
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
         Left            =   240
         TabIndex        =   19
         Top             =   4800
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Selected Item"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Items"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   3480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin Crystal.CrystalReport r1 
         Left            =   4200
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   4200
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4200
         Width           =   4935
      End
      Begin VB.TextBox Text1 
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
         Left            =   2160
         TabIndex        =   5
         Top             =   3570
         Width           =   1575
      End
      Begin VB.TextBox Text4 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4935
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
         Left            =   2160
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   60293123
         CurrentDate     =   39506
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   60293123
         CurrentDate     =   39506
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Press F1 For Search List"
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
         Left            =   540
         TabIndex        =   15
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label5 
         Caption         =   "Quality Info"
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
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Dying Info"
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
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "To"
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
         Left            =   840
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "From"
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
         Left            =   840
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Give Starting and Last Date For Preview"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "vour5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New Bloom_r
Private blm As New bloom1
Private Sub Command1_Click()
Dim F As String
Screen.MousePointer = vbHourglass

    If Val(Text5.Text) = 1 Then
    If Check1.Value = 0 Then
    If Option1 = True Then
        Blmr.DyeReport DTPicker1.Value, DTPicker2.Value, Val(Text3.Text)
    End If
    If Option2 = True Then
        Blmr.DyeReport DTPicker1.Value, DTPicker2.Value, Val(Text3.Text), Val(Text1.Text)
    End If
    End If
    If Check1.Value = 1 Then
    If Option1 = True Then
        Blmr.DyeReport DTPicker1.Value, DTPicker2.Value, Val(Text3.Text), , True
    End If
    If Option2 = True Then
        Blmr.DyeReport DTPicker1.Value, DTPicker2.Value, Val(Text3.Text), Val(Text1.Text), True
    End If
    End If
        r1.ReportFileName = App.Path & "\Reports\Dyingrpt.rpt"
        r1.DataFiles(0) = App.Path & "\Book.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = F
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
'DTPicker1.SetFocus
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search1.Text3.Text = 9
        Search1.Show
    End If
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
Dim B As Boolean
If Val(Text1.Text) > 0 Then
    Text2.Text = blm.FillCloth1(Val(Text1.Text))
    If Text2.Text = "NOT FOUND" Then
        MsgBox "Invalid Cloth Quality Code...."
        Cancel = True
    End If
Else
    'MsgBox "Please Give Some Cloth Quality Code...."
    'Cancel = True
End If

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
        Search2.Text3.Text = 5
        Search2.Show
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
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

Private Sub Text3_Validate(Cancel As Boolean)
Dim B As Boolean
If Val(Text3.Text) > 0 Then
    Text4.Text = blm.Dying(Val(Text3.Text))
    If Text4.Text = "NOT FOUND" Then
        MsgBox "Invalid Dying Code...."
        Cancel = True
    End If
Else
    MsgBox "Please Give Some Dying Code...."
    Cancel = True
End If

End Sub
