VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Inv1 
   Caption         =   "Sale Invoice"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   10125
      TabIndex        =   12
      Top             =   5895
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   7485
      MaxLength       =   10
      TabIndex        =   15
      Top             =   7890
      Width           =   1215
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   6975
      MaxLength       =   50
      TabIndex        =   14
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "A/c Info"
      Height          =   1695
      Left            =   360
      TabIndex        =   42
      Top             =   5640
      Width           =   5415
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text12 
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
         Left            =   960
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox Text13 
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
         Left            =   2880
         TabIndex        =   46
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label25 
         Caption         =   "City List"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Party List"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Party Name"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Party Code"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport r1 
      Left            =   0
      Top             =   4575
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text16 
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
      Left            =   7440
      TabIndex        =   53
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text15 
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
      Left            =   6975
      TabIndex        =   51
      Top             =   5940
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   3960
      TabIndex        =   47
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Inv1.frx":0000
         Left            =   2880
         List            =   "Inv1.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   54
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   45023235
         CurrentDate     =   37158
      End
      Begin VB.Label Label26 
         Caption         =   "Month"
         Height          =   375
         Left            =   2280
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Invocie #"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text11 
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
      Left            =   10095
      MaxLength       =   15
      TabIndex        =   16
      Top             =   6555
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
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
      Left            =   10125
      MaxLength       =   2
      TabIndex        =   11
      Top             =   5580
      Width           =   1230
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   10260
      TabIndex        =   40
      Top             =   7875
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
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
      Left            =   10125
      MaxLength       =   15
      TabIndex        =   13
      Top             =   6210
      Width           =   1215
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
      Left            =   6975
      TabIndex        =   37
      Top             =   5580
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2415
      Left            =   360
      TabIndex        =   35
      Top             =   3120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   8454143
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item Information"
      Height          =   1575
      Left            =   360
      TabIndex        =   24
      Top             =   1440
      Width           =   11175
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   10320
         TabIndex        =   64
         Text            =   "1"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   9480
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   58
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   661
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   7
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(*) Exp Search"
               TextSave        =   "(*) Exp Search"
               Key             =   ""
               Object.Tag             =   ""
               Object.ToolTipText     =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "(/) Finish"
               TextSave        =   "(/) Finish"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(F3) to Enter Party"
               TextSave        =   "(F3) to Enter Party"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "(F5) New"
               TextSave        =   "(F5) New"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "(F6) Update"
               TextSave        =   "(F6) Update"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(F7) Save"
               TextSave        =   "(F7) Save"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(F8) Reset"
               TextSave        =   "(F8) Reset"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   9000
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text17 
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
         Left            =   5940
         TabIndex        =   56
         Top             =   1230
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2640
         Top             =   1200
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Clear Entry"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   975
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
         Left            =   10080
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Left            =   7440
         TabIndex        =   4
         Top             =   720
         Width           =   735
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
         Left            =   6585
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text2 
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
         Left            =   1320
         TabIndex        =   27
         Top             =   720
         Width           =   5205
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   360
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "110001"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "INV_TYPE=1 FOR BUNYAN SAL,2 for Socks and 3 for Towels"
         Height          =   375
         Left            =   4680
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "GST (%)"
         Height          =   495
         Left            =   9600
         TabIndex        =   62
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "(F10) to Clear This Entry"
         Height          =   255
         Left            =   3120
         TabIndex        =   59
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Discount (%)"
         Height          =   495
         Left            =   8880
         TabIndex        =   57
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Cur. Stock"
         Height          =   255
         Left            =   5805
         TabIndex        =   55
         Top             =   1125
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
         Height          =   255
         Left            =   10320
         TabIndex        =   32
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Rate/Unit"
         Height          =   255
         Left            =   7440
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Unit"
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   465
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Exp. Name"
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Exp. Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Actions"
      Height          =   1215
      Left            =   7800
      TabIndex        =   23
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2400
         Picture         =   "Inv1.frx":006B
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1560
         Picture         =   "Inv1.frx":05CC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   855
         Left            =   720
         Picture         =   "Inv1.frx":0AD3
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1215
      Left            =   360
      TabIndex        =   20
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1920
         Picture         =   "Inv1.frx":0FC1
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   480
         Picture         =   "Inv1.frx":14E7
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Label Label29 
      Caption         =   "Service Charges"
      Height          =   285
      Left            =   8910
      TabIndex        =   70
      Top             =   5895
      Width           =   1260
   End
   Begin VB.Label Label28 
      Caption         =   "Change"
      Height          =   255
      Left            =   8895
      TabIndex        =   69
      Top             =   6570
      Width           =   1095
   End
   Begin VB.Label Label27 
      Caption         =   "Remarks "
      Height          =   255
      Left            =   5895
      TabIndex        =   68
      Top             =   6300
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Amount Less Discount %"
      Height          =   375
      Left            =   5895
      TabIndex        =   65
      Top             =   6660
      Width           =   1455
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6960
      TabIndex        =   60
      Top             =   6630
      Width           =   1710
   End
   Begin VB.Label Label16 
      Caption         =   "Total Item Types"
      Height          =   255
      Left            =   5880
      TabIndex        =   52
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Total Units"
      Height          =   255
      Left            =   5895
      TabIndex        =   50
      Top             =   5940
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Discount %"
      Height          =   255
      Left            =   8895
      TabIndex        =   41
      Top             =   5580
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Change"
      Height          =   255
      Left            =   8880
      TabIndex        =   39
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Amount Rec."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8895
      TabIndex        =   38
      Top             =   6210
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Total Amount"
      Height          =   255
      Left            =   5895
      TabIndex        =   36
      Top             =   5580
      Width           =   975
   End
End
Attribute VB_Name = "Inv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
'Private blmr As New bloom_r
Dim discountitm As Currency
Dim balamt As Currency

Private Function CheckDiscount()
Dim tb As Recordset
Dim db As Database
Dim ssql As String
Dim amt As Currency
Dim diff As Single
Dim orgamt As Currency

Set db = OpenDatabase(Blm1.pathMain)
ssql = "select * from item where Code = " & Val(Text1.Text)
Set tb = db.OpenRecordset(ssql)
If discountitm > 0 Then

If Not tb.EOF Then

    amt = tb.Fields("Purrate").Value * Val(Text5.Text)
    orgamt = Val(Text4.Text) * Val(Text5.Text)
    diff = orgamt - amt
    'MsgBox "Purchase amt " & amt & "  Sale Amt " & orgamt & " Diff " & diff & " Discount " & discountitm
    
    If discountitm > diff Then
    Text18.Text = vbNullString
    MsgBox "You cannot Allow Discount at this Ratio to This Item"
    End If
    
End If
End If

End Function

Private Sub SpeakBill()
Dim i As Long
For i = 1 To GRID1.Rows - 1

    curagent.Speak GRID1.TextMatrix(R, 2) & " Quantity " & GRID1.TextMatrix(R, 5)
    
Next i
End Sub
Private Sub LoadChar()
Dim strchar As String
strchar = "C:\WINDOWS\Msagent\CHARS\Genie.acs"
Call Agent1.Characters.Load("Genie", strchar)
Set curagent = Agent1.Characters("Genie")
Call curagent.Show
curagent.Speak "Welcome to Shop Right "
End Sub
Private Sub PrintReport()
Dim R As Long

        f = "{Sale_Vw_Final.Inv_no} = " & Val(Text14.Text)
        f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        r1.ReportFileName = App.Path & "\Invoice.rpt"
        r1.DataFiles(0) = App.Path & "\BLOOM.MDB"
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
End Sub
Private Sub edit1()
Dim db As Database
Dim tb As Recordset
Dim DI As Single
Dim ssql As String
Set db = OpenDatabase(Blm1.pathMain)
ssql = "select * from sale_1 where inv_type=" & Val(Text20.Text) & " and inv_no = " & Val(Text14.Text)
ssql = ssql & " And MNO = " & Combo3.ItemData(Combo3.ListIndex)


Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    Date1.Value = tb.Fields("Inv_date").Value
    Text12.Text = tb.Fields("party").Value
    Text13.Text = Blm1.party1(tb.Fields("Party").Value)
    Text10.Text = tb.Fields("Discount").Value
    Text9.Text = tb.Fields("AmountRec").Value
    Text11.Text = tb.Fields("Change").Value
    Text23.Text = tb.Fields("Services").Value & ""
    Text7.Text = tb.Fields("Total").Value
    Text21.Text = tb.Fields("Goods").Value & ""
    Text22.Text = tb.Fields("Bilty").Value & ""
Else
    MsgBox "Invalid Number in The Selected Month..."
End If
tb.Close

ssql = "select * from sale_2 where inv_no = " & Val(Text14.Text)
ssql = ssql & " and inv_type=" & Val(Text20.Text)
ssql = ssql & " And MNO = " & Combo3.ItemData(Combo3.ListIndex)

Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    Do While Not tb.EOF
        With GRID1
            .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = tb.Fields("item").Value
                .TextMatrix(.Rows - 1, 2) = Blm1.Item1(tb.Fields("item").Value)
                .TextMatrix(.Rows - 1, 3) = UnitRet(tb.Fields("item").Value)
                .TextMatrix(.Rows - 1, 4) = Format(tb.Fields("Rate").Value, "#.00")
                .TextMatrix(.Rows - 1, 5) = Format(tb.Fields("QTY").Value, "#.00")
                DI = (tb.Fields("Rate").Value * tb.Fields("QTY").Value) * tb.Fields("DiscountItm").Value / 100
                .TextMatrix(.Rows - 1, 6) = Format(tb.Fields("DiscountItm").Value, "#.00")
                .TextMatrix(.Rows - 1, 7) = Format(tb.Fields("GST").Value, "#.00")
                .TextMatrix(.Rows - 1, 8) = Format((Val(.TextMatrix(.Rows - 1, 4)) * Val(.TextMatrix(.Rows - 1, 5))) - DI, "#.00")
                DI = 0
        End With
        tb.MoveNext
    Loop
End If
tb.Close
db.Close
End Sub
Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim i As Long
If GRID1.Rows > 1 Then
    
    Set db = OpenDatabase(Blm1.pathMain)
If Option2 = True Then
ssql = "Delete from Sale_1 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
ssql = ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute ssql

ssql = "Delete from Sale_2 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
ssql = ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute ssql


ssql = "Delete from voucher where inv_type=" & Val(Text20.Text) & " and e_type=3 and ent_no = " & Val(Text14.Text)
ssql = ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute ssql
End If

Set tb = db.OpenRecordset("Voucher", dbOpenTable)
tb.AddNew
    tb.Fields("ent_no").Value = Val(Text14.Text)
    tb.Fields("E_TYpe").Value = 3
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("Party").Value = Val(Text12.Text)
    tb.Fields("Debit").Value = Val(Label20.Caption)
    tb.Fields("Remarks").Value = "Bill # " & Val(Text14.Text) & " On " & Format(Date1.Value, "dd/MMM/yyyy")
    If Val(Text10.Text) > 0 Then
        tb.Fields("Remarks").Value = tb.Fields("Remarks").Value & " Less " & Text10.Text & "%"
    End If
    tb.Fields("Credit").Value = 0
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text14.Text)
    tb.Fields("E_TYpe").Value = 3
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("Party").Value = 5000 + Val(Text20.Text) '"Vest Sales"
    tb.Fields("Credit").Value = Val(Label20.Caption)
    tb.Fields("Remarks").Value = "Bill # " & Val(Text14.Text) & " On " & Format(Date1.Value, "dd/MMM/yyyy")
    If Val(Text10.Text) > 0 Then
        tb.Fields("Remarks").Value = tb.Fields("Remarks").Value & " Less " & Text10.Text & "%"
    End If
    tb.Fields("Debit").Value = 0
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update

tb.Close

Set tb = db.OpenRecordset("Sale_1", dbOpenTable)
tb.AddNew
    tb.Fields("Inv_no").Value = Val(Text14.Text)
    tb.Fields("Inv_date").Value = Date1.Value
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("Party").Value = Val(Text12.Text)
    tb.Fields("Discount").Value = Val(Text10.Text)
    tb.Fields("AmountRec").Value = Fix(Val(Text9.Text))
    tb.Fields("Change").Value = Val(Text11.Text)
    tb.Fields("Services").Value = Val(Text23.Text)
    tb.Fields("Total").Value = Val(Label20.Caption)
    tb.Fields("S_Type").Value = 1
    tb.Fields("Goods").Value = Text21.Text
    tb.Fields("Bilty").Value = Text22.Text
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update
tb.Close

Set tb = db.OpenRecordset("Sale_2", dbOpenTable)
For i = 1 To GRID1.Rows - 1
With GRID1

tb.AddNew
    tb.Fields("Inv_no").Value = Val(Text14.Text)
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("Inv_date").Value = Date1.Value
    tb.Fields("Item").Value = Val(.TextMatrix(i, 1))
    tb.Fields("rate").Value = Val(.TextMatrix(i, 4))
    tb.Fields("Sr_No").Value = Val(.TextMatrix(i, 0))
    
'    MsgBox tb.Fields("rate").Value
    tb.Fields("QTY").Value = Val(.TextMatrix(i, 5))
    tb.Fields("DiscountItm").Value = Val(.TextMatrix(i, 6))
    tb.Fields("GST").Value = Val(.TextMatrix(i, 7))
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update
End With
Blm1.UpdateStock Val(GRID1.TextMatrix(i, 1)), Blm1.ITEMstocks(Val(GRID1.TextMatrix(i, 1)), Date1.Value), Date1.Value
Next i

tb.Close
db.Close

Text14.Tag = Text14.Text
i = MsgBox("Do You want to Print the Bill", vbYesNo, "Print Bill")
If i = 6 Then PrintReport
Command2_Click
Option1 = True
End If
End Sub

Private Sub Clear1()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text6.Text = vbNullString
Text18.Text = vbNullString
Text19.Text = vbNullString
Text17.Text = vbNullString



Text1.Text = "110001"
End Sub
Private Sub Transfer1()
Dim i As Long
Dim R As Long
With GRID1
    For R = 1 To .Rows - 1
        If .TextMatrix(R, 1) = Text1.Text Then
            'MsgBox Left(Text5.Text, 1)
            If Val(Text5.Text) < 0 Then
 '           MsgBox Val(.TextMatrix(r, 5)) & "   " & Val(Text5.Text)
            
            .TextMatrix(R, 5) = Val(.TextMatrix(R, 5)) - (Val(Text5.Text) * -1)
            
'            MsgBox .TextMatrix(r, 5)
            End If
            If Val(Text5.Text) > 0 Then .TextMatrix(R, 5) = Val(.TextMatrix(R, 5)) + Val(Text5.Text)
            .TextMatrix(R, 7) = Val(.TextMatrix(R, 4)) * Val(.TextMatrix(R, 5))
            If Val(.TextMatrix(R, 5)) <= 0 Then
                If .Rows > 2 Then
                    .RemoveItem R
                Else
                    .Rows = 1
                End If
            End If
            Exit Sub
        End If
    Next R
End With
With GRID1
    .Rows = .Rows + 1
    i = .Rows - 1
    .TextMatrix(i, 0) = i
    .TextMatrix(i, 1) = Text1.Text
    .TextMatrix(i, 2) = Text2.Text
    .TextMatrix(i, 3) = Text3.Text
    .TextMatrix(i, 4) = Format(Val(Text4.Text), "#.00")
    .TextMatrix(i, 5) = Format(Val(Text5.Text), "#.00")
    .TextMatrix(i, 6) = Format(Val(Text18.Text), "#.00")
    .TextMatrix(i, 7) = Format(Val(Text19.Text), "#.00")
    .TextMatrix(i, 8) = Format(Val(Text6.Text), "#.00")
    
    
End With
'GRID1.TopRow = GRID1.Rows - 1

End Sub

Private Sub Flex1()
With GRID1
    .Rows = 1
    .Cols = 9
    .ColWidth(0) = 500
    .TextMatrix(0, 0) = "Sr #"
    .ColWidth(1) = 1000
    .TextMatrix(0, 1) = "Item Code"
    .ColWidth(2) = 4200
    .TextMatrix(0, 2) = "Item Name"
    .ColWidth(3) = 10
    .TextMatrix(0, 3) = "Unit"
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "Rate/U"
    .ColWidth(5) = 1000
    .TextMatrix(0, 5) = "Quantity"
    .ColWidth(6) = 1000
    .TextMatrix(0, 6) = "Discount %"
    .ColWidth(7) = 800
    .TextMatrix(0, 7) = "GST %"
    .ColWidth(8) = 1500
    .TextMatrix(0, 8) = "Amount"
    
End With
End Sub
Private Function UnitRet(c As Long) As String
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(Blm1.pathMain)
ssql = "select unit from item where code = " & c
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    UnitRet = tb.Fields("Unit").Value & ""
    
End If
tb.Close
db_m.Close
End Function
Private Function RatePerUnit(c As Long) As Double
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(Blm1.pathMain)
ssql = "select Rate from item where code = " & c
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    If Not IsNull(tb.Fields("Rate").Value) Then
        RatePerUnit = tb.Fields("Rate").Value
    Else
        RatePerUnit = 0
    End If
Else
    RatePerUnit = 0
End If
tb.Close
db_m.Close
End Function

Private Function RateRet(c As Long) As Currency
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(Blm1.pathMain)
ssql = "select Rate from item where code = " & c
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    RateRet = tb.Fields("Rate").Value
    
End If
tb.Close
db_m.Close
End Function

Private Function max1() As Long
Dim db As Database
Dim tb As Recordset
Dim ssql As String
ssql = "select MAX(inv_no) AS C FROM sale_1 where inv_Type = " & Val(Text20.Text)
ssql = ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
'MsgBox ssql
Set db = OpenDatabase(Blm1.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
db.Close
End Function

Private Sub Check1_Click()
Combo1.Visible = Not Combo1.Visible
If Check1.Value = 1 Then
Text12.Enabled = True
Text13.Enabled = True
Else
Text12.Enabled = False
Text13.Enabled = False
Text12.Text = "110010002"
Text13.Text = Blm1.party1(110010002)
End If
End Sub

Private Sub Combo1_Click()
Dim R As Long
If Combo1.ListIndex > -1 Then
Text12.Text = Combo1.ItemData(Combo1.ListIndex)
Text13.Text = Combo1.Text
For R = 0 To Combo2.ListCount - 1
    If Combo2.ItemData(R) = Val(Mid(Text12.Text, 1, 2)) Then
        Combo2.ListIndex = R
        Exit For
    End If
Next R
End If

End Sub

Private Sub Combo2_LostFocus()
If Combo2.ListIndex > -1 Then
Screen.MousePointer = vbHourglass
Dim ssql As String

ssql = "Select * from Parties Where CCode = " & Combo2.ItemData(Combo2.ListIndex)
Blm1.fill_comb ssql, Combo1, "Name", "Code"
Me.Refresh
Screen.MousePointer = vbDefault
End If
End Sub

Private Sub Combo3_Click()
If Combo3.ListIndex > -1 Then
    Text14.Text = max1
End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Command1.Enabled = False
Dim i As Integer
save
Screen.MousePointer = vbDefault
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Clear1
Text9.Text = vbNullString
Text11.Text = vbNullString
Text10.Text = vbNullString
Text14.Text = vbNullString
Text23.Text = ""
Text1.Text = "110001"
Date1.Value = Date
Flex1
If Option1 = True Then
Text14.Text = max1
Text1.SetFocus
Else
Text14.Enabled = True
Text14.SetFocus

End If
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me

End Sub

Private Sub Command5_Click()
Transfer1
Clear1
Text1.SetFocus
End Sub

Private Sub Command6_Click()
If Option2 = True Then
    Blm1.LessStock Val(Text1.Text), Val(Text5.Text) * -1, Date1.Value
End If
Clear1
Text1.SetFocus
End Sub

Private Sub Command7_Click()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

Set db = OpenDatabase(Blm1.pathMain)
If Option2 = True Then
ssql = "Delete from Sale_1 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
ssql = ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute ssql

ssql = "Delete from Sale_2 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
ssql = ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute ssql


ssql = "Delete from voucher where inv_type=" & Val(Text20.Text) & " and e_type=3 and ent_no = " & Val(Text14.Text)
ssql = ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute ssql

End If
db.Close
Command2_Click

End Sub

Private Sub Form_Activate()
Dim ssql As String


ssql = "SELECT * FROM Parties oRDER BY Code"
Blm1.fill_comb ssql, Combo1, "NAME", "CODE"

ssql = "Select * from City Order by Name"
Blm1.fill_comb ssql, Combo2, "Name", "Code"
If Combo3.ListIndex > -1 Then
    Text14.Text = max1
End If
Text1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 42 Then
    Search1.Text3.Text = 1
    Search1.Show
End If
If KeyCode = 111 Then Text10.SetFocus
If KeyCode = vbKeyF5 Then Option1.Value = True
If KeyCode = vbKeyF6 Then Option2.Value = True
If KeyCode = vbKeyF7 Then Command1_Click
If KeyCode = vbKeyF8 Then Command2_Click
If KeyCode = vbKeyF10 Then Command6_Click
If KeyCode = vbKeyF3 Then Text12.SetFocus

End Sub

Private Sub Form_Load()
Dim ssql As String
Date1.Value = Date

Flex1
Combo3.ListIndex = 0


ssql = "SELECT * FROM Parties oRDER BY NAME"
Blm1.fill_comb ssql, Combo1, "NAME", "CODE"

ssql = "Select * from City Order by Name"
Blm1.fill_comb ssql, Combo2, "Name", "Code"
'LoadChar
End Sub

Private Sub Grid1_DblClick()
Dim thisrow As Long

If crow > 0 Then
    thisrow = crow
Else
    thisrow = GRID1.Row
End If
If Val(Text5.Text) > 0 Then
    MsgBox "You Already Have Entry There"
Else
With GRID1
Text1.Text = .TextMatrix(thisrow, 1)
Text17.Text = Format(Blm1.ITEMstocks(Val(.TextMatrix(thisrow, 1)), Date1.Value), "#.000")
Text2.Text = .TextMatrix(thisrow, 2)
Text3.Text = .TextMatrix(thisrow, 3)
Text4.Text = .TextMatrix(thisrow, 4)
Text5.Text = .TextMatrix(thisrow, 5)
Text18.Text = .TextMatrix(thisrow, 6)
Text19.Text = .TextMatrix(thisrow, 7)
Text6.Text = .TextMatrix(thisrow, 8)
End With
'MsgBox thisrow
If GRID1.Rows = 2 Then
    GRID1.Rows = 1
Else
    GRID1.RemoveItem (thisrow)
End If
Dim i As Long

For i = 1 To GRID1.Rows - 1
    GRID1.TextMatrix(i, 0) = i
Next i
End If
Text4.SetFocus

End Sub

Private Sub Option1_Click()
Command2_Click
Command7.Visible = False
Text14.Enabled = False
Text1.SetFocus
'date1.SetFocus
End Sub

Private Sub Option2_Click()
Command2_Click
Command7.Visible = True
Text14.Enabled = True
Text14.SetFocus
Date1.Enabled = True
End Sub

Private Sub Text1_GotFocus()
On Error Resume Next
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
      '  If KeyAscii = 46 Then
     '       Text10.SetFocus
      '  Else
            If KeyAscii = 42 Then
                Screen.MousePointer = vbHourglass
                Search1.Text3.Text = 1
                Search1.Show
                Screen.MousePointer = vbDefault
            Else
                Beep
                KeyAscii = 0
            End If
        'End If
    End If
End If
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim b As Boolean

If Val(Text1.Text) > 0 Then
    Text2.Text = Blm1.Item1(Val(Text1.Text))
    If Text2.Text = "NOT FOUND" Then
        MsgBox "Invalid Item Code...."
        Text1.Text = ""
        Cancel = True
    Else
        Text3.Text = UnitRet(Val(Text1.Text))
        Text4.Text = RatePerUnit(Val(Text1.Text))
        Text17.Text = Format(Blm1.ITEMstocks(Val(Text1.Text), Date1.Value), "#.000")
    End If
Else

    MsgBox "Please Give Some Item Code...."
    Cancel = True
End If
End Sub

Private Sub Text10_GotFocus()
On Error Resume Next
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text10_LostFocus()
Text10.BackColor = vbWhite
End Sub

Private Sub Text11_GotFocus()
On Error Resume Next
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text11_LostFocus()
Text11.BackColor = vbWhite
End Sub

Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12.Text)
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
If KeyCode = vbKeyF2 Then
    Search2.Text3.Text = 1
    Search2.Show
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
Dim b As Boolean

If Val(Text12.Text) > 0 Then
    Text13.Text = Blm1.party1(Val(Text12.Text))
    If Text13.Text = "NOT FOUND" Then
        MsgBox "Invalid A/c Code...."
        Cancel = True
    Else
        'ledgerbal
    End If
Else
    MsgBox "Please Give Some A/c Code...."
    Cancel = True
End If

End Sub

Private Sub Text14_GotFocus()
On Error Resume Next
Text14.SelStart = 0
Text14.SelLength = Len(Text14.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
        
        'Grid1.SetFocus
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text14_LostFocus()
'Me.ActiveControl.BackColor = vbWhite
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
If Option2 = True Then
If Val(Text14.Text) > 0 Then
    edit1
End If
End If
End Sub

Private Sub Text18_GotFocus()
On Error Resume Next
Text18.SelStart = 0
Text18.SelLength = Len(Text18.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text18_LostFocus()
Text18.BackColor = vbWhite
End Sub

Private Sub Text19_GotFocus()
On Error Resume Next
Text19.SelStart = 0
Text19.SelLength = Len(Text19.Text)
Me.ActiveControl.BackColor = vbYellow

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text19_LostFocus()
Text19.BackColor = vbWhite
End Sub

Private Sub Text21_GotFocus()
On Error Resume Next
Text21.SelStart = 0
Text21.SelLength = Len(Text21.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text21_LostFocus()
Text21.BackColor = vbWhite

End Sub

Private Sub Text22_GotFocus()
On Error Resume Next
Text22.SelStart = 0
Text22.SelLength = Len(Text22.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text22_LostFocus()
Text22.BackColor = vbWhite

End Sub

Private Sub Text23_GotFocus()
On Error Resume Next
Text23.SelStart = 0
Text23.SelLength = Len(Text23.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text23_LostFocus()
Text23.BackColor = vbWhite
End Sub

Private Sub Text4_GotFocus()
On Error Resume Next
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = vbWhite

End Sub

Private Sub Text5_GotFocus()
On Error Resume Next
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
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

Private Sub Text5_LostFocus()
Text5.BackColor = vbWhite
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
Dim R As Long
Dim qtyinbill As Currency

For R = 1 To GRID1.Rows - 1
    If Val(GRID1.TextMatrix(R, 1)) = Val(Text1.Text) Then
        qtyinbill = Val(GRID1.TextMatrix(R, 5))
        Exit For
    End If

Next R
'MsgBox Text17.Text & "   " & Text5.Text

If Len(Text5.Text) > 0 And Val(Text5.Text) <> 0 Then
    If Val(Text5.Text) <= (Val(Text17.Text) - qtyinbill) Then
        Cancel = False
    Else
        
'        Cancel = True
'        MsgBox "No Stock Available of this Item"
    End If
    'MsgBox "Text"
Else
    MsgBox "Please Give Some Qunatity"
    Cancel = True
End If
End Sub

Private Sub Text7_Change()
'Text9.Text = Text7.Text
End Sub

Private Sub Text9_GotFocus()
On Error Resume Next
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
Me.ActiveControl.BackColor = vbYellow
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text9_LostFocus()
Text9.BackColor = vbWhite
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
If Val(Text9.Text) <= 0 Then
 '   MsgBox "Please Give Recvd. Amount"
'    Cancel = True
End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo ErrHand

Dim amt As Currency
Dim amtdue As Currency
Dim tu As Currency
Dim itypes As Integer
Dim i As Long
Dim disamt As Currency
Dim ditm As Currency
Dim CNTL As Control
Dim SrvCharges As Double
'For Each CNTL In Me.Controls
'    If TypeOf CNTL Is TextBox Then
'    CNTL.BackColor = vbWhite
'    End If
'    If TypeOf CNTL Is CommandButton Then
'    CNTL.BackColor = vbWhite
''    End If
'    If TypeOf CNTL Is OptionButton Then
'    CNTL.BackColor = vbWhite
'    End If
'Next

'Me.ActiveControl.BackColor = vbYellow

If Val(Text4.Text) > 0 And Val(Text5.Text) > 0 Then
ditm = (Val(Text4.Text) * Val(Text5.Text)) * Val(Text18.Text) / 100
'MsgBox ditm
Text6.Text = (Val(Text4.Text) * Val(Text5.Text)) - ditm
Text6.Text = Val(Text6.Text) + (Val(Text6.Text) * Val(Text19.Text) / 100)
End If
If GRID1.Rows > 1 Then
    Command1.Enabled = True
    Command7.Enabled = True
Else
    Command1.Enabled = False
    Command7.Enabled = False
End If
amt = 0

For i = 1 To GRID1.Rows - 1
    amt = amt + Val(GRID1.TextMatrix(i, 8))
    tu = tu + Val(GRID1.TextMatrix(i, 5))
    'itypes = itypes + Val(Grid1.TextMatrix(i, 0))
    
Next i
balamt = 0
Text7.Text = Format(amt, "#.00")
Text15.Text = Format(tu, "#.00")
Text16.Text = Format(GRID1.Rows - 1, "#.00")
disamt = (Val(Text7.Text) / 100) * Val(Text10.Text)
amtdue = Val(Text7.Text) - disamt
SrvCharges = (amtdue * Val(Text23.Text)) / 100
balamt = amtdue + Val(SrvCharges)
amtdue = amtdue + Val(SrvCharges)
Label20.Caption = amtdue
Text11.Text = Format(amtdue - Val(Text9.Text), "#.00")

If Len(Text1.Text) = 6 Then
    Command5.Enabled = True
Else
    Command5.Enabled = False
    Exit Sub
End If

If Text2.Text = "NOT FOUND" Then
    Command5.Enabled = False
    Exit Sub
Else
    Command5.Enabled = True
End If

If Val(Text5.Text) = 0 Then
    Command5.Enabled = False
    Exit Sub
Else
    Command5.Enabled = True
End If
If Len(Text12.Text) > 1 Then
    If Text13.Text <> "NOT" Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
    
Else
    Command1.Enabled = False
End If
Exit Sub

ErrHand:
If Err.Number = 91 Then Resume Next

End Sub
