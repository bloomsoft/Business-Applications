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
   ScaleHeight     =   6300
   ScaleWidth      =   8130
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
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
      Left            =   5760
      MaxLength       =   15
      TabIndex        =   80
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   7320
      TabIndex        =   75
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   9960
      MaxLength       =   15
      TabIndex        =   74
      Top             =   8040
      Visible         =   0   'False
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
      Left            =   7080
      MaxLength       =   2
      TabIndex        =   73
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   9960
      TabIndex        =   72
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Totals"
      Height          =   2295
      Left            =   5880
      TabIndex        =   57
      Top             =   5040
      Width           =   5655
      Begin VB.TextBox Text30 
         Height          =   405
         Left            =   1200
         ScrollBars      =   2  'Vertical
         TabIndex        =   88
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   1200
         TabIndex        =   87
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text28 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   85
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   1200
         TabIndex        =   83
         Top             =   720
         Width           =   1695
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
         Height          =   405
         Left            =   1200
         TabIndex        =   61
         Top             =   240
         Width           =   1695
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
         Left            =   4080
         TabIndex        =   60
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   4080
         MaxLength       =   50
         TabIndex        =   59
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   58
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Caption         =   "Note"
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
         Left            =   120
         TabIndex        =   89
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label39 
         Caption         =   "Total Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "Brokery"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Brokery Rate"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Amount W/O Brokery"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Total Bales"
         Height          =   255
         Left            =   3000
         TabIndex        =   66
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Total Weight"
         Height          =   255
         Left            =   3000
         TabIndex        =   65
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Truck Lorry"
         Height          =   255
         Left            =   3000
         TabIndex        =   64
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Bilty Number"
         Height          =   255
         Left            =   3000
         TabIndex        =   63
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label30 
         Height          =   330
         Left            =   3960
         TabIndex        =   62
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Actions"
      Height          =   1215
      Left            =   8520
      TabIndex        =   50
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "&Save"
         Height          =   855
         Left            =   120
         Picture         =   "Inv1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000E&
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1080
         Picture         =   "Inv1.frx":2471
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000E&
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2040
         Picture         =   "Inv1.frx":4AEE
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "A/c Info"
      Height          =   2295
      Left            =   360
      TabIndex        =   27
      Top             =   5040
      Width           =   5415
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   2880
         TabIndex        =   71
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   960
         TabIndex        =   69
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Text            =   "Combo2"
         Top             =   1200
         Width           =   4095
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
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1680
         Width           =   4095
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
         TabIndex        =   31
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label36 
         Caption         =   "Bro Name"
         Height          =   255
         Left            =   1920
         TabIndex        =   70
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label35 
         Caption         =   "Bro Code"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "City List"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Party List"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Party Name"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Party Code"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport r1 
      Left            =   0
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   3120
      TabIndex        =   32
      Top             =   120
      Width           =   5295
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   2760
         TabIndex        =   49
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Inv1.frx":550A
         Left            =   2760
         List            =   "Inv1.frx":5535
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3840
         Picture         =   "Inv1.frx":5575
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker Date1 
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   37158
      End
      Begin VB.Label Label32 
         Caption         =   "Job Balance"
         Height          =   255
         Left            =   3840
         TabIndex        =   54
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Job No."
         Height          =   255
         Left            =   2040
         TabIndex        =   48
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Month"
         Height          =   375
         Left            =   2160
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Invocie #"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2055
      Left            =   360
      TabIndex        =   26
      Top             =   3000
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   3625
      _Version        =   393216
      BackColor       =   8454143
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item Information"
      Height          =   1575
      Left            =   360
      TabIndex        =   16
      Top             =   1440
      Width           =   11175
      Begin VB.TextBox Text31 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10500
         TabIndex        =   90
         Top             =   720
         Width           =   585
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   10080
         TabIndex        =   56
         Top             =   720
         Width           =   450
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   8760
         TabIndex        =   8
         Text            =   "Combo4"
         Top             =   720
         Width           =   1365
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   10320
         TabIndex        =   43
         Text            =   "1"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   7440
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   661
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   7
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(*) Item Search"
               TextSave        =   "(*) Item Search"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "(/) Finish"
               TextSave        =   "(/) Finish"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(F3) to Enter Party"
               TextSave        =   "(F3) to Enter Party"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "(F5) New"
               TextSave        =   "(F5) New"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "(F6) Update"
               TextSave        =   "(F6) Update"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(F7) Save"
               TextSave        =   "(F7) Save"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Text            =   "(F8) Reset"
               TextSave        =   "(F8) Reset"
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Top             =   720
         Width           =   840
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
         Left            =   4320
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2625
         Top             =   1200
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Clear Entry"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   7920
         TabIndex        =   7
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
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   6
         Top             =   720
         Width           =   1005
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
         Left            =   5280
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
         Left            =   3720
         TabIndex        =   21
         Top             =   720
         Width           =   630
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
         TabIndex        =   19
         Top             =   720
         Width           =   2415
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
      Begin VB.Label Label34 
         Caption         =   "L-Kami"
         Height          =   255
         Left            =   10320
         TabIndex        =   55
         Top             =   465
         Width           =   495
      End
      Begin VB.Label Label29 
         Caption         =   "WareHouse"
         Height          =   195
         Left            =   9000
         TabIndex        =   47
         Top             =   450
         Width           =   885
      End
      Begin VB.Label Label18 
         Caption         =   "Bales"
         Height          =   240
         Left            =   6240
         TabIndex        =   46
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label23 
         Caption         =   "INV_TYPE=1 FOR BUNYAN SAL,2 for Socks and 3 for Towels"
         Height          =   375
         Left            =   4680
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "GST (%)"
         Height          =   495
         Left            =   7560
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "(F10) to Clear This Entry"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label17 
         Caption         =   "Cur. Stock"
         Height          =   255
         Left            =   4440
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
         Height          =   255
         Left            =   7920
         TabIndex        =   24
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
         Left            =   7080
         TabIndex        =   23
         Top             =   465
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Rate"
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Unit"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Item Code"
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
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1215
      Left            =   360
      TabIndex        =   13
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000E&
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1440
         Picture         =   "Inv1.frx":5E29
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "Inv1.frx":634F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
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
      Left            =   4440
      TabIndex        =   81
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
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
      Height          =   375
      Left            =   5280
      TabIndex        =   79
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Amount Less Discount %"
      Height          =   375
      Left            =   3960
      TabIndex        =   78
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Total Item Types"
      Height          =   255
      Left            =   5880
      TabIndex        =   77
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Change"
      Height          =   255
      Left            =   8760
      TabIndex        =   76
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
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
Private Sub JobBalance()
Dim db As Database
Dim tb As Recordset
Dim tbP As Recordset
Dim Ssql As String
Set db = OpenDatabase(blm.pathMain)

Ssql = "Select Sum(Quantity-LKamiValue) as Bal from In_DTL where JobNo=" & Val(Text23.Text)
Set tbP = db.OpenRecordset(Ssql)

Ssql = "select * from PContract where Cont_no = " & Val(Text23.Text) & " and SellerCode=" & Val(Text12.Text)
Set tb = db.OpenRecordset(Ssql)
If Not tb.EOF Then
    If Not IsNull(tbP.Fields("Bal")) Then
        Label32.Caption = tb.Fields("Quantity") - tbP.Fields("Bal").Value
    Else
        Label32.Caption = tb.Fields("Quantity")
    End If
Else
    MsgBox "Invalid Job No. or Don't Belong to the Selected party"
End If
tbP.Close
tb.Close
db.Close
End Sub

Private Function CheckDiscount()
'Dim tb As Recordset
'Dim db As Database
'Dim ssql As String
'Dim amt As Currency
'Dim diff As Single
'Dim orgamt As Currency
'
'Set db = OpenDatabase(Blm1.pathMain)
'ssql = "select * from item where Code = " & Val(Text1.Text)
'Set tb = db.OpenRecordset(ssql)
'If discountitm > 0 Then
'
'If Not tb.EOF Then
'
'    amt = tb.Fields("Purrate").Value * Val(Text5.Text)
'    orgamt = Val(Text4.Text) * Val(Text5.Text)
'    diff = orgamt - amt
'    'MsgBox "Purchase amt " & amt & "  Sale Amt " & orgamt & " Diff " & diff & " Discount " & discountitm
'
'    If discountitm > diff Then
'    Text18.Text = vbNullString
'    MsgBox "You cannot Allow Discount at this Ratio to This Item"
'    End If
'
'End If
'End If

End Function

Private Sub SpeakBill()
Dim i As Long
For i = 1 To Grid1.Rows - 1

    curagent.Speak Grid1.TextMatrix(R, 2) & " Quantity " & Grid1.TextMatrix(R, 5)
    
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
Load InvPrint
InvPrint.Text2.Text = Text20.Text
InvPrint.Text1.Text = Val(Text14.Text)

For R = 0 To InvPrint.Combo3.ListCount - 1
    If InvPrint.Combo3.ItemData(R) = Combo3.ItemData(Combo3.ListIndex) Then
        InvPrint.Combo3.ListIndex = R
        Exit For
    End If
Next R

Me.WindowState = 1
InvPrint.Show
End Sub
Private Sub edit1()
Dim db As Database
Dim tb As Recordset
Dim DI As Single
Dim Ssql As String
Set db = OpenDatabase(Blm1.pathMain)
Ssql = "select * from sale_1 where inv_type=" & Val(Text20.Text) & " and inv_no = " & Val(Text14.Text)
Ssql = Ssql & " And MNO = " & Combo3.ItemData(Combo3.ListIndex)


Set tb = db.OpenRecordset(Ssql)
If Not tb.EOF Then
    Date1.Value = tb.Fields("Inv_date").Value
    Text12.Text = tb.Fields("party").Value
    Text13.Text = Blm1.party1(tb.Fields("Party").Value)
    Text25.Text = tb.Fields("BrokerCode").Value & ""
    Text26.Text = Blm1.broker1(Val(tb.Fields("BrokerCode").Value & ""))
    Text10.Text = tb.Fields("Discount").Value
    Text9.Text = tb.Fields("AmountRec").Value
    Text11.Text = tb.Fields("Change").Value
    Text7.Text = tb.Fields("Total").Value
    Text21.Text = tb.Fields("Goods").Value & ""
    Text22.Text = tb.Fields("Bilty").Value & ""
Else
    MsgBox "Invalid Number in The Selected Month..."
End If
tb.Close

Ssql = "select * from sale_2 where inv_no = " & Val(Text14.Text)
Ssql = Ssql & " and inv_type=" & Val(Text20.Text)
Ssql = Ssql & " And MNO = " & Combo3.ItemData(Combo3.ListIndex)

Set tb = db.OpenRecordset(Ssql)
If Not tb.EOF Then
    Do While Not tb.EOF
        With Grid1
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
                .TextMatrix(.Rows - 1, 9) = tb.Fields("WareHouse").Value
                .TextMatrix(.Rows - 1, 10) = Blm1.WareHouse(tb.Fields("WareHouse").Value)
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
Dim Ssql As String
Dim i As Long
If Grid1.Rows > 1 Then
    If Val(Text9.Text) > 0 Then
    Set db = OpenDatabase(Blm1.pathMain)
If Option2 = True Then
Ssql = "Delete from Sale_1 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
Ssql = Ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute Ssql

Ssql = "Delete from Sale_2 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
Ssql = Ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute Ssql


Ssql = "Delete from voucher where inv_type=" & Val(Text20.Text) & " and e_type=3 and ent_no = " & Val(Text14.Text)
Ssql = Ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute Ssql
End If

Set tb = db.OpenRecordset("Voucher", dbOpenTable)
tb.AddNew
    tb.Fields("ent_no").Value = Val(Text14.Text)
    tb.Fields("E_TYpe").Value = 3
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("Party").Value = Val(Text12.Text)
    tb.Fields("BrokerCode").Value = Val(Text25.Text)
    tb.Fields("Debit").Value = Val(Label20.Caption)
    tb.Fields("Remarks").Value = "Bill # " & Val(Text14.Text) & " Total Wt: " & Text15.Text & " Total Bales: " & Label30.Caption
    tb.Fields("Credit").Value = 0
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text14.Text)
    tb.Fields("E_TYpe").Value = 3
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("Party").Value = 5000 + Val(Text20.Text) '"Vest Sales"
    tb.Fields("BrokerCode").Value = Val(Text25.Text)
    tb.Fields("Credit").Value = Val(Label20.Caption)
    tb.Fields("Remarks").Value = Text13.Text & " Bill # " & Val(Text14.Text) & " Total Wt: " & Text15.Text & " Total Bales: " & Label30.Caption
    tb.Fields("Debit").Value = 0
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text14.Text)
    tb.Fields("E_TYpe").Value = 3
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("Party").Value = Val(Text25.Text)
    tb.Fields("BrokerCode").Value = -1
    tb.Fields("Debit").Value = 0
    tb.Fields("Credit").Value = Val(Text28.Text)
    tb.Fields("Remarks").Value = "Brokerage on Bill # " & Val(Text14.Text) & " Total Wt: " & Text15.Text & " Total Bales: " & Label30.Caption
    tb.Fields("Credit").Value = 0
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update

tb.AddNew
    tb.Fields("ent_no").Value = Val(Text14.Text)
    tb.Fields("E_TYpe").Value = 3
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("v_date").Value = Date1.Value
    tb.Fields("BrokerCode").Value = -1
    tb.Fields("Party").Value = Val(Text25.Text) 'Brokery Debit A/c
    tb.Fields("Debit").Value = Val(Text28.Text)
    tb.Fields("Credit").Value = 0
    tb.Fields("Remarks").Value = "Brokerage on  Bill # " & Val(Text14.Text) & " Total Wt: " & Text15.Text & " Total Bales: " & Label30.Caption
    tb.Fields("Credit").Value = 0
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update

tb.Close

Set tb = db.OpenRecordset("Sale_1", dbOpenTable)
tb.AddNew
    tb.Fields("Inv_no").Value = Val(Text14.Text)
    tb.Fields("Inv_date").Value = Date1.Value
    tb.Fields("Inv_type").Value = Val(Text20.Text)
    tb.Fields("Party").Value = Val(Text12.Text)
    tb.Fields("BrokerCode").Value = Val(Text25.Text)
    tb.Fields("Discount").Value = Val(Text10.Text)
    tb.Fields("AmountRec").Value = Fix(Text9.Text)
    tb.Fields("Change").Value = Val(Text11.Text)
    tb.Fields("Total").Value = Val(Label20.Caption)
    tb.Fields("S_Type").Value = 1
    tb.Fields("Goods").Value = Text21.Text
    tb.Fields("Bilty").Value = Text22.Text
    tb.Fields("MNO").Value = Combo3.ItemData(Combo3.ListIndex)
tb.Update
tb.Close

Set tb = db.OpenRecordset("Sale_2", dbOpenTable)
For i = 1 To Grid1.Rows - 1
With Grid1

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
    tb.Fields("WareHouse").Value = Val(.TextMatrix(i, 9))
tb.Update
End With
Blm1.UpdateStock Val(Grid1.TextMatrix(i, 1)), Blm1.ITEMstocks(Val(Grid1.TextMatrix(i, 1)), Date1.Value), Date1.Value
Next i

tb.Close
db.Close

Text14.Tag = Text14.Text
i = MsgBox("Do You want to Print the Bill", vbYesNo, "Print Bill")
If i = 6 Then PrintReport
Command2_Click
Option1 = True
Else
    MsgBox "Please Give Recv. Amount"
    Text9.SetFocus
End If
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
With Grid1
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
With Grid1
    .Rows = .Rows + 1
    i = .Rows - 1
    .TextMatrix(i, 0) = i
    .TextMatrix(i, 1) = Text1.Text
    .TextMatrix(i, 2) = Text2.Text
    .TextMatrix(i, 3) = Text3.Text
    .TextMatrix(i, 4) = Format(Val(Text4.Text), "#.00")
    .TextMatrix(i, 5) = Format(Val(Text5.Text), "#.00")
    .TextMatrix(i, 6) = Text18.Text
    .TextMatrix(i, 7) = Format(Val(Text19.Text), "#.00")
    .TextMatrix(i, 8) = Format(Val(Text6.Text), "#.00")
    .TextMatrix(i, 9) = Combo4.ItemData(Combo4.ListIndex)
    .TextMatrix(i, 10) = Combo4.Text
    
End With
'GRID1.TopRow = GRID1.Rows - 1

End Sub

Private Sub Flex1()
With Grid1
    .Rows = 1
    .Cols = 11
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
    .TextMatrix(0, 6) = "Bales"
    .ColWidth(7) = 0
    .TextMatrix(0, 7) = "GST %"
    .ColWidth(8) = 1500
    .TextMatrix(0, 8) = "Amount"
    .ColWidth(9) = 0
    .TextMatrix(0, 9) = "WareHouseCode"
    .ColWidth(10) = 1000
    .TextMatrix(0, 10) = "WareHouse"
End With
End Sub
Private Function UnitRet(c As Long) As String
Dim DB_M As Database
Dim tb As Recordset
Dim Ssql As String

Set DB_M = OpenDatabase(Blm1.pathMain)
Ssql = "select unit from item where code = " & c
Set tb = DB_M.OpenRecordset(Ssql)
If Not tb.EOF Then
    UnitRet = tb.Fields("Unit").Value & ""
    
End If
tb.Close
DB_M.Close
End Function
Private Function RatePerUnit(c As Long) As Double
Dim DB_M As Database
Dim tb As Recordset
Dim Ssql As String

Set DB_M = OpenDatabase(Blm1.pathMain)
Ssql = "select Rate from item where code = " & c
Set tb = DB_M.OpenRecordset(Ssql)
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
DB_M.Close
End Function

Private Function RateRet(c As Long) As Currency
Dim DB_M As Database
Dim tb As Recordset
Dim Ssql As String

Set DB_M = OpenDatabase(Blm1.pathMain)
Ssql = "select Rate from item where code = " & c
Set tb = DB_M.OpenRecordset(Ssql)
If Not tb.EOF Then
    RateRet = tb.Fields("Rate").Value
    
End If
tb.Close
DB_M.Close
End Function

Private Function max1() As Long
Dim db As Database
Dim tb As Recordset
Dim Ssql As String
Ssql = "select MAX(inv_no) AS C FROM sale_1 where inv_Type = " & Val(Text20.Text)
Ssql = Ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
'MsgBox ssql
Set db = OpenDatabase(Blm1.pathMain)
Set tb = db.OpenRecordset(Ssql)
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
Dim Ssql As String

Ssql = "Select * from Parties Where CCode = " & Combo2.ItemData(Combo2.ListIndex)
Blm1.fill_comb Ssql, Combo1, "Name", "Code"
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

Private Sub Combo4_KeyPress(KeyAscii As Integer)
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
Dim Ssql As String

Set db = OpenDatabase(Blm1.pathMain)
If Option2 = True Then
Ssql = "Delete from Sale_1 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
Ssql = Ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute Ssql

Ssql = "Delete from Sale_2 where inv_type=" & Val(Text20.Text) & " and Inv_No = " & Text14.Text
Ssql = Ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute Ssql


Ssql = "Delete from voucher where inv_type=" & Val(Text20.Text) & " and e_type=3 and ent_no = " & Val(Text14.Text)
Ssql = Ssql & " and MNO = " & Combo3.ItemData(Combo3.ListIndex)
db.Execute Ssql

End If
db.Close
Command2_Click

End Sub

Private Sub Form_Activate()
Dim Ssql As String

Text14.Text = max1
Ssql = "SELECT * FROM Parties oRDER BY NAME"
Blm1.fill_comb Ssql, Combo1, "NAME", "CODE"

Ssql = "Select * from City Order by Name"
Blm1.fill_comb Ssql, Combo2, "Name", "Code"

Ssql = "Select * from WareHouse Order by Name"
Blm1.fill_comb Ssql, Combo4, "Name", "Code"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 42 Then
    Search1.Text3.Text = 1
    Search1.Show
End If

If KeyCode = vbKeyF5 Then Option1.Value = True
If KeyCode = vbKeyF6 Then Option2.Value = True
If KeyCode = vbKeyF7 Then Command1_Click
If KeyCode = vbKeyF8 Then Command2_Click
If KeyCode = vbKeyF10 Then Command6_Click
If KeyCode = vbKeyF3 Then Text12.SetFocus

End Sub

Private Sub Form_Load()
Dim Ssql As String
Date1.Value = Date

Flex1
Combo3.ListIndex = 0


Ssql = "SELECT * FROM Parties oRDER BY NAME"
Blm1.fill_comb Ssql, Combo1, "NAME", "CODE"

Ssql = "Select * from City Order by Name"
Blm1.fill_comb Ssql, Combo2, "Name", "Code"
'LoadChar
End Sub

Private Sub Grid1_DblClick()
Dim thisrow As Long
Dim R As Integer
If crow > 0 Then
    thisrow = crow
Else
    thisrow = Grid1.Row
End If
If Val(Text5.Text) > 0 Then
    MsgBox "You Already Have Entry There"
Else
With Grid1
Text1.Text = .TextMatrix(thisrow, 1)
Text17.Text = Format(Blm1.ITEMstocks(Val(.TextMatrix(thisrow, 1)), Date1.Value), "#.000")
Text2.Text = .TextMatrix(thisrow, 2)
Text3.Text = .TextMatrix(thisrow, 3)
Text4.Text = .TextMatrix(thisrow, 4)
Text5.Text = .TextMatrix(thisrow, 5)
Text18.Text = .TextMatrix(thisrow, 6)
Text19.Text = .TextMatrix(thisrow, 7)
Text6.Text = .TextMatrix(thisrow, 8)
For R = 0 To Combo4.ListCount - 1
    If Combo4.ItemData(R) = Val(.TextMatrix(thisrow, 9)) Then
        Combo4.ListIndex = R
        Exit For
    End If
Next R
End With
'MsgBox thisrow
If Grid1.Rows = 2 Then
    Grid1.Rows = 1
Else
    Grid1.RemoveItem (thisrow)
End If
Dim i As Long

For i = 1 To Grid1.Rows - 1
    Grid1.TextMatrix(i, 0) = i
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

Private Sub Text23_Validate(Cancel As Boolean)
If Val(Text23.Text) > 0 Then
    JobBalance
End If
End Sub

Private Sub Text25_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF1 Then
            Load Search3
            Search3.Text3.Text = 3
            Search3.Show vbModal
        End If
End Sub

Private Sub Text25_Validate(Cancel As Boolean)
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

Private Sub Text4_GotFocus()
On Error Resume Next
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
Me.ActiveControl.BackColor = vbYellow
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

For R = 1 To Grid1.Rows - 1
    If Val(Grid1.TextMatrix(R, 1)) = Val(Text1.Text) Then
        qtyinbill = Val(Grid1.TextMatrix(R, 5))
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
Text9.Text = Text7.Text
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
    MsgBox "Please Give Recvd. Amount"
    Cancel = True
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
Dim cntl As Control
Dim Qty As Double
Dim p As Integer
Dim f As Integer, s As Integer

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

If Len(Text24.Text) > 0 Then
If Trim(Text24.Text) <> "NIL" Then
p = InStr(1, Text24.Text, "/", vbBinaryCompare)
f = Val(Mid(Text24.Text, 1, p - 1))
End If
End If
'MsgBox "First " & f
s = Val(Mid(Text24.Text, p + 1, Len(Text24.Text) - (p)))
'MsgBox "Second " & s
'If Trim(Text10.Text) = "3/5" Then
'    MsgBox Val(Text7.Text) & " " & Val(Text9.Text)
If f > 0 And s > 0 Then
    If s = 5 Then
        Text31.Text = Round((Val(Text5.Text) / 400) * f)
    Else
        Text31.Text = Round((Val(Text5.Text) / 800) * f)
        'MsgBox "Test"
    End If
    
End If


If Val(Text4.Text) > 0 And Val(Text5.Text) > 0 Then
ditm = 0
'MsgBox ditm
Text6.Text = (Val(Text4.Text) * Val(Text5.Text)) - ditm
Text6.Text = Val(Text6.Text) + (Val(Text6.Text) * Val(Text19.Text) / 100)
End If
If Grid1.Rows > 1 Then
    Command1.Enabled = True
    Command7.Enabled = True
Else
    Command1.Enabled = False
    Command7.Enabled = False
End If
amt = 0

For i = 1 To Grid1.Rows - 1
    amt = amt + Val(Grid1.TextMatrix(i, 8))
    tu = tu + Val(Grid1.TextMatrix(i, 5))
    Qty = Qty + Val(Grid1.TextMatrix(i, 6))
    'itypes = itypes + Val(Grid1.TextMatrix(i, 0))
    
Next i
balamt = 0
Text7.Text = Format(amt, "#.00")
Text15.Text = Format(tu, "#.00")
Label30.Caption = Qty
Text16.Text = Format(Grid1.Rows - 1, "#.00")
'disamt = (Val(Text7.Text) / 100) * Val(Text10.Text)
amtdue = Val(Text7.Text) - disamt
balamt = amtdue
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
