VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Pur Job"
   ClientHeight    =   6180
   ClientLeft      =   330
   ClientTop       =   720
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11220
   Begin VB.Frame Frame2 
      Caption         =   "Actions"
      Height          =   2415
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command6 
         Caption         =   "&Update"
         Height          =   855
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&New"
         Height          =   855
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "E&xit"
         Height          =   855
         Left            =   2760
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "&Delete"
         Height          =   975
         Left            =   5400
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3720
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   64946177
         CurrentDate     =   39394
      End
      Begin VB.Label Label4 
         Caption         =   "Party Name"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Party Code"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Job No."
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
