VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4035
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6330
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2785.029
   ScaleMode       =   0  'User
   ScaleWidth      =   5944.197
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Contact Us :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   4935
      Begin VB.Label Label4 
         Caption         =   "Phone No              :  041-538500"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   3870
      End
      Begin VB.Label Label3 
         Caption         =   "Senior Programmer : Muhammad Sajjad"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "Softlogic@yahoo.com"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   " Emial          : arshad_paracha@yahoo.com"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   975
         Width           =   3255
      End
      Begin VB.Label lblDisclaimer 
         Caption         =   "Project Manager    :  Muhammad Arshad Farooq Paracha        "
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5400
      TabIndex        =   0
      Top             =   3120
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5859.682
      Y1              =   1408.044
      Y2              =   1408.044
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":030A
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Export System "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5859.682
      Y1              =   1408.044
      Y2              =   1408.044
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   1440
      TabIndex        =   4
      Top             =   780
      Width           =   1005
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
