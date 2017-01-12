VERSION 5.00
Begin VB.Form ActivationFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activation Form"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "SoftLogic Software System :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.TextBox txt2 
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
         Left            =   3360
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt3 
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
         Left            =   4560
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton ExitBtn 
         Caption         =   "EXIT"
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
         Left            =   4680
         TabIndex        =   4
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton OkBtn 
         Caption         =   "OK"
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
         Left            =   2880
         TabIndex        =   3
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txt1 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   $"ActivationFrm.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label ActivationLbl 
         Caption         =   "Softwate Key :"
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
         Left            =   720
         TabIndex        =   2
         Top             =   2640
         Width           =   1335
      End
   End
End
Attribute VB_Name = "ActivationFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SoftwareKey As String

Private Sub ExitBtn_Click()
    End
End Sub

Private Sub Form_Load()

SoftwareKey = "MIS1AQS2RAZ3"
Dim RegKey As String
   ' SaveSetting "Explorer", "Author", "Key", "www-13"
    RegKey = GetSetting("Explorer", "Author", "Key")
    
    If RegKey = SoftwareKey Then
        Unload Me
       frmPassward.Show
         
    Else
         ActivationFrm.Show
    End If
    
    
End Sub

Private Sub OkBtn_Click()
    Dim UserKey As String
    
    
    SoftwareKey = "MIS1AQS2RAZ3"
    UserKey = txt1.Text + txt2.Text + txt3.Text
    
    If UserKey = SoftwareKey Then
        
        SaveSetting "Explorer", "Author", "Key", SoftwareKey
        frmPassward.Show
        Unload Me
    Else
        MsgBox "Invalid Software Key,Please contact the software vendor", vbCritical
        
    End If
    
End Sub

Private Sub txt1_Change()
If Len(txt1) > Val(4) Then
    txt2.SetFocus
End If
End Sub

Private Sub txt2_Change()
    If Len(txt2) > Val(4) Then
        Me.txt3.SetFocus
    End If
End Sub
