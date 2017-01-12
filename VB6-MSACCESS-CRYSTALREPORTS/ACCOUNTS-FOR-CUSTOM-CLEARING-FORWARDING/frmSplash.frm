VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6195
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblSerialNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2006-2007 BloomSoft Technologies. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   4515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Clearing Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   3990
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRegisterTo 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered to: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   4935
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BloomSoft"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1020
      Left            =   0
      TabIndex        =   0
      Top             =   4245
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   3765
      Left            =   255
      Picture         =   "frmSplash.frx":000C
      Top             =   405
      Width           =   3705
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blm As Bloom_r

Private Sub Form_Activate()
Dim d As Integer
d = Second(Time())
Do While Not ((Second(Time()) - d) >= 4 Or (Second(Time()) - d) <= -4)
    DoEvents
Loop
Me.Hide
Unload Me
Load MDIForm1
   MDIForm1.Show
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Unload Me
    'MDIForm1.Show
End Sub

Private Sub Form_Load()
Set blm = New Bloom_r
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
'    img1.Picture = LoadPicture(blm.report_path & "weave.gif")
    lblRegisterTo.Caption = lblRegisterTo.Caption & " MAYO ENTERPRISES"
    'lblSerialNo.Caption = lblSerialNo.Caption & CPUID
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub Frame1_Click()
Me.Hide
Unload Me
Load MDIForm1
   MDIForm1.Show
End Sub

Private Sub lblLicenseTo_Click()

End Sub

Private Sub ll_Click()

End Sub

