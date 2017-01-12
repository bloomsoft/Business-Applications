VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6450
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   6450
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:0092-300-9660066  Off: 0092-41-614804"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2007-2015 BloomSoft Technologies. All Rights Reserved."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1860
      TabIndex        =   2
      Top             =   5970
      Width           =   5145
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts && Inventory Management"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1365
      Left            =   720
      TabIndex        =   1
      Top             =   4800
      Width           =   6945
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BloomSoft"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   2970
      TabIndex        =   0
      Top             =   4320
      Width           =   2610
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Blm As bloom_r

Private Sub Form_Activate()
Dim d As Integer
d = Second(Time())
Do While Not ((Second(Time()) - d) >= 4 Or (Second(Time()) - d) <= -4)
    DoEvents
Loop
'Unload Me
 '   MDIForm1.Show
 Load login2
 login2.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Unload Me
    'MDIForm1.Show
End Sub

Private Sub Form_Load()
Set Blm = New bloom_r
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
'    img1.Picture = LoadPicture(blm.report_path & "weave.gif")
    'll.Caption = ll.Caption & " H.Azam. Enterprises, Faisalabad Pakistan"
    Label4.Caption = "[ " & YearN & "-" & YearN + 1 & " ]"
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set Blm = Nothing
End Sub

Private Sub Frame1_Click()
    Unload Me
    MDIForm1.Show
End Sub

Private Sub lblLicenseTo_Click()

End Sub

