VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5205
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   2415
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:0092-41-639075 E:mail:bloomsoft@onebox.com"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:0092-300-9660066  Off: 0092-41-614804"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2000-2001 BloomSoft Technologies. All Rights Reserved."
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
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   5145
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Knitting Unit Manager"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   585
      Left            =   -135
      TabIndex        =   5
      Top             =   3840
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   5640
      Picture         =   "frmSplash.frx":37696
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   4
      Top             =   4800
      Width           =   1515
   End
   Begin VB.Label ll 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Registered to:                  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Image img1 
      Height          =   585
      Left            =   120
      Picture         =   "frmSplash.frx":3D678
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3120
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BloomSoft"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   3300
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blm As bloom_r

Private Sub Form_Activate()
Dim d As Integer
d = Second(Time())
Do While Not ((Second(Time()) - d) >= 4 Or (Second(Time()) - d) <= -4)
    DoEvents
Loop
Unload Me
    MDIForm1.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    MDIForm1.Show
End Sub

Private Sub Form_Load()
Set blm = New bloom_r
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
'    img1.Picture = LoadPicture(blm.report_path & "weave.gif")
    ll.Caption = ll.Caption & " Alpha Knitting Faisalabad, Pakistan. "
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub Frame1_Click()
    Unload Me
    MDIForm1.Show
End Sub

Private Sub lblLicenseTo_Click()

End Sub

