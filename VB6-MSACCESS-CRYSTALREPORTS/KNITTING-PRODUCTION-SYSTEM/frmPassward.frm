VERSION 5.00
Begin VB.Form frmPassward 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALPHA KNITTING (pvt) LTD User Security..."
   ClientHeight    =   3570
   ClientLeft      =   2295
   ClientTop       =   1125
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5010
   Begin VB.CommandButton Command2 
      Caption         =   "Ca&ncel"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Security System :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Passward :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Enter user Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label Label5 
      Caption         =   "SoftLogic Software Systems (pvt) Ltd."
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Don't Try to Login Unless you are  an Authorized user."
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Warning :"
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
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "frmPassward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    
    If UCase$(txtName.Text) = "RIZWAN" And UCase$(txtPass.Text) = "RFAZU" Then
        frmMain.Show
        Unload Me
    Else
        MsgBox "Invalid User Name or Passward ,Try Again", vbInformations
        Me.txtName.SetFocus
    End If
    
End Sub

Private Sub Command2_Click()
    Me.txtName.Text = ""
    Me.txtPass.Text = ""
    Me.txtName.SetFocus
End Sub

Private Sub Command3_Click()
    End
    
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPass.SetFocus
    End If
    
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If UCase$(Me.txtName.Text) = "RIZWAN" And UCase$(Me.txtPass.Text) = "RFAZU" Then
            frmMain.Show
            Unload Me
        Else
            MsgBox "Invalid User Name or Passward ,Try Again", vbInformation
            Me.txtName.SetFocus
        End If
  End If
        
End Sub
