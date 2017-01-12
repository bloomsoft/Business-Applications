VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Specific Camera"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Set"
      Height          =   345
      Left            =   3240
      TabIndex        =   7
      Top             =   1110
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Server IP"
      Height          =   405
      Left            =   1770
      TabIndex        =   6
      Top             =   645
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1245
      TabIndex        =   4
      Text            =   "210.56.22.122"
      Top             =   1125
      Width           =   1620
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   405
      Left            =   3270
      TabIndex        =   3
      Top             =   630
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   405
      Left            =   255
      TabIndex        =   2
      Top             =   645
      Width           =   1125
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1635
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   165
      Width           =   2760
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2820
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Server IP"
      Height          =   255
      Left            =   165
      TabIndex        =   5
      Top             =   1125
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Select Camera"
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   1530
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim S As String
Dim p As Integer
S = Text1.Text
Winsock1.Close
Sleep 2
Winsock1.Connect S, 9000
Sleep 4
If Winsock1.State = sckConnected Then
    'MsgBox Combo1.ItemData(Combo1.ListIndex)
    Winsock1.SendData "C:" & Combo1.ItemData(Combo1.ListIndex)
End If
Sleep 2
Winsock1.Close
Me.Hide
Unload Me
End Sub
Private Sub Sleep(Seconds As Integer)
Dim OLDDate As Date
OLDDate = Now
Dim D As Long

Do
D = DateDiff("s", Now, OLDDate)
If Abs(D) >= Seconds Then
    
    Exit Do
End If
DoEvents
Loop

End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command3_Click()
Me.Height = Me.Height + 400

End Sub

Private Sub Command4_Click()
Me.Height = Me.Height - 400
End Sub

Private Sub Form_Load()
Text1.Text = ServerIP
Dim R As Integer
Combo1.Clear
For R = 1 To 32
    Combo1.AddItem "Camera " & R
    Combo1.ItemData(Combo1.NewIndex) = R
Next R
Combo1.ListIndex = 0

End Sub
