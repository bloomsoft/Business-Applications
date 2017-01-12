VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Timer Setter"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2520
      TabIndex        =   4
      Top             =   690
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   390
      Left            =   240
      TabIndex        =   3
      Top             =   645
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2745
      MaxLength       =   2
      TabIndex        =   1
      Top             =   285
      Width           =   510
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1815
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Seconds"
      Height          =   240
      Left            =   3345
      TabIndex        =   2
      Top             =   300
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Interval Between Camera Change"
      Height          =   225
      Left            =   225
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim S As String
Dim p As Integer
S = ServerIP
Winsock1.Close
Sleep 2
'MsgBox S
Winsock1.Connect S, 9000
Sleep 4
If Winsock1.State = sckConnected Then
    
    Winsock1.SendData "S:" & Val(Text1.Text)
    
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
