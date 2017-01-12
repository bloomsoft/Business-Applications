VERSION 5.00
Begin VB.Form newapp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a New Application Like It"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Create"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Give the Title of New Application"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "newapp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
'Dim fs As New FileSystemObject
If fs.FileExists("C:\Soft\" & Text1.Text) Then
    MsgBox "With this Name Application Already Exist...."
Else
    fs.CopyFolder "C:\Soft\AccountsNew", "C:\Soft\" & Text1.Text, True
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub
