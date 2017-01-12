VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4875
      TabIndex        =   7
      Top             =   4080
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4095
      TabIndex        =   6
      Top             =   4080
      Width           =   720
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   4095
      Width           =   1545
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3960
      Top             =   1440
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Drive"
      Height          =   210
      Left            =   1215
      TabIndex        =   4
      Top             =   4125
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "BloomSoft Technologies"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 4.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3720
      TabIndex        =   3
      Top             =   2400
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed to : Mr. Qasim Afzaal."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   -15
      TabIndex        =   2
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grey, Processing  && Dying"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   15
      TabIndex        =   1
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   4170
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shown As Boolean
Private FS As New FileSystemObject
Private Sub SavePath()
    Dim TS As TextStream
    
    Set TS = FS.CreateTextFile(App.Path & "\Path.txt", True)
    TS.Write Drive1.Drive
    TS.Close
    
End Sub
Private Function GetPath() As String
    Dim TS As TextStream
    Dim K As String
    If FS.FileExists(App.Path & "\Path.txt") Then
    Set TS = FS.OpenTextFile(App.Path & "\Path.txt", ForReading)
    K = TS.ReadAll
    TS.Close
    End If
    GetPath = K
End Function

Private Sub Command1_Click()
DatabaseDrive = Drive1.Drive
SavePath
Me.Hide
MDIForm1.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
Dim F As String
F = GetPath
If Len(F) > 0 Then
    Drive1.Drive = F
End If
End Sub

