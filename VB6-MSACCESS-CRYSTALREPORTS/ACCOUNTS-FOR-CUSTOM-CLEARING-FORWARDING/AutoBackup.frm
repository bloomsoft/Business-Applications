VERSION 5.00
Begin VB.Form AutoBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Backup Path Settings"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4950
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   900
      TabIndex        =   5
      Top             =   645
      Width           =   3885
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   210
      Width           =   3825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   795
      Left            =   2730
      TabIndex        =   2
      Top             =   4275
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   795
      Left            =   540
      TabIndex        =   1
      Top             =   4275
      Width           =   1470
   End
   Begin VB.Label Label2 
      Caption         =   "Folder"
      Height          =   285
      Left            =   165
      TabIndex        =   4
      Top             =   690
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Drive"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   960
   End
End
Attribute VB_Name = "AutoBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\User.mdb")
Ssql = "Delete from Backup"
DB.Execute Ssql

Set TB = DB.OpenRecordset("Backup", dbOpenTable)
TB.AddNew
    TB.Fields("Path").Value = Dir1.Path
TB.Update
TB.Close
DB.Close
BackupPath = Dir1.Path
MsgBox "Location for AutoBackup Saved!"

End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
End Sub
