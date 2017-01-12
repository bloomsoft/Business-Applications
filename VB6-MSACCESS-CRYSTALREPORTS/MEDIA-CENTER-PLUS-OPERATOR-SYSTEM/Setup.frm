VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Setup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup "
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive3 
      Height          =   315
      Left            =   4140
      TabIndex        =   26
      Top             =   6075
      Width           =   1965
   End
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   2175
      TabIndex        =   25
      Top             =   6075
      Width           =   1995
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   195
      TabIndex        =   24
      Top             =   6075
      Width           =   1995
   End
   Begin VB.DirListBox Dir3 
      Height          =   2565
      Left            =   4170
      TabIndex        =   23
      Top             =   3465
      Width           =   1890
   End
   Begin VB.DirListBox Dir2 
      Height          =   2565
      Left            =   2175
      TabIndex        =   22
      Top             =   3450
      Width           =   1965
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   180
      TabIndex        =   21
      Top             =   3450
      Width           =   1965
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   780
      Left            =   4395
      Picture         =   "Setup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   780
      Left            =   3165
      Picture         =   "Setup.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1455
      TabIndex        =   8
      Top             =   195
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   3960
      TabIndex        =   9
      Top             =   195
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   300
      Left            =   1455
      TabIndex        =   10
      Top             =   555
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   915
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   300
      Left            =   1440
      TabIndex        =   12
      Top             =   1275
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   300
      Left            =   1440
      TabIndex        =   13
      Top             =   1635
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin MSComCtl2.DTPicker DTPicker7 
      Height          =   300
      Left            =   1440
      TabIndex        =   14
      Top             =   1995
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin MSComCtl2.DTPicker DTPicker8 
      Height          =   300
      Left            =   1425
      TabIndex        =   15
      Top             =   2355
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515074
      CurrentDate     =   39401
   End
   Begin VB.Label Label11 
      Caption         =   "Songs Folder"
      Height          =   270
      Left            =   4275
      TabIndex        =   20
      Top             =   3105
      Width           =   1110
   End
   Begin VB.Label Label10 
      Caption         =   "Naat Folder"
      Height          =   270
      Left            =   2205
      TabIndex        =   19
      Top             =   3105
      Width           =   1110
   End
   Begin VB.Label Label9 
      Caption         =   "Quran Folder"
      Height          =   270
      Left            =   255
      TabIndex        =   18
      Top             =   3105
      Width           =   1110
   End
   Begin VB.Label Label8 
      Caption         =   "Juma"
      Height          =   270
      Left            =   300
      TabIndex        =   7
      Top             =   2355
      Width           =   1110
   End
   Begin VB.Label Label7 
      Caption         =   "Isha"
      Height          =   270
      Left            =   300
      TabIndex        =   6
      Top             =   1995
      Width           =   1110
   End
   Begin VB.Label Label6 
      Caption         =   "Maghrib"
      Height          =   270
      Left            =   300
      TabIndex        =   5
      Top             =   1635
      Width           =   1110
   End
   Begin VB.Label Label5 
      Caption         =   "Asr"
      Height          =   270
      Left            =   300
      TabIndex        =   4
      Top             =   1275
      Width           =   1110
   End
   Begin VB.Label Label4 
      Caption         =   "Duhar"
      Height          =   270
      Left            =   300
      TabIndex        =   3
      Top             =   930
      Width           =   1110
   End
   Begin VB.Label Label3 
      Caption         =   "Fajar"
      Height          =   270
      Left            =   300
      TabIndex        =   2
      Top             =   585
      Width           =   1110
   End
   Begin VB.Label Label2 
      Caption         =   "Stop Time"
      Height          =   270
      Left            =   2895
      TabIndex        =   1
      Top             =   195
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Start Time"
      Height          =   270
      Left            =   285
      TabIndex        =   0
      Top             =   210
      Width           =   1110
   End
End
Attribute VB_Name = "Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Edit1()
On Error Resume Next
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\Main.mdb")
Set TB = DB.OpenRecordset("Select * from Setup")
If Not TB.EOF Then
    DTPicker1.Value = CDate(Date & " " & TB.Fields("StartTime").Value)
    DTPicker2.Value = CDate(Date & " " & TB.Fields("EndTime").Value)
    DTPicker3.Value = CDate(Date & " " & TB.Fields("Fajar").Value)
    DTPicker4.Value = CDate(Date & " " & TB.Fields("Duhar").Value)
    DTPicker5.Value = CDate(Date & " " & TB.Fields("Asr").Value)
    DTPicker6.Value = CDate(Date & " " & TB.Fields("Maghrib").Value)
    DTPicker7.Value = CDate(Date & " " & TB.Fields("Isha").Value)
    DTPicker8.Value = CDate(Date & " " & TB.Fields("Juma").Value)
    
    If Not IsNull(TB.Fields("Quran").Value) Then Dir1.Path = TB.Fields("Quran").Value
    If Not IsNull(TB.Fields("Naat").Value) Then Dir2.Path = TB.Fields("Naat").Value
    If Not IsNull(TB.Fields("Songs").Value) Then Dir3.Path = TB.Fields("Songs").Value
    
    Drive1.Drive = Left(Dir1.Path, 4)
    Drive2.Drive = Left(Dir2.Path, 4)
    Drive3.Drive = Left(Dir3.Path, 4)
End If

TB.Close
DB.Close
End Sub
Private Sub Save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\Main.mdb")
Ssql = "Delete from Setup"
DB.Execute Ssql

Set TB = DB.OpenRecordset("Setup", dbOpenTable)
TB.AddNew
    TB.Fields("StartTime").Value = DTPicker1.Hour & ":" & DTPicker1.Minute & ":" & DTPicker1.Second
    TB.Fields("EndTime").Value = DTPicker2.Hour & ":" & DTPicker2.Minute & ":" & DTPicker2.Second
    TB.Fields("fajar").Value = DTPicker3.Hour & ":" & DTPicker3.Minute & ":" & DTPicker3.Second
    TB.Fields("Duhar").Value = DTPicker4.Hour & ":" & DTPicker4.Minute & ":" & DTPicker4.Second
    TB.Fields("Asr").Value = DTPicker5.Hour & ":" & DTPicker5.Minute & ":" & DTPicker5.Second
    TB.Fields("Maghrib").Value = DTPicker6.Hour & ":" & DTPicker6.Minute & ":" & DTPicker6.Second
    TB.Fields("Isha").Value = DTPicker7.Hour & ":" & DTPicker7.Minute & ":" & DTPicker7.Second
    TB.Fields("Juma").Value = DTPicker8.Hour & ":" & DTPicker8.Minute & ":" & DTPicker8.Second
    
    TB.Fields("Quran").Value = Dir1.Path
    TB.Fields("Naat").Value = Dir2.Path
    TB.Fields("Songs").Value = Dir3.Path
TB.Update
TB.Close
DB.Close
End Sub
Private Sub Command1_Click()
Save
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive2_Change()
Dir2.Path = Drive2.Drive
End Sub

Private Sub Drive3_Change()
Dir3.Path = Drive3.Drive
End Sub

Private Sub Form_Load()
Edit1
End Sub
