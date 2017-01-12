VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlarm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Schedual"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   855
      Left            =   210
      Picture         =   "frmAlrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   855
      Left            =   1590
      Picture         =   "frmAlrm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   855
      Left            =   3000
      Picture         =   "frmAlrm.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   3195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3945
      Begin VB.TextBox Text1 
         Height          =   1965
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1080
         Width           =   2955
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   870
         TabIndex        =   4
         Top             =   630
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         _Version        =   393216
         Format          =   66912258
         CurrentDate     =   39237
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   870
         TabIndex        =   2
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   66912259
         CurrentDate     =   39237
      End
      Begin VB.Label Label3 
         Caption         =   "Note"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Time"
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim B As Boolean
Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set TB = DB.OpenRecordset("Alarm", dbOpenTable)
TB.AddNew
    TB.Fields("vdate").Value = DTPicker1.Value
    TB.Fields("vTime").Value = DTPicker2.Value
    TB.Fields("Note").Value = Text1.Text
    
TB.Update
TB.Close
DB.Close
End Sub

Private Sub Command2_Click()
DTPicker1.Value = Date
DTPicker2.Value = Now
Text1.Text = ""

End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
DTPicker2.Value = DTPicker1.Value
End Sub

