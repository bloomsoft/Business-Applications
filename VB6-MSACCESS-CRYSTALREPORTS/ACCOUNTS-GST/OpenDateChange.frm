VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form OpenDateChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening balance Date Change"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4260
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   495
      Left            =   885
      Picture         =   "OpenDateChange.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   690
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   2565
      Picture         =   "OpenDateChange.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   225
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   50069507
      CurrentDate     =   39132
   End
   Begin VB.Label Label1 
      Caption         =   "New Date"
      Height          =   300
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   1170
   End
End
Attribute VB_Name = "OpenDateChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Change the Date of Opening Balances", vbYesNo)
If Result = vbYes Then
    Dim DB As Database
    Dim Ssql As String
    Set DB = OpenDatabase(App.path & "\Bloom.mdb")
    Ssql = "Update Acchart Set OpDate=#" & DTPicker1.Value & "#"
    DB.Execute Ssql
    Ssql = "Update vouDTL Set v_Date=#" & DTPicker1.Value & "# where V_type=10"
    DB.Execute Ssql
    DB.Close
    MsgBox "Date of All of Opening Balances Has been Updated"
    
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
DTPicker1.Value = Date
End Sub
