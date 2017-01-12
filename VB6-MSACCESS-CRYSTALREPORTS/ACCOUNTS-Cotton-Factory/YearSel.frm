VERSION 5.00
Begin VB.Form YearSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Financial Year Selector"
   ClientHeight    =   1650
   ClientLeft      =   5445
   ClientTop       =   5370
   ClientWidth     =   4680
   Icon            =   "YearSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   825
      Left            =   2520
      Picture         =   "YearSel.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   780
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   825
      Left            =   1470
      Picture         =   "YearSel.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   780
      Width           =   945
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   180
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Select Year"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   1695
   End
End
Attribute VB_Name = "YearSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
YearN = Combo1.ItemData(Combo1.ListIndex)
Me.Hide
frmSplash.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If

End Sub

Private Sub Form_Load()
Dim R As Long
Combo1.clear
For R = 2007 To 2020
    Combo1.AddItem "Jul - " & R & " To Jun - " & R + 1
    Combo1.ItemData(Combo1.NewIndex) = R
Next R

Combo1.ListIndex = (Year(Now)) - Year(Now)
'Combo1.ListIndex = 1
End Sub
