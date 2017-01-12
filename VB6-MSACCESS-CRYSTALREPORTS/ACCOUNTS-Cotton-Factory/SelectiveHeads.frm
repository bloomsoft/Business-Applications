VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SelectiveHeads 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "SelectiveHeads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7155
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   4380
      TabIndex        =   10
      Top             =   420
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   66912259
      CurrentDate     =   38224
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   6390
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin Crystal.CrystalReport R1 
      Left            =   1830
      Top             =   3630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      Left            =   180
      TabIndex        =   2
      Top             =   5040
      Width           =   6855
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         CausesValidation=   0   'False
         Height          =   915
         Left            =   5220
         Picture         =   "SelectiveHeads.frx":74F2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   1185
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refill"
         Height          =   915
         Left            =   3900
         Picture         =   "SelectiveHeads.frx":77FC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Preview"
         Height          =   915
         Left            =   2670
         Picture         =   "SelectiveHeads.frx":7C3E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Un Select All"
         Height          =   915
         Left            =   1530
         Picture         =   "SelectiveHeads.frx":8908
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Select All"
         Height          =   915
         Left            =   390
         Picture         =   "SelectiveHeads.frx":95D2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.ListBox List1 
      Height          =   4110
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   780
      Width           =   6795
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   4380
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   66912259
      CurrentDate     =   38224
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Up To :"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Heads "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2955
   End
End
Attribute VB_Name = "SelectiveHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New bloom_r
Private Blm1 As New bloom1
Private Sub FillListBox()
Dim RST As Recordset
Dim DBM As Database
Dim Ssql As String
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Ssql = "Select * from Heads where Code<99 Order by Name"
Set RST = DBM.OpenRecordset(Ssql)
List1.clear
If Not RST.EOF Then
    Do While Not RST.EOF
        List1.AddItem RST.Fields("Code").Value & " - " & RST.Fields("Name").Value & ""
        List1.ItemData(List1.NewIndex) = RST.Fields("Code").Value
        RST.MoveNext
    Loop
End If
RST.Close

End Sub

Private Sub Command1_Click()
Dim R As Integer
For R = 0 To List1.ListCount - 1
    List1.Selected(R) = True
Next R
End Sub

Private Sub Command2_Click()
Dim R As Integer
For R = 0 To List1.ListCount - 1
    List1.Selected(R) = False
Next R

End Sub

Private Sub Command3_Click()
Dim R As Integer
Dim F As String
Dim S As String
Screen.MousePointer = vbHourglass
For R = 0 To List1.ListCount - 1
    If List1.Selected(R) = True Then
        S = S & List1.ItemData(R) & ","
    End If
Next R
If Len(S) <= 0 Then Exit Sub
F = Left(S, Len(S) - 1)
If Val(Text1.Text) = 1 Then
    Blmr.trial2 DTPicker2.Value, DTPicker1.Value, ProgressBar1, F
    r1.ReportFileName = App.path & "\Trial3.rpt"
    r1.DataFiles(0) = App.path & "\Book.mdb"
    r1.WindowTitle = "Closing Balances of All Accounts in a Control Head"
    r1.Action = 1
End If
    
Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()
FillListBox
End Sub

Private Sub Command5_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
FillListBox
DTPicker1.Value = Date
DTPicker2.Value = FStartDate
End Sub

