VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SelectiveSubHeads 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "SelectiveSubHeads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7155
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      Left            =   180
      TabIndex        =   10
      Top             =   5400
      Width           =   6855
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Select All"
         Height          =   915
         Left            =   390
         Picture         =   "SelectiveSubHeads.frx":74F2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Un Select All"
         Height          =   915
         Left            =   1530
         Picture         =   "SelectiveSubHeads.frx":81BC
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Preview"
         Height          =   915
         Left            =   2670
         Picture         =   "SelectiveSubHeads.frx":8E86
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
         Picture         =   "SelectiveSubHeads.frx":9B50
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   210
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         CausesValidation=   0   'False
         Height          =   915
         Left            =   5220
         Picture         =   "SelectiveSubHeads.frx":9F92
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   1185
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   6855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   4380
      TabIndex        =   6
      Top             =   420
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   66912259
      CurrentDate     =   38224
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   6720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1980
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1155
   End
   Begin Crystal.CrystalReport R1 
      Left            =   570
      Top             =   3720
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
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1620
      Width           =   6795
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   4380
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   66912259
      CurrentDate     =   38224
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Sub Heads"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Up To :"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Heads "
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   540
      Width           =   2955
   End
End
Attribute VB_Name = "SelectiveSubHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New bloom_r
Private Blm As New bloom1
Dim SHCodes() As String
Private Sub FillListBox()
Dim RST As Recordset
Dim DBM As Database
Dim Ssql As String
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Ssql = "Select * from Heads Where Code > 99 "
If Combo1.ListIndex > -1 Then
    Ssql = Ssql & " and Mid(Code,1,2)='" & Combo1.ItemData(Combo1.ListIndex) & "'"
End If
Ssql = Ssql & " Order By Code"
Set RST = DBM.OpenRecordset(Ssql)
List1.clear
If Not RST.EOF Then
    RST.MoveLast
    ReDim SHCodes(RST.RecordCount)
    RST.MoveFirst
    
    Do While Not RST.EOF
        List1.AddItem RST.Fields("Code").Value & " - " & RST.Fields("Name").Value & ""
        SHCodes(List1.NewIndex) = RST.Fields("Code").Value
        RST.MoveNext
    Loop
End If
RST.Close

End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex > -1 Then
FillListBox
End If
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
        S = S & SHCodes(R) & ","
    End If
Next R
'If Len(s) <= 0 Then Exit Sub

If Val(Text1.Text) = 1 And Len(S) > 0 Then
F = Left(S, Len(S) - 1)
Blmr.trial2 DTPicker2.Value, DTPicker1.Value, ProgressBar1, , F
r1.ReportFileName = App.path & "\Trial3.rpt"
r1.DataFiles(0) = App.path & "\Book.mdb"
r1.WindowTitle = "Closing Balances of All Accounts in a Sub-Head"
r1.Action = 1
End If
    
If Val(Text1.Text) = 2 Then
    If Len(S) > 0 Then S = Left(S, Len(S) - 1)
    F = "{SalVoucher.SDate}=Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
    F = F & " and {SalVoucher.EDate}=Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
    If Len(S) > 0 Then
        F = F & " and {SalVoucher.SHCode} in [" & S & "]"
    End If
'    MsgBox F
    r1.ReportFileName = App.path & "\SalSheet.rpt"
    r1.DataFiles(0) = App.path & "\Years\" & YearN & "\Bloom.mdb"
    r1.ReportTitle = "From : " & Format(DTPicker2.Value, "dd-MMM-yyyy") & " To : " & Format(DTPicker1.Value, "dd-MMM-yyyy")
    r1.SelectionFormula = F
    r1.WindowTop = 0
    r1.WindowLeft = 0
    r1.WindowState = crptMaximized
    r1.Action = 1
End If
    
    
    
'If Val(Text1.Text) = 2 Then
'If Check1.Value = 0 Then
'    Blmr.TrialMonths DTPicker2.Value, DTPicker1.Value, s, F
'Else
'    Blmr.TrialMonths DTPicker2.Value, DTPicker1.Value, s, F, Combo2.ItemData(Combo2.ListIndex)
'End If
'With Blmr
'    R1.LogOnServer dll:=.DllName, server:=.DSNName, Database:=.ServiceName, userid:=.UserName, Password:=.REPAssword
'End With
'R1.ReportFileName = App.path & "\Reports\MonthTrial.rpt"
'R1.DataFiles(0) = OrgInfo & UserN & ".MONTHLYTRIAL"
'R1.WindowTitle = "Monthly Trial of All Accounts"
'R1.Action = 1
'End If
    

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
Dim Ssql As String
Ssql = "Select * from Heads where Code <= 99 Order by Code"
Blm.fill_comb2 Ssql, Combo1, "Name", "Code"

FillListBox
DTPicker1.Value = Date
DTPicker2.Value = FStartDate
End Sub
