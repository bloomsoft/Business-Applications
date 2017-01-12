VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PItem 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   4455
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   855
         Left            =   3240
         Picture         =   "PItem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   1680
         Picture         =   "PItem.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Preview"
         Height          =   855
         Left            =   240
         Picture         =   "PItem.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   4455
      Begin Crystal.CrystalReport R1 
         Left            =   3120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Item Code"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Name"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3480
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Party Name"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Party Code"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56492035
         CurrentDate     =   37718
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56492035
         CurrentDate     =   37718
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label7 
      Caption         =   "(F2) to Search Party, (F3) to Search Item"
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   5640
      Width           =   3015
   End
End
Attribute VB_Name = "PItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1

Private Sub Command1_Click()
Dim f As String
Screen.MousePointer = vbHourglass
If Val(Text5.Text) = 1 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Sale_VW_Final.party}=" & Val(Text1.Text)
        f = f & " and {Sale_VW_Final.Item}=" & Val(Text4.Text)
'        MsgBox f
        r1.ReportFileName = App.Path & "\SaleItempartyW.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
End If

If Val(Text5.Text) = 2 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Sale_VW_Final.party}=" & Val(Text1.Text)
        f = f & " and {Sale_VW_Final.Item}=" & Val(Text4.Text)
'        MsgBox f
        r1.ReportFileName = App.Path & "\SaleItempartyS.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
End If

If Val(Text5.Text) = 3 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Sale_VW_Final.party}=" & Val(Text1.Text)
        f = f & " and {Sale_VW_Final.Item}=" & Val(Text4.Text)
'        MsgBox f
        r1.ReportFileName = App.Path & "\SaleItempartyT.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    Load Search2
    Search2.Text3.Text = 11
    Search2.Show
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    Text2.Text = Blm1.party1(Val(Text1.Text))
    If Text2.Text = "NOT" Then
        MsgBox "Invalid Account Code"
        
    End If
    
End If

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Load Search1
    Search1.Text3.Text = 11
    Search1.Show
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) > 0 Then
    Text3.Text = Blm1.Item1(Val(Text4.Text))
    If Text3.Text = "NOT" Then
        MsgBox "Invalid Item Code"
        
    End If
    
End If
End Sub
