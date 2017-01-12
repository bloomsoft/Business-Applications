VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form InvPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "InvPrint.frx":0000
      Left            =   2040
      List            =   "InvPrint.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2040
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "To Invoice #"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "From Invoice #"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   975
      Left            =   3120
      Picture         =   "InvPrint.frx":006B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Prev"
      Height          =   975
      Left            =   240
      Picture         =   "InvPrint.frx":05CC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice No."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Select Month"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "InvPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function CheckGST(b As Long, invtype As Byte) As Boolean
Dim db As Database
Dim tb As Recordset

Dim ssql As String
Set db = OpenDatabase(App.Path & "\Bloom.mdb")
ssql = "select Sum(GST) as G from Sale_2 where Inv_no=" & b & " and Inv_type= " & invtype
'MsgBox ssql
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("G").Value) Then

    If tb.Fields("G").Value > 0 Then
        CheckGST = True
    Else
        CheckGST = False
    End If
Else
    CheckGST = False
End If
tb.Close
db.Close

End Function
Private Sub Command1_Click()
Dim f As String
Dim b As Boolean
Screen.MousePointer = vbHourglass
    If Val(Text2.Text) = 1 Then
        f = "{Sale_Vw_Final.Inv_no} = " & Val(Text1.Text)
        f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        r1.ReportFileName = App.Path & "\Invoice.rpt"
        r1.DataFiles(0) = App.Path & "\BLOOM.MDB"
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    
    If Val(Text2.Text) = 4 Then
        f = "{Sale_Vw_Final.Inv_no} >= " & Val(Text3.Text)
        f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        f = f & " and {Sale_Vw_Final.Inv_no} <= " & Val(Text4.Text)
        r1.ReportFileName = App.Path & "\Invoice.rpt"
        r1.DataFiles(0) = App.Path & "\BLOOM.MDB"
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text2.Text) = 8 Then
        f = "{PContractVW.Cont_No} = " & Val(Text1.Text)
         'f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        r1.ReportFileName = App.Path & "\PContract.rpt"
        r1.DataFiles(0) = App.Path & "\BLOOM.MDB"
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text2.Text) = 9 Then
        f = "{SContractVW.Cont_No} = " & Val(Text1.Text)
        'f = f & " and {Sale_Vw_Final.MNO} = " & Combo3.ItemData(Combo3.ListIndex)
        r1.ReportFileName = App.Path & "\SContract.rpt"
        r1.DataFiles(0) = App.Path & "\BLOOM.MDB"
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
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Combo3.ListIndex = 0
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub
