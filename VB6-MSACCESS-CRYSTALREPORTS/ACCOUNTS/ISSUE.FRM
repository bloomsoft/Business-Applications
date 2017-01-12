VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form issue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Issue Note"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   4455
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   960
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   36755
      End
      Begin VB.Label Label6 
         Caption         =   "Purpose"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Rate"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Issue #"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   600
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private Function edit() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select * from issue where issue_no = " & Val(Text1.Text)

Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("v_date").Value
    Text1.Text = tb.Fields("issue_no").Value
    Text2.Text = tb.Fields("quantity").Value
    Text3.Text = tb.Fields("rate").Value
    Text4.Text = tb.Fields("purpose").Value
    Combo1.ListIndex = tb.Fields("item").Value - 1
    edit = True
Else
    MsgBox "Not Found ...!"
    edit = False
End If
tb.Close
db.Close

    
End Function

Private Sub clear()
Dim cntl As Control
Dim ssql As String
For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
ssql = "select * from item "
blm.fill_comb ssql, Combo1, "Name", "Code"
End Sub

Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

    
If Option2 = True Then
   Set db = OpenDatabase(blm.patHmain)
    ssql = "delete from issue where issue_no = " & Val(Text1.Text)
    db.Execute ssql
   db.Close
End If
Set db = OpenDatabase(blm.patHmain)
    Set tb = db.OpenRecordset("issue", dbOpenTable)
        tb.AddNew
            tb.Fields("v_date").Value = date1.Value
            tb.Fields("issue_no").Value = Val(Text1.Text)
            tb.Fields("item").Value = Combo1.ItemData(Combo1.ListIndex)
            tb.Fields("quantity").Value = Val(Text2.Text)
            tb.Fields("rate").Value = Val(Text3.Text)
            tb.Fields("purpose").Value = UCase(CStr(Text4.Text))
        tb.Update
    tb.Close
    db.Close

End Sub
Private Function max1() As Double
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
    
    ssql = "select max(issue_no)as c from issue"
    Set db = OpenDatabase(blm.patHmain)
    Set tb = db.OpenRecordset(ssql)
    If IsNull(tb.Fields("c").Value) = False Then
        max1 = tb.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    tb.Close
    db.Close
End Function

Private Sub Combo1_GotFocus()
If Option1 = True Then
    Text1.Text = max1
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command1_Click()
Call save
Command2_Click
End Sub

Private Sub Command2_Click()
Call clear
If Option1 = True Then
    Combo1.SetFocus
Else
Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Dim ssql As String

Me.Top = ((Screen.Height - Me.Height) / 2) - 500
Me.Left = (Screen.Width - Me.Width) / 2

ssql = "select * from item "
blm.fill_comb ssql, Combo1, "Name", "Code"

End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Combo1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text1.Text) > 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim b As Boolean
If Val(Text1.Text) > 0 Then
    b = edit
    If b = False Then
        Cancel = True
    End If
Else
   ' Cancel = True
End If
    
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text2.Text) > 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Exit Sub
Else
    Cancel = True
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text3.Text) > 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Val(Text3.Text) > 0 Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
    Cancel = True
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub
