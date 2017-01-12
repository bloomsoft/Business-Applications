VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cont_p_c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Contract Entry (Dyeing)"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "CONT_3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   34
      Top             =   6240
      Width           =   6735
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3600
         TabIndex        =   38
         Text            =   "Combo3"
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   360
         TabIndex        =   37
         Text            =   "Combo2"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label16 
         Caption         =   "Item List"
         Height          =   255
         Left            =   3600
         TabIndex        =   36
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "A/c List"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   6735
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   13
         Top             =   3960
         Width           =   5055
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   3840
         TabIndex        =   12
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "CONT_3.frx":030A
         Left            =   1440
         List            =   "CONT_3.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Co&mplete"
         Height          =   375
         Left            =   5640
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1800
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   4440
         Width           =   4215
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   2520
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   2520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   49610755
         CurrentDate     =   36801
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   49610755
         CurrentDate     =   36801
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "GST Reg #"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Gst Ratio"
         Height          =   255
         Left            =   3000
         TabIndex        =   43
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "GST"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   3525
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Quantity Cloth"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Reference"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Color Description"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Payment"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Del Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Rate"
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Fabric Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Fabric Code"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   3360
      TabIndex        =   21
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   2400
         Picture         =   "CONT_3.frx":0321
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Clear"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   1320
         Picture         =   "CONT_3.frx":062B
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   975
         Left            =   240
         Picture         =   "CONT_3.frx":0935
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   1920
         Picture         =   "CONT_3.frx":0C3F
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   360
         Picture         =   "CONT_3.frx":0F49
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
End
Attribute VB_Name = "cont_p_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private Sub clear()
Dim cntl As Control
For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
    
    
Next
combs
Check2.Value = 0
Text1.Text = max1
Combo4.ListIndex = 0
End Sub
Private Sub edit1()
Dim tb As ADODB.Recordset
Dim ssql As String
ssql = "SELECT * FROM CONT_1 WHERE E_Type = 3 and CONT_NO = " & Val(Text1.Text)
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("V_DATE").Value
    date2.Value = tb.Fields("DEL_DATE").Value
    Text2.Text = tb.Fields("PARTY").Value
    Text3.Text = blm.party1(tb.Fields("PARTY").Value)
    Text4.Text = tb.Fields("ITEM").Value
    Text5.Text = blm.Cloth(tb.Fields("ITEM").Value)
    Text6.Text = tb.Fields("RATE").Value
    Text9.Text = tb.Fields("CQUANTITY").Value
    Text10.Text = tb.Fields("PAYMENT").Value & ""
    Text11.Text = tb.Fields("REMARKS").Value & ""
    If Not IsNull(tb.Fields("REFERENCE").Value) Then
    Combo1.ListIndex = tb.Fields("REFERENCE").Value - 1
    End If
    If Not IsNull(tb.Fields("Complete").Value) Then
        Check2.Value = tb.Fields("Complete").Value
    End If
'    MsgBox tb.Fields("GST").Value
    Combo4.ListIndex = tb.Fields("GST").Value - 1
     Text12.Text = tb.Fields("GST_Rate").Value
     Text13.Text = tb.Fields("GST_No").Value
    
End If
tb.Close

End Sub
Private Sub save()
Dim tb As New ADODB.Recordset
If Option2 = True Then
    Dim ssql As String
        ssql = "DELETE FROM CONT_1 WHERE E_Type=3 and CONT_NO = " & Val(Text1.Text)
        CN.Execute ssql
End If

tb.Open "CONT_1", CN, 0, 3, 0
tb.AddNew
    tb.Fields("CONT_NO").Value = Val(Text1.Text)
    tb.Fields("E_TYPE").Value = 3
    tb.Fields("V_DATE").Value = date1.Value
    tb.Fields("DEL_DATE").Value = date2.Value
    tb.Fields("PARTY").Value = Val(Text2.Text)
    tb.Fields("ITEM").Value = Val(Text4.Text)
    tb.Fields("RATE").Value = Val(Text6.Text)
    tb.Fields("CQUANTITY").Value = Val(Text9.Text)
    tb.Fields("PAYMENT").Value = Text10.Text
    tb.Fields("REMARKS").Value = UCase(CStr(Text11.Text))
    If Combo1.ListIndex > -1 Then
    tb.Fields("REFERENCE").Value = Combo1.ItemData(Combo1.ListIndex)
    End If
    tb.Fields("Complete").Value = Check2.Value
    tb.Fields("GST").Value = Combo4.ItemData(Combo4.ListIndex)
    tb.Fields("GST_Rate").Value = Val(Text12.Text)
    tb.Fields("GST_No").Value = Text13.Text
    
tb.Update
tb.Close

End Sub
Private Function max1() As Long
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select MAX(CONT_NO) AS C FROM CONT_1 where e_type=3"
Set tb = CN.Execute(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
End Function

Private Sub combs()
Dim ssql As String

ssql = "select * from emp1 order by Emp_no"
blm.fill_comb ssql, Combo1, "name", "Emp_no"

ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo2, "name", "code"

ssql = "select * from Cloth order by code"
blm.fill_comb ssql, Combo3, "name", "code" ', "wIDTH"

End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
If Check1.Value = 0 Then
    Text2.Text = Combo2.ItemData(Combo2.ListIndex)
    Text3.Text = Combo2.Text
End If
If Check1.Value = 1 Then
    Text8.Text = Combo2.ItemData(Combo2.ListIndex)
    Text9.Text = Combo2.Text
End If

End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Check1.Value = 0 Then
    Text2.SetFocus
Else
    Text8.SetFocus
End If
End If
End Sub

Private Sub Combo2_LostFocus()
Check1.Value = 0
End Sub

Private Sub Combo3_Click()
If Combo3.ListCount > 0 Then
    Text4.Text = Combo3.ItemData(Combo3.ListIndex)
    Text5.Text = Combo3.Text
End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If

End Sub

Private Sub Command1_Click()
Call save
Command2_Click
End Sub

Private Sub Command2_Click()
Call clear

If Option2 = True Then
    Text1.SetFocus
Else
    Text2.SetFocus
End If

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If

End Sub

Private Sub Form_Load()
Me.Top = ((Screen.Height - Me.Height) / 2) - 500
Me.Left = (Screen.Width - Me.Width) / 2
combs
Text1.Text = max1
date1.Value = Date
date2.Value = Date
Combo4.ListIndex = 0
End Sub

Private Sub Option1_Click()
clear
Check2.Visible = False
Text1.Enabled = False
Text2.SetFocus
End Sub

Private Sub Option2_Click()
clear
Check2.Visible = True
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Option2 = True Then
    If Val(Text1.Text) > 0 Then
        Call edit1
    Else
        Cancel = True
    End If
End If
End Sub

Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
End Sub

Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12.Text)

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        KeyAscii = 0
        Beep
End If

End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text3.Text = blm.party1(Val(Text2.Text))
        If Text3.Text = "NOT" Then
            Cancel = True
        End If
Else
        Cancel = True
End If
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo3.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) > 0 Then
    Text5.Text = blm.Cloth(Val(Text4.Text))
        If Text5.Text = "NOT" Then
            Cancel = True
        End If
Else
        Cancel = True
End If

End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        KeyAscii = 0
        Beep
End If

End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Val(Text6.Text) > 0 Then
    Exit Sub
Else
    Cancel = True
End If
End Sub


Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        KeyAscii = 0
        Beep
End If


End Sub
