VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form kachi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kachi Parchi"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "&Change"
      Height          =   255
      Left            =   4200
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&New"
      Height          =   255
      Left            =   2040
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   6615
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5160
         Top             =   3240
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   4680
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   375
         Left            =   3960
         TabIndex        =   43
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Top             =   3720
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   6000
         TabIndex        =   10
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5640
         TabIndex        =   7
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   2280
         Width           =   5175
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   37
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   36749
      End
      Begin VB.Label Label19 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Contract Date"
         Height          =   255
         Left            =   2760
         TabIndex        =   42
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Broker A/c Title"
         Height          =   255
         Left            =   2760
         TabIndex        =   38
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Net Gazana"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Less Gazana"
         Height          =   255
         Left            =   2880
         TabIndex        =   35
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Pakki/P. Date"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Pakki Parchi #"
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Rejection"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Gazan Rec."
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "L/Kami ratio"
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Thans"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Width"
         Height          =   255
         Left            =   5040
         TabIndex        =   28
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Quality"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Broker A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Kachi Parchi #"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "kachi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1


Private Function edit1() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select * from contract where cont_no = " & Val(Text2.Text)
ssql = ssql & " and e_type = 2"

Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date3.Value = tb.Fields("v_date").Value
    Text3.Text = tb.Fields("party").Value
    Text4.Text = blm.party1(tb.Fields("party").Value)
    If tb.Fields("BROKER").Value > 0 Then
            Text5.Text = tb.Fields("BROKER").Value
            Text6.Text = blm.party1(tb.Fields("broker").Value)
    End If
    Text7.Text = tb.Fields("item").Value
    Text8.Text = blm.item1(tb.Fields("item").Value)
    edit1 = True
Else
    MsgBox "Not Found ...!"
    edit1 = False
End If
tb.Close
db.Close

    
End Function
Private Function max1() As Double
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
    
    ssql = "select max(kachi_no)as c from kachi"
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
Private Function edit_kachi() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select * from kachi where kachi_no = " & Val(Text1.Text)

Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("v_date").Value
    Text2.Text = tb.Fields("cont_no").Value
    Text10.Text = tb.Fields("thans").Value
    Text11.Text = tb.Fields("gazana").Value
    Text14.Text = tb.Fields("pakki_no").Value
    date2.Value = tb.Fields("pakki_date").Value
    Text13.Text = tb.Fields("rejection").Value
    Text12.Text = tb.Fields("lessration").Value
    Text15.Text = tb.Fields("less").Value
    Text17.Text = tb.Fields("remarks").Value
    edit_kachi = True
Else
    MsgBox "Not Found ...!"
    edit_kachi = False
End If
tb.Close
db.Close

    
End Function

Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Set db = OpenDatabase(blm.patHmain)
If Option2 = True Then
    ssql = "delete from kachi where kachi_no = " & Val(Text1.Text)
    db.Execute ssql
End If
db.Close
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset("kachi", dbOpenTable)
tb.AddNew
    tb.Fields("kachi_no").Value = Val(Text1.Text)
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("cont_no").Value = Val(Text2.Text)
    tb.Fields("thans").Value = Val(Text10.Text)
    tb.Fields("gazana").Value = Val(Text11.Text)
    tb.Fields("pakki_no").Value = Val(Text14.Text)
    tb.Fields("pakki_date").Value = date2.Value
    tb.Fields("rejection").Value = Val(Text13.Text)
    tb.Fields("lessration").Value = CStr(Text12.Text)
    tb.Fields("less").Value = Val(Text15.Text)
    tb.Fields("remarks").Value = CStr(Text17.Text)
tb.Update
tb.Close
db.Close

    
End Sub

Private Sub clear()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
End Sub

Private Sub Command1_Click()
Call save
Command2_Click

End Sub

Private Sub Command2_Click()
Call clear
Text2.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide

End Sub

Private Sub Form_Load()
date1.Value = Date
date2.Value = Date
date3.Value = Date
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Text2.SetFocus
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
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim b As Boolean
If Val(Text1.Text) > 0 Then
    If Option2 = True Then
        b = edit_kachi
        If b = False Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text10_Validate(Cancel As Boolean)
If Val(Text10.Text) > 0 Then
    Exit Sub
Else
    Cancel = True
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
If Val(Text11.Text) > 0 Then
    Exit Sub
Else
    Cancel = True
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Text2_GotFocus()
If Option1 = True Then
    Text1.Text = max1
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim b As Boolean

If Val(Text2.Text) > 0 Then
    b = edit1
    If b = False Then
        Cancel = True
    End If
End If
End Sub

Private Sub Timer1_Timer()
Text16.Text = Val(Text11.Text) - Val(Text13.Text) - Val(Text15.Text)
End Sub
