VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form pakki 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pakki Parchi"
   ClientHeight    =   7200
   ClientLeft      =   1080
   ClientTop       =   540
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   6480
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Change"
      Height          =   255
      Left            =   5640
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&New"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   240
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Width           =   9615
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3240
         Top             =   1920
      End
      Begin VB.TextBox Text20 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   48
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox Text19 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   46
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text18 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   44
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   16
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   3840
         Width           =   4575
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   10
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   7
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2400
         Width           =   8055
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   8040
         TabIndex        =   43
         Top             =   960
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1440
         TabIndex        =   42
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   36749
      End
      Begin VB.Label Label21 
         Caption         =   "Broker Name"
         Height          =   255
         Left            =   3600
         TabIndex        =   47
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Broker A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Amount Payable"
         Height          =   255
         Left            =   3360
         TabIndex        =   41
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Deduction"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Total Amount"
         Height          =   255
         Left            =   3600
         TabIndex        =   39
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Rate"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   3840
         TabIndex        =   37
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Net Gazana"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Less Gazana"
         Height          =   255
         Left            =   7080
         TabIndex        =   35
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "L/Kami"
         Height          =   255
         Left            =   3960
         TabIndex        =   34
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Rejection"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Rec. Thans"
         Height          =   255
         Left            =   3600
         TabIndex        =   32
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Rec. Gazana"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   3840
         TabIndex        =   29
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Contract No"
         Height          =   255
         Left            =   7080
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Kachi Parchi Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Kachi Parchi #"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Pakki Parchi #"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "pakki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim ssql As String

If Option2 = True Then
Set db = OpenDatabase(blm.patHmain)
ssql = "delete from pakki where pakki_no = " & Val(Text1.Text)
db.Execute ssql
db.Close
End If

Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset("pakki", dbOpenTable)
tb.AddNew
    tb.Fields("pakki_date").Value = date1.Value
    tb.Fields("pakki_no").Value = Val(Text1.Text)
    tb.Fields("kachi_no").Value = Val(Text2.Text)
    tb.Fields("kachi_date").Value = date2.Value
    tb.Fields("cont_no").Value = Val(Text3.Text)
    tb.Fields("party").Value = Val(Text4.Text)
    tb.Fields("item").Value = Val(Text18.Text)
    tb.Fields("rec_gazana").Value = Val(Text7.Text)
    tb.Fields("rec_thans").Value = Val(Text8.Text)
    tb.Fields("rejection").Value = Val(Text9.Text)
    tb.Fields("less").Value = Val(Text10.Text)
    tb.Fields("remarks").Value = CStr(Text13.Text)
    tb.Fields("rate").Value = Val(Text14.Text)
    tb.Fields("deduction").Value = Val(Text16.Text)
tb.Update
tb.Close
db.Close
End Sub
Private Function edit_kachi() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select * from kachi where kachi_no = " & Val(Text2.Text)

Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date2.Value = tb.Fields("v_date").Value
    Text3.Text = tb.Fields("cont_no").Value
    edit1 (Val(Text3.Text))
    Text8.Text = tb.Fields("thans").Value
    Text7.Text = tb.Fields("gazana").Value
    edit_kachi = True
Else
    MsgBox "Not Found ...!"
    edit_kachi = False
End If
tb.Close
db.Close

    
End Function
Private Function edit_pakki() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select * from pakki where pakki_no = " & Val(Text1.Text)

Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("pakki_date").Value
    Text2.Text = tb.Fields("kachi_no").Value
    Text2_Validate (False)
    Text3.Text = tb.Fields("cont_no").Value
    edit1 (Val(Text3.Text))
    Text8.Text = tb.Fields("rec_thans").Value
    Text7.Text = tb.Fields("rec_gazana").Value
    Text9.Text = tb.Fields("rejection").Value
    Text10.Text = tb.Fields("less").Value
    Text14.Text = tb.Fields("rate").Value
    Text16.Text = tb.Fields("deduction").Value
    Text13.Text = tb.Fields("remarks").Value
    edit_pakki = True
Else
    MsgBox "Not Found ...!"
    edit_pakki = False
End If
tb.Close
db.Close

    
End Function

Private Sub edit1(c As Long)




Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select * from contract where cont_no = " & Val(Text2.Text)
ssql = ssql & " and e_type = 2"

Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
   
    Text4.Text = tb.Fields("party").Value
    Text5.Text = blm.party1(tb.Fields("party").Value)
    If tb.Fields("BROKER").Value > 0 Then
            Text19.Text = tb.Fields("BROKER").Value
            Text20.Text = blm.party1(tb.Fields("broker").Value)
    End If
    Text14.Text = tb.Fields("rate").Value
    Text18.Text = tb.Fields("item").Value
    Text6.Text = blm.item1(tb.Fields("item").Value)
    
End If
tb.Close
db.Close
End Sub

Private Function max1() As Double
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
    
    ssql = "select max(pakki_no)as c from pakki"
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
If Option1 = True Then
Text2.SetFocus
End If
If Option2 = True Then
Text1.SetFocus
End If

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Option1_Click()
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
        b = edit_pakki
        If b = False Then
            Cancel = True
            Text1.Text = vbNullString
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

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text8.Text) > 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 And Val(Text8.Text) > 0 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

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
    
        b = edit_kachi
        If b = False Then
            Cancel = True
        End If
    End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
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

Private Sub Timer1_Timer()
Text11.Text = Val(Text9.Text) + Val(Text10.Text)
Text12.Text = Val(Text7.Text) - Val(Text9.Text) - Val(Text10.Text)
Text15.Text = Val(Text12.Text) * Val(Text14.Text)
Text17.Text = Val(Text15.Text) - Val(Text16.Text)
End Sub
