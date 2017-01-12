VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form sale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sale Voucher"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   6375
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   35
         Text            =   "Combo2"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3840
         TabIndex        =   34
         Text            =   "Combo1"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "Items List"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "A/c List"
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   6375
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Stock A/c Title"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Stock A/c Code"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   6375
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1200
         TabIndex        =   33
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24510467
         CurrentDate     =   36764
      End
      Begin VB.Label Label9 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Inv. No."
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Amount"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Voucher #"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   3720
      TabIndex        =   14
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "sale"
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
    ssql = "delete from voumst where v_date = #" & date1.Value & "#"
    ssql = ssql & " and v_type = 5"
    ssql = ssql & " and v_no = " & Val(Text1.Text)
    db.Execute ssql
    ssql = "delete from voudtl where v_date = #" & date1.Value & "#"
    ssql = ssql & " and v_type = 5"
    ssql = ssql & " and v_no = " & Val(Text1.Text)
    db.Execute ssql
    db.Close
End If
Set db = OpenDatabase(blm.patHmain)
Set tb = db.OpenRecordset("voumst", dbOpenTable)
tb.AddNew
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("v_no").Value = Val(Text1.Text)
    tb.Fields("v_type").Value = 5
    tb.Fields("narration").Value = "Sale Voucher " & Val(Text1.Text)
tb.Update
tb.Close
Set tb = db.OpenRecordset("voudtl", dbOpenTable)
tb.AddNew
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("v_no").Value = Val(Text1.Text)
    tb.Fields("v_type").Value = 5
    tb.Fields("party").Value = Val(Text2.Text)
    tb.Fields("item").Value = Val(Text4.Text)
    tb.Fields("debit").Value = Val(Text6.Text)
    tb.Fields("credit").Value = 0
    tb.Fields("remarks").Value = UCase(CStr(Text8.Text))
    tb.Fields("inv_no").Value = CStr(Text7.Text)
    tb.Fields("STOCK").Value = 0
tb.Update
tb.AddNew
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("v_no").Value = Val(Text1.Text)
    tb.Fields("v_type").Value = 5
    tb.Fields("party").Value = Val(Text9.Text)
    tb.Fields("item").Value = Val(Text4.Text)
    tb.Fields("debit").Value = 0
    tb.Fields("credit").Value = Val(Text6.Text)
    tb.Fields("remarks").Value = UCase(CStr(Text8.Text))
    tb.Fields("inv_no").Value = CStr(Text7.Text)
    tb.Fields("STOCK").Value = 1
tb.Update

tb.Close
db.Close
End Sub

Private Function edit1() As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim tb2 As Recordset

Set db = OpenDatabase(blm.patHmain)
ssql = "select * from voumst where v_type = 5"
ssql = ssql & " and v_no = " & Val(Text1.Text)
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("v_date").Value
    
        ssql = "select * from voudtl where v_type = 5"
        ssql = ssql & " and v_no = " & Val(Text1.Text)
        Set tb2 = db.OpenRecordset(ssql)
        If Not tb2.EOF Then
            Do While Not tb2.EOF
            If tb2.Fields("STOCK").Value = 0 Then
            
           ' MsgBox tb2.Fields("PARTY").Value & " " & tb2.Fields("STOCK").Value
                Text2.Text = tb2.Fields("party").Value
                Text3.Text = blm.party1(tb2.Fields("party").Value)
                Text4.Text = tb2.Fields("ITEM").Value
                Text5.Text = blm.item1(tb2.Fields("ITEM").Value)
                Text6.Text = Format(tb2.Fields("debit").Value, "#.00")
                Text7.Text = Format(tb2.Fields("inv_no").Value, "#.00")
                Text8.Text = tb2.Fields("remarks").Value
            End If
            If tb2.Fields("STOCK").Value = 1 Then
            'MsgBox tb2.Fields("PARTY").Value & " " & tb2.Fields("STOCK").Value
                Text9.Text = tb2.Fields("PARTY").Value
                Text10.Text = blm.party1(tb2.Fields("PARTY").Value)
            End If
            tb2.MoveNext
            Loop
        End If
        tb2.Close
        edit1 = True
Else
    MsgBox "No Voucher With this No. in This Type..."
    edit1 = False
End If
tb.Close
db.Close
End Function

Private Sub combs()
Dim ssql As String
ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo1, "name", "code"
ssql = "select * from item order by code"
blm.fill_comb ssql, Combo2, "name", "code"
End Sub
Private Function max1() As Double
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
    
    ssql = "select max(v_no)as c from voumst where v_type = 5"
    
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

Private Sub Combo1_Click()
If Combo1.ListCount > 0 Then
    Text2.Text = Combo1.ItemData(Combo1.ListIndex)
    Text3.Text = Combo1.Text
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
    Text4.Text = Combo2.ItemData(Combo2.ListIndex)
    Text5.Text = Combo2.Text
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub clearfull()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
combs
End Sub

Private Sub Command1_Click()
Call save
Command2_Click
End Sub

Private Sub Command2_Click()
Call clearfull
Text2.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
combs
End Sub

Private Sub Option1_Click()
Call clearfull
Text1.Enabled = False
Text2.SetFocus

End Sub

Private Sub Option2_Click()
Call clearfull
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
    If KeyAscii = 13 And Val(Text1.Text) <> 0 Then
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
    b = edit1
        If b = False Then
            Cancel = True
        End If
End If

End Sub

Private Sub Text2_GotFocus()
If Option1 = True Then
    Text1.Text = max1
End If
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo1.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) <> 0 Then
    Text3.Text = blm.party1(Val(Text2.Text))
    If Text3.Text = "NOT" Then
        Cancel = True
    End If
        
End If

End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) <> 0 Then
    Text5.Text = blm.item1(Val(Text4.Text))
    If Text5.Text = "NOT" Then
        Cancel = True
    End If
        
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
    Else
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub Text9_Validate(Cancel As Boolean)
If Val(Text9.Text) <> 0 Then
    Text10.Text = blm.party1(Val(Text9.Text))
    If Text10.Text = "NOT" Then
        Cancel = True
    End If
        
End If

End Sub
