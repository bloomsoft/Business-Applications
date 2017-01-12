VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Item1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shipment Expence Definition"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Supporting Items"
      Height          =   2535
      Left            =   240
      TabIndex        =   17
      Top             =   7755
      Visible         =   0   'False
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   1335
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "(F1 to Clear, F2 to Accept,F3 to Search)"
         Height          =   375
         Left            =   4680
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Per"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Qty"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   2505
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         Height          =   975
         Left            =   3720
         Picture         =   "Item1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         Height          =   975
         Left            =   2520
         Picture         =   "Item1.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   975
         Left            =   1320
         Picture         =   "Item1.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1050
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   6255
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label12 
         Caption         =   "(F4) to Search Expence to Edit"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Exp. Description"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Exp. Code"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   6255
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         Height          =   915
         Left            =   3720
         Picture         =   "Item1.frx":0F56
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         Height          =   915
         Left            =   1320
         Picture         =   "Item1.frx":147C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   7920
      Width           =   735
   End
End
Attribute VB_Name = "Item1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Dim PDESC As String
Private Sub Clear1()
Text4.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = vbNullString

End Sub

Private Sub save()
Dim db As Database
Dim tb As Recordset
Dim R As Integer
Dim ssql As String
Set db = OpenDatabase(blm.pathMain)
If Option2 = True Then
    ssql = "delete from Item where code = " & Val(Text1.Text)
    db.Execute ssql
    
    
End If


Set tb = db.OpenRecordset("Item", dbOpenTable)
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    
    
tb.Update
tb.Close

db.Close

End Sub
Private Function check(s As String) As Boolean
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "SELECT * FROM Item WHERE NAME = '" & UCase(s) & "'"
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)

If Not tb.EOF Then
    check = True
Else
    check = False
End If
tb.Close
db.Close
End Function
Private Function max1() As Long
Dim db As Database
Dim tb As Recordset
Dim ssql As String
ssql = "select MAX(CODE) AS C FROM Item"
Set db = OpenDatabase(blm.pathMain)
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
db.Close
End Function
Private Sub Command1_Click()
Call save

Command2_Click

If Option1 = True Then
Text2.SetFocus
Else
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
PDESC = Text2.Text

Text2.Text = vbNullString


Clear1
If Option2 = True Then
    Text1.Enabled = True
    Text1.Text = vbNullString
Else
    Text1.Enabled = False
    Text1.Text = max1

End If
Command1.Enabled = False
If Option1 = True Then
Text2.SetFocus
Else
Text1.SetFocus
End If

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Combs()
End Sub

Private Sub Form_Activate()
Me.Top = 10
Me.Left = 10

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
Clear1
'Text4.SetFocus
End If
If KeyCode = vbKeyF2 Then
    
    Clear1
End If
If KeyCode = vbKeyF3 Then
                Screen.MousePointer = vbHourglass
                Search1.Text3.Text = 3
                Search1.Show
                Screen.MousePointer = vbDefault
End If
If Option2 = True Then
If KeyCode = vbKeyF4 Then
                Screen.MousePointer = vbHourglass
                Search1.Text3.Text = 2
                Search1.Show
                Screen.MousePointer = vbDefault
End If
End If
If KeyCode = vbKeyF12 Then


    Text2.Text = PDESC
End If
    
End Sub

Private Sub Form_Load()
Set blm = New bloom1
Text1.Text = max1
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub Grid1_DblClick()
If GRID1.Rows > 1 Then
With GRID1
    Text4.Text = .TextMatrix(.Row, 0)
    Text6.Text = .TextMatrix(.Row, 1)
    Text7.Text = .TextMatrix(.Row, 2)
    Text8.Text = .TextMatrix(.Row, 3)
    If GRID1.Rows = 2 Then
        GRID1.Rows = 1
    Else
        GRID1.RemoveItem .Row
    End If
End With
End If
End Sub

Private Sub Option1_Click()
Command2_Click
If Option1 = True Then
    Text1.Enabled = False
    Label12.Visible = False
    Text1.Text = max1
    Text2.Text = vbNullString
    Text2.SetFocus
    
Else
    Text1.Enabled = True
End If
End Sub

Private Sub Option2_Click()
Command2_Click
If Option2 = True Then

    'Combo2.Enabled = False
    Text1.Enabled = True
    Label12.Visible = True
    Text1.Text = vbNullString
     Text1.SetFocus
    
Else
    Text1.Enabled = False
End If

End Sub

Private Function edit1() As Boolean
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String

Set db_m = OpenDatabase(blm.pathMain)
ssql = "select * from item where code = " & Val(Text1.Text)
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    Text1.Text = tb.Fields("Code").Value & ""
    Text2.Text = tb.Fields("Name").Value & ""
   
    edit1 = False
Else
    MsgBox "Invalid Expence Code....."
    edit1 = True
    Exit Function
End If
tb.Close
db_m.Close

End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
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
Dim i As Long
If Option2 = True Then
If Val(Text1.Text) > 0 Then
        Cancel = edit1
End If
End If
End Sub

Private Sub Text2_Change()
If Text2.Text <> vbNullString Or Text2.Text <> "" Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")

End Sub

Private Sub Text2_LostFocus()
If Option1 = True Then
Dim b As Boolean
If Len(Text2.Text) > 0 Then
b = check(UCase(CStr(Text2.Text)))
If b = True Then
    MsgBox "ITEM ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) > 0 Then
    Text6.Text = blm.Item1(Val(Text4.Text))
    If Text6.Text = "NOT" Then
        MsgBox "Invalid Supporting Item Code...."
        Cancel = True
    End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys ("{TAB}")
        End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub
