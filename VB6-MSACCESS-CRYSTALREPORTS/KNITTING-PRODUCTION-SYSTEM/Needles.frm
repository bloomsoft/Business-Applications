VERSION 5.00
Begin VB.Form Needles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Needles And Sinkers Information"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   Icon            =   "Needles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   1080
      TabIndex        =   15
      Top             =   4320
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   2760
         Picture         =   "Needles.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   1560
         Picture         =   "Needles.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   855
         Left            =   360
         Picture         =   "Needles.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   2895
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   6135
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Needles.frx":0C28
         Left            =   1800
         List            =   "Needles.frx":0C32
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         MaxLength       =   18
         TabIndex        =   6
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         MaxLength       =   69
         TabIndex        =   2
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Part Type"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "No. of Machines"
         Height          =   375
         Left            =   840
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Count in Set"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Part No."
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Part Name + Make"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Part Code"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1335
      Left            =   1080
      TabIndex        =   10
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   2400
         Picture         =   "Needles.frx":0C4A
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   480
         Picture         =   "Needles.frx":0F54
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "Needles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1


Private Sub save()
Dim tb As New ADODB.Recordset
Dim ssql As String
If Option2 = True Then
    ssql = "delete from Parts where Partcode = " & Val(Text1.Text)
    CN.Execute ssql
End If


tb.Open "Parts", CN, 0, 3, 0
tb.AddNew
    tb.Fields("PartCODE").Value = Val(Text1.Text)
    tb.Fields("PartNAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("PartType").Value = Val(Left(Combo2.Text, 1))
    tb.Fields("PartNo").Value = Text3.Text
    tb.Fields("SetCount").Value = Text4.Text
    tb.Fields("Machines").Value = Val(Text6.Text)
    
tb.Update
tb.Close


End Sub
Private Function Check(S As String) As Boolean

Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM Parts WHERE PartNAME = '" & S & "'"
Set tb = CN.Execute(ssql)

If Not tb.EOF Then
    Check = True
Else
    Check = False
End If
tb.Close

End Function
Private Function max1() As Long
Dim tb As ADODB.Recordset
Dim ssql As String
ssql = "select MAX(PARTCODE) AS C FROM Parts"
Set tb = CN.Execute(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
End Function

Private Sub Combo1_Change()
If Option2 = True Then
Text2.Text = Combo1.Text
End If
End Sub

Private Sub Combo1_Click()
If Option2 = True Then
Text1.Text = Combo1.ItemData(Combo1.ListIndex)
Text1.Enabled = False
Text2.Text = Combo1.Text
Call edit1

End If
End Sub

Private Sub Command1_Click()
Call save
combs
Command2_Click
If Option1 = True Then
Text2.SetFocus
Else
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = vbNullString
Text2.Text = vbNullString
If Option2 = True Then
    Text1.Enabled = True
    Text1.Text = vbNullString
    Combo1.Visible = True
    Text2.Visible = False
Else
    Text1.Enabled = False
    Text1.Text = max1
    Combo1.Visible = False
    Text2.Visible = True
End If
Text3.Text = vbNullString
Text4.Text = vbNullString
Text6.Text = vbNullString
Command1.Enabled = False
If Option1 = True Then
Text2.SetFocus
Else
Combo1.SetFocus
End If

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub combs()
Dim ssql As String
ssql = "Select * from Parts order by PartName"
blm.fill_comb ssql, Combo1, "PartName", "PartCode"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If
End Sub

Private Sub Form_Load()
Set blm = New bloom1

combs
Text1.Text = max1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub Option1_Click()
Command2_Click
If Option1 = True Then
'    Combo2.Enabled = True
    Text1.Enabled = False
    Text1.Text = max1
    Text2.Visible = True
    Text2.Text = vbNullString
    Text2.SetFocus
    Combo1.Visible = False
    
Else
    Text1.Enabled = True
End If
End Sub

Private Sub Option2_Click()
Command2_Click
If Option2 = True Then
combs

    Text1.Enabled = True
    Text1.Text = vbNullString
    
    
    Combo1.Visible = True
    Text2.Visible = False
    Combo1.SetFocus
    
Else
    Text1.Enabled = False
End If

End Sub

Private Sub edit1()
Dim tb As ADODB.Recordset
Dim ssql As String
Dim i As Long
For i = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(i) = Val(Text1.Text) Then
        Combo1.ListIndex = i
        Exit For
    End If
Next i

ssql = "SELECT * FROM Parts WHERE Partcode = " & Val(Text1.Text)
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    Text2.Text = tb.Fields("Partname").Value
    Text1.Enabled = False
    Combo2.ListIndex = tb.Fields("PartType").Value
    Text3.Text = tb.Fields("PartNo").Value & ""
    Text4.Text = tb.Fields("SetCount").Value & ""
    Text6.Text = tb.Fields("Machines").Value & ""
    
    
Else
    MsgBox "Invalid Part Code"
    Combo1.SetFocus
End If
tb.Close
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
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
If KeyAscii = 13 Then Text3.SetFocus

End Sub

Private Sub Text2_LostFocus()
If Option1 = True Then
Dim b As Boolean

b = Check(UCase(CStr(Text2.Text)))
If b = True Then
    MsgBox "PART ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub


