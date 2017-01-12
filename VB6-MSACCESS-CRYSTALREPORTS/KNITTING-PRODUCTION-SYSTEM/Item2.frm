VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Item2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cloths Information"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "Item2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4920
      TabIndex        =   25
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   480
      TabIndex        =   24
      Top             =   3480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&OK"
      Height          =   495
      Left            =   6120
      TabIndex        =   23
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3960
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   20
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   480
      TabIndex        =   18
      Top             =   3120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid G1 
      Height          =   1335
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   1320
      TabIndex        =   11
      Top             =   4920
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   2760
         Picture         =   "Item2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   1560
         Picture         =   "Item2.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   975
         Left            =   360
         Picture         =   "Item2.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   6255
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Width"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cloth Description"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cloth Code"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1215
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   2400
         Picture         =   "Item2.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   480
         Picture         =   "Item2.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Waste %"
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "%Age"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Yarn Name"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Yarn Code"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   6360
      Width           =   735
   End
End
Attribute VB_Name = "Item2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As bloom1
Private Sub Flex1()
With G1
    .Rows = 1
    .Cols = 4
    .ColWidth(0) = 1400
    .TextMatrix(0, 0) = "Yarn Code"
    .ColWidth(1) = 1900
    .TextMatrix(0, 1) = "Yarn Name"
    .ColWidth(2) = 500
    .TextMatrix(0, 2) = "%Age"
    .ColWidth(3) = 500
    .TextMatrix(0, 3) = "%Waste"
End With
End Sub
Private Sub save()
Dim tb As New ADODB.Recordset
Dim ssql As String
Dim R As Long
If Option2 = True Then
    ssql = "delete from Cloth where code = " & Val(Text1.Text)
    CN.Execute ssql
    ssql = "delete from ClothRatio where Clothcode = " & Val(Text1.Text)
    CN.Execute ssql
    
End If


tb.Open "Cloth", CN, 0, 3, 0
tb.AddNew
    tb.Fields("CODE").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("Width").Value = CStr(Text3.Text)
    
tb.Update
tb.Close

tb.Open "ClothRatio", CN, 0, 3, 0
For R = 1 To G1.Rows - 1
tb.AddNew
    tb.Fields("ClothCODE").Value = Val(Text1.Text)
    tb.Fields("YarnCode").Value = Val(G1.TextMatrix(R, 0))
    tb.Fields("Percentage").Value = Val(G1.TextMatrix(R, 2))
    tb.Fields("Waste").Value = Val(G1.TextMatrix(R, 3))
tb.Update
Next R
tb.Close

End Sub
Private Function Check(S As String) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM Cloth WHERE NAME = '" & S & "'"
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
ssql = "select MAX(CODE) AS C FROM Cloth"
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
If Not IsNull(tb.Fields("C")) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
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
If Combo1.ListCount > 0 Then
Text1.Text = Combo1.ItemData(Combo1.ListIndex)
Text1.Enabled = False
Text2.Text = Combo1.Text
Text3.Text = blm.ClothWidth(Combo1.ItemData(Combo1.ListIndex))
edit1
End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus 'SendKeys ("{TAB}")
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
Text3.Text = vbNullString
G1.Rows = 1
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

Private Sub combs()
Dim ssql As String


ssql = "select * from Cloth order by name"
blm.fill_comb ssql, Combo1, "Name", "Code"
End Sub

Private Sub Command4_Click()
G1.Rows = G1.Rows + 1
G1.TextMatrix(G1.Rows - 1, 0) = Text4.Text
G1.TextMatrix(G1.Rows - 1, 1) = Text5.Text
G1.TextMatrix(G1.Rows - 1, 2) = Text6.Text
G1.TextMatrix(G1.Rows - 1, 3) = Text7.Text
Text4.SetFocus
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
Flex1
Text1.Text = max1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set blm = Nothing
End Sub

Private Sub G1_DblClick()
If G1.Rows > 1 Then
    Text4.Text = G1.TextMatrix(G1.Row, 0)
    Text5.Text = G1.TextMatrix(G1.Row, 1)
    Text6.Text = G1.TextMatrix(G1.Row, 2)
    Text7.Text = G1.TextMatrix(G1.Row, 3)
    Text4.SetFocus
    If G1.Rows > 2 Then
        G1.RemoveItem G1.Row
    Else
        G1.Rows = 1
    End If
End If
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
    Text4.Text = List1.ItemData(List1.ListIndex)
    Text5.Text = List1.Text
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
Text4.SetFocus

End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub Option1_Click()
Command2_Click
If Option1 = True Then

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
    'Combo2.Enabled = False
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
For i = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(i) = Val(Text1.Text) Then
        Combo1.ListIndex = i
        Exit For
    End If
        
Next i
Dim RST As ADODB.Recordset
Dim S As String
S = "Select a.YarnCode,b.Name,a.Percentage,a.Waste from ClothRatio a,Yarn b Where a.YarnCode=b.Code and a.ClothCode=" & Val(Text1.Text)

Set RST = CN.Execute(S)
If Not RST.EOF Then
    G1.Rows = 1
    Do While Not RST.EOF
        G1.Rows = G1.Rows + 1
        G1.TextMatrix(G1.Rows - 1, 0) = RST.Fields("YarnCode")
        G1.TextMatrix(G1.Rows - 1, 1) = RST.Fields("Name")
        G1.TextMatrix(G1.Rows - 1, 2) = RST.Fields("Percentage")
        G1.TextMatrix(G1.Rows - 1, 3) = RST.Fields("Waste")
        RST.MoveNext
    Loop
End If
RST.Close

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
        edit1
'        MsgBox group(Val(Text1.Text))
        
        
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


Private Sub Text2_LostFocus()
If Option1 = True Then
Dim b As Boolean

b = Check(UCase(CStr(Text2.Text)))
If b = True Then
    MsgBox "ITEM ALREADY EXIST,,,,"
    Text2.SetFocus
End If
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim S As String

S = "Select * from Yarn where Y_type=1 Order By Code"
blm.fill_comb S, List1, "Name", "Code"

List1.Visible = True
List1.SetFocus
End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) > 0 Then
    Text5.Text = blm.Yarn(Val(Text4.Text))
    If Text5.Text = "NOT" Then
        MsgBox "Wrong Yarn Code"
    End If
End If
End Sub
