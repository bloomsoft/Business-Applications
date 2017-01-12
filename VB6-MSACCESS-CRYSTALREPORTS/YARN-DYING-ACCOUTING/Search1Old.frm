VERSION 5.00
Begin VB.Form Search1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Search Wizard"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Select Catagory"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Refresh List"
      Height          =   1095
      Left            =   1560
      Picture         =   "Search1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   3960
      Picture         =   "Search1.frx":04EF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   1095
      Left            =   2760
      Picture         =   "Search1.frx":0A50
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   6735
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Item List"
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6735
      Begin VB.ListBox List1 
         Height          =   3180
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item To Search"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6375
      End
   End
End
Attribute VB_Name = "Search1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Dim tb As Recordset
Dim db As Database
Private Sub ItemListCatgs(CatgCode As Long)
Dim tbf As Recordset
Dim s As String
List1.Clear
'MsgBox itname
If CatgCode > 0 Then
    s = "GroupC = " & CatgCode
    tb.Filter = s
    Set tbf = tb.OpenRecordset()
    If Not tbf.EOF Then
        Do While Not tbf.EOF
            List1.AddItem tbf.Fields("Name").Value & "  [" & tbf.Fields("Unit").Value & "]"
            List1.ItemData(List1.NewIndex) = tbf.Fields("Code").Value
            tbf.MoveNext
        Loop
    End If
    
Else
    tb.MoveFirst
    If Not tb.EOF Then
    Do While Not tb.EOF
            List1.AddItem tb.Fields("Name").Value & "  [" & tb.Fields("Unit").Value & "]"
            List1.ItemData(List1.NewIndex) = tb.Fields("Code").Value
            tb.MoveNext
    Loop
    End If
End If

End Sub
Private Sub ItemListOLDC(CatgCode As Long)
Dim tbf As Recordset
Dim s As String
List1.Clear
'MsgBox itname
If CatgCode > 0 Then
    s = "OLDCODE = " & CatgCode
    tb.Filter = s
    Set tbf = tb.OpenRecordset()
    If Not tbf.EOF Then
        Do While Not tbf.EOF
            List1.AddItem tbf.Fields("Name").Value & "  [" & tbf.Fields("Unit").Value & "]"
            List1.ItemData(List1.NewIndex) = tbf.Fields("Code").Value
            tbf.MoveNext
        Loop
    End If
    
Else
    tb.MoveFirst
    If Not tb.EOF Then
    Do While Not tb.EOF
            List1.AddItem tb.Fields("Name").Value & "  [" & tb.Fields("Unit").Value & "]"
            List1.ItemData(List1.NewIndex) = tb.Fields("Code").Value
            tb.MoveNext
    Loop
    End If
End If

End Sub

Private Sub Combs()
Dim ssql As String

ssql = "Select * from Groups Order by Name"
Blm1.fill_comb ssql, Combo1, "Name", "Code"
End Sub
Private Sub ItemList(itname As String)
Dim tbf As Recordset
List1.Clear
'MsgBox itname
If Len(itname) > 0 Then
    tb.Filter = "Name Like '" & itname & "*'"
    Set tbf = tb.OpenRecordset()
    If Not tbf.EOF Then
        Do While Not tbf.EOF
            List1.AddItem tbf.Fields("Name").Value & "  [" & tbf.Fields("Unit").Value & "]"
            List1.ItemData(List1.NewIndex) = tbf.Fields("Code").Value
            tbf.MoveNext
        Loop
    End If
    
Else
    If Len(itename) = 0 Then
 '   MsgBox "yes"
    tb.MoveFirst
    If Not tb.EOF Then
    Do While Not tb.EOF
            List1.AddItem tb.Fields("Name").Value '& "  [" & tb.Fields("Unit").Value & "]"
            List1.ItemData(List1.NewIndex) = tb.Fields("Code").Value
            tb.MoveNext
    Loop
    End If
    End If
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1.SetFocus
End Sub

Private Sub Combo1_LostFocus()
Dim ssql As Long
If Combo1.ListIndex > -1 Then
    ssql = Combo1.ItemData(Combo1.ListIndex)
    ItemListCatgs ssql
    ssql = 0
End If

End Sub

Private Sub Command1_Click()
If Text3.Text = 1 Then
YarnIssue.Text4.Text = Label1.Caption
YarnIssue.Text5.Text = Label2.Caption
Me.Hide
YarnIssue.Text4.SetFocus
End If

Text1.Text = vbNullString
'Text2.Text = vbNullString
End Sub

Private Sub Command2_Click()
Me.Hide

End Sub

Private Sub Command3_Click()
Dim ssql As String
ssql = ""
tb.Requery
ItemList ssql
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then List1.SetFocus
If KeyCode = vbKeyF1 Then Combo1.SetFocus
If KeyCode = vbKeyF2 Then Text2.SetFocus

End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
Dim ssql As String
Set db = OpenDatabase(Blm1.pathMain)
ssql = "select * from CLOTHS order by Name"
Set tb = db.OpenRecordset(ssql, dbOpenDynaset)
ssql = ""
ItemList ssql
Combs


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
tb.Close
db.Close
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
    Label1.Caption = List1.ItemData(List1.ListIndex)
    Label2.Caption = List1.Text
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Text1.SetFocus
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Text1_Change()

ItemList UCase(Text1.Text)
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim s As String
s = ""
If KeyAscii = 34 Or KeyAscii = 39 Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 8 And Len(Text1.Text) = 0 Then
    ItemList s
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then List1.SetFocus
If KeyCode = vbKeyF1 Then Combo1.SetFocus
If KeyCode = vbKeyF2 Then Text1.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
    If KeyAscii = 13 Then
        ItemListOLDC Val(Text2.Text)
    Else
        Beep
        KeyAscii = 0
    End If
End If

End Sub
