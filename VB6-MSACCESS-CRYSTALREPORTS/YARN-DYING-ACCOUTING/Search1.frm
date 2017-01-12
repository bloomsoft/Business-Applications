VERSION 5.00
Begin VB.Form Search1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cloth Quality Search Wizard"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Refresh List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      Picture         =   "Search1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4110
      Picture         =   "Search1.frx":04EF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6465
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2895
      Picture         =   "Search1.frx":0A50
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6465
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   7095
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3960
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cloth Quality List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   7095
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cloth Quality To Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6615
      End
   End
End
Attribute VB_Name = "Search1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Dim tb As Recordset
Dim DB As Database
Private Sub ItemListCatgs(CatgCode As Long)
Dim tbf As Recordset
Dim s As String
List1.clear
'MsgBox itname
If CatgCode > 0 Then
    s = "CCode = " & CatgCode
    tb.Filter = s
    Set tbf = tb.OpenRecordset()
    If Not tbf.EOF Then
        Do While Not tbf.EOF
            List1.AddItem tbf.Fields("Name").Value
            List1.ItemData(List1.NewIndex) = tbf.Fields("Code").Value
            tbf.MoveNext
        Loop
    End If
    
Else
    tb.MoveFirst
    If Not tb.EOF Then
    Do While Not tb.EOF
            List1.AddItem tb.Fields("Name").Value
            List1.ItemData(List1.NewIndex) = tb.Fields("Code").Value
            tb.MoveNext
    Loop
    End If
End If

End Sub
Private Sub ItemListOLDC(CatgCode As Long)
Dim tbf As Recordset
Dim s As String
List1.clear
'MsgBox itname
If CatgCode > 0 Then
    s = "OLDCODE = " & CatgCode
    tb.Filter = s
    Set tbf = tb.OpenRecordset()
    If Not tbf.EOF Then
        Do While Not tbf.EOF
            List1.AddItem tbf.Fields("Name").Value
            List1.ItemData(List1.NewIndex) = tbf.Fields("Code").Value
            tbf.MoveNext
        Loop
    End If
    
Else
    tb.MoveFirst
    If Not tb.EOF Then
    Do While Not tb.EOF
            List1.AddItem tb.Fields("Name").Value
            List1.ItemData(List1.NewIndex) = tb.Fields("Code").Value
            tb.MoveNext
    Loop
    End If
End If

End Sub

Private Sub Combs()
End Sub
Private Sub ItemList(itname As String)
Dim tbf As Recordset
List1.clear
'MsgBox itname
If Len(itname) > 0 Then
    tb.Filter = "Name Like '" & itname & "*'"
    Set tbf = tb.OpenRecordset()
    If Not tbf.EOF Then
        Do While Not tbf.EOF
            List1.AddItem tbf.Fields("Name").Value
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
            List1.AddItem tb.Fields("Name").Value & ""
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
Dim Ssql As Long
If Combo1.ListIndex > -1 Then
    Ssql = Combo1.ItemData(Combo1.ListIndex)
    ItemListCatgs Ssql
    Ssql = 0
End If

End Sub

Private Sub Command1_Click()
If Text3.Text = 1 Then
YarnIssue.Text4.Text = Label1.Caption
YarnIssue.Text5.Text = Label2.Caption
Me.Hide
YarnIssue.Text4.SetFocus
End If

If Text3.Text = 2 Then
ClothRec.Text4.Text = Label1.Caption
ClothRec.Text5.Text = Label2.Caption
Me.Hide
ClothRec.Text4.SetFocus
End If

If Text3.Text = 3 Then
vouloompay.Text4.Text = Label1.Caption
vouloompay.Text5.Text = Label2.Caption
Me.Hide
vouloompay.Text4.SetFocus
End If

If Text3.Text = 4 Then
LotPacking.Text4.Text = Label1.Caption
LotPacking.Text5.Text = Label2.Caption
Me.Hide
LotPacking.Text4.SetFocus
End If

If Text3.Text = 5 Then
voudyingpay.Text4.Text = Label1.Caption
voudyingpay.Text5.Text = Label2.Caption
Me.Hide
voudyingpay.Text4.SetFocus
End If

If Text3.Text = 6 Then
vour4.Text1.Text = Label1.Caption
vour4.Text2.Text = Label2.Caption
Me.Hide
vour4.Text1.SetFocus
End If

If Text3.Text = 7 Then
vour2.Text1.Text = Label1.Caption
vour2.Text2.Text = Label2.Caption
Me.Hide
vour2.Text1.SetFocus
End If

If Text3.Text = 8 Then
vour2.Text1.Text = Label1.Caption
vour2.Text2.Text = Label2.Caption
Me.Hide
vour2.Text1.SetFocus
End If


If Text3.Text = 9 Then
vour5.Text1.Text = Label1.Caption
vour5.Text2.Text = Label2.Caption
Me.Hide
vour5.Text1.SetFocus
End If

Text1.Text = vbNullString
'Text2.Text = vbNullString
End Sub

Private Sub Command2_Click()
Me.Hide

End Sub

Private Sub Command3_Click()
Dim Ssql As String
Ssql = ""
tb.Requery
ItemList Ssql
End Sub

Private Sub Form_Activate()
'Text1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then List1.SetFocus
If KeyCode = vbKeyF1 Then Combo1.SetFocus
If KeyCode = vbKeyF2 Then Text2.SetFocus

End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
Dim Ssql As String
Set DB = OpenDatabase(Blm.pathMain)
Ssql = "select * from CLOTHS order by Name"
Set tb = DB.OpenRecordset(Ssql, dbOpenDynaset)
Ssql = ""
ItemList Ssql
Combs


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
tb.Close
DB.Close
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
