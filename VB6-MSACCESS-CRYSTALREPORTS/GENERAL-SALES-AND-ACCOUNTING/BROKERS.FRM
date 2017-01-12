VERSION 5.00
Begin VB.Form Brokers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brokers Informations"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   1680
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      MaxLength       =   150
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2760
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Update"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&New"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   360
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Address"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Phone"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Broker Name"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Broker Code"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Brokers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As Recordset
Dim DB As Database
Private Blm1 As New bloom1
Private Function Edit1() As Boolean
Dim RST As Recordset
Dim ssql As String

ssql = "Select * from Brokers where Code = " & Val(Text1.Text)
Set RST = DB.OpenRecordset(ssql)
If Not RST.EOF Then
    Text2.Text = RST.Fields("Name").Value & ""
    Text3.Text = RST.Fields("Phone").Value & ""
    Text4.Text = RST.Fields("Address").Value & ""
    Edit1 = False
Else
    MsgBox "Invalid Brokers Code"
    Edit1 = True
End If
RST.Close

End Function
Private Function Max1() As Long
Dim RST As Recordset
Dim ssql As String

ssql = "Select Max(Code) as C from Brokers"
Set RST = DB.OpenRecordset(ssql)
If Not IsNull(RST.Fields("C").Value) Then
    Max1 = RST.Fields("C").Value + 1
Else
    Max1 = 1
End If
RST.Close

End Function
Private Sub Save()
If Option2 = True Then
    ssql = "Delete from Brokers Where Code = " & Val(Text1.Text)
    DB.Execute ssql
    Rs.Close
    Set Rs = DB.OpenRecordset("Brokers", dbOpenDynaset)
End If
'MsgBox "Test"
Rs.AddNew
    Rs.Fields("Code").Value = Val(Text1.Text)
    Rs.Fields("Name").Value = Text2.Text
    Rs.Fields("Phone").Value = Text3.Text
    Rs.Fields("Address").Value = Text4.Text
Rs.Update
Rs.Close
Set Rs = DB.OpenRecordset("Brokers", dbOpenDynaset)
End Sub

Private Sub Command1_Click()
Save
Command2_Click
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.Text = Max1
If Option1 = True Then
    Text2.SetFocus
Else
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

Dim i As Long
'If TypeOf Me.ActiveControl Is DTPicker Then MsgBox KeyAscii
 Do While i < 100000

        i = i + 1
Loop
If KeyAscii = 13 Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If
End Sub

Private Sub Form_Load()
Set DB = OpenDatabase(Blm1.pathMain)
Set Rs = DB.OpenRecordset("Brokers", dbOpenDynaset)
Text1.Text = Max1
End Sub

Private Sub List1_Click()
If List1.ListCount > 0 Then
    Text1.Text = List1.ItemData(List1.ListIndex)
    Text2.Text = List1.Text & ""
    Edit1
End If
End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Command2_Click
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Command2_Click
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
If Option2 = True Then
    Dim ssql As String
    Dim RST As Recordset
    
    ssql = "Select * from Brokers Order By Code"
    Set RST = DB.OpenRecordset(ssql)
    If RST.EOF Then
        List1.clear
    Else
        List1.clear
        Do While Not RST.EOF
            List1.AddItem RST.Fields("Name") & ""
            List1.ItemData(List1.NewIndex) = RST.Fields("Code").Value
            RST.MoveNext
        Loop
        List1.ListIndex = 0
    End If
    RST.Close
    List1.Visible = True
    List1.SetFocus
End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Enabled = False
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    Cancel = Edit1
End If
End Sub

Private Sub Timer1_Timer()
If Len(Text2.Text) > 0 Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
If Me.ActiveControl <> "List1" Then
    'List1.Visible = False
End If
End Sub
