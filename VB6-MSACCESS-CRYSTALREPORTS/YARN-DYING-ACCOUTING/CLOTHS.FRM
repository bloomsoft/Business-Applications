VERSION 5.00
Begin VB.Form Cloths 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cloth Qualities Informations"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "CLOTHS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   2415
      TabIndex        =   25
      Top             =   2505
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   8055
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   1560
         Picture         =   "CLOTHS.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   5160
         Picture         =   "CLOTHS.frx":0C00
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Quality Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   8055
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4320
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1680
         Top             =   1800
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2265
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "CLOTHS.frx":1645
         Left            =   5880
         List            =   "CLOTHS.frx":165B
         TabIndex        =   7
         Top             =   1815
         Width           =   1815
      End
      Begin VB.Label Label9 
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
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Quality Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Reed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Pick"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Warp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Weft"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "PC/Cotton"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   8055
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   600
         Picture         =   "CLOTHS.frx":1685
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3240
         Picture         =   "CLOTHS.frx":1F88
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6000
         Picture         =   "CLOTHS.frx":2872
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   7080
      TabIndex        =   23
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "Cloths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As Recordset
Dim DB As Database
Private blm1 As New bloom1
Private Function CheckCode() As Boolean
Dim RST As Recordset
Dim Ssql As String
Dim R As Long
Ssql = "Select * from Cloths where Code = " & Val(Text1.Text)
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    MsgBox "Cloth Quality Code Already Exist"
    CheckCode = True
Else
    
    CheckCode = False
End If
RST.Close

End Function

Private Function edit1() As Boolean
Dim RST As Recordset
Dim Ssql As String
Dim R As Long
Ssql = "Select * from Cloths where Code = " & Val(Text1.Text)
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    Text2.Text = RST.Fields("Reed").Value & ""
    Text3.Text = RST.Fields("Pick").Value & ""
    Text4.Text = RST.Fields("Warp").Value & ""
    Text5.Text = RST.Fields("Weft").Value & ""
    Text6.Text = RST.Fields("Width").Value & ""
    Text7.Text = RST.Fields("NAME").Value & ""
    For R = 0 To Combo1.ListCount - 1
    If Combo1.List(R) = RST.Fields("CType").Value Then
        Combo1.ListIndex = R
        Exit For
    End If
    Next R
    edit1 = False
Else
    MsgBox "Invalid Cloth Quality Code"
    edit1 = True
End If
RST.Close

End Function
Private Function Max1() As Long
'Dim RST As Recordset
'Dim Ssql As String
'
'Ssql = "Select Max(Code) as C from Cloths"
'Set RST = DB.OpenRecordset(Ssql)
'If Not IsNull(RST.Fields("C").Value) Then
'    Max1 = RST.Fields("C").Value + 1
'Else
'    Max1 = 1
'End If
'RST.Close

End Function
Private Sub Save()
If Option2 = True Then
    Ssql = "Delete from Cloths Where Code = " & Val(Text1.Text)
    DB.Execute Ssql
    RS.Close
    Set RS = DB.OpenRecordset("Cloths", dbOpenDynaset)
End If
'MsgBox "Test"
RS.AddNew
    RS.Fields("Code").Value = Val(Text1.Text)
    RS.Fields("NAME").Value = Text7.Text
    RS.Fields("Reed").Value = Val(Text2.Text)
    RS.Fields("Pick").Value = Val(Text3.Text)
    RS.Fields("Warp").Value = Val(Text4.Text)
    RS.Fields("Weft").Value = Val(Text5.Text)
    RS.Fields("Width").Value = Val(Text6.Text)
    RS.Fields("CType").Value = Combo1.Text
RS.Update
RS.Close
Set RS = DB.OpenRecordset("Cloths", dbOpenDynaset)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys ("{TAB}")
End If
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
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

If Option1 = True Then
'    Text1.Text = Max1
    Text7.SetFocus
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
Set DB = OpenDatabase(blm1.pathMain)
Set RS = DB.OpenRecordset("Cloths", dbOpenDynaset)
'Text1.Text = Max1
Me.Top = 10
Me.Left = 10
End Sub

Private Sub List1_Click()
If List1.ListCount > 0 Then
    Text1.Text = List1.ItemData(List1.ListIndex)
    Text7.Text = List1.Text
    edit1
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7.SetFocus
End If
End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub Option1_Click()
'Text1.Enabled = False
Command2_Click
End Sub

Private Sub Option2_Click()
'Text1.Enabled = True
Command2_Click
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
If Option2 = True Then
    Dim Ssql As String
    Dim RST As Recordset
    
    Ssql = "Select * from Cloths Order By Code"
    Set RST = DB.OpenRecordset(Ssql)
    If RST.EOF Then
        List1.clear
    Else
        List1.clear
        Do While Not RST.EOF
            List1.AddItem RST.Fields("Reed") & " * " & RST.Fields("Pick") & " / " & RST.Fields("Warp") & " * " & RST.Fields("Weft") & "  " & RST.Fields("Width") & "'" & "  " & RST.Fields("CType") & "'" & RST.Fields("NAME")
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
'If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
'    Exit Sub
'Else
'    KeyAscii = 0
'End If
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Then

Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text1_LostFocus()
'Text1.Enabled = False
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Option1 = True Then
    If Val(Text1.Text) > 0 Then
        Cancel = CheckCode
    End If
End If

If Option2 = True Then
    If Val(Text1.Text) > 0 Then
        Cancel = edit1()
    End If
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
ElseIf KeyAscii = 13 Then
    SendKeys ("{TAB}")
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Timer1_Timer()
If Len(Text2.Text) <= 0 Then
    Command1.Enabled = False
ElseIf Len(Text3.Text) <= 0 Then
    Command1.Enabled = False
ElseIf Len(Text4.Text) <= 0 Then
    Command1.Enabled = False
ElseIf Len(Text5.Text) <= 0 Then
    Command1.Enabled = False
ElseIf Len(Text6.Text) <= 0 Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If

If Me.ActiveControl <> "List1" Then
    'List1.Visible = False
End If
End Sub
