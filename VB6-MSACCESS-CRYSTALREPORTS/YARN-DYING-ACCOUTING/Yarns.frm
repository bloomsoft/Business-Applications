VERSION 5.00
Begin VB.Form Yarns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Yarn Informations"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "Yarns.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   7215
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1200
         Picture         =   "Yarns.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4920
         Picture         =   "Yarns.frx":0C00
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Yarn Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   7215
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   240
         Top             =   360
      End
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
         Height          =   1860
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
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
         Height          =   525
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Press Down Arrow For Edition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Yarn Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Yarn Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actions"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         Picture         =   "Yarns.frx":1645
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2880
         Picture         =   "Yarns.frx":1F48
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5160
         Picture         =   "Yarns.frx":2832
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label8 
      Caption         =   "BloomSoft"
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "Yarns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As Recordset
Dim DB As Database
Private blm1 As New bloom1
Private Function edit1() As Boolean
Dim RST As Recordset
Dim Ssql As String

Ssql = "Select * from Yarns where Code = " & Val(Text1.Text)
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    Text2.Text = RST.Fields("Name").Value & ""
    edit1 = False
Else
    MsgBox "Invalid Yarn Code"
    edit1 = True
End If
RST.Close

End Function
Private Function Max1() As Long
Dim RST As Recordset
Dim Ssql As String

Ssql = "Select Max(Code) as C from Yarns"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("C").Value) Then
    Max1 = RST.Fields("C").Value + 1
Else
    Max1 = 1
End If
RST.Close

End Function
Private Sub Save()
If Option2 = True Then
    Ssql = "Delete from Yarns Where Code = " & Val(Text1.Text)
    DB.Execute Ssql
    RS.Close
    Set RS = DB.OpenRecordset("Yarns", dbOpenDynaset)
End If
'MsgBox "Test"
RS.AddNew
    RS.Fields("Code").Value = Val(Text1.Text)
    RS.Fields("Name").Value = Text2.Text
RS.Update
RS.Close
Set RS = DB.OpenRecordset("Yarns", dbOpenDynaset)
End Sub

Private Sub Command1_Click()
Save
Command2_Click
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
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
Set DB = OpenDatabase(blm1.pathMain)
Set RS = DB.OpenRecordset("Yarns", dbOpenDynaset)
Text1.Text = Max1
End Sub

Private Sub List1_Click()
If List1.ListCount > 0 Then
    Text1.Text = List1.ItemData(List1.ListIndex)
    Text2.Text = List1.Text & ""
End If
End Sub

Private Sub List1_LostFocus()
List1.Visible = False
Text1.SetFocus
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Command2_Click
Label3.Visible = False
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Command2_Click
Label3.Visible = True
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
If Option2 = True Then
    Dim Ssql As String
    Dim RST As Recordset
    
    Ssql = "Select * from Yarns Order By Code"
    Set RST = DB.OpenRecordset(Ssql)
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
'Text1.Enabled = False
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    Cancel = edit1
End If

End Sub

Private Sub Timer1_Timer()
If Me.ActiveControl <> "List1" Then
    'List1.Visible = False
End If
End Sub
