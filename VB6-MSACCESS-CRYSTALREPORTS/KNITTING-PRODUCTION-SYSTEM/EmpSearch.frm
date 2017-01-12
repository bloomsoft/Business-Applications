VERSION 5.00
Begin VB.Form EmpSearch 
   Caption         =   "Employees Search"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "Employees"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "EmpSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FillList(Optional N As String)
Dim S As String
Dim Tb As New ADODB.Recordset

S = "Select * from SW2005.EmpProfile"
If Len(N) > 0 Then
    S = S & " where Upper(F_name) Like '%" & UCase(N) & "%' Order by F_Name"
End If

Set Tb = CN.Execute(S)
List1.clear
If Not Tb.EOF Then
    Do While Not Tb.EOF
        List1.AddItem Tb.Fields("F_Name") & " " & Tb.Fields("L_Name") & " S/O " & Tb.Fields("Father").Value
        Tb.MoveNext
    Loop
End If
Tb.Close

End Sub

Private Sub Command1_Click()
SelEmpName = List1.Text
Me.Hide
Unload Me
End Sub

Private Sub Command2_Click()
SelEmpName = ""
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
FillList
End Sub

Private Sub List1_DblClick()
Command1_Click
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
    FillList Text1.Text
Else
    FillList
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    List1.SetFocus
End If
End Sub
