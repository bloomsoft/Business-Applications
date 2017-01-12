VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form emp1 
   Caption         =   "Employee's Information"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11880
   Icon            =   "EMP1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD3 
      Left            =   9600
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select &Picture"
      Height          =   735
      Left            =   9360
      TabIndex        =   18
      Top             =   7080
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   5880
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1200
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   7920
      TabIndex        =   14
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   1095
         Left            =   2640
         Picture         =   "EMP1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   1095
         Left            =   1440
         Picture         =   "EMP1.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   1095
         Left            =   240
         Picture         =   "EMP1.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   915
         Left            =   3360
         Picture         =   "EMP1.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   915
         Left            =   960
         Picture         =   "EMP1.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Information"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   11655
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "&Husband Name"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&Father Name"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1560
         MaxLength       =   75
         TabIndex        =   3
         Top             =   1320
         Width           =   9975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   2
         Top             =   840
         Width           =   7815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Phone P.P."
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Phone"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Emp Name"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Emp No."
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Image Image3 
      Height          =   2805
      Left            =   9360
      Picture         =   "EMP1.frx":123C
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2490
   End
   Begin VB.Image Image2 
      Height          =   3615
      Left            =   4680
      Picture         =   "EMP1.frx":9259
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   120
      Picture         =   "EMP1.frx":34373
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   4575
   End
End
Attribute VB_Name = "emp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private Sub edit1()
'On Error Resume Next
Dim tb As ADODB.Recordset
Dim tbv As Recordset
Dim ssql As String

ssql = "SELECT * FROM Emp1 WHERE Emp_no = " & Val(Text1.Text)
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    Text2.Text = tb.Fields("name").Value
    Text1.Enabled = False
    Text3.Text = tb.Fields("F_Name").Value & ""
    Text4.Text = tb.Fields("Address").Value & ""
    Text5.Text = tb.Fields("Phone").Value & ""
    Text8.Text = tb.Fields("Phone_p").Value & ""
    If Not IsNull(tb.Fields("p_1")) Then
    Image1.Picture = LoadPicture(tb.Fields("P_1").Value)
    End If
    If Not IsNull(tb.Fields("p_2")) Then
    Image2.Picture = LoadPicture(tb.Fields("P_2").Value)
    End If
    If Not IsNull(tb.Fields("p_3")) Then
    Image3.Picture = LoadPicture(tb.Fields("P_3").Value)
    End If
    cd1.FileName = tb.Fields("P_1").Value & ""
    cd2.FileName = tb.Fields("P_2").Value & ""
    CD3.FileName = tb.Fields("P_3").Value & ""
    If tb.Fields("G_type").Value = 1 Then Option3 = True
    If tb.Fields("G_type").Value = 2 Then Option4 = True
Else
    MsgBox "Invalid Employee's Code"
    
End If
tb.Close
End Sub

Private Sub clear()
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Text4.Text = vbNullString
Text5.Text = vbNullString
Text8.Text = vbNullString
cd1.FileName = ""
cd2.FileName = ""
CD3.FileName = ""
Image1.Picture = LoadPicture()
Image2.Picture = LoadPicture()
Image3.Picture = LoadPicture()
If Option1 = True Then
    Text2.SetFocus
End If
If Option2 = True Then
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub
Private Sub save()
Dim tb As New ADODB.Recordset
Dim ssql As String
If Option2 = True Then
    ssql = "delete from Emp1 where Emp_no = " & Val(Text1.Text)
    CN.Execute ssql
End If


tb.Open "Emp1", CN, 0, 3, 0
tb.AddNew
    tb.Fields("Emp_no").Value = Val(Text1.Text)
    tb.Fields("NAME").Value = UCase(CStr(Text2.Text))
    tb.Fields("F_Name").Value = Text3.Text
    tb.Fields("Address").Value = Text4.Text
    tb.Fields("Phone").Value = Text5.Text
    tb.Fields("Phone_p").Value = Text8.Text
    tb.Fields("P_1").Value = cd1.FileName
    tb.Fields("P_2").Value = cd2.FileName
    tb.Fields("P_3").Value = CD3.FileName
    If Option3 = True Then tb.Fields("G_Type").Value = 1
    If Option4 = True Then tb.Fields("G_Type").Value = 2
tb.Update
tb.Close

End Sub
Private Function Check(S As String) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM Emp1 WHERE NAME = '" & S & "'"
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
ssql = "select MAX(Emp_no) AS C FROM Emp1"
Set tb = CN.Execute(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
End Function

Private Sub Command1_Click()
save
clear
End Sub

Private Sub Command2_Click()
clear
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Command4_Click()

    CD3.InitDir = "c:\soft\Bilal"
    CD3.Filter = "Bitmaps|*.bmp|JPEGS|*.jpg|Gifs|*.gif|ALL Image Types|*.bmp;*.jpg;*.gif"
    CD3.FilterIndex = 4
    CD3.ShowOpen
    If CD3.FileName <> "" Or CD3.FileName <> vbNullString Then
        Image3.Picture = LoadPicture(CD3.FileName)
    End If


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If

End Sub

Private Sub Form_Load()
Text1.Text = max1

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    cd1.InitDir = "c:\soft\Bilal"
    cd1.Filter = "Bitmaps|*.bmp|JPEGS|*.jpg|Gifs|*.gif|ALL Image Types|*.bmp;*.jpg;*.gif"
    cd1.FilterIndex = 4
    cd1.ShowOpen
    If cd1.FileName <> "" Or cd1.FileName <> vbNullString Then
        Image1.Picture = LoadPicture(cd1.FileName)
    End If
End If
        
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    cd2.InitDir = "c:\soft\Bilal"
    cd2.Filter = "Bitmaps|*.bmp|JPEGS|*.jpg|Gifs|*.gif|ALL Image Types|*.bmp;*.jpg;*.gif"
    cd2.FilterIndex = 4
    cd2.ShowOpen
    If cd2.FileName <> "" Or cd2.FileName <> vbNullString Then
        Image2.Picture = LoadPicture(cd2.FileName)
    End If
End If

End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Text1.Text = max1
Text2.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Option3_Click()
Text3.SetFocus
End Sub

Private Sub Option4_Click()
Text3.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
    
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    If Option2 = True Then
        edit1
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Cancel = Check(Text2.Text)
End Sub

Private Sub Text6_Change()
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
    
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
    
End If

End Sub

