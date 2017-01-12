VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EmpAtd 
   Caption         =   "Employee Attendance Entry"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "EmpAtd.frx":0000
      Left            =   1140
      List            =   "EmpAtd.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1740
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid G1 
      Height          =   5055
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8916
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command4 
         Caption         =   "&Exit"
         Height          =   855
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   855
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   855
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "&Open"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74973187
         CurrentDate     =   37709
      End
      Begin VB.Label Label1 
         Caption         =   "Select Date"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "EmpAtd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm As New bloom1
Private Sub save()
Dim Rs As Recordset
Dim DB As Database
Dim Ssql As String
Dim R As Long
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "Delete from EmpATD where A_Date=#" & DTPicker1.Value & "#"
DB.Execute Ssql

Set Rs = DB.OpenRecordset("EmpATD", dbOpenTable)

For R = 1 To G1.Rows - 1
    Rs.AddNew
        Rs.Fields("A_Date").Value = DTPicker1.Value
        Rs.Fields("AC_Code").Value = Val(G1.TextMatrix(R, 1))
        Rs.Fields("Status").Value = G1.TextMatrix(R, 3)
        
    Rs.Update
Next R
Rs.Close
DB.Close

End Sub
Private Sub FillGrid()
Dim RST As Recordset
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)
Dim Ssql As String
G1.Rows = 1
Ssql = "Select * from Acchart where Status=0 and Mid(Code,1,2)='" & EmpHead & "' Order by Code"
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    Do While Not RST.EOF
        With G1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = RST.Fields("Code").Value
            .TextMatrix(.Rows - 1, 2) = RST.Fields("Name").Value
            .TextMatrix(.Rows - 1, 3) = "P"
        End With
        RST.MoveNext
    Loop
    
End If
RST.Close
DB.Close
End Sub
Private Sub flex1()
With G1
    .Rows = 1
    .Cols = 4
    .ColWidth(0) = 1200
    .TextMatrix(0, 0) = "Sr#"
    
    .ColWidth(1) = 2500
    .TextMatrix(0, 1) = "Emp. A/c Code"
    
    .ColWidth(2) = 3500
    .TextMatrix(0, 2) = "Emp. A/c Name"
    
    .ColWidth(3) = 1500
    .TextMatrix(0, 3) = "Status"
End With
End Sub

Private Sub Command1_Click()
Dim RST As Recordset
Dim Ssql As String
Dim DB As Database

Set DB = OpenDatabase(Blm.patHmain)
Ssql = "Select a.*,b.Name from EmpAtd a,Acchart b where a.Ac_Code=b.Code and a.A_Date = #" & DTPicker1.Value & "# Order by A.Ac_Code"
Set RST = DB.OpenRecordset(Ssql)
G1.Rows = 1
'MsgBox ssql
If Not RST.EOF Then
    Do While Not RST.EOF
            
        With G1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = RST.Fields("Ac_Code").Value
            .TextMatrix(.Rows - 1, 2) = RST.Fields("Name").Value & ""
            .TextMatrix(.Rows - 1, 3) = RST.Fields("Status").Value
        End With
        RST.MoveNext
    Loop
Else
    FillGrid
End If
RST.Close
DB.Close
End Sub

Private Sub Command2_Click()
If G1.Rows > 1 Then
    Screen.MousePointer = vbHourglass
    save
    Screen.MousePointer = vbDefault
    Command3_Click
    
End If
End Sub

Private Sub Command3_Click()
G1.Rows = 1
Combo1.Visible = False
Combo1.ListIndex = 0
End Sub

Private Sub Command4_Click()
Me.Hide
Unload Me
End Sub

Private Sub DTPicker1_Change()
Command1.Enabled = True

End Sub

Private Sub DTPicker1_LostFocus()
    If DTPicker1.Value >= FStartDate And DTPicker1.Value <= FEndDate Then
    '    Text1.Text = max1
        'edit1
    Else
        MsgBox "Please Select Proper Date in the Financial Year"
    End If
End Sub

Private Sub Form_Load()

DTPicker1.Value = Date
flex1
Combo1.AddItem "P"
Combo1.AddItem "A"
Combo1.AddItem "L"
Combo1.AddItem "*"
Combo1.AddItem "C"
Combo1.AddItem "Y"

End Sub

Private Sub G1_EnterCell()
Dim R As Long
If G1.Col = 3 Then
    Combo1.Top = G1.CellTop + G1.Top
    Combo1.Left = G1.CellLeft + G1.Left
    Combo1.Width = G1.CellWidth
    Combo1.Visible = True
    For R = 0 To Combo1.ListCount - 1
    If Combo1.List(R) = G1.Text Then
        Combo1.ListIndex = R
        Exit For
    End If
    Next R
    Combo1.SetFocus
End If
End Sub

Private Sub G1_LeaveCell()
If G1.Col = 3 Then
    G1.Text = Combo1.Text
    Combo1.Visible = False

End If
End Sub
