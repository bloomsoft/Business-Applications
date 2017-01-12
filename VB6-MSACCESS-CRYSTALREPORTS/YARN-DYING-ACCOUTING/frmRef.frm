VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLots 
   Caption         =   "Dying Lots Settings"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8715
      TabIndex        =   8
      Top             =   90
      Width           =   1275
   End
   Begin VB.TextBox txtDyeName 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2115
      TabIndex        =   7
      Top             =   540
      Width           =   6480
   End
   Begin VB.TextBox txtDyeCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2115
      TabIndex        =   6
      Top             =   45
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2055
      Top             =   2955
   End
   Begin VB.TextBox txtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   3
      Top             =   1605
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid G1 
      Height          =   4380
      Left            =   120
      TabIndex        =   0
      Top             =   1215
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   7726
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   6195
      TabIndex        =   2
      Top             =   5760
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   420
      Left            =   4005
      TabIndex        =   1
      Top             =   5760
      Width           =   1320
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3405
      TabIndex        =   9
      Top             =   75
      Width           =   5190
   End
   Begin VB.Label Label2 
      Caption         =   "Dying Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   105
      TabIndex        =   5
      Top             =   525
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Dying Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Top             =   45
      Width           =   1995
   End
End
Attribute VB_Name = "frmLots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim Filling As Boolean
Private Function LOTCheck(S As String) As Boolean
Dim DB As Database
Dim tb As Recordset
Dim R As Long
Dim B As Boolean
Dim Ssql As String
lblError.Caption = ""

For R = 1 To G1.Rows - 2
    If R <> G1.Row Then
        If Val(G1.TextMatrix(R, 6)) = Val(S) Then
            lblError.Caption = "This Lot No " & S & " alreasy Exist"
            LOTCheck = True
            Exit Function
        End If
    End If
Next R

Set DB = OpenDatabase(blm.pathMain)
Ssql = "select * from ClothRec where LOT_NO =  " & Val(S) & " and DYING_CODE =  " & Val(txtDyeCode.Text) & ""

Set tb = DB.OpenRecordset(Ssql)

If Not tb.EOF Then
    lblError.Caption = "This Lot No " & S & " alreasy Exist"
    LOTCheck = True
End If
tb.Close
DB.Close
End Function

Private Sub Save()
Dim RST As Recordset
Dim DB As Database
Dim R As Integer

Set DB = OpenDatabase(blm.pathMain)
For R = 1 To G1.Rows - 2
    With G1
        If Val(.TextMatrix(R, 6)) > 0 Then
            Ssql = "Update Clothrec Set Lot_No=" & Val(.TextMatrix(R, 6))
            If Len(.TextMatrix(R, 7)) > 0 Then
                Ssql = Ssql & ",Program='" & .TextMatrix(R, 7) & "'"
            End If
            Ssql = Ssql & " where Rec_No=" & Val(.TextMatrix(R, 0))
            DB.Execute Ssql
            
        End If
    End With
Next R
DB.Close
MsgBox "All Lot Nos Saved!"
End Sub
Private Sub FillGrid()
Filling = True
Flex1
Dim Ssql As String
Dim RST As Recordset
Dim DB As Database
Dim B As Boolean
Set DB = OpenDatabase(blm.pathMain)
Ssql = "Select * from ClothRec where Dying_Code=" & txtDyeCode.Text & " Order By Rec_No"

Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    txtGrid.Text = ""
    Do While Not RST.EOF
        If RST.Fields("Lot_No").Value = 0 Then
            B = True
        ElseIf Len(RST.Fields("Program").Value) = 0 Then
            B = True
        Else
            B = False
        End If
        If B = True Then
        With G1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = RST.Fields("Rec_No").Value
            .TextMatrix(.Rows - 1, 1) = Format(RST.Fields("Date_RECIEVE").Value, "dd-MMM-yyyy")
            .TextMatrix(.Rows - 1, 2) = blm.Factory(RST.Fields("Fac_Code").Value)
            .TextMatrix(.Rows - 1, 3) = blm.FillCloth1(RST.Fields("Cloth_Code").Value)
            .TextMatrix(.Rows - 1, 4) = RST.Fields("Gazana").Value
            .TextMatrix(.Rows - 1, 5) = RST.Fields("Thans").Value
            If RST.Fields("Lot_No").Value <> 0 Then
                .TextMatrix(.Rows - 1, 6) = RST.Fields("Lot_No").Value
                .TextMatrix(.Rows - 1, 8) = "1"
            Else
                .TextMatrix(.Rows - 1, 8) = "0"
            End If
            If Not IsNull(RST.Fields("Program").Value) Then .TextMatrix(.Rows - 1, 7) = RST.Fields("Program").Value
        End With
        End If
    RST.MoveNext
    Loop
    RST.MoveFirst
    txtGrid.Text = RST.Fields("Rec_No").Value & ""
End If
RST.Close
DB.Close
G1.Rows = G1.Rows + 1
Filling = False
End Sub
Private Sub Flex1()
With G1
    .Rows = 1
    .Cols = 9
    .ColWidth(0) = 800
    .TextMatrix(0, 0) = "Receipt#"
    
    .ColWidth(1) = 1500
    .TextMatrix(0, 1) = "Date"
    
    .ColWidth(2) = 2500
    .TextMatrix(0, 2) = "Factory Name"
    
    .ColWidth(3) = 1800
    .TextMatrix(0, 3) = "Quality"
    
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "Gazana"
    
    .ColWidth(5) = 900
    .TextMatrix(0, 5) = "Thans"
    
    .ColWidth(6) = 1300
    .TextMatrix(0, 6) = "Lot No"
    
    .ColWidth(7) = 1300
    .TextMatrix(0, 7) = "Program"
    
    .ColWidth(8) = 1
    .TextMatrix(0, 8) = ""
    
End With
End Sub

Private Sub cmbParty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Command1_Click()
If G1.Rows > 1 Then
    Save
    Me.Hide
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Command3_Click()
If Val(txtDyeCode.Text) > 0 And Len(txtDyeName.Text) > 0 Then
    FillGrid
End If
End Sub

Private Sub Command4_Click()
Dim R As VbMsgBoxResult
Dim CR As Integer
Dim Ssql As String
If G1.Rows > 2 Then
    R = MsgBox("Do You Realy Want to Delete This Adjustment Entry", vbYesNo)
    If R = vbYes Then
        CR = Val(G1.TextMatrix(G1.Row, 11))
        If CR = 0 Then
            Ssql = "Delete from Refs where SerialNo=" & Val(G1.TextMatrix(G1.Row, 10))
            Cn.Execute Ssql
            txtGrid.Text = ""
            Command3_Click
        Else
            MsgBox "You Can only Delete Blind Adjustments Here"
        End If
    End If
End If

End Sub

Private Sub G1_EnterCell()
If G1.Col > 7 Then Exit Sub
If Filling = True Then Exit Sub
If G1.Row <= 0 Then Exit Sub
Dim R As Integer

txtGrid.Top = G1.CellTop + G1.Top
txtGrid.Left = G1.CellLeft + G1.Left
txtGrid.Height = G1.CellHeight
txtGrid.Width = G1.CellWidth
txtGrid.Text = G1.Text
txtGrid.Visible = True
If txtGrid.Visible = True Then txtGrid.SetFocus

End Sub

Private Sub G1_GotFocus()
DoEvents
If G1.Row = 0 And G1.Col = 0 Then
    If G1.Rows >= 2 Then
        G1.Row = 1
        G1.Col = 6
        DoEvents
    End If
End If

End Sub

Private Sub G1_LeaveCell()
If Filling = True Then Exit Sub
If G1.Row <= 0 Then Exit Sub
txtGrid.Visible = False
G1.Text = txtGrid.Text
End Sub

Private Sub G1_Scroll()
txtGrid.Visible = False
End Sub

Private Sub txtDyeCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Search2.Text3.Text = 6
        Search2.Show
    End If
End Sub

Private Sub txtDyeCode_KeyPress(KeyAscii As Integer)
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

Private Sub txtDyeCode_Validate(Cancel As Boolean)
Dim B As Boolean
If Val(txtDyeCode.Text) > 0 Then
    txtDyeName.Text = blm.Dying(Val(txtDyeCode.Text))
    If txtDyeName.Text = "NOT FOUND" Then
        MsgBox "Invalid Dying Code...."
        Cancel = True
    End If
Else
    MsgBox "Please Give Some Dying Code...."
    Cancel = True
End If

End Sub

Private Sub txtGrid_Change()
G1.Text = txtGrid.Text
End Sub

Private Sub txtGrid_GotFocus()
txtGrid.SelStart = 0
txtGrid.SelLength = Len(txtGrid.Text)
If G1.Col = 7 Then
    txtGrid.MaxLength = 20
Else
    txtGrid.MaxLength = 0
End If
End Sub

Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
    If G1.Col > 6 Then G1.Col = G1.Col - 1
End If
If KeyCode = vbKeyRight Then
    If G1.Col >= 6 And G1.Col < 7 Then G1.Col = G1.Col + 1
End If
End Sub

Private Sub txtGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NC As Integer
Dim BC As Boolean
'txtGrid_Validate (BC)
'If BC = True Then Exit Sub
If G1.Col < 6 Then G1.Col = 6
If KeyCode = vbKeyDown Then
    If G1.Row < (G1.Rows - 1) Then
        G1.Row = G1.Row + 1
    End If
End If
If KeyCode = vbKeyUp Then
    If G1.Row > 1 Then
        G1.Row = G1.Row - 1
    End If
End If
If KeyCode = vbKeyReturn Then
    DoEvents
        If G1.Row < (G1.Rows - 1) Then
                If G1.Col < (G1.Cols - 2) Then
                    G1.Col = G1.Col + 1
                Else
                    G1.Row = G1.Row + 1
                    G1.Col = 6
                End If
                DoEvents
        End If
    
End If

End Sub

Private Sub txtGrid_Validate(Cancel As Boolean)
If G1.Col = 6 Then
    If Len(txtGrid.Text) > 0 And Val(G1.TextMatrix(G1.Row, 8)) = 0 Then
        Cancel = LOTCheck(txtGrid.Text)
        If Cancel = True Then txtGrid.Text = ""
        Exit Sub
    End If
End If
If G1.Col = 7 Then
    If Len(txtGrid.Text) > 15 Then
        Cancel = True
        Exit Sub
    End If
End If
End Sub
