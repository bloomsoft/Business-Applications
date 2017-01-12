VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NeedleOut 
   Caption         =   "Needles Out To Machine"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7095
   Icon            =   "NeedleOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   1530
      TabIndex        =   43
      Top             =   2280
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.ListBox List4 
      Height          =   1815
      Left            =   1530
      TabIndex        =   22
      Top             =   3030
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4140
      TabIndex        =   38
      Top             =   3750
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   240
      TabIndex        =   23
      Top             =   1110
      Width           =   11415
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10350
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3870
         TabIndex        =   4
         Top             =   810
         Width           =   945
      End
      Begin VB.TextBox Text4 
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
         Left            =   5850
         TabIndex        =   5
         Top             =   810
         Width           =   3780
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1290
         TabIndex        =   3
         Top             =   810
         Width           =   1275
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   3870
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1230
         Width           =   7335
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   1290
         TabIndex        =   9
         Top             =   1650
         Width           =   1275
      End
      Begin VB.TextBox Text20 
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
         Left            =   3870
         TabIndex        =   10
         Top             =   1650
         Width           =   7320
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Accept Entry"
         Height          =   375
         Left            =   9960
         TabIndex        =   11
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1290
         TabIndex        =   7
         Top             =   1230
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   3870
         TabIndex        =   1
         Top             =   360
         Width           =   945
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cancel this Inwardi"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   3960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6240
         Top             =   2280
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   5850
         TabIndex        =   2
         Top             =   390
         Width           =   5355
      End
      Begin MSComCtl2.DTPicker date4 
         Height          =   375
         Left            =   5160
         TabIndex        =   25
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   44433411
         CurrentDate     =   36921
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   44433411
         CurrentDate     =   36749
      End
      Begin VB.Label Label9 
         Caption         =   "Part Name"
         Height          =   255
         Left            =   4980
         TabIndex        =   42
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "Part No"
         Height          =   255
         Left            =   9720
         TabIndex        =   41
         Top             =   870
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Part Type"
         Height          =   255
         Left            =   2730
         TabIndex        =   40
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Part Code"
         Height          =   255
         Left            =   90
         TabIndex        =   39
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "Employee"
         Height          =   255
         Left            =   2730
         TabIndex        =   37
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Machine Code"
         Height          =   255
         Left            =   90
         TabIndex        =   36
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Machine Name"
         Height          =   255
         Left            =   2700
         TabIndex        =   35
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label15 
         Height          =   255
         Left            =   9240
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   6840
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "No Of Sets"
         Height          =   255
         Left            =   3090
         TabIndex        =   32
         Top             =   2220
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   31
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   4980
         TabIndex        =   30
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Width"
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   90
         TabIndex        =   28
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Outward #"
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "A/C List"
      Height          =   1005
      Left            =   3285
      TabIndex        =   20
      Top             =   45
      Visible         =   0   'False
      Width           =   4830
      Begin VB.ComboBox Combo2 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   450
         TabIndex        =   21
         Text            =   "Combo2"
         Top             =   405
         Width           =   3900
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      CausesValidation=   0   'False
      Height          =   3105
      Left            =   240
      TabIndex        =   19
      Top             =   3690
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5477
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1095
      Left            =   8160
      TabIndex        =   18
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2280
         Picture         =   "NeedleOut.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1200
         Picture         =   "NeedleOut.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   120
         Picture         =   "NeedleOut.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1800
         Picture         =   "NeedleOut.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   240
         Picture         =   "NeedleOut.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "NeedleOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim org_q As Currency
Dim rej As Currency

Private Sub edit1Cont(R As Long, c As Long, e As Integer)
Dim Tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from cont_1 where cont_no = " & c
ssql = ssql & " and e_type = " & e
org_q = 0
Set Tb = CN.Execute(ssql)
If Not Tb.EOF Then
     
     Grid1.TextMatrix(R, 1) = Format(Tb.Fields("v_dATE").Value, "dd/MM/yyyy")
Else
    MsgBox "Not Found ...!"
    
End If
Tb.Close
End Sub

Private Sub Transfer1()
With Grid1
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = Val(Text3.Text)
    .TextMatrix(.Rows - 1, 1) = Text6.Text
    .TextMatrix(.Rows - 1, 2) = Text4.Text
    .TextMatrix(.Rows - 1, 3) = Text7.Text
    .TextMatrix(.Rows - 1, 4) = Val(Text5.Text)
    .TextMatrix(.Rows - 1, 5) = Text22.Text
    .TextMatrix(.Rows - 1, 6) = Val(Text21.Text)
    .TextMatrix(.Rows - 1, 7) = Text20.Text
    
End With
End Sub

Private Sub Flex1()
With Grid1
    .Rows = 1
    .Cols = 8
    .ColWidth(0) = 1000
    .TextMatrix(0, 0) = "Part Code"
    .ColWidth(1) = 1000
    .TextMatrix(0, 1) = "Part Type"
    .ColWidth(2) = 2000
    .TextMatrix(0, 2) = "Part Name"
    .ColWidth(3) = 1000
    .TextMatrix(0, 3) = "Part No"
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "Quantity"
    .ColWidth(5) = 2000
    .TextMatrix(0, 5) = "Employee"
    .ColWidth(6) = 1100
    .TextMatrix(0, 6) = "Machine.Code"
    .ColWidth(7) = 2200
    .TextMatrix(0, 7) = "Machine.Name"
    
End With
End Sub

Private Function Check(c As Long) As Boolean
Dim Tb As ADODB.Recordset
Dim ssql As String
    
ssql = "select * from needlesout where in_no = " & c
ssql = ssql & " and E_type=2"
Set Tb = CN.Execute(ssql)
If Not Tb.EOF Then
    MsgBox "needlesout No already Exist..."
    Check = True
Else
    Check = False
End If
Tb.Close
End Function


Private Function edit1() As Boolean
Dim Tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from cont_1 where cont_no = " & Val(Text2.Text)
ssql = ssql & " and e_type = 2"
org_q = 0
Set Tb = CN.Execute(ssql)

If Not Tb.EOF Then
    date3.Value = Tb.Fields("v_date").Value
    Text3.Text = Tb.Fields("party").Value
    Text4.Text = blm.party1(Tb.Fields("party").Value)
    Label21.Caption = Format(Tb.Fields("del_date").Value, "dd/MM/yyyy")
    Label23.Caption = Format(Tb.Fields("Rate").Value, "#.00")
    org_q = Tb.Fields("Cquantity").Value
    Label13.Caption = Format(Tb.Fields("CQuantity").Value, "#.00")
    Label15.Caption = Format(Tb.Fields("YQuantity").Value, "#.00")
    
    Text12.Text = Tb.Fields("yARNcOUNT").Value
    Text13.Text = blm.Yarn(Tb.Fields("yARNcOUNT").Value)
    If Not IsNull(Tb.Fields("complete").Value) Then
    If Tb.Fields("Complete").Value = 1 Then
        MsgBox "You Have Marked this Contract as Completed....."
        
    End If
    End If
    edit1 = True
Else
    MsgBox "Not Found ...!"
    edit1 = False
End If
Tb.Close
End Function
Private Function max1() As Double
    Dim ssql As String
    Dim Tb As ADODB.Recordset
    
    ssql = "select max(outno)as c from needlesout"
    Set Tb = CN.Execute(ssql)
    If IsNull(Tb.Fields("c").Value) = False Then
        max1 = Tb.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    Tb.Close
End Function
Private Function edit_kachi() As Boolean
Dim Tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from needlesout where outno = " & Val(Text1.Text)

Set Tb = CN.Execute(ssql)
If Not Tb.EOF Then
    date1.Value = Tb.Fields("outdate").Value
    Text17.Text = Tb.Fields("remarks").Value & ""
          
   Do While Not Tb.EOF
   With Grid1
   .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = Tb.Fields("partcode").Value
     blm.PartDet Val(Tb.Fields("partcode").Value), Text6, Text4, Text7
    .TextMatrix(.Rows - 1, 1) = Text6.Text
    .TextMatrix(.Rows - 1, 2) = Text4.Text
    .TextMatrix(.Rows - 1, 3) = Text7.Text
    Text6.Text = ""
    Text4.Text = ""
    Text7.Text = ""
    
    .TextMatrix(.Rows - 1, 4) = Tb.Fields("qty").Value & ""
    .TextMatrix(.Rows - 1, 5) = Tb.Fields("emp").Value
    .TextMatrix(.Rows - 1, 6) = Tb.Fields("machinecode").Value
    .TextMatrix(.Rows - 1, 7) = blm.machine(Tb.Fields("machinecode").Value)
    End With
    Tb.MoveNext
    Loop

    edit_kachi = True
Else
    MsgBox "Not Found ...!"
    edit_kachi = False
End If
Tb.Close
    
End Function

Private Sub save()
Dim Tb As New ADODB.Recordset
Dim i As Long
Dim ssql As String
If Option2 = True Then
    ssql = "delete from needlesout where outno = " & Val(Text1.Text)
    CN.Execute ssql
End If
Tb.Open "needlesout", CN, 0, 3, 0
For i = 1 To Grid1.Rows - 1

With Grid1
Tb.AddNew
    Tb.Fields("outno").Value = Val(Text1.Text)
    Tb.Fields("outdate").Value = date1.Value
    Tb.Fields("remarks").Value = Text17.Text
        
    Tb.Fields("partcode").Value = Val(.TextMatrix(i, 0))
    Tb.Fields("qty").Value = Val(.TextMatrix(i, 4))
    Tb.Fields("emp").Value = .TextMatrix(i, 5)
    Tb.Fields("machinecode").Value = Val(.TextMatrix(i, 6))
End With

Tb.Update
Next i
Tb.Close
End Sub

Private Sub clear()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
Label23.Caption = vbNullString
If Option1 = True Then
    Text1.Text = max1
End If

End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
    Text3.Text = Combo2.ItemData(Combo2.ListIndex)
    Text4.Text = Combo2.Text
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim R As VbMsgBoxResult

If Grid1.Rows = 1 Then
MsgBox "Please Complete The Entery", , BLOOMSOFT
Exit Sub
End If

Call save
R = MsgBox("Want to Print", vbYesNo)
If R = vbYes Then
   Load vour
    vour.Caption = "Needles and Sinkers OutWard"
    vour.Text2.Text = 65
    vour.Text1.Text = Text1.Text
    vour.Label1.Caption = "OutWard #"
    vour.Show
End If
If R = vbNo Then
Command2_Click
Option1 = True
End If
If R = vbYes Then
vour.Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Call clear
Flex1
date1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide

End Sub

Private Sub Command4_Click()
If Val(Text3.Text) > 0 Then
Transfer1

Text5.Text = ""
Text2.Text = ""
Text21.Text = ""
Text20.Text = ""
Text22.Text = ""
Text3.Text = ""
Text6.Text = ""
Text4.Text = ""
Text7.Text = ""

Else
    MsgBox "Please Complete the Entry"
End If
Text3.SetFocus
End Sub

Private Sub date1_LostFocus()
If Option1 = True Then
    Text1.Text = max1
End If
'Lostf date1
End Sub

Private Sub date3_GotFocus()
'GOTF date3
End Sub

Private Sub date3_LostFocus()
'Lostf date3
End Sub

Private Sub Form_Activate()
If Me.Visible = True Then
Command2_Click
Option1 = True
Me.WindowState = vbMaximized
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If

End Sub

Private Sub Form_Load()
Dim ssql As String
Me.Top = ((Screen.Height - Me.Height) / 2) - 1000
Me.Left = (Screen.Width - Me.Width) / 2
date1.Value = Date
Flex1


ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo2, "name", "code"
Text1.Text = max1

End Sub

Private Sub Grid1_DblClick()
If Grid1.Rows > 1 Then
    
    With Grid1
        Text3.Text = .TextMatrix(.Row, 0)
        Text6.Text = .TextMatrix(.Row, 1)
        Text4.Text = .TextMatrix(.Row, 2)
        Text7.Text = .TextMatrix(.Row, 3)
        Text5.Text = .TextMatrix(.Row, 4)
        Text22.Text = .TextMatrix(.Row, 5)
        Text21.Text = .TextMatrix(.Row, 6)
        Text20.Text = .TextMatrix(.Row, 7)
        
    End With
End If
If Grid1.Rows = 2 Then
    Grid1.Rows = 1
Else
'    Grid1.Rows = Grid1.Rows - 1
    Grid1.RemoveItem Grid1.Row

End If
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
Text3.Text = List1.ItemData(List1.ListIndex)
Text4.Text = List1.Text
End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
List1.Visible = False
End If

End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub List2_Click()
If List2.ListIndex > -1 Then
Text14.Text = List2.ItemData(List2.ListIndex)
Text15.Text = List2.Text
End If

End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text14.SetFocus
List2.Visible = False
End If

End Sub

Private Sub List2_LostFocus()
List2.Visible = False
End Sub

Private Sub List4_Click()
If List4.ListIndex > -1 Then
Text21.Text = List4.ItemData(List4.ListIndex)
Text20.Text = List4.Text
End If

End Sub

Private Sub List4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text21.SetFocus
List4.Visible = False
End If

End Sub

Private Sub List4_LostFocus()
List4.Visible = False
End Sub

Private Sub List3_Click()
If List3.ListIndex > -1 Then
Text2.Text = List3.ItemData(List3.ListIndex)
End If

End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Grid1.Rows > 1 Then
    If Text4.Text <> Mid(List3.Text, 1, Len(Text4.Text)) Then
    MsgBox "Please Select Same Contrect Party Name"
    Exit Sub
    End If
End If

Text2.SetFocus
List3.Visible = False
End If

End Sub

Private Sub List3_LostFocus()
List3.Visible = False
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Check1.Visible = False
date4.Visible = False
Command2_Click
date1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Check1.Visible = True
date4.Visible = True
Text1.SetFocus

End Sub

Private Sub Text1_GotFocus()
GOTF Text1
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
    
End If

End Sub

Private Sub Text1_LostFocus()
Lostf Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim b As Boolean
If Val(Text1.Text) > 0 Then
'    If Option1 = True Then
'        b = Check(Val(Text1.Text))
'        Cancel = b
'    End If
    If Option2 = True Then
    Grid1.Rows = 1
        
        b = edit_kachi
        If b = False Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub Text10_GotFocus()
'GOTF Text10
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
    
End If

End Sub

Private Sub Text10_LostFocus()
'Lostf Text10
End Sub

Private Sub Text11_GotFocus()
'GOTF Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
    
End If
End Sub

Private Sub Text11_LostFocus()
'Lostf Text11
End Sub

Private Sub Text12_GotFocus()
'GOTF Text12
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim Tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from yarn where y_type=1"
Set Tb = CN.Execute(S)
List1.Visible = True
If Not Tb.EOF Then
List1.clear
Do While Not Tb.EOF
List1.AddItem Tb.Fields("name").Value
List1.ItemData(List1.NewIndex) = Tb.Fields("code").Value
Tb.MoveNext
Loop
List1.ListIndex = 0
End If
Tb.Close
Set Tb = Nothing
List1.SetFocus
End If

End Sub


Private Sub Text12_LostFocus()
'Lostf Text12
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
If Val(Text12.Text) > 0 Then
    Text13.Text = blm.Yarn(Val(Text12.Text))
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub Text14_GotFocus()
'OTF Text14
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim Tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from yarn where y_type=2"
Set Tb = CN.Execute(S)
List2.Visible = True
If Not Tb.EOF Then
List2.clear
Do While Not Tb.EOF
List2.AddItem Tb.Fields("name").Value
List2.ItemData(List2.NewIndex) = Tb.Fields("code").Value
Tb.MoveNext
Loop
List2.ListIndex = 0
List2.SetFocus
End If
Tb.Close
Set Tb = Nothing
End If

End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub Text14_LostFocus()
'ostf Text14
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
If Val(Text14.Text) > 0 Then
    Text15.Text = blm.Lycra(Val(Text14.Text))
    Else
    Text15.Text = ""
End If

End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text16_GotFocus()
'OTF Text16
Text16.SelStart = 0
Text16.SelLength = Len(Text16.Text)
End Sub

Private Sub Text16_LostFocus()
'Lostf Text16
End Sub

Private Sub Text17_GotFocus()
GOTF Text1
'gotfocused Text17
End Sub


Private Sub Text17_LostFocus()
Lostf Text17
End Sub

Private Sub Text19_GotFocus()
'GOTF Text19
End Sub

Private Sub Text19_LostFocus()
'Lostf Text19
End Sub

Private Sub Text2_Change()
Dim b As Boolean
If Option2 = True Then
If Val(Text2.Text) > 0 Then
    b = edit1
End If
End If
End Sub

Private Sub Text2_GotFocus()
GOTF Text2
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim Tb As ADODB.Recordset
Dim i As Integer
Dim S As String

S = "select * from cont_1 where e_type=2 order by party"
Set Tb = CN.Execute(S)
List3.Visible = True
If Not Tb.EOF Then
List3.clear
Do While Not Tb.EOF
aa = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
bb = blm.party1(Tb.Fields("party").Value) & aa
CC = Mid(bb, 1, 28)
List3.AddItem CC & " " & Tb.Fields("cont_no").Value & "               " & blm.Yarn(Tb.Fields("yarncount").Value) & "              " & blm.Cloth(Tb.Fields("item").Value) & Tb.Fields("MGuage").Value

'List3.AddItem blm.party1(tb.Fields("party").Value) & "                " & tb.Fields("cont_no").Value & "               " & blm.Yarn(tb.Fields("yarncount").Value) & "              " & blm.Cloth(tb.Fields("item").Value) & "              " & tb.Fields("MGuage").Value
List3.ItemData(List3.NewIndex) = Tb.Fields("cont_no").Value
Tb.MoveNext
Loop
List3.ListIndex = 0
End If
Tb.Close
Set Tb = Nothing
List3.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text2_LostFocus()
Lostf Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim b As Boolean

If Val(Text2.Text) > 0 Then
    b = edit1
    If b = False Then
        Cancel = True
    End If
End If
End Sub

Private Sub Text21_GotFocus()
GOTF Text21
End Sub

Private Sub Text21_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim Tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from Machine where Status=0"
Set Tb = CN.Execute(S)
List4.Visible = True
If Not Tb.EOF Then
List4.clear
Do While Not Tb.EOF
List4.AddItem Tb.Fields("name").Value
List4.ItemData(List4.NewIndex) = Tb.Fields("code").Value
Tb.MoveNext
Loop
List4.ListIndex = 0
End If
Tb.Close
Set Tb = Nothing
List4.SetFocus
End If

End Sub

Private Sub Text21_Validate(Cancel As Boolean)
If Val(Text21.Text) > 0 Then
    Text20.Text = blm.machine(Val(Text21.Text))
        If Text20.Text = "NOT FOUND" Then
            Cancel = True
        End If
Else
        Cancel = True
End If
End Sub

Private Sub Text22_Change()
GOTF Text22
End Sub

Private Sub Text22_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load EmpSearch
    EmpSearch.Show vbModal
    Text22.Text = SelEmpName
End If

End Sub

Private Sub Text3_GotFocus()
GOTF Text3
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim Tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from parts"
Set Tb = CN.Execute(S)
List1.Visible = True
If Not Tb.EOF Then
List1.clear
Do While Not Tb.EOF
List1.AddItem Tb.Fields("partname").Value
List1.ItemData(List1.NewIndex) = Tb.Fields("partcode").Value
Tb.MoveNext
Loop
List1.ListIndex = 0
End If
Tb.Close
Set Tb = Nothing
List1.SetFocus
End If

End Sub

Private Sub Text3_LostFocus()
Lostf Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Val(Text3.Text) > 0 Then
blm.PartDet Val(Text3.Text), Text6, Text4, Text7
Else
        Cancel = True
End If

End Sub

Private Sub Text5_GotFocus()
GOTF Text5
End Sub

Private Sub Text5_LostFocus()
Lostf Text5
End Sub

Private Sub Text8_GotFocus()
'GOTF Text8
End Sub

Private Sub Text8_LostFocus()
'Lostf Text8
End Sub

Private Sub Text9_GotFocus()
'GOTF Text9
End Sub

Private Sub Text9_LostFocus()
'Lostf Text9
End Sub

Private Sub Timer1_Timer()
Dim f As Integer, S As Integer
'Label27.Caption = Format(Val(Text10.Text) * Val(Text11.Text) + Val(Text16.Text), "#.000")

'Text16.Text = Val(Text11.Text) - Val(Text13.Text) - Val(Text18.Text)
'Text15.Text = Val(Text13.Text) + Val(Text18.Text)
End Sub
