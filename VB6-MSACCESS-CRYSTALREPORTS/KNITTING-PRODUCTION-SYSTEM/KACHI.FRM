VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form kachi 
   Caption         =   "Yarn Inward"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7095
   Icon            =   "KACHI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ListBox List3 
      Height          =   4935
      Left            =   1560
      TabIndex        =   66
      Top             =   2280
      Visible         =   0   'False
      Width           =   9345
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   1575
      TabIndex        =   61
      Top             =   3285
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   1560
      TabIndex        =   60
      Top             =   3000
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Frame Frame4 
      Caption         =   "A/C List"
      Height          =   1005
      Left            =   3285
      TabIndex        =   54
      Top             =   45
      Width           =   4830
      Begin VB.ComboBox Combo2 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   450
         TabIndex        =   55
         Text            =   "Combo2"
         Top             =   405
         Width           =   3900
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2775
      Left            =   240
      TabIndex        =   53
      Top             =   3960
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4895
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1095
      Left            =   8160
      TabIndex        =   44
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2280
         Picture         =   "KACHI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1200
         Picture         =   "KACHI.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   735
         Left            =   120
         Picture         =   "KACHI.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   41
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1800
         Picture         =   "KACHI.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   240
         Picture         =   "KACHI.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   240
      TabIndex        =   22
      Top             =   1080
      Width           =   11415
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   8160
         TabIndex        =   18
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   7560
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   7560
         TabIndex        =   11
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox Text15 
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
         Left            =   3960
         TabIndex        =   14
         Top             =   1920
         Width           =   1590
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text13 
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
         Left            =   3960
         TabIndex        =   9
         Top             =   1560
         Width           =   1590
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Accept Entry"
         Height          =   375
         Left            =   9240
         TabIndex        =   20
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   9240
         TabIndex        =   17
         Text            =   "24"
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   6840
         TabIndex        =   15
         Top             =   1920
         Width           =   615
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
         Left            =   3960
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker date4 
         Height          =   375
         Left            =   5160
         TabIndex        =   36
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   36921
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cancel this Inwardi"
         Height          =   255
         Left            =   2760
         TabIndex        =   35
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
         Left            =   1320
         TabIndex        =   19
         Top             =   2280
         Width           =   6615
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   375
         Left            =   9240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58916867
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         TabIndex        =   12
         Text            =   "44.2"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   9240
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   4230
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6960
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58916867
         CurrentDate     =   36749
      End
      Begin VB.Label Label30 
         Caption         =   "+Wt."
         Height          =   255
         Left            =   7560
         TabIndex        =   65
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label28 
         Caption         =   "Weight"
         Height          =   255
         Left            =   7560
         TabIndex        =   64
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "0.000"
         Height          =   255
         Left            =   10440
         TabIndex        =   63
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Weight"
         Height          =   255
         Left            =   9840
         TabIndex        =   62
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Lycra Name"
         Height          =   255
         Left            =   2760
         TabIndex        =   59
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Lycra Code"
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Yarn Name"
         Height          =   255
         Left            =   2760
         TabIndex        =   57
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Yarn Code"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Left            =   9240
         TabIndex        =   52
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label14 
         Caption         =   "Yarn Quantity"
         Height          =   255
         Left            =   8040
         TabIndex        =   51
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   6840
         TabIndex        =   50
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   5640
         TabIndex        =   49
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Wt./Bag"
         Height          =   255
         Left            =   8040
         TabIndex        =   48
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Lycra Bags"
         Height          =   255
         Left            =   5640
         TabIndex        =   47
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Del. Date"
         Height          =   255
         Left            =   8040
         TabIndex        =   40
         Top             =   840
         Width           =   735
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
         TabIndex        =   39
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Rate"
         Height          =   255
         Left            =   5040
         TabIndex        =   38
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   9240
         TabIndex        =   37
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Contract Date"
         Height          =   255
         Left            =   8040
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Wt./Bag"
         Height          =   255
         Left            =   8040
         TabIndex        =   31
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Yarn Bags"
         Height          =   255
         Left            =   5640
         TabIndex        =   30
         Top             =   1560
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
         Caption         =   "Quality"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   5640
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Inward #"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "kachi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim org_q As Currency
Dim rej As Currency

Private Sub edit1Cont(R As Long, c As Long, e As Integer)
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from cont_1 where cont_no = " & c
ssql = ssql & " and e_type = " & e
org_q = 0
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
     
     Grid1.TextMatrix(R, 1) = Format(tb.Fields("v_dATE").Value, "dd/MM/yyyy")
Else
    MsgBox "Not Found ...!"
    
End If
tb.Close
End Sub

Private Sub Transfer1()
With Grid1
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 0) = Text2.Text
    .TextMatrix(.Rows - 1, 1) = Format(date3.Value, "dd/MMM/yyyy")
    .TextMatrix(.Rows - 1, 2) = Text10.Text
    .TextMatrix(.Rows - 1, 3) = Label27.Caption
    .TextMatrix(.Rows - 1, 4) = Text5.Text
    .TextMatrix(.Rows - 1, 5) = Text19.Text
    .TextMatrix(.Rows - 1, 6) = Text17.Text
    .TextMatrix(.Rows - 1, 7) = Text12.Text
    .TextMatrix(.Rows - 1, 8) = Text13.Text
    .TextMatrix(.Rows - 1, 9) = Text14.Text
    .TextMatrix(.Rows - 1, 10) = Text15.Text
    .TextMatrix(.Rows - 1, 11) = Text11.Text
    .TextMatrix(.Rows - 1, 12) = Text6.Text
    .TextMatrix(.Rows - 1, 13) = Text16.Text
    
End With
End Sub

Private Sub Flex1()
With Grid1
    .Rows = 1
    .Cols = 14
    .ColWidth(0) = 1000
    .TextMatrix(0, 0) = "Cont. No."
    .ColWidth(1) = 1200
    .TextMatrix(0, 1) = "Cont.Date"
    .ColWidth(2) = 1200
    .TextMatrix(0, 2) = "Yarn Bags"
    .ColWidth(3) = 1200
    .TextMatrix(0, 3) = "Yarn Weight"
    .ColWidth(4) = 1200
    .TextMatrix(0, 4) = "Lycra Bags"
    .ColWidth(5) = 1200
    .TextMatrix(0, 5) = "Lycra Weight"
    .ColWidth(6) = 2500
    .TextMatrix(0, 6) = "Remarks"
    .ColWidth(7) = 500
    .TextMatrix(0, 7) = "Y.Code"
    .ColWidth(8) = 800
    .TextMatrix(0, 8) = "Y.Name"
    .ColWidth(9) = 500
    .TextMatrix(0, 9) = "L.Code"
    .ColWidth(10) = 500
    .TextMatrix(0, 10) = "L.Name"
    .ColWidth(11) = 1000
    .TextMatrix(0, 11) = "Yarn Wt/Pack"
    .ColWidth(12) = 1000
    .TextMatrix(0, 12) = "Lycra Wt/Pack"
    .ColWidth(13) = 1000
    .TextMatrix(0, 13) = "Add. Wt."
    
End With
End Sub

Private Function Check(c As Long) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String
    
ssql = "select * from inward where in_no = " & c
ssql = ssql & " and E_type=2"
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    MsgBox "Inward No already Exist..."
    Check = True
Else
    Check = False
End If
tb.Close
End Function


Private Function edit1() As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from cont_1 where cont_no = " & Val(Text2.Text)
ssql = ssql & " and e_type = 2"
org_q = 0
Set tb = CN.Execute(ssql)

If Not tb.EOF Then
    date3.Value = tb.Fields("v_date").Value
    Text3.Text = tb.Fields("party").Value
    Text4.Text = blm.party1(tb.Fields("party").Value)
    Label21.Caption = Format(tb.Fields("del_date").Value, "dd/MM/yyyy")
    Label23.Caption = Format(tb.Fields("Rate").Value, "#.00")
    org_q = tb.Fields("Cquantity").Value
    Label13.Caption = Format(tb.Fields("CQuantity").Value, "#.00")
    Label15.Caption = Format(tb.Fields("YQuantity").Value, "#.00")
    
    Text12.Text = tb.Fields("yARNcOUNT").Value
    Text13.Text = blm.Yarn(tb.Fields("yARNcOUNT").Value)
    If Not IsNull(tb.Fields("complete").Value) Then
    If tb.Fields("Complete").Value = 1 Then
        MsgBox "You Have Marked this Contract as Completed....."
        
    End If
    End If
    edit1 = True
Else
    MsgBox "Not Found ...!"
    edit1 = False
End If
tb.Close
End Function
Private Function max1() As Double
    Dim ssql As String
    Dim tb As ADODB.Recordset
    
    ssql = "select max(in_no)as c from inward where e_type=2"
    Set tb = CN.Execute(ssql)
    If IsNull(tb.Fields("c").Value) = False Then
        max1 = tb.Fields("c").Value + 1
    Else
        max1 = 1
    End If
    tb.Close
End Function
Private Function edit_kachi() As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from inward where in_no = " & Val(Text1.Text)
ssql = ssql & " and e_type=2"

Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    Text3.Text = tb.Fields("accode").Value & ""
    Text4.Text = tb.Fields("acname").Value & ""
    Text8.Text = tb.Fields("quality").Value & ""
    date1.Value = tb.Fields("in_date").Value
   Do While Not tb.EOF
   With Grid1
   .Rows = .Rows + 1
   If tb.Fields("cont_no").Value > 0 Then
    .TextMatrix(.Rows - 1, 0) = tb.Fields("cont_no").Value
    edit1Cont .Rows - 1, tb.Fields("Cont_no").Value, 2
   End If
    .TextMatrix(.Rows - 1, 2) = tb.Fields("Bags").Value
    .TextMatrix(.Rows - 1, 3) = tb.Fields("Quantity").Value
    .TextMatrix(.Rows - 1, 6) = tb.Fields("remarks").Value & ""
    .TextMatrix(.Rows - 1, 4) = tb.Fields("Lycra").Value
    .TextMatrix(.Rows - 1, 5) = tb.Fields("L_Bags").Value
    .TextMatrix(.Rows - 1, 7) = tb.Fields("Yarn").Value & ""
    .TextMatrix(.Rows - 1, 8) = blm.Yarn(tb.Fields("Yarn").Value & "")
    .TextMatrix(.Rows - 1, 9) = tb.Fields("LycraCode").Value & ""
    .TextMatrix(.Rows - 1, 10) = blm.Lycra(tb.Fields("LycraCode").Value & "")
    .TextMatrix(.Rows - 1, 11) = tb.Fields("WtPerPackYarn").Value & ""
    .TextMatrix(.Rows - 1, 12) = tb.Fields("WtPerpackLycra").Value & ""
    .TextMatrix(.Rows - 1, 13) = tb.Fields("AddWt").Value & ""
    If Not IsNull(tb.Fields("cancel").Value) Then
        Check1.Value = tb.Fields("cancel").Value
    End If
    If Not IsNull(tb.Fields("c_date").Value) Then
        date4.Value = tb.Fields("c_date").Value
    End If
    End With
    tb.MoveNext
    Loop

    edit_kachi = True
Else
    MsgBox "Not Found ...!"
    edit_kachi = False
End If
tb.Close
    
End Function

Private Sub save()
Dim tb As New ADODB.Recordset
Dim i As Long
Dim ssql As String
If Option2 = True Then
    ssql = "delete from Inward where in_no = " & Val(Text1.Text)
    ssql = ssql & " and e_type = 2"
    CN.Execute ssql
End If
tb.Open "inward", CN, 0, 3, 0
For i = 1 To Grid1.Rows - 1

With Grid1
tb.AddNew
    tb.Fields("in_no").Value = Val(Text1.Text)
    tb.Fields("in_date").Value = date1.Value
    tb.Fields("accode").Value = Text3.Text
    tb.Fields("acname").Value = Text4.Text
    tb.Fields("quality").Value = Text8.Text
    tb.Fields("E_Type").Value = 2
    
    tb.Fields("cont_no").Value = Val(.TextMatrix(i, 0))
    tb.Fields("bags").Value = Val(.TextMatrix(i, 2))
    tb.Fields("quantity").Value = Val(.TextMatrix(i, 3))
    tb.Fields("remarks").Value = CStr(.TextMatrix(i, 6))
    tb.Fields("Lycra").Value = Val(.TextMatrix(i, 4))
    tb.Fields("L_Bags").Value = Val(.TextMatrix(i, 5))
    tb.Fields("Yarn").Value = Val(.TextMatrix(i, 7))
    tb.Fields("LycraCode").Value = Val(.TextMatrix(i, 9))
    tb.Fields("WtPerPackYarn").Value = Val(.TextMatrix(i, 11))
    tb.Fields("WtPerpackLycra").Value = Val(.TextMatrix(i, 12))
    tb.Fields("AddWt").Value = Val(.TextMatrix(i, 13))
    If Option2 = True Then
        tb.Fields("cancel").Value = Check1.Value
        tb.Fields("c_date").Value = date4.Value
    End If
End With

tb.Update
Next i
tb.Close
End Sub

Private Sub clear()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
Label23.Caption = vbNullString
Label21.Caption = vbNullString
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
If Len(Text3.Text) <= 0 Then
MsgBox "Please Select Party", , BLOOMSOFT
Exit Sub
End If

If Grid1.Rows = 1 Then
MsgBox "Please Complete The Entery", , BLOOMSOFT
Exit Sub
End If

Call save
R = MsgBox("Want to Print", vbYesNo)
If R = vbYes Then
   Load vour
    vour.Caption = "Inward for Knitting Sale Contract"
    vour.Text2.Text = 1
    vour.Text1.Text = Text1.Text
    vour.Label1.Caption = "Inward #"
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
If Val(Text12.Text) > 0 Or Val(Text14.Text) > 0 Then
Transfer1
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text10.Text = ""
Text5.Text = ""
Text16.Text = ""
Text19.Text = ""
Text17.Text = ""
Label27.Caption = ""
Text2.Text = ""
Else
    MsgBox "Please Complete the Entry"
End If
Text2.SetFocus
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
date3.Value = Date


ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo2, "name", "code"
Text1.Text = max1

End Sub

Private Sub Grid1_DblClick()
If Grid1.Rows > 1 Then
    With Grid1
        Text2.Text = .TextMatrix(.Row, 0)
        If Val(Text2.Text) > 0 Then
        date3.Value = CDate(.TextMatrix(.Row, 1))
        End If
        Text10.Text = .TextMatrix(.Row, 2)
        Label27.Caption = .TextMatrix(.Row, 3)
        Text5.Text = .TextMatrix(.Row, 4)
        Text19.Text = .TextMatrix(.Row, 5)
        Text17.Text = .TextMatrix(.Row, 6)
        Text12.Text = .TextMatrix(.Row, 7)
        Text13.Text = .TextMatrix(.Row, 8)
        Text14.Text = .TextMatrix(.Row, 9)
        Text15.Text = .TextMatrix(.Row, 10)
        Text11.Text = .TextMatrix(.Row, 11)
        Text6.Text = .TextMatrix(.Row, 12)
        Text16.Text = .TextMatrix(.Row, 13)
        
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
Text12.Text = List1.ItemData(List1.ListIndex)
Text13.Text = List1.Text
End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text12.SetFocus
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
    If Option1 = True Then
        b = Check(Val(Text1.Text))
        Cancel = b
    End If
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
GOTF Text10
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
Lostf Text10
End Sub

Private Sub Text11_GotFocus()
GOTF Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
    
End If
End Sub

Private Sub Text11_LostFocus()
Lostf Text11
End Sub

Private Sub Text12_GotFocus()
GOTF Text12
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from yarn where y_type=1"
Set tb = CN.Execute(S)
List1.Visible = True
If Not tb.EOF Then
List1.clear
Do While Not tb.EOF
List1.AddItem tb.Fields("name").Value
List1.ItemData(List1.NewIndex) = tb.Fields("code").Value
tb.MoveNext
Loop
List1.ListIndex = 0
End If
tb.Close
Set tb = Nothing
List1.SetFocus
End If

End Sub


Private Sub Text12_LostFocus()
Lostf Text12
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
GOTF Text14
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from yarn where y_type=2"
Set tb = CN.Execute(S)
List2.Visible = True
If Not tb.EOF Then
List2.clear
Do While Not tb.EOF
List2.AddItem tb.Fields("name").Value
List2.ItemData(List2.NewIndex) = tb.Fields("code").Value
tb.MoveNext
Loop
List2.ListIndex = 0
List2.SetFocus
End If
tb.Close
Set tb = Nothing
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
Lostf Text14
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
GOTF Text16
Text16.SelStart = 0
Text16.SelLength = Len(Text16.Text)
End Sub

Private Sub Text16_LostFocus()
Lostf Text16
End Sub

Private Sub Text17_GotFocus()
GOTF Text1
'gotfocused Text17
End Sub


Private Sub Text17_LostFocus()
Lostf Text17
End Sub

Private Sub Text19_GotFocus()
GOTF Text19
End Sub

Private Sub Text19_LostFocus()
Lostf Text19
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
Dim tb As ADODB.Recordset
Dim i As Integer
Dim S As String

S = "select * from cont_1 where e_type=2 order by party"
Set tb = CN.Execute(S)
List3.Visible = True
If Not tb.EOF Then
List3.clear
Do While Not tb.EOF
aa = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
bb = blm.party1(tb.Fields("party").Value) & aa
CC = Mid(bb, 1, 28)
List3.AddItem CC & " " & tb.Fields("cont_no").Value & "               " & blm.Yarn(tb.Fields("yarncount").Value) & "              " & blm.Cloth(tb.Fields("item").Value) & tb.Fields("MGuage").Value

'List3.AddItem blm.party1(tb.Fields("party").Value) & "                " & tb.Fields("cont_no").Value & "               " & blm.Yarn(tb.Fields("yarncount").Value) & "              " & blm.Cloth(tb.Fields("item").Value) & "              " & tb.Fields("MGuage").Value
List3.ItemData(List3.NewIndex) = tb.Fields("cont_no").Value
tb.MoveNext
Loop
List3.ListIndex = 0
End If
tb.Close
Set tb = Nothing
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

Private Sub Text3_GotFocus()
GOTF Text3
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
End Sub

Private Sub Text3_LostFocus()
Lostf Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Val(Text3.Text) > 0 Then
    Text4.Text = blm.party1(Val(Text3.Text))
        If Text4.Text = "NOT FOUND" Then
            Cancel = True
            
        End If
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
GOTF Text8
End Sub

Private Sub Text8_LostFocus()
Lostf Text8
End Sub

Private Sub Text9_GotFocus()
GOTF Text9
End Sub

Private Sub Text9_LostFocus()
Lostf Text9
End Sub

Private Sub Timer1_Timer()
Dim f As Integer, S As Integer
Label27.Caption = Format(Val(Text10.Text) * Val(Text11.Text) + Val(Text16.Text), "#.000")

'Text16.Text = Val(Text11.Text) - Val(Text13.Text) - Val(Text18.Text)
'Text15.Text = Val(Text13.Text) + Val(Text18.Text)
End Sub
