VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form OutWarrdSKC 
   Caption         =   "Cloth Outward"
   ClientHeight    =   7815
   ClientLeft      =   2280
   ClientTop       =   630
   ClientWidth     =   7095
   Icon            =   "OutWardSKC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   7095
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   3840
      Top             =   8040
   End
   Begin VB.ListBox List3 
      Height          =   4155
      Left            =   1560
      TabIndex        =   64
      Top             =   1860
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   8340
      TabIndex        =   63
      Top             =   2160
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Frame Frame5 
      Caption         =   "A/C List"
      Height          =   1005
      Left            =   3285
      TabIndex        =   58
      Top             =   90
      Width           =   4830
      Begin VB.ComboBox Combo2 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   450
         TabIndex        =   59
         Text            =   "Combo2"
         Top             =   405
         Width           =   3900
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4080
      Top             =   480
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   240
      TabIndex        =   55
      Top             =   3240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sent Items Info"
      Height          =   1335
      Left            =   240
      TabIndex        =   38
      Top             =   6240
      Width           =   11415
      Begin VB.Label Label38 
         Height          =   195
         Left            =   10680
         TabIndex        =   57
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "Balance"
         Height          =   255
         Left            =   9900
         TabIndex        =   56
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         Height          =   255
         Left            =   9360
         TabIndex        =   54
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label35 
         Caption         =   "Fabric"
         Height          =   255
         Left            =   8220
         TabIndex        =   53
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         Height          =   255
         Left            =   6360
         TabIndex        =   52
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "Yarn Count"
         Height          =   255
         Left            =   5400
         TabIndex        =   51
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Label32"
         Height          =   255
         Left            =   3720
         TabIndex        =   50
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Cloth Rolls Sent"
         Height          =   375
         Left            =   2280
         TabIndex        =   49
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Label30"
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Cloth Sent Qty."
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "Label28"
         Height          =   255
         Left            =   9240
         TabIndex        =   46
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label27 
         Caption         =   "Lycra Bags Rec."
         Height          =   255
         Left            =   7920
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Label26"
         Height          =   255
         Left            =   6900
         TabIndex        =   44
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Lycra Rec. Qty."
         Height          =   255
         Left            =   5520
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Yarn Bags Rec."
         Height          =   255
         Left            =   2280
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Yarn Rec. Qty."
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1095
      Left            =   8160
      TabIndex        =   34
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2280
         Picture         =   "OutWardSKC.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1200
         Picture         =   "OutWardSKC.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   735
         Left            =   120
         Picture         =   "OutWardSKC.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   31
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1800
         Picture         =   "OutWardSKC.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   240
         Picture         =   "OutWardSKC.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   240
      TabIndex        =   17
      Top             =   1080
      Width           =   11415
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1320
         MaxLength       =   250
         TabIndex        =   13
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   9660
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1560
         Width           =   1545
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   8085
         TabIndex        =   11
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9675
         TabIndex        =   12
         Top             =   1170
         Width           =   1545
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   8100
         TabIndex        =   6
         Top             =   840
         Width           =   540
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5880
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&OK"
         Height          =   375
         Left            =   10320
         TabIndex        =   16
         Top             =   360
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5640
         Top             =   4680
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   5880
         TabIndex        =   14
         Top             =   1560
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   330
         Left            =   8160
         TabIndex        =   3
         Top             =   405
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   36749
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9240
         TabIndex        =   7
         Top             =   840
         Width           =   1995
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   840
         Width           =   3015
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
         Height          =   285
         Left            =   6060
         TabIndex        =   2
         Top             =   480
         Width           =   795
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   450
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20643843
         CurrentDate     =   36749
      End
      Begin VB.Label Label24 
         Caption         =   "Size/Guage"
         Height          =   255
         Left            =   360
         TabIndex        =   68
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Van No"
         Height          =   240
         Left            =   8700
         TabIndex        =   65
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label10 
         Caption         =   "Lycra%"
         Height          =   285
         Left            =   6960
         TabIndex        =   62
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Lycra Used"
         Height          =   240
         Left            =   8715
         TabIndex        =   61
         Top             =   1230
         Width           =   825
      End
      Begin VB.Label Label8 
         Caption         =   "Fabric Code"
         Height          =   255
         Left            =   6960
         TabIndex        =   60
         Top             =   900
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
         TabIndex        =   30
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Rate"
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label21 
         Height          =   255
         Left            =   6780
         TabIndex        =   28
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Contract Date"
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Cloth Quantity"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Cloth Rolls"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         Height          =   255
         Left            =   8700
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   5100
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Out ward #"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label15 
      Caption         =   "0.00"
      Height          =   255
      Left            =   6000
      TabIndex        =   67
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Total Wt."
      Height          =   255
      Left            =   5160
      TabIndex        =   66
      Top             =   5880
      Width           =   975
   End
End
Attribute VB_Name = "OutWarrdSKC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim org_q As Currency
Dim rej As Currency
Private Type ContInfo
    clothsent As Long
    yarnsent As Long
    clothrec As Long
    yarnrec As Long
    lycrarec As Long
    lbagsrec As Long
    lbagssent As Long
    lycrasent As Long
    sclothrolls As Long
    syarnbags As Long
    rclothrolls As Long
    ryarnbags As Long
    YarnCount As String
    Cloth As String
    
End Type
Private Sub clear()
Text5.Text = vbNullString
Text6.Text = vbNullString
Text7.Text = vbNullString
Text8.Text = vbNullString
Text9.Text = vbNullString
Text10.Text = vbNullString
Text11.Text = vbNullString
Text17.Text = vbNullString
Text13.Text = vbNullString

End Sub
Private Sub Transfer1()
    With Grid1
        .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Text2.Text
            .TextMatrix(.Rows - 1, 1) = Format(date3.Value, "dd/MM/yyyy")
            .TextMatrix(.Rows - 1, 2) = Text8.Text
            .TextMatrix(.Rows - 1, 3) = Format(Label21.Caption, "dd/MM/yyyy")
            .TextMatrix(.Rows - 1, 4) = Text10.Text
            .TextMatrix(.Rows - 1, 5) = Text11.Text
            .TextMatrix(.Rows - 1, 6) = Text5.Text
            .TextMatrix(.Rows - 1, 7) = Text17.Text
            .TextMatrix(.Rows - 1, 8) = Text6.Text
            .TextMatrix(.Rows - 1, 9) = Val(Text7.Text)
            .TextMatrix(.Rows - 1, 10) = Val(Text9.Text)
            .TextMatrix(.Rows - 1, 11) = Text13.Text
            
    End With
End Sub
    
Private Sub Flex1()
With Grid1
    .Rows = 1
    .Cols = 12
    .ColWidth(0) = 1000
    .TextMatrix(0, 0) = "Cont_no"
    .ColWidth(1) = 1000
    .TextMatrix(0, 1) = "Cont Date"
    .ColWidth(2) = 1500
    .TextMatrix(0, 2) = "Fabric"
    .ColWidth(3) = 1000
    .TextMatrix(0, 3) = "Del Date"
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "Rolls"
    .ColWidth(5) = 1000
    .TextMatrix(0, 5) = "Quantity"
    .ColWidth(6) = 1000
    .TextMatrix(0, 6) = "Rate"
    .ColWidth(7) = 3000
    .TextMatrix(0, 7) = "Remarks"
    .ColWidth(8) = 3
    .TextMatrix(0, 8) = "Fabric Code"
    .ColWidth(9) = 600
    .TextMatrix(0, 9) = "Lycra%"
    .ColWidth(10) = 800
    .TextMatrix(0, 10) = "Lycra Used"
    .ColWidth(11) = 1500
    .TextMatrix(0, 11) = "Size / Guage "
    
    
End With
End Sub
Private Function Getinfo(c As Long, e As Byte) As ContInfo
    Dim tb As ADODB.Recordset
    Dim ssql As String
    Dim CN2 As ContInfo
    
    
    ssql = "select sum(quantity)as q,sum(bags)as b,sum(lycra)as l,sum(l_bags)as lb from inward where cont_no = " & c
    ssql = ssql & " and e_type= " & e
    
    
    Set tb = CN.Execute(ssql)
    If Not IsNull(tb.Fields("Q").Value) Then
        CN2.yarnrec = tb.Fields("Q").Value
    End If
    If Not IsNull(tb.Fields("B").Value) Then
        CN2.ryarnbags = tb.Fields("B").Value
    End If
    If Not IsNull(tb.Fields("l").Value) Then
        CN2.lycrarec = tb.Fields("l").Value
    End If
    If Not IsNull(tb.Fields("lb").Value) Then
        CN2.lbagsrec = tb.Fields("lb").Value
    End If
    
    
    tb.Close
    
    ssql = "select sum(quantity)as q,sum(rolls)as r from Outward where cont_no = " & c
    ssql = ssql & " and e_type= " & e
    
    
    Set tb = CN.Execute(ssql)
    If Not IsNull(tb.Fields("Q").Value) Then
        CN2.clothsent = tb.Fields("Q").Value
    End If
    If Not IsNull(tb.Fields("r").Value) Then
        CN2.sclothrolls = tb.Fields("r").Value
    End If
    tb.Close
    Getinfo = CN2

End Function


Private Function Check(c As Long) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String
    
ssql = "select * from outward where out_no = " & c
ssql = ssql & " and E_type=2"
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    MsgBox "Outward No already Exist..."
    Check = True
Else
    Check = False
End If
tb.Close
End Function


Private Function edit1() As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String
Dim u As ContInfo

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
    Text5.Text = Format(tb.Fields("Rate").Value, "#.00")
    Text7.Text = Format(tb.Fields("Lycra").Value, "#.00")
    Text13.Text = Format(tb.Fields("MGUAGE").Value)
    org_q = tb.Fields("Cquantity").Value
'    Label13.Caption = Format(tb.Fields("CQuantity").Value, "#.00")
'    Label15.Caption = Format(tb.Fields("YQuantity").Value, "#.00")
    Label34.Caption = blm.Yarn(tb.Fields("YarnCount").Value)
    Label36.Caption = blm.Cloth(tb.Fields("Item").Value)
    Text6.Text = tb.Fields("item").Value
    Text8.Text = blm.Cloth(tb.Fields("item").Value)
    If Not IsNull(tb.Fields("complete").Value) Then
    If tb.Fields("Complete").Value = 1 Then
        MsgBox "You Have Marked this Contract as Completed....."
        
    End If
    End If
    edit1 = True
    u = Getinfo(Val(Text2.Text), 2)
    Label16.Caption = Format(u.yarnrec, "#.00")
    Label20.Caption = Format(u.ryarnbags, "#.00")
    Label26.Caption = Format(u.lycrarec, "#.00")
    Label28.Caption = Format(u.lbagsrec, "#.00")
    Label30.Caption = Format(u.clothsent, "#.00")
    Label32.Caption = Format(u.sclothrolls, "#.00")
    
Else
    MsgBox "Not Found ...!"
    edit1 = False
End If
tb.Close
End Function
Private Sub edit1Cont(R As Long, c As Long)
Dim tb As ADODB.Recordset
Dim ssql As String
Dim u As ContInfo

ssql = "select * from cont_1 where cont_no = " & c
ssql = ssql & " and e_type = 2"
org_q = 0
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
     Grid1.TextMatrix(R, 0) = tb.Fields("cont_no").Value
     Grid1.TextMatrix(R, 1) = Format(tb.Fields("v_dATE").Value, "dd/MM/yyyy")
     Grid1.TextMatrix(R, 2) = blm.Cloth(tb.Fields("item").Value)
     Grid1.TextMatrix(R, 3) = Format(tb.Fields("del_dATE").Value, "dd/MM/yyyy")
     Grid1.TextMatrix(R, 6) = Format(tb.Fields("Rate").Value, "#.00")
    ' Grid1.TextMatrix(r, 2) = blm.Cloth(tb.Fields("Item").Value)
    'Text9.Text = tb.Fields("item").Value
    
    If Not IsNull(tb.Fields("complete").Value) Then
    If tb.Fields("Complete").Value = 1 Then
        MsgBox "You Have Marked this Contract as Completed....."
        
    End If
    End If
    
    u = Getinfo(Val(Text2.Text), 2)
    Label16.Caption = Format(u.yarnrec, "#.00")
    Label20.Caption = Format(u.ryarnbags, "#.00")
    Label26.Caption = Format(u.lycrarec, "#.00")
    Label28.Caption = Format(u.lbagsrec, "#.00")
    Label30.Caption = Format(u.clothsent, "#.00")
    Label32.Caption = Format(u.sclothrolls, "#.00")
    
Else
    MsgBox "Not Found ...!"
    
End If
tb.Close
End Sub

Private Function max1() As Double
    Dim ssql As String
    Dim tb As ADODB.Recordset
    
    ssql = "select max(out_no)as c from outward where e_type=2"
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
Dim i As Integer

ssql = "select * from outward where out_no = " & Val(Text1.Text)
ssql = ssql & " and e_type=2"

Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("v_date").Value
    Text3.Text = tb.Fields("accode").Value & ""
    Text4.Text = tb.Fields("acname").Value & ""
    Text12.Text = tb.Fields("vanno").Value & ""
    Text17.Text = tb.Fields("remarks").Value & ""
    If Not IsNull(tb.Fields("cancel").Value) Then
        Check1.Value = tb.Fields("cancel").Value
    End If
    If Not IsNull(tb.Fields("c_date").Value) Then
        date4.Value = tb.Fields("c_date").Value
    End If
    
Do While Not tb.EOF
    Grid1.Rows = Grid1.Rows + 1
        If tb.Fields("cont_no") > 0 Then
        edit1Cont Grid1.Rows - 1, tb.Fields("Cont_no").Value
        'Grid1.TextMatrix(Grid1.Rows - 1, 0) = tb.Fields("cont_no").Value
        End If
        Grid1.TextMatrix(Grid1.Rows - 1, 4) = tb.Fields("rolls").Value
        Grid1.TextMatrix(Grid1.Rows - 1, 5) = tb.Fields("Quantity").Value
        Grid1.TextMatrix(Grid1.Rows - 1, 6) = tb.Fields("Rate").Value & ""
        Grid1.TextMatrix(Grid1.Rows - 1, 7) = tb.Fields("Remarks").Value & ""
        Grid1.TextMatrix(Grid1.Rows - 1, 8) = tb.Fields("Item").Value
        Grid1.TextMatrix(Grid1.Rows - 1, 2) = tb.Fields("quality").Value & ""
                If Not IsNull(tb.Fields("lyc_per")) Then
        Grid1.TextMatrix(Grid1.Rows - 1, 9) = tb.Fields("lyc_per").Value
        Grid1.TextMatrix(Grid1.Rows - 1, 10) = tb.Fields("lyc_used").Value & ""
        Grid1.TextMatrix(Grid1.Rows - 1, 11) = tb.Fields("size1").Value & ""
      
      End If
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
Dim ssql As String
Dim i As Integer
If Option2 = True Then
    ssql = "delete from Outward where out_no = " & Val(Text1.Text)
    ssql = ssql & " and e_type = 2"
    CN.Execute ssql
End If
tb.Open "Outward", CN, 0, 3, 0
For i = 1 To Grid1.Rows - 1
tb.AddNew
    tb.Fields("out_no").Value = Val(Text1.Text)
    tb.Fields("v_date").Value = date1.Value
    tb.Fields("accode").Value = Text3.Text
    tb.Fields("acname").Value = Text4.Text
    tb.Fields("vanno").Value = Text12.Text
    tb.Fields("E_Type").Value = 2
    
    With Grid1
    tb.Fields("cont_no").Value = Val(.TextMatrix(i, 0))
    tb.Fields("rolls").Value = Val(.TextMatrix(i, 4))
    tb.Fields("quantity").Value = Val(.TextMatrix(i, 5))
    tb.Fields("Rate").Value = Val(.TextMatrix(i, 6))
    tb.Fields("remarks").Value = .TextMatrix(i, 7)
    tb.Fields("Item").Value = Val(.TextMatrix(i, 8))
    tb.Fields("quality").Value = .TextMatrix(i, 2)
    tb.Fields("lyc_per").Value = Val(.TextMatrix(i, 9))
    tb.Fields("lyc_used").Value = Val(.TextMatrix(i, 10))
    tb.Fields("size1").Value = .TextMatrix(i, 11)
    
    End With
    
    If Option2 = True Then
'        tb.Fields("cancel").Value = Check1.Value
'        tb.Fields("c_date").Value = date4.Value
    End If

tb.Update
Next i
tb.Close
End Sub

Private Sub Clearfull()
Dim cntl As Control

For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
Next
Label23.Caption = vbNullString
Label21.Caption = vbNullString
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
R = MsgBox("Do You Want To Print", vbYesNo)
If R = vbYes Then
Load vour
vour.Caption = "Outward for Knitting Sale Contract"
vour.Text2.Text = 4
vour.Frame2.Visible = True
vour.Label1.Caption = "Outward #"
vour.Text1.Text = Text1.Text
vour.Show
vour.Text1.SetFocus
End If
If R = vbNo Then
Command2_Click
Option1 = True
End If
End Sub

Private Sub Command2_Click()
Call Clearfull
Flex1
'Text1.Enabled = False
If Option1 = True Then
Text1.Text = max1
End If
date1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide

End Sub

Private Sub Command4_Click()
If Val(Text10.Text) > 0 Then
Transfer1
clear
Text2.Text = ""
Else
    MsgBox "Please Complete The Entry"
End If
Text2.SetFocus
End Sub

Private Sub Form_Activate()
'Command2_Click
If Me.Visible = True Then
Option2 = True
Text1.Text = ""
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

date3.Value = Date
Flex1

ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo2, "name", "code"
Text1.Text = max1
End Sub

Private Sub Grid1_DblClick()
If Grid1.Rows > 1 Then
    With Grid1
    Text2.Text = .TextMatrix(.Row, 0)
    If Val(Text2.Text) > 0 Then
    date3.Value = .TextMatrix(.Row, 1)
    End If
    Text8.Text = .TextMatrix(.Row, 2)
    
    Label21.Caption = .TextMatrix(.Row, 3)
    Text10.Text = .TextMatrix(.Row, 4)
    Text11.Text = .TextMatrix(.Row, 5)
    Label23.Caption = .TextMatrix(.Row, 6)
    Text5.Text = .TextMatrix(.Row, 6)
    Text17.Text = .TextMatrix(.Row, 7)
    Text6.Text = .TextMatrix(.Row, 8)
    Text7.Text = .TextMatrix(.Row, 9)
    Text9.Text = .TextMatrix(.Row, 10)
    Text13.Text = .TextMatrix(.Row, 11)
    
    If .Rows > 2 Then
        .RemoveItem .Row
    Else
        .Rows = 1
    End If
    End With
    Text10.SetFocus
End If
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
Text6.Text = List1.ItemData(List1.ListIndex)
Text8.Text = List1.Text
End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
List1.Visible = False
End If

End Sub

Private Sub List1_LostFocus()
List1.Visible = False
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
Command2_Click
Text1.Text = max1
date1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
GOTF Text1
If Option1 = True Then
    Text1.Text = max1
End If
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

Private Sub Text10_Validate(Cancel As Boolean)
If Val(Text10.Text) > 0 Then
    Exit Sub
Else
'    Cancel = True
End If
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
Text9.Text = Val(Text7.Text) * Val(Text11.Text) / 100
Lostf Text11
End Sub


Private Sub Text12_GotFocus()
GOTF Text12
End Sub

Private Sub Text12_LostFocus()
Lostf Text12
End Sub

Private Sub Text17_GotFocus()
GOTF Text17
End Sub

Private Sub Text17_LostFocus()
Lostf Text17
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
aa = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
bb = blm.party1(tb.Fields("party").Value) & aa
CC = Mid(bb, 1, 28)
List3.AddItem CC & " " & tb.Fields("cont_no").Value & "               " & blm.Yarn(tb.Fields("yarncount").Value) & "              " & blm.Cloth(tb.Fields("item").Value) & tb.Fields("MGUage").Value

'List3.AddItem blm.party1(tb.Fields("party").Value) & "                " & tb.Fields("cont_no").Value & "               " & blm.Cloth(tb.Fields("item").Value) & "            " & blm.Yarn(tb.Fields("yarncount").Value) & "           " & tb.Fields("MGUage").Value
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

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub Text5_LostFocus()
Lostf Text5
End Sub

Private Sub Text6_GotFocus()
GOTF Text6
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim tb As New ADODB.Recordset
Dim i As Integer
tb.Open "cloth", CN, 0, 3, 0
List1.Visible = True
'Dim S As String
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

Private Sub Text6_LostFocus()
Lostf Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Val(Text6.Text) > 0 Then
    Text8.Text = blm.Cloth(Val(Text6.Text))
End If
End Sub


Private Sub Text7_GotFocus()
GOTF Text7
End Sub

Private Sub Text7_LostFocus()
Lostf Text7
End Sub

Private Sub Text9_GotFocus()
GOTF Text9
End Sub

Private Sub Text9_LostFocus()
Lostf Text9
End Sub

Private Sub Timer1_Timer()
Dim f As Integer, S As Integer
Text9.Text = Round(Val(Text11.Text) * Val(Text7.Text) / (100 + Val(Text7.Text)), 3)
'Text16.Text = Val(Text11.Text) - Val(Text13.Text) - Val(Text18.Text)
'Text15.Text = Val(Text13.Text) + Val(Text18.Text)
End Sub

Private Sub Timer2_Timer()
Label38.Caption = Val(Label16.Caption) - Val(Label30.Caption)
End Sub

Private Sub Timer3_Timer()
Dim R As Integer
Dim TQty As Double
For R = 1 To Grid1.Rows - 1
    TQty = TQty + Val(Grid1.TextMatrix(R, 5))
Next R
Label15.Caption = Format(TQty, "#.000")
End Sub
