VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cont_PNN 
   Caption         =   "Purchase Contract Entry (Knitting)"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   Icon            =   "CONT_PN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo7 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   78
      Text            =   "Combo7"
      Top             =   960
      Width           =   4335
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   51
      Top             =   6120
      Width           =   11415
      Begin VB.ComboBox Combo6 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   9240
         TabIndex        =   76
         Text            =   "Combo6"
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox Combo5 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   6720
         TabIndex        =   73
         Text            =   "Combo5"
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   3600
         TabIndex        =   55
         Text            =   "Combo3"
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo2 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   360
         TabIndex        =   54
         Text            =   "Combo2"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label31 
         Caption         =   "Lycra Counts"
         Height          =   255
         Left            =   9240
         TabIndex        =   75
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "Yarns Counts"
         Height          =   255
         Left            =   6720
         TabIndex        =   74
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Item List"
         Height          =   255
         Left            =   3600
         TabIndex        =   53
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "A/c List"
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   120
      TabIndex        =   38
      Top             =   1320
      Width           =   11415
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   5580
         TabIndex        =   30
         Top             =   4380
         Width           =   915
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   3480
         TabIndex        =   29
         Top             =   4380
         Width           =   975
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   4380
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   1440
         TabIndex        =   82
         Top             =   1860
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   1440
         TabIndex        =   81
         Top             =   2520
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   9720
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   7680
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   7920
         TabIndex        =   22
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox Text21 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   15
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   7920
         TabIndex        =   14
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   10200
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   7920
         TabIndex        =   18
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Text18 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text16 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   11160
         Top             =   120
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   26
         Top             =   3960
         Width           =   5055
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   10080
         TabIndex        =   25
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "CONT_PN.frx":030A
         Left            =   7680
         List            =   "CONT_PN.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Co&mplete"
         Height          =   375
         Left            =   5640
         TabIndex        =   57
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1800
         TabIndex        =   56
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7680
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   3960
         Width           =   3615
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   3480
         Width           =   5055
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3000
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20578307
         CurrentDate     =   36801
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   7920
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3840
         TabIndex        =   17
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text5 
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
         Left            =   9720
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   7920
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text3 
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
         Left            =   3840
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20578307
         CurrentDate     =   36801
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Opening Lycra"
         Height          =   255
         Left            =   4500
         TabIndex        =   85
         Top             =   4380
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Opening Fabric"
         Height          =   255
         Left            =   2340
         TabIndex        =   84
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "Opening Yarn"
         Height          =   252
         Left            =   324
         TabIndex        =   83
         Top             =   4380
         Width           =   1008
      End
      Begin VB.Label Label34 
         Caption         =   "St. Len."
         Height          =   255
         Left            =   8760
         TabIndex        =   80
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Grey GSM"
         Height          =   255
         Left            =   6600
         TabIndex        =   79
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Machine Guage Size"
         Height          =   495
         Left            =   6600
         TabIndex        =   72
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "Machine Name"
         Height          =   255
         Left            =   8520
         TabIndex        =   71
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Machine Code"
         Height          =   375
         Left            =   6600
         TabIndex        =   70
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Daily Prod. Kgs."
         Height          =   255
         Left            =   8520
         TabIndex        =   69
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Machine Booked"
         Height          =   255
         Left            =   6600
         TabIndex        =   68
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Lycra Count"
         Height          =   255
         Left            =   2880
         TabIndex        =   67
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Lycra Code"
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Yarn Count"
         Height          =   255
         Left            =   2880
         TabIndex        =   65
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label22 
         Height          =   255
         Left            =   10680
         TabIndex        =   64
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Yarn Code"
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "GST Reg #"
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Gst Ratio"
         Height          =   255
         Left            =   8880
         TabIndex        =   61
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "GST"
         Height          =   255
         Left            =   6600
         TabIndex        =   60
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Quantity Cloth"
         Height          =   255
         Left            =   8520
         TabIndex        =   59
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Lycra %"
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Reference"
         Height          =   255
         Left            =   6600
         TabIndex        =   50
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Payment"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Del Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   47
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Quantity Yarn"
         Height          =   255
         Left            =   6600
         TabIndex        =   46
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Rate"
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Fabric Name"
         Height          =   255
         Left            =   8520
         TabIndex        =   44
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Fabric Code"
         Height          =   252
         Left            =   6612
         TabIndex        =   43
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label4 
         Caption         =   "A/c Title"
         Height          =   255
         Left            =   3000
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "A/c Code"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   40
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Contract No."
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   8040
      TabIndex        =   37
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   2400
         Picture         =   "CONT_PN.frx":0321
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Clear"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   1320
         Picture         =   "CONT_PN.frx":062B
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   975
         Left            =   240
         Picture         =   "CONT_PN.frx":0935
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "&Update"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   1920
         Picture         =   "CONT_PN.frx":0C3F
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   975
         Left            =   360
         Picture         =   "CONT_PN.frx":0F49
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Label Label32 
      Caption         =   "Machine List"
      Height          =   255
      Left            =   3480
      TabIndex        =   77
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "cont_PNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private Sub clear()
Dim cntl As Control
For Each cntl In Me.Controls
    If TypeOf cntl Is TextBox Then cntl.Text = vbNullString
    If TypeOf cntl Is DTPicker Then cntl.Value = Date
    
    
Next
combs
Check2.Value = 0
Text1.Text = max1
Combo4.ListIndex = 0
End Sub
Private Sub edit1()
Dim tb As ADODB.Recordset
Dim ssql As String
ssql = "SELECT * FROM CONT_1 WHERE E_Type = 1 and CONT_NO = " & Val(Text1.Text)

Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("V_DATE").Value
    date2.Value = tb.Fields("DEL_DATE").Value
    Text2.Text = tb.Fields("PARTY").Value
    Text3.Text = blm.party1(tb.Fields("PARTY").Value)
    Text4.Text = tb.Fields("ITEM").Value
    Text5.Text = blm.Cloth(tb.Fields("ITEM").Value)
    Text6.Text = tb.Fields("RATE").Value
    Text8.Text = tb.Fields("LYCRA").Value
    Text7.Text = tb.Fields("YQUANTITY").Value
    Text9.Text = tb.Fields("CQUANTITY").Value
    Text10.Text = tb.Fields("PAYMENT").Value & ""
    Text11.Text = tb.Fields("REMARKS").Value & ""
    If Combo1.ListIndex > -1 Then
    Combo1.ListIndex = tb.Fields("REFERENCE").Value - 1
    End If
    If Not IsNull(tb.Fields("Complete").Value) Then
        Check2.Value = tb.Fields("Complete").Value
    End If
'    MsgBox tb.Fields("GST").Value
    Combo4.ListIndex = tb.Fields("GST").Value - 1
     Text12.Text = tb.Fields("GST_Rate").Value
     Text13.Text = tb.Fields("GST_No").Value & ""
     If Not IsNull(tb.Fields("YarnCount").Value) Then
     Text14.Text = tb.Fields("YarnCount").Value
     Text16.Text = blm.Yarn(tb.Fields("YarnCount").Value)
     End If
     If Not IsNull(tb.Fields("LycraCount").Value) Then
     Text17.Text = tb.Fields("LycraCount").Value
     Text18.Text = blm.Lycra(tb.Fields("LycraCount").Value)
     End If
     If Not IsNull(tb.Fields("Machine").Value) Then
     Text20.Text = tb.Fields("Machine").Value
     Text21.Text = blm.machine(tb.Fields("Machine").Value)
     End If
     Text23.Text = tb.Fields("greygsm").Value & ""
     Text24.Text = tb.Fields("st_len").Value & ""
     If Not IsNull(tb.Fields("MBOOKED").Value) Then
     Text15.Text = tb.Fields("MBOOKED").Value & ""
     End If
     Text22.Text = tb.Fields("MGUAGE").Value & ""
     If Not IsNull(tb.Fields("DPROD").Value) Then
     Text19.Text = tb.Fields("DPROD").Value
     End If
     If Not IsNull(tb.Fields("opbal")) Then
     Text25.Text = tb.Fields("opbal").Value
     Else
     Text25.Text = ""
     End If
     If Not IsNull(tb.Fields("opfabric")) Then
     Text26.Text = tb.Fields("opfabric").Value
     Else
     Text26.Text = ""
     End If
     If Not IsNull(tb.Fields("oplycra")) Then
     Text27.Text = tb.Fields("oplycra").Value
     Else
     Text27.Text = ""
     End If
Else
MsgBox "Invalid Contract No"
Command2_Click
End If
tb.Close

End Sub
Private Sub save()
Dim tb As New ADODB.Recordset
If Option2 = True Then
    Dim ssql As String
        ssql = "DELETE FROM CONT_1 WHERE E_Type=1 and CONT_NO = " & Val(Text1.Text)
        CN.Execute ssql
End If

tb.Open "CONT_1", CN, 0, 3, 0
tb.AddNew
    tb.Fields("CONT_NO").Value = Val(Text1.Text)
    tb.Fields("E_TYPE").Value = 1
    tb.Fields("V_DATE").Value = date1.Value
    tb.Fields("DEL_DATE").Value = date2.Value
    tb.Fields("PARTY").Value = Val(Text2.Text)
    tb.Fields("ITEM").Value = Val(Text4.Text)
    tb.Fields("RATE").Value = Val(Text6.Text)
    tb.Fields("YQUANTITY").Value = Val(Text7.Text)
    tb.Fields("CQUANTITY").Value = Val(Text9.Text)
    tb.Fields("Lycra").Value = Val(Text8.Text)
    tb.Fields("PAYMENT").Value = Text10.Text
    tb.Fields("REMARKS").Value = UCase(CStr(Text11.Text))
    If Combo1.ListIndex > -1 Then
    tb.Fields("REFERENCE").Value = Combo1.ItemData(Combo1.ListIndex)
    End If
    tb.Fields("Complete").Value = Check2.Value
    tb.Fields("GST").Value = Combo4.ItemData(Combo4.ListIndex)
    tb.Fields("GST_Rate").Value = Val(Text12.Text)
    tb.Fields("GST_No").Value = Text13.Text
    tb.Fields("yARNcOUNT").Value = Val(Text14.Text)
    tb.Fields("LYCRACOUNT").Value = Val(Text17.Text)
    tb.Fields("MACHINE").Value = Val(Text20.Text)
    tb.Fields("greygsm").Value = Text23.Text
    tb.Fields("st_len").Value = Text24.Text
    tb.Fields("MBooked").Value = Text15.Text
    tb.Fields("MGuage").Value = Text22.Text
    tb.Fields("DProd").Value = Val(Text19.Text)
    tb.Fields("opbal").Value = Val(Text25.Text)
    tb.Fields("opfabric").Value = Val(Text26.Text)
    tb.Fields("oplycra").Value = Val(Text27.Text)

tb.Update
tb.Close

End Sub
Private Function max1() As Long
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select MAX(CONT_NO) AS C FROM CONT_1 where e_type=1"
Set tb = CN.Execute(ssql)
If Not IsNull(tb.Fields("C").Value) Then
    max1 = tb.Fields("C").Value + 1
Else
    max1 = 1
End If
tb.Close
End Function

Private Sub combs()
Dim ssql As String

ssql = "select * from emp1 order by Emp_no"
blm.fill_comb ssql, Combo1, "name", "Emp_no"

ssql = "select * from acchart order by name"
blm.fill_comb ssql, Combo2, "name", "code"

ssql = "select * from Cloth order by code"
blm.fill_comb ssql, Combo3, "name", "code" ', "wIDTH"

ssql = "select * from Yarn where Y_type=1 order by code"
blm.fill_comb ssql, Combo5, "name", "code" ', "wIDTH"

ssql = "select * from Yarn where Y_type=2 order by code"
blm.fill_comb ssql, Combo6, "name", "code" ', "wIDTH"

ssql = "select * from Machine order by code"
blm.fill_comb ssql, Combo7, "name", "code" ', "wIDTH"


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Combo2_Click()
If Combo2.ListCount > 0 Then
If Check1.Value = 0 Then
    Text2.Text = Combo2.ItemData(Combo2.ListIndex)
    Text3.Text = Combo2.Text
End If
If Check1.Value = 1 Then
    Text8.Text = Combo2.ItemData(Combo2.ListIndex)
    Text9.Text = Combo2.Text
End If

End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Check1.Value = 0 Then
    Text2.SetFocus
Else
    Text8.SetFocus
End If
End If
End Sub

Private Sub Combo2_LostFocus()
Check1.Value = 0
End Sub

Private Sub Combo3_Click()
If Combo3.ListCount > 0 Then
    Text4.Text = Combo3.ItemData(Combo3.ListIndex)
    Text5.Text = Combo3.Text
End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub Combo5_Click()
Text14.Text = Combo5.ItemData(Combo5.ListIndex)
Text16.Text = Combo5.Text

End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text14.SetFocus
End Sub

Private Sub Combo6_Click()
Text17.Text = Combo6.ItemData(Combo6.ListIndex)
Text18.Text = Combo6.Text
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text17.SetFocus
End Sub

Private Sub Combo7_Click()
Text20.Text = Combo7.ItemData(Combo7.ListIndex)
Text21.Text = Combo7.Text
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text20.SetFocus
End Sub

Private Sub Command1_Click()
Call save
Command2_Click
End Sub

Private Sub Command2_Click()
Call clear

If Option2 = True Then
    Text1.SetFocus
Else
    Text2.SetFocus
End If

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    keybd_event VK_TAB, 0, 1, 0
    keybd_event VK_TAB, 0, 3, 0
End If
End Sub

Private Sub Form_Load()
Me.Top = ((MDIForm1.Height - Me.Height) / 2)
Me.Left = (MDIForm1.Width - Me.Width) / 2
combs
Text1.Text = max1
date1.Value = Date
date2.Value = Date
Combo4.ListIndex = 0
End Sub

Private Sub List1_Click()
If List1.ListIndex > -1 Then
Text14.Text = List1.ItemData(List1.ListIndex)
Text16.Text = List1.Text
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text14.SetFocus
List1.Visible = False
End If
End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub List2_Click()
If List2.ListIndex > -1 Then
Text17.Text = List2.ItemData(List2.ListIndex)
Text18.Text = List2.Text
End If
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text17.SetFocus
List2.Visible = False
End If
End Sub

Private Sub List2_LostFocus()
List2.Visible = False
End Sub

Private Sub Option1_Click()
clear
Check2.Visible = False
Text1.Enabled = False
Text2.SetFocus
End Sub

Private Sub Option2_Click()
clear
Check2.Visible = True
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
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

Private Sub Text1_Validate(Cancel As Boolean)
If Option2 = True Then
    If Val(Text1.Text) > 0 Then
        Call edit1
    Else
        Cancel = True
    End If
End If
End Sub

Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
End Sub

Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12.Text)

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        KeyAscii = 0
        Beep
End If

End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo5.SetFocus
Dim ssql As String
If KeyCode = vbKeyDown Then
Dim tb As ADODB.Recordset
Dim i As Integer
ssql = "select * from yarn where y_type=1 order by code"
Set tb = CN.Execute(ssql)
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

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text14_Validate(Cancel As Boolean)
If Val(Text14.Text) > 0 Then
    Text16.Text = blm.Yarn(Val(Text14.Text))
    Else
    Text16.Text = ""
End If
End Sub

Private Sub Text17_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo6.SetFocus
Dim ssql As String
If KeyCode = vbKeyDown Then
Dim tb As ADODB.Recordset
Dim i As Integer
ssql = "select * from yarn where y_type=2 order by code"
Set tb = CN.Execute(ssql)
List2.Visible = True
'Dim S As String
If Not tb.EOF Then
List2.clear
Do While Not tb.EOF
List2.AddItem tb.Fields("name").Value
List2.ItemData(List2.NewIndex) = tb.Fields("code").Value
tb.MoveNext
Loop
List2.ListIndex = 0
End If
tb.Close
Set tb = Nothing
List2.SetFocus
End If

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text17_Validate(Cancel As Boolean)
If Val(Text17.Text) > 0 Then
    Text18.Text = blm.Lycra(Val(Text17.Text))
    Else
    Text18.Text = ""

End If
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Val(Text2.Text) > 0 Then
    Text3.Text = blm.party1(Val(Text2.Text))
        If Text3.Text = "NOT" Then
            Cancel = True
        Else
            Text13.Text = blm.GST1(Val(Text2.Text))
        End If
Else
        Cancel = True
End If
End Sub

Private Sub Text20_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo7.SetFocus
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text20_Validate(Cancel As Boolean)
If Val(Text20.Text) Then
    Text21.Text = blm.machine(Val(Text20.Text))
End If
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Combo3.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) > 0 Then
    Text5.Text = blm.Cloth(Val(Text4.Text))
        If Text5.Text = "NOT" Then
            Cancel = True
        End If
Else
        Cancel = True
End If

End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        KeyAscii = 0
        Beep
End If

End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Val(Text6.Text) > 0 Then
    Exit Sub
Else
    Cancel = True
End If
End Sub

Private Sub Text7_GotFocus()
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        KeyAscii = 0
        Beep
End If

End Sub

Private Sub Text8_GotFocus()
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        Beep
        KeyAscii = 0
End If

End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 8 Then

Else
        KeyAscii = 0
        Beep
End If


End Sub

Private Sub Timer1_Timer()
Text9.Text = Format(Val(Text7.Text) * 44.2, "#.0")
End Sub
