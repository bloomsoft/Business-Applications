VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Ins 
   Caption         =   "Needles Out To Machine"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7095
   Icon            =   "Ins.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   7095
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   555
      Left            =   240
      TabIndex        =   45
      Top             =   2910
      Width           =   11505
      Begin VB.Line Line15 
         BorderColor     =   &H00808080&
         X1              =   8730
         X2              =   8730
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00808080&
         X1              =   7710
         X2              =   7710
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00808080&
         X1              =   10110
         X2              =   10110
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         X1              =   7260
         X2              =   7260
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   6270
         X2              =   6270
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   5490
         X2              =   5490
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   4740
         X2              =   4740
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   4020
         X2              =   4020
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   3330
         X2              =   3330
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   2670
         X2              =   2670
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   1740
         X2              =   1740
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   990
         X2              =   990
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   420
         X2              =   420
         Y1              =   90
         Y2              =   540
      End
      Begin VB.Label Label26 
         Caption         =   "Remarks"
         Height          =   315
         Left            =   10440
         TabIndex        =   59
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label27 
         Caption         =   " îŒBèe ¹ÍiBIBÌ¿"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8730
         TabIndex        =   58
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label Label25 
         Caption         =   "  BÄ¼†B· ÓÍÌm"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7650
         TabIndex        =   57
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label24 
         Caption         =   " îƒ Êj¸Î»"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5490
         TabIndex        =   56
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   " “iBq Êj¸Î»"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6270
         TabIndex        =   55
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label21 
         Caption         =   "ºÌ"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   54
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label20 
         Caption         =   "AjíŒ î‘‚"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4800
         TabIndex        =   53
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label18 
         Caption         =   "B¸¼è î‘‚"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   52
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "BÃjMóA jÎ‚"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3390
         TabIndex        =   51
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label16 
         Caption         =   " fÄI îÄ¿"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2670
         TabIndex        =   50
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label14 
         Caption         =   "ÔiBèe ½ÎM"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   49
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "îJèe ½ÎM"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   48
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label11 
         Caption         =   " ÆkË"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   510
         TabIndex        =   47
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label10 
         Caption         =   " ¾Ëi"
         BeginProperty Font 
            Name            =   "AlKatib1"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   46
         Top             =   120
         Width           =   435
      End
   End
   Begin VB.CheckBox C1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1500
      TabIndex        =   44
      Top             =   4740
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox T1 
      Height          =   285
      Left            =   810
      TabIndex        =   43
      Top             =   4110
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.ListBox List4 
      Height          =   450
      Left            =   1530
      TabIndex        =   22
      Top             =   2670
      Visible         =   0   'False
      Width           =   4575
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
      Height          =   1725
      Left            =   240
      TabIndex        =   23
      Top             =   1110
      Width           =   11505
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   12810
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   810
         Width           =   1515
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
         Left            =   8310
         TabIndex        =   10
         Top             =   2430
         Visible         =   0   'False
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
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2520
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   1290
         TabIndex        =   6
         Top             =   1230
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
         Left            =   4200
         TabIndex        =   7
         Top             =   1230
         Width           =   7020
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Accept Entry"
         Height          =   375
         Left            =   9180
         TabIndex        =   8
         Top             =   1980
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   7110
         TabIndex        =   5
         Top             =   810
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   4200
         TabIndex        =   1
         Top             =   390
         Width           =   1515
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
         Left            =   7110
         TabIndex        =   2
         Top             =   390
         Width           =   4095
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
         Format          =   20578307
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
         Format          =   20578307
         CurrentDate     =   36749
      End
      Begin VB.Label Label9 
         Caption         =   "Part Name"
         Height          =   255
         Left            =   7440
         TabIndex        =   42
         Top             =   2400
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "Part No"
         Height          =   255
         Left            =   12180
         TabIndex        =   41
         Top             =   2430
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Fabric Construction"
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "Shift"
         Height          =   255
         Left            =   90
         TabIndex        =   39
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "Employee"
         Height          =   255
         Left            =   3540
         TabIndex        =   37
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Machine Code"
         Height          =   255
         Left            =   90
         TabIndex        =   36
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Machine Name"
         Height          =   255
         Left            =   2760
         TabIndex        =   35
         Top             =   1260
         Width           =   1215
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
         Caption         =   "Inspection M/S"
         Height          =   255
         Left            =   5880
         TabIndex        =   30
         Top             =   390
         Width           =   1155
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
         Caption         =   "Operator Name"
         Height          =   255
         Left            =   5850
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Receipt #"
         Height          =   255
         Left            =   2790
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
      Left            =   3510
      TabIndex        =   20
      Top             =   -900
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
      Height          =   3195
      Left            =   225
      TabIndex        =   9
      Top             =   3150
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5636
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      Height          =   1095
      Left            =   8250
      TabIndex        =   19
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Exit"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2280
         Picture         =   "Ins.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Reset"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1200
         Picture         =   "Ins.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   120
         Picture         =   "Ins.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Change"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   1800
         Picture         =   "Ins.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&New"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   240
         Picture         =   "Ins.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Ins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Dim org_q As Currency
Dim rej As Currency
Dim FRolls As Double
Dim FWt As Double
Dim CHPREins As Boolean
Private Sub PREINS()
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from INS where Ino = " & Val(Text1.Text) & " order by roll"
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
CHPREins = True

'MsgBox "This Inspection Is Already Exist"
Else
CHPREins = False
End If
End Sub


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
    .Cols = 15
    .ColWidth(0) = 1
    .TextMatrix(0, 0) = "Sr.No"
    .ColWidth(1) = 400
    .TextMatrix(0, 1) = "Roll"
    .ColWidth(2) = 570
    .TextMatrix(0, 2) = "Weight"
    .ColWidth(3) = 750
    .TextMatrix(0, 3) = "1"
    .ColWidth(4) = 935
    .TextMatrix(0, 4) = "2"
    .ColWidth(5) = 665
    .TextMatrix(0, 5) = "3"
    .ColWidth(6) = 690
    .TextMatrix(0, 6) = "4"
    .ColWidth(7) = 715
    .TextMatrix(0, 7) = "5"
    .ColWidth(8) = 750
    .TextMatrix(0, 8) = "6"
    .ColWidth(9) = 780
    .TextMatrix(0, 9) = "7"
    .ColWidth(10) = 990
    .TextMatrix(0, 10) = "8"
    .ColWidth(11) = 450
    .TextMatrix(0, 11) = "9"
    .ColWidth(12) = 1020
    .TextMatrix(0, 12) = "10"
    .ColWidth(13) = 1380
    .TextMatrix(0, 13) = "11"
    .ColWidth(14) = 1340
    .TextMatrix(0, 14) = "Remarks"
    
End With
Dim Pwt As Double
If FRolls > 0 Then
Pwt = 0
Pwt = FWt / FRolls
Grid1.Rows = 1
For i = 1 To FRolls
Grid1.Rows = Grid1.Rows + 1
Grid1.TextMatrix(i, 1) = i
Grid1.TextMatrix(i, 2) = Round(Pwt, 3)
Next i
Else
'MsgBox "Invalid Receipt No"
End If
End Sub
Private Sub Flx()
With Grid1
    .Rows = 1
    .Cols = 15
    .ColWidth(0) = 1
    .TextMatrix(0, 0) = "Sr.No"
    .ColWidth(1) = 900
    .TextMatrix(0, 1) = "Roll"
    .ColWidth(2) = 1000
    .TextMatrix(0, 2) = "Weight"
    .ColWidth(3) = 1000
    .TextMatrix(0, 3) = "Oil/Dust Spot"
    .ColWidth(4) = 1000
    .TextMatrix(0, 4) = "Oil Line"
    .ColWidth(5) = 1000
    .TextMatrix(0, 5) = "Pulling Thread"
    .ColWidth(6) = 1000
    .TextMatrix(0, 6) = "Bowing"
    .ColWidth(7) = 1000
    .TextMatrix(0, 7) = "Patta"
    .ColWidth(8) = 1000
    .TextMatrix(0, 8) = "Contamination"
    .ColWidth(9) = 1000
    .TextMatrix(0, 9) = "Needle/Sinker Line"
    .ColWidth(10) = 1000
    .TextMatrix(0, 10) = "Knit Hole"
    .ColWidth(11) = 1000
    .TextMatrix(0, 11) = "Lycra Short"
    .ColWidth(12) = 1000
    .TextMatrix(0, 12) = "Misc Prob"
    .ColWidth(13) = 1500
    .TextMatrix(0, 13) = "Discription Of Misc.Prob"
    .ColWidth(14) = 2000
    .TextMatrix(0, 14) = "Remarks"
    
End With
End Sub
Private Function Check(c As Long) As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String
    
ssql = "select * from INS where in_no = " & c
ssql = ssql & " and E_type=2"
Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    MsgBox "INS No already Exist..."
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
Private Sub FRec()
FRolls = 0
FWt = 0
    Dim ssql As String
    Dim tb As ADODB.Recordset
    
    ssql = "select rolls as c, quantity as w from fabricrcvd where out_no=" & Val(Text1.Text)
    Set tb = CN.Execute(ssql)
    If Not tb.EOF Then
    FRolls = tb.Fields("c").Value
    FWt = tb.Fields("w").Value
    Else
    FRolls = 0
    End If
    tb.Close
Flex1
End Sub

Private Function max1() As Double
'    Dim ssql As String
'    Dim tb As ADODB.Recordset
'
'    ssql = "select max(Ino)as c from INS"
'    Set tb = CN.Execute(ssql)
'    If IsNull(tb.Fields("c").Value) = False Then
'        max1 = tb.Fields("c").Value + 1
'    Else
'        max1 = 1
'    End If
'    tb.Close
End Function
Private Function edit_kachi() As Boolean
Dim tb As ADODB.Recordset
Dim ssql As String

ssql = "select * from INS where Ino = " & Val(Text1.Text) & " order by roll"

Set tb = CN.Execute(ssql)
If Not tb.EOF Then
    date1.Value = tb.Fields("Idate").Value
    Text17.Text = tb.Fields("ims").Value & ""
    Text3.Text = tb.Fields("shift").Value & ""
    Text6.Text = tb.Fields("fabric_cons").Value & ""
    Text5.Text = tb.Fields("opname").Value & ""
    Text21.Text = tb.Fields("machinecode").Value
    Text20.Text = blm.machine(tb.Fields("machinecode").Value)
          
   Do While Not tb.EOF
   With Grid1
   .Rows = .Rows + 1
    
    .TextMatrix(.Rows - 1, 1) = tb.Fields("roll").Value
    .TextMatrix(.Rows - 1, 2) = tb.Fields("wt").Value
    .TextMatrix(.Rows - 1, 3) = tb.Fields("spot").Value
    .TextMatrix(.Rows - 1, 4) = tb.Fields("oilline").Value
    .TextMatrix(.Rows - 1, 5) = tb.Fields("pulling").Value
    .TextMatrix(.Rows - 1, 6) = tb.Fields("bowing").Value
    .TextMatrix(.Rows - 1, 7) = tb.Fields("patta").Value
    .TextMatrix(.Rows - 1, 8) = tb.Fields("contami").Value
    .TextMatrix(.Rows - 1, 9) = tb.Fields("n_s_line").Value
    .TextMatrix(.Rows - 1, 10) = tb.Fields("hole").Value
    .TextMatrix(.Rows - 1, 11) = tb.Fields("short").Value
    .TextMatrix(.Rows - 1, 12) = tb.Fields("misc").Value
    .TextMatrix(.Rows - 1, 13) = tb.Fields("misc_des").Value & ""
    .TextMatrix(.Rows - 1, 14) = tb.Fields("remarks").Value & ""
    
    
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
    ssql = "delete from INS where Ino = " & Val(Text1.Text)
    CN.Execute ssql
End If
tb.Open "INS", CN, 0, 3, 0
For i = 1 To Grid1.Rows - 1

With Grid1
tb.AddNew
    tb.Fields("Ino").Value = Val(Text1.Text)
    tb.Fields("Idate").Value = date1.Value
    tb.Fields("ims").Value = Text17.Text
    tb.Fields("shift").Value = Text3.Text
    tb.Fields("fabric_cons").Value = Text6.Text
    tb.Fields("opname").Value = Text5.Text
    tb.Fields("machinecode").Value = Val(Text21.Text)
        
    tb.Fields("roll").Value = Val(.TextMatrix(i, 1))
    tb.Fields("wt").Value = Val(.TextMatrix(i, 2))
    
    tb.Fields("spot").Value = Val(.TextMatrix(i, 3))
    tb.Fields("oilline").Value = Val(.TextMatrix(i, 4))
    tb.Fields("pulling").Value = Val(.TextMatrix(i, 5))
    tb.Fields("bowing").Value = Val(.TextMatrix(i, 6))
    tb.Fields("patta").Value = Val(.TextMatrix(i, 7))
    tb.Fields("contami").Value = Val(.TextMatrix(i, 8))
    tb.Fields("n_s_line").Value = Val(.TextMatrix(i, 9))
    tb.Fields("hole").Value = Val(.TextMatrix(i, 10))
    tb.Fields("short").Value = Val(.TextMatrix(i, 11))
    tb.Fields("misc").Value = Val(.TextMatrix(i, 12))
    tb.Fields("misc_des").Value = Val(.TextMatrix(i, 13))
    tb.Fields("remarks").Value = .TextMatrix(i, 14)

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
If Option1 = True Then
   ' Text1.Text = max1
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
    vour.Caption = "Inspection Report"
    vour.Text2.Text = 68
    vour.Text1.Text = Text1.Text
    vour.Label1.Caption = "Inspection #"
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
CHPREins = False
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
   ' Text1.Text = max1
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
CHPREins = False
Me.Top = ((Screen.Height - Me.Height) / 2) - 1000
Me.Left = (Screen.Width - Me.Width) / 2
date1.Value = Date
Flx
FRolls = 0
FWt = 0
ssql = "select * from acchart order by name"
'blm.fill_comb ssql, Combo2, "name", "code"
'Text1.Text = max1

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

Private Sub Grid1_EnterCell()
Dim R As Long
If Grid1.Col > 1 Then
If Grid1.Col = 7 Or Grid1.Col = 8 Then
    C1.Top = Grid1.CellTop + Grid1.Top
    C1.Left = Grid1.CellLeft + Grid1.Left
    C1.Width = Grid1.CellWidth
    C1.Height = Grid1.CellHeight
    C1.Visible = True
    C1.Value = Val(Grid1.Text)
    C1.SetFocus

Else
    T1.Top = Grid1.CellTop + Grid1.Top
    T1.Left = Grid1.CellLeft + Grid1.Left
    T1.Width = Grid1.CellWidth
    T1.Height = Grid1.CellHeight
    T1.Visible = True
    T1.Text = Grid1.Text
    T1.SetFocus

End If
End If
End Sub

Private Sub Grid1_LeaveCell()

If Grid1.Col > 1 Then

If Grid1.Col = 7 Or Grid1.Col = 8 Then
    Grid1.Text = C1.Value
    C1.Visible = False

Else
    Grid1.Text = T1.Text
    T1.Visible = False

End If

End If

End Sub

Private Sub Grid1_Scroll()
C1.Visible = False
T1.Visible = False
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
'Text1.Enabled = False
Check1.Visible = False
date4.Visible = False
Command2_Click
CHPREins = False
date1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.Enabled = True
Check1.Visible = True
date4.Visible = True
'Text1.SetFocus

End Sub

Private Sub T1_GotFocus()
T1.SelStart = 0
T1.SelLength = Len(T1.Text)
End Sub

Private Sub Text1_Change()
PREINS
    If CHPREins = True Then
    Option2 = True
    Else
'    Option1 = True
    End If

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

Public Sub Text1_Validate(Cancel As Boolean)
Dim b As Boolean
If Val(Text1.Text) > 0 Then
    
    If Option1 = True Then
        FRec
    End If
    If Option2 = True Then
    Grid1.Rows = 1
        
        b = edit_kachi
        If b = False Then
            Cancel = True
        End If
    End If
Else
Cancel = True
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
'GOTF Text14
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
'Lostf Text14
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
'GOTF Text16
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

Private Sub Text21_GotFocus()
GOTF Text21
End Sub

Private Sub Text21_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from Machine where Status=0"
Set tb = CN.Execute(S)
List4.Visible = True
If Not tb.EOF Then
List4.clear
Do While Not tb.EOF
List4.AddItem tb.Fields("name").Value
List4.ItemData(List4.NewIndex) = tb.Fields("code").Value
tb.MoveNext
Loop
List4.ListIndex = 0
End If
tb.Close
Set tb = Nothing
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

Private Sub Text3_GotFocus()
GOTF Text3
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
Dim tb As ADODB.Recordset
Dim i As Integer
Dim S As String
S = "select * from parts"
Set tb = CN.Execute(S)
List1.Visible = True
If Not tb.EOF Then
List1.clear
Do While Not tb.EOF
List1.AddItem tb.Fields("partname").Value
List1.ItemData(List1.NewIndex) = tb.Fields("partcode").Value
tb.MoveNext
Loop
List1.ListIndex = 0
End If
tb.Close
Set tb = Nothing
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
