VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form newDocs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Clearing Documents Definition"
   ClientHeight    =   6135
   ClientLeft      =   2115
   ClientTop       =   840
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   705
      Left            =   6015
      Picture         =   "NewDocs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   30
      Width           =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      CausesValidation=   0   'False
      Height          =   645
      Left            =   6015
      Picture         =   "NewDocs.frx":04EE
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   735
      Width           =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      CausesValidation=   0   'False
      Height          =   600
      Left            =   6015
      Picture         =   "NewDocs.frx":09F5
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1380
      Width           =   840
   End
   Begin VB.Frame Frame4 
      Height          =   4125
      Left            =   4560
      TabIndex        =   79
      Top             =   1980
      Width           =   2370
      Begin VB.TextBox txtPCT 
         Height          =   285
         Left            =   1020
         MaxLength       =   100
         TabIndex        =   43
         Top             =   3810
         Width           =   1260
      End
      Begin VB.TextBox txtExamineDate 
         Height          =   285
         Left            =   90
         MaxLength       =   100
         TabIndex        =   42
         Top             =   3540
         Width           =   2190
      End
      Begin VB.TextBox txtDocumentRecDate 
         Height          =   285
         Left            =   90
         MaxLength       =   100
         TabIndex        =   41
         Top             =   3075
         Width           =   2205
      End
      Begin VB.TextBox txtShippingBill 
         Height          =   285
         Left            =   90
         MaxLength       =   100
         TabIndex        =   40
         Top             =   2610
         Width           =   2205
      End
      Begin VB.TextBox txtwdbmemo 
         Height          =   285
         Left            =   90
         MaxLength       =   100
         TabIndex        =   39
         Top             =   2160
         Width           =   2220
      End
      Begin VB.TextBox txtBL 
         Height          =   285
         Left            =   90
         MaxLength       =   100
         TabIndex        =   38
         Top             =   1515
         Width           =   2220
      End
      Begin VB.TextBox txtInvoicePackingList 
         Height          =   285
         Left            =   1650
         MaxLength       =   100
         TabIndex        =   37
         Top             =   990
         Width           =   660
      End
      Begin VB.TextBox txtBETriplicate 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   36
         Top             =   720
         Width           =   1230
      End
      Begin VB.TextBox txtFileNo 
         Height          =   285
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   35
         Top             =   450
         Width           =   1260
      End
      Begin VB.TextBox txtRefNo 
         Height          =   285
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   34
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label Label27 
         Caption         =   "PCT"
         Height          =   195
         Left            =   90
         TabIndex        =   89
         Top             =   3840
         Width           =   765
      End
      Begin VB.Label Label26 
         Caption         =   "Consignment Examined on"
         Height          =   195
         Left            =   75
         TabIndex        =   88
         Top             =   3345
         Width           =   2130
      End
      Begin VB.Label Label25 
         Caption         =   "Documents Received on"
         Height          =   195
         Left            =   90
         TabIndex        =   87
         Top             =   2895
         Width           =   2130
      End
      Begin VB.Label Label24 
         Caption         =   "Shipping Bill 'E' Form No."
         Height          =   195
         Left            =   90
         TabIndex        =   86
         Top             =   2430
         Width           =   2130
      End
      Begin VB.Label Label23 
         Caption         =   "Wharfage/Demurrage/Barge Memo"
         Height          =   360
         Left            =   90
         TabIndex        =   85
         Top             =   1785
         Width           =   2220
      End
      Begin VB.Label Label22 
         Caption         =   "B/L Original / Copies No."
         Height          =   195
         Left            =   90
         TabIndex        =   84
         Top             =   1320
         Width           =   1830
      End
      Begin VB.Label Label21 
         Caption         =   "Invoice, Packing List"
         Height          =   195
         Left            =   90
         TabIndex        =   83
         Top             =   1065
         Width           =   1560
      End
      Begin VB.Label Label20 
         Caption         =   "B/E Triplicate"
         Height          =   195
         Left            =   90
         TabIndex        =   82
         Top             =   795
         Width           =   1125
      End
      Begin VB.Label Label19 
         Caption         =   "File No."
         Height          =   195
         Left            =   90
         TabIndex        =   81
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Ref. No"
         Height          =   195
         Left            =   105
         TabIndex        =   80
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Customs Details"
      Height          =   4125
      Left            =   60
      TabIndex        =   54
      Top             =   1980
      Width           =   4485
      Begin VB.TextBox txtGoodsDelTo 
         Height          =   285
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   32
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtBondBillNR 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   30
         Top             =   3180
         Width           =   1365
      End
      Begin VB.TextBox txtCHCRNo 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   28
         Top             =   2910
         Width           =   1365
      End
      Begin VB.TextBox txtCAARNo 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   26
         Top             =   2640
         Width           =   1365
      End
      Begin VB.TextBox txtBargememo 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   24
         Top             =   2370
         Width           =   1365
      End
      Begin VB.TextBox txtDemurrage 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2100
         Width           =   1365
      End
      Begin VB.TextBox txtDMI 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1830
         Width           =   1365
      End
      Begin VB.TextBox txtWharfage 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1560
         Width           =   1365
      End
      Begin VB.TextBox txtImportValue 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1305
         Width           =   1365
      End
      Begin VB.TextBox txtBondCashNo 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1020
         Width           =   1365
      End
      Begin VB.TextBox txtIndex 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   14
         Top             =   765
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   255
         Left            =   3180
         TabIndex        =   12
         Top             =   510
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin VB.TextBox txtIGMNo 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   11
         Top             =   495
         Width           =   1365
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   3180
         MaxLength       =   100
         TabIndex        =   10
         Top             =   225
         Width           =   1260
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1305
         MaxLength       =   100
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   255
         Left            =   3180
         TabIndex        =   16
         Top             =   1050
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   255
         Left            =   3180
         TabIndex        =   19
         Top             =   1590
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   255
         Left            =   3180
         TabIndex        =   21
         Top             =   1875
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   255
         Left            =   3180
         TabIndex        =   23
         Top             =   2130
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker8 
         Height          =   255
         Left            =   3180
         TabIndex        =   25
         Top             =   2385
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker9 
         Height          =   255
         Left            =   3180
         TabIndex        =   27
         Top             =   2625
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker10 
         Height          =   255
         Left            =   3165
         TabIndex        =   29
         Top             =   2910
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker11 
         Height          =   255
         Left            =   3165
         TabIndex        =   31
         Top             =   3180
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin MSComCtl2.DTPicker DTPicker12 
         Height          =   255
         Left            =   3165
         TabIndex        =   33
         Top             =   3465
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   13
         Left            =   2745
         TabIndex        =   78
         Top             =   3495
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   12
         Left            =   2745
         TabIndex        =   77
         Top             =   3225
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   11
         Left            =   2745
         TabIndex        =   76
         Top             =   2955
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   10
         Left            =   2745
         TabIndex        =   75
         Top             =   2670
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   9
         Left            =   2745
         TabIndex        =   74
         Top             =   2415
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   8
         Left            =   2745
         TabIndex        =   73
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   7
         Left            =   2745
         TabIndex        =   72
         Top             =   1875
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   6
         Left            =   2745
         TabIndex        =   71
         Top             =   1605
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   5
         Left            =   2745
         TabIndex        =   70
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
         Height          =   225
         Index           =   4
         Left            =   2730
         TabIndex        =   69
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   225
         Index           =   3
         Left            =   2730
         TabIndex        =   68
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "GOODS DELIVERED TO"
         Height          =   225
         Index           =   2
         Left            =   75
         TabIndex        =   67
         Top             =   3510
         Width           =   2130
      End
      Begin VB.Label Label7 
         Caption         =   "Bond Bill NR."
         Height          =   225
         Index           =   1
         Left            =   75
         TabIndex        =   66
         Top             =   3165
         Width           =   1005
      End
      Begin VB.Label Label17 
         Caption         =   "C.H.C. R/NO."
         Height          =   225
         Left            =   90
         TabIndex        =   65
         Top             =   2925
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "C.A.A. R/NO."
         Height          =   225
         Left            =   90
         TabIndex        =   64
         Top             =   2655
         Width           =   1005
      End
      Begin VB.Label Label15 
         Caption         =   "Barg Memo No."
         Height          =   225
         Left            =   90
         TabIndex        =   63
         Top             =   2385
         Width           =   1005
      End
      Begin VB.Label Label14 
         Caption         =   "DEMURRAGE"
         Height          =   225
         Left            =   90
         TabIndex        =   62
         Top             =   2130
         Width           =   1110
      End
      Begin VB.Label Label13 
         Caption         =   "DMI"
         Height          =   225
         Left            =   105
         TabIndex        =   61
         Top             =   1875
         Width           =   1005
      End
      Begin VB.Label Label12 
         Caption         =   "WHARFAGE"
         Height          =   225
         Left            =   90
         TabIndex        =   60
         Top             =   1650
         Width           =   1005
      End
      Begin VB.Label Label11 
         Caption         =   "IMPORT VALUE"
         Height          =   225
         Left            =   90
         TabIndex        =   59
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Label Label10 
         Caption         =   "Bond/Cash No."
         Height          =   225
         Left            =   90
         TabIndex        =   58
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label9 
         Caption         =   "INDEX"
         Height          =   225
         Left            =   90
         TabIndex        =   57
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "IGM No."
         Height          =   225
         Left            =   90
         TabIndex        =   56
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "From"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   55
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   75
      TabIndex        =   51
      Top             =   1140
      Width           =   5880
      Begin VB.CheckBox Check1 
         Caption         =   "Chec&k if This Documents Cleared or Closed"
         Height          =   240
         Left            =   2370
         TabIndex        =   93
         Top             =   180
         Width           =   3375
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2940
         Top             =   180
      End
      Begin VB.TextBox txtPartyName 
         Height          =   285
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   510
         Width           =   4395
      End
      Begin VB.TextBox txtPartyCode 
         Height          =   285
         Left            =   270
         TabIndex        =   7
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Party Name"
         Height          =   225
         Left            =   1335
         TabIndex        =   53
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Party Code"
         Height          =   255
         Left            =   255
         TabIndex        =   52
         Top             =   225
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   90
      TabIndex        =   13
      Top             =   -45
      Width           =   5850
      Begin VB.TextBox txtClear 
         Height          =   285
         Left            =   4425
         TabIndex        =   6
         Top             =   870
         Width           =   1305
      End
      Begin VB.TextBox txtGoods 
         Height          =   285
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   5
         Top             =   870
         Width           =   3075
      End
      Begin VB.TextBox txtPackages 
         Height          =   285
         Left            =   225
         MaxLength       =   50
         TabIndex        =   4
         Top             =   870
         Width           =   1080
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3195
         TabIndex        =   2
         Top             =   375
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin VB.TextBox txtLCNo 
         Height          =   285
         Left            =   1365
         TabIndex        =   1
         Top             =   390
         Width           =   1830
      End
      Begin VB.TextBox txtSerialNo 
         Height          =   285
         Left            =   225
         TabIndex        =   0
         Top             =   390
         Width           =   1080
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   4485
         TabIndex        =   3
         Top             =   375
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39141
      End
      Begin VB.Label Label30 
         Caption         =   "Clear/Ship/S.S."
         Height          =   225
         Left            =   4500
         TabIndex        =   92
         Top             =   690
         Width           =   1245
      End
      Begin VB.Label Label29 
         Caption         =   "Goods"
         Height          =   240
         Left            =   1395
         TabIndex        =   91
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label Label28 
         Caption         =   "Packages"
         Height          =   195
         Left            =   255
         TabIndex        =   90
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "Karachi Date"
         Height          =   180
         Left            =   4530
         TabIndex        =   50
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   180
         Left            =   3255
         TabIndex        =   49
         Top             =   135
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "L/C No."
         Height          =   240
         Left            =   1365
         TabIndex        =   48
         Top             =   150
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Serial No."
         Height          =   210
         Left            =   240
         TabIndex        =   47
         Top             =   150
         Width           =   885
      End
   End
End
Attribute VB_Name = "newDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditMode As Boolean
Private Blm1 As New bloom1
Private Sub Edit1()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\Bloom.mdb")
Ssql = "Select * from Docs where SrNo=" & Val(txtSerialNo.Text)
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    EditMode = True
    Check1.Value = TB.Fields("Status").Value
    txtSerialNo.Text = TB.Fields("SrNo").Value & ""
    txtLCNo.Text = TB.Fields("LCNo").Value & ""
    DTPicker1.Value = TB.Fields("LCDate").Value
    DTPicker2.Value = TB.Fields("KarachiDate").Value
    txtPackages.Text = TB.Fields("Packages").Value & ""
    txtGoods.Text = TB.Fields("Goods").Value & ""
    txtClear.Text = TB.Fields("PerSS").Value & ""
    txtPartyCode.Text = TB.Fields("PartyCode").Value
    txtPartyName.Text = Blm1.party1(TB.Fields("partyCode").Value)
    txtFrom.Text = TB.Fields("ShipFrom").Value & ""
    txtTo.Text = TB.Fields("Shipto").Value & ""
    txtIGMNo.Text = TB.Fields("IGMNo").Value & ""
    DTPicker3.Value = TB.Fields("IGMDate").Value
    txtIndex.Text = TB.Fields("IndexNo").Value & ""
    txtBondCashNo.Text = TB.Fields("BondNo").Value & ""
    DTPicker4.Value = TB.Fields("BondDate").Value
    txtImportValue.Text = TB.Fields("ImportValue").Value & ""
    txtWharfage.Text = TB.Fields("Wharfage").Value & ""
    DTPicker5.Value = TB.Fields("WharfageDate").Value
    txtDMI.Text = TB.Fields("DMI").Value & ""
    DTPicker6.Value = TB.Fields("DMIDate").Value
    txtDemurrage.Text = TB.Fields("Demurrage").Value & ""
    DTPicker7.Value = TB.Fields("DemurrageDate").Value
    txtBargememo.Text = TB.Fields("BargMemo").Value & ""
    DTPicker8.Value = TB.Fields("BargDate").Value
    txtCAARNo.Text = TB.Fields("CAANo").Value & ""
    DTPicker9.Value = TB.Fields("CAADate")
    txtCHCRNo.Text = TB.Fields("CHCNo").Value & ""
    DTPicker10.Value = TB.Fields("CHCDate").Value
    txtBondBillNR.Text = TB.Fields("BondBillno").Value & ""
    DTPicker11.Value = TB.Fields("BondBillDate").Value & ""
    txtGoodsDelTo.Text = TB.Fields("GoodsDel").Value & ""
    DTPicker12.Value = TB.Fields("GoodsDelDate").Value
    txtRefNo.Text = TB.Fields("RefNo").Value & ""
    txtFileNo.Text = TB.Fields("FileNo").Value & ""
    txtBETriplicate.Text = TB.Fields("BeTriplicate").Value & ""
    txtInvoicePackingList.Text = TB.Fields("InvoicePackingList").Value & ""
    txtBL.Text = TB.Fields("BL").Value & ""
    txtwdbmemo.Text = TB.Fields("WDBMemo").Value & ""
    txtShippingBill.Text = TB.Fields("Shipbill").Value & ""
    txtDocumentRecDate.Text = TB.Fields("DocumentsRecOn") & ""
    txtExamineDate.Text = TB.Fields("ConsignmentExam").Value & ""
    txtPCT.Text = TB.Fields("PCT").Value & ""
End If
TB.Close
DB.Close
End Sub

Private Sub Save()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\Bloom.mdb")
If EditMode = True Then
    Ssql = "Delete from Docs where SrNo=" & Val(txtSerialNo.Text)
    DB.Execute Ssql
End If
Set TB = DB.OpenRecordset("Docs", dbOpenTable)
TB.AddNew
    TB.Fields("Status").Value = Check1.Value
    TB.Fields("SrNo").Value = Val(txtSerialNo.Text)
    TB.Fields("LCNo").Value = txtLCNo.Text
    TB.Fields("LCDate").Value = DTPicker1.Value
    TB.Fields("KarachiDate").Value = DTPicker2.Value
    TB.Fields("Packages").Value = txtPackages.Text
    TB.Fields("Goods").Value = txtGoods.Text
    TB.Fields("PerSS").Value = txtClear.Text
    TB.Fields("PartyCode").Value = Val(txtPartyCode.Text)
    TB.Fields("ShipFrom").Value = txtFrom.Text
    TB.Fields("Shipto").Value = txtTo.Text
    TB.Fields("IGMNo").Value = txtIGMNo.Text
    TB.Fields("IGMDate").Value = DTPicker3.Value
    TB.Fields("IndexNo").Value = txtIndex.Text
    TB.Fields("BondNo").Value = txtBondCashNo.Text
    TB.Fields("BondDate").Value = DTPicker4.Value
    TB.Fields("ImportValue").Value = Val(txtImportValue.Text)
    TB.Fields("Wharfage").Value = txtWharfage.Text
    TB.Fields("WharfageDate").Value = DTPicker5.Value
    TB.Fields("DMI").Value = txtDMI.Text
    TB.Fields("DMIDate").Value = DTPicker6.Value
    TB.Fields("Demurrage").Value = txtDemurrage.Text
    TB.Fields("DemurrageDate").Value = DTPicker7.Value
    TB.Fields("BargMemo").Value = txtBargememo.Text
    TB.Fields("BargDate").Value = DTPicker8.Value
    TB.Fields("CAANo").Value = txtCAARNo.Text
    TB.Fields("CAADate") = DTPicker9.Value
    TB.Fields("CHCNo").Value = txtCHCRNo.Text
    TB.Fields("CHCDate").Value = DTPicker10.Value
    TB.Fields("BondBillno").Value = txtBondBillNR.Text
    TB.Fields("BondBillDate").Value = DTPicker11.Value
    TB.Fields("GoodsDel").Value = txtGoodsDelTo.Text
    TB.Fields("GoodsDelDate").Value = DTPicker12.Value
    TB.Fields("RefNo").Value = txtRefNo.Text
    TB.Fields("FileNo").Value = txtFileNo.Text
    TB.Fields("BeTriplicate").Value = txtBETriplicate.Text
    TB.Fields("InvoicePackingList").Value = txtInvoicePackingList.Text
    TB.Fields("BL").Value = txtBL.Text
    TB.Fields("WDBMemo").Value = txtwdbmemo.Text
    TB.Fields("Shipbill").Value = txtShippingBill.Text
    TB.Fields("DocumentsRecOn") = txtDocumentRecDate.Text
    TB.Fields("ConsignmentExam").Value = txtExamineDate.Text
    TB.Fields("PCT").Value = txtPCT.Text
TB.Update
TB.Close
DB.Close
End Sub
Private Sub Max1()
Dim Ssql As String
Dim DB As Database
Dim TB As Recordset

Ssql = "Select Max(SrNo) as S from Docs"
Set DB = OpenDatabase(App.Path & "\Bloom.mdb")
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("S").Value) Then
    txtSerialNo.Text = TB.Fields("S").Value + 1
Else
    txtSerialNo.Text = 1
End If
TB.Close
DB.Close
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Save
Command2_Click
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
EditMode = False
FullClear
Max1
txtLCNo.SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
SendKeys ("{TAB}")
End If
End Sub
Private Sub initializeDates()
Dim c As Control
For Each c In Me.Controls
    If TypeOf c Is DTPicker Then c.Value = Date
Next
End Sub
Private Sub FullClear()
Dim c As Control
For Each c In Me.Controls
    If TypeOf c Is TextBox Then c.Text = ""
Next
initializeDates
End Sub
Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
Max1
initializeDates
End Sub

Private Sub Timer1_Timer()
If Val(txtSerialNo.Text) > 0 And Len(txtLCNo.Text) > 0 And Len(txtPartyCode.Text) > 0 And Len(txtPartyName.Text) > 0 Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub

Private Sub txtPartyCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Load Search2
    Search2.Text3.Text = 1
    Search2.Show vbModal
End If
End Sub

Private Sub txtSerialNo_Validate(Cancel As Boolean)
If Val(txtSerialNo.Text) > 0 Then
    Edit1
End If
End Sub
