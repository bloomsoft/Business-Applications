VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALPHA kNITTING (pvt) Ltd..(Bank Form)"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Bank Form :"
      ForeColor       =   &H00FF0000&
      Height          =   4815
      Left            =   720
      TabIndex        =   24
      Top             =   480
      Width           =   8175
      Begin VB.TextBox txtEdate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   9
         Tag             =   "9"
         Top             =   1920
         Width           =   2880
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Tag             =   "8"
         Top             =   1920
         Width           =   2715
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Tag             =   "3"
         Top             =   840
         Width           =   2880
      End
      Begin VB.TextBox txtsrDate 
         DataField       =   "Sr_No"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1200
         Width           =   2880
      End
      Begin VB.TextBox txtSrNo 
         DataField       =   "Date"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Tag             =   "4"
         Top             =   1200
         Width           =   2715
      End
      Begin VB.TextBox txtInvVal 
         DataField       =   "BillNO"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Tag             =   "2"
         Top             =   840
         Width           =   2715
      End
      Begin VB.TextBox txtBillNo 
         DataField       =   "BillDate"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Tag             =   "6"
         Top             =   1560
         Width           =   2715
      End
      Begin VB.TextBox txtInvNo 
         DataField       =   "InvNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Tag             =   "0"
         Top             =   360
         Width           =   2715
      End
      Begin VB.TextBox txtInv_Date 
         DataField       =   "Inv_Date"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   1
         Tag             =   "1"
         Top             =   360
         Width           =   2880
      End
      Begin VB.TextBox txtBillDate 
         DataField       =   "Inv_Value"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Tag             =   "7"
         Top             =   1560
         Width           =   2880
      End
      Begin VB.TextBox txtShip_Date 
         DataField       =   "Ship_Date"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Tag             =   "10"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtTerms 
         DataField       =   "Terms"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3615
         TabIndex        =   11
         Tag             =   "11"
         Top             =   2400
         Width           =   4205
      End
      Begin VB.TextBox txtShip_BillDate 
         DataField       =   "Ship_Bill_No"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   13
         Tag             =   "13"
         Top             =   2760
         Width           =   1620
      End
      Begin VB.TextBox txtShip_Bill_No 
         DataField       =   "Ship_Bill_Date"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Tag             =   "12"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtBill_Amount 
         DataField       =   "Bill_Amount"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   14
         Tag             =   "14"
         Top             =   2760
         Width           =   1554
      End
      Begin VB.TextBox txtCommission 
         DataField       =   "Commission"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Tag             =   "15"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtB_Charges 
         DataField       =   "B_Charges"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Tag             =   "16"
         Top             =   3240
         Width           =   1260
      End
      Begin VB.TextBox txtOther_Charges 
         DataField       =   "Other_Charges"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   17
         Tag             =   "17"
         Top             =   3240
         Width           =   1554
      End
      Begin VB.TextBox txtExch_Rate 
         DataField       =   "Exch_Rate"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   18
         Tag             =   "18"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtNet_Amount 
         DataField       =   "Net_Amount"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0000;(0.0000)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3960
         TabIndex        =   19
         Tag             =   "19"
         Top             =   3600
         Width           =   1260
      End
      Begin VB.TextBox txtReal_Date 
         DataField       =   "Real_Date"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   20
         Tag             =   "20"
         Top             =   3600
         Width           =   1560
      End
      Begin VB.TextBox txtMonth 
         DataField       =   "Month"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   22
         Tag             =   "22"
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtCity 
         DataField       =   "City"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Tag             =   "21"
         Top             =   4080
         Width           =   2715
      End
      Begin VB.Label Label3 
         Caption         =   "Form E No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "E Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   47
         ToolTipText     =   "Form E Date"
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Form E No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -1800
         TabIndex        =   46
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4320
         TabIndex        =   45
         ToolTipText     =   "M/S "
         Top             =   840
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4320
         TabIndex        =   44
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Serial No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   480
         TabIndex        =   43
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoice Val"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   705
         TabIndex        =   41
         ToolTipText     =   "Bill Date"
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoice No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   375
         TabIndex        =   40
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inv Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   4185
         TabIndex        =   39
         ToolTipText     =   "Invoice Date"
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   4200
         TabIndex        =   38
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ship Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   480
         TabIndex        =   37
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Terms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   2925
         TabIndex        =   36
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   2880
         TabIndex        =   35
         ToolTipText     =   "Shipment Bill Date"
         Top             =   2760
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ship Bill No "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   360
         TabIndex        =   34
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   5400
         TabIndex        =   33
         ToolTipText     =   "Bill Amount"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comision"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   435
         TabIndex        =   32
         Top             =   3240
         Width           =   795
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "B_Charges"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   2880
         TabIndex        =   31
         ToolTipText     =   "Bank Charges"
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   5520
         TabIndex        =   30
         ToolTipText     =   "Other Chargers"
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exch_Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   285
         TabIndex        =   29
         ToolTipText     =   "Exchange Rate"
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Net_Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   2880
         TabIndex        =   28
         ToolTipText     =   "Net Amount Equivalent in Rupees"
         Top             =   3600
         Width           =   1020
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Real_Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   5325
         TabIndex        =   27
         ToolTipText     =   "Date Of Realization"
         Top             =   3600
         Width           =   795
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   4245
         TabIndex        =   26
         Top             =   4080
         Width           =   525
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   720
         TabIndex        =   25
         Top             =   4080
         Width           =   315
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":0AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":0BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":0D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":0EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":13F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":193A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":1A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":1C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":1D7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "new"
            Object.ToolTipText     =   "Add New Record"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "save"
            Object.ToolTipText     =   "Save Record"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancel"
            Object.ToolTipText     =   "Cancel Record"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "first"
            Object.ToolTipText     =   "First Record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "next"
            Object.ToolTipText     =   "Move Next"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "previous"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "last"
            Object.ToolTipText     =   "Move Last"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   "Find Record"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print Record"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuForms 
      Caption         =   "&Forms"
      Begin VB.Menu Invoice 
         Caption         =   "&Invoice Form"
      End
      Begin VB.Menu export 
         Caption         =   "&Export License"
      End
      Begin VB.Menu paking 
         Caption         =   "&Paking List Form"
      End
      Begin VB.Menu Combined 
         Caption         =   "&GSP Form"
      End
      Begin VB.Menu Certificate 
         Caption         =   "&Cerificate Of Origin (Textile)"
      End
      Begin VB.Menu performa 
         Caption         =   "&Performa Invoice Form"
      End
      Begin VB.Menu phma 
         Caption         =   "&PHMA Form"
      End
      Begin VB.Menu quota 
         Caption         =   "&Quota Transfer Form"
      End
      Begin VB.Menu release 
         Caption         =   "&Release And Undertaking"
      End
      Begin VB.Menu main 
         Caption         =   "&Main Form"
      End
   End
   Begin VB.Menu Record_menu 
      Caption         =   "&Record"
      Begin VB.Menu First 
         Caption         =   "&First"
         Shortcut        =   ^F
      End
      Begin VB.Menu previous 
         Caption         =   "&Previous"
         Shortcut        =   ^P
      End
      Begin VB.Menu next 
         Caption         =   "&Next"
         Shortcut        =   ^T
      End
      Begin VB.Menu last 
         Caption         =   "&Last"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu cancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu query 
         Caption         =   "&Query"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
      Begin VB.Menu yes 
         Caption         =   "&Yes"
         Shortcut        =   ^Y
      End
      Begin VB.Menu no 
         Caption         =   "&No"
      End
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Object Variables to Create Instance of ADO

Dim rsMas As ADODB.Recordset
Dim rsDet As ADODB.Recordset
Dim Cn As ADODB.Connection
Dim Cmd As ADODB.Command
Dim sSqlMas, sSqlDet, sTemp, sTemp1 As String
Dim bNew, bNewMas, rsDetail, rsQry As Boolean

'**************************************************************
'Subroutine to open Master Connection
Sub OpenMasConnection()

    
    
    Set Cn = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    Set rsMas = New ADODB.Recordset
    With Cn
        .Provider = "MICROSOFT.JET.OLEDB.3.51"
        .ConnectionString = App.Path & "\Export.mdb"
        .Open
    End With
          
End Sub

'****************************************************************
'Subroutine to open the Detail Recordset

Sub OpenMasRecordSet()
    
    sSqlMas = "Select * from BankTbl"
    rsMas.Open sSqlMas, Cn, adOpenKeyset, adLockPessimistic
End Sub

'**************************************************************
' Subroutine to find a record on the given criteria

Sub FindRecord()
    
       sTemp = Me.txtInvNo.Text
       rsMas.Find "InvNo ='" & sTemp & "'"
       If rsMas.EOF Then
             MsgBox "Record does not exist", vbInformation
             rsMas.MoveFirst
             ClearMas
             
             Me.save.Enabled = False
             Me.txtInvNo.SetFocus
         
        Else
            FillMas
            
       End If
      
End Sub

'************************************************************
Private Sub cancel_Click()
    If bNew = True Then
        rsMas.CancelUpdate
        Me.new.Enabled = True
        
     ElseIf bNewMas = True Then
         rsMas.CancelUpdate
         
         Me.new.Enabled = True
    End If
      Me.txtInvNo.SetFocus
      Me.cancel.Enabled = False
      Me.query.Enabled = True
      Me.mnufind.Enabled = True
      
      
              'Menu options disabled and enabled
              Me.First.Enabled = True
              Me.next.Enabled = True
              Me.previous.Enabled = True
              Me.last.Enabled = True
              Me.mnufind.Enabled = False
      
      Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(6).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    
    Toolbar1.Buttons(8).Enabled = False
    ClearMas
    EnableMas
End Sub

'*******************************************************
Sub LockRecord()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 22 Then
             mControl.Locked = True
        End If
    Next
End Sub

'*******************************************************
Sub UnLockRecord()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 22 Then
             mControl.Locked = False
        End If
    Next
End Sub

Private Sub Certificate_Click()
    
    frmCertificate.Show
    Unload Me
End Sub
 
Private Sub Combined_Click()
    frmGenSystem.Show
    Unload Me
End Sub

Private Sub delete_Click()
Dim Respose As String
    
      Response = MsgBox("Do you want to delete this record", vbYesNo)
     
     If Response = vbYes Then
        rsMas.delete adAffectCurrent
        ClearMas
        If rsMas.EOF <> True Then
            rsMas.MoveNext
        End If
        
     End If
    Me.delete.Enabled = False
End Sub
Private Sub export_Click()
    frmExpLicense.Show
    Unload Me
End Sub

Private Sub find_Click()
    If Me.txtInvNo = "" Then
        MsgBox "Enter Invoice No. to find Record", vbInformation
     Else
        FindRecord
     End If
    
End Sub

Private Sub First_Click()
          
            rsMas.MoveFirst
            EnableMas
            FillMas
              
              'Toolbar options enabled and disabled
              Toolbar1.Buttons(4).Enabled = False
              Toolbar1.Buttons(5).Enabled = True
              Toolbar1.Buttons(6).Enabled = False
              Toolbar1.Buttons(7).Enabled = True
              
              'Menu options disabled and enabled
              Me.First.Enabled = False
              Me.next.Enabled = True
              Me.previous.Enabled = False
              Me.last.Enabled = True
              Me.save.Enabled = False
              Me.cancel.Enabled = True
      
End Sub

'*********************************************************
Private Sub Form_Load()
                
'* Call subroutine to open the Master and Detail Connection
    
    OpenMasConnection
    OpenMasRecordSet
    rsDetail = True
    DisableMas
       
    Me.new.Enabled = True
    Me.save.Enabled = False
    Me.mnufind.Enabled = False
    Me.cancel.Enabled = False
    
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(6).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(8).Enabled = False
        
End Sub

Private Sub Invoice_Click()
    frmInvoice.Show
    Unload Me
    
End Sub

Private Sub last_Click()
            rsMas.MoveLast
            FillMas
            'To enable or disable Toolbar options
              Toolbar1.Buttons(4).Enabled = True
              Toolbar1.Buttons(5).Enabled = False
              Toolbar1.Buttons(6).Enabled = True
              Toolbar1.Buttons(7).Enabled = False
              
              'Menu options disabled and enabled
              Me.First.Enabled = True
              Me.next.Enabled = False
              Me.previous.Enabled = True
              Me.last.Enabled = False
              Me.save.Enabled = False
              Me.cancel.Enabled = True
End Sub

Private Sub main_Click()
        frmMain.Show
        Unload Me
        
End Sub

Private Sub mnufind_Click()
    
    If Me.txtInvNo.Text = "" Then
        MsgBox "Enter Invoice No. to find Record", vbInformation
        
    Else
        FindRecord            'Subroutine to find a record
    End If
        
    
        Me.query.Enabled = True
        Me.new.Enabled = True
        Me.save.Enabled = True
        Me.mnufind.Enabled = False
        Me.cancel.Enabled = False
    
    
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(4).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = True
    
    rsQry = False
End Sub

'***************************************************************
'****** Open the Recordset*********************
Private Sub new_Click()
      
    rsMas.AddNew
    bNewMas = True
    
    ' To enable and clear all the controls
    
    EnableMas     'Enable Master Fields
    ClearMas      'Clear Master Fields
    
    Me.new.Enabled = False
    Me.save.Enabled = True
           
    Me.query.Enabled = False
    Me.mnufind.Enabled = False
    
    Me.txtInvNo.SetFocus
    Me.cancel.Enabled = True
   
    FillDate     ' Fill Date in Master Fields
    
    
              'Menu options disabled and enabled
              Me.First.Enabled = False
              Me.next.Enabled = False
              Me.previous.Enabled = False
              Me.last.Enabled = False

    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(3).Enabled = True
     Toolbar1.Buttons(9).Enabled = False
End Sub

Sub FillDate()
    
    
    Me.txtInv_Date.Text = Date
    Me.txtBillDate = Date
    Me.txtEdate = Date
    Me.txtReal_Date = Date
    Me.txtShip_BillDate = Date
    Me.txtShip_Date = Date
    Me.txtsrDate = Date
    
End Sub


Private Sub Packing_Click()
    frmPacking.Show
    Unload Me
End Sub

Private Sub next_Click()
            On Error GoTo Errorhandler
            If rsMas.EOF Then
                MsgBox "You are at the last Record", vbInformation
                rsMas.MoveLast
                'To enable or disable Toolbar options
              Toolbar1.Buttons(4).Enabled = True
              Toolbar1.Buttons(5).Enabled = False
              Toolbar1.Buttons(6).Enabled = True
              Toolbar1.Buttons(7).Enabled = False
              
              'Menu options disabled and enabled
              Me.First.Enabled = True
              Me.next.Enabled = False
              Me.previous.Enabled = True
              Me.last.Enabled = False
              Me.save.Enabled = False
              Me.cancel.Enabled = True
        Else
            rsMas.MoveNext
            FillMas
                Toolbar1.Buttons(4).Enabled = True
                Toolbar1.Buttons(5).Enabled = True
                Toolbar1.Buttons(6).Enabled = True
                Toolbar1.Buttons(7).Enabled = True
                'Menu options disabled and enabled
                Me.First.Enabled = True
                Me.next.Enabled = True
                Me.previous.Enabled = True
                Me.last.Enabled = True
        End If
Errorhandler:
0
End Sub

Private Sub paking_Click()
    frmPacking.Show
    Unload Me
    
End Sub

Private Sub performa_Click()
    frmPerforma.Show
    Unload Me
    
End Sub

Private Sub previous_Click()
On Error GoTo Errorhandler
              If rsMas.BOF Then
                    MsgBox "You are at the first record", vbInformation
                    rsMas.MoveFirst
                    'Toolbar options enabled and disabled
                    Toolbar1.Buttons(4).Enabled = False
                    Toolbar1.Buttons(5).Enabled = True
                    Toolbar1.Buttons(6).Enabled = False
                    Toolbar1.Buttons(7).Enabled = True
              
                'Menu options disabled and enabled
                    Me.First.Enabled = False
                    Me.next.Enabled = True
                    Me.previous.Enabled = False
                    Me.last.Enabled = True
                    Me.save.Enabled = False
                    
              Else
                    rsMas.MovePrevious
                    FillMas
                    Toolbar1.Buttons(4).Enabled = True
                    Toolbar1.Buttons(5).Enabled = True
                    Toolbar1.Buttons(6).Enabled = True
                    Toolbar1.Buttons(7).Enabled = True
                    
                    'Menu options disabled and enabled
                    Me.First.Enabled = True
                    Me.next.Enabled = True
                    Me.previous.Enabled = True
                    Me.last.Enabled = True
                    Me.cancel.Enabled = True
            End If
Errorhandler:
0
End Sub

Private Sub query_Click()
    
    
    'Call subroutine to clear and enable controls
    Me.query.Enabled = False
    Me.new.Enabled = False
    Me.save.Enabled = False
    Me.mnufind.Enabled = True
    Me.cancel.Enabled = True
    
    Me.next.Enabled = False
    Me.previous.Enabled = False
    Me.last.Enabled = False
    Me.First.Enabled = False
    
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
            
    EnableMas
    ClearMas
    
    Me.txtInvNo.SetFocus
       
    rsQry = True
    
End Sub

Private Sub quota_Click()
    frmQuota.Show
    Unload Me
    
End Sub

Private Sub release_Click()
    frmRelease.Show
    Unload Me
    
End Sub

'******************************************************

Private Sub save_Click()
    
  On Error GoTo Errorhandler
        SaveMasRecord
        DisableMas
        
        Me.new.Enabled = True
        
        Me.save.Enabled = False
        Me.cancel.Enabled = False
        Me.query.Enabled = True
        
        
        bNewMas = False
        
              'Menu options disabled and enabled
              Me.First.Enabled = True
              Me.next.Enabled = True
              Me.previous.Enabled = True
              Me.last.Enabled = True
        
        Toolbar1.Buttons(4).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = True

Errorhandler:
        
        Select Case Err.Number
        
            Case -2147217887
                MsgBox "Record Already exist", vbInformation
                 
             Case -2147352571
                MsgBox "There is Invalid Data In some Fields,Record Cann't be saved ", vbInformation
                               
                 
        End Select
    
 
    
End Sub

'****************************************************

Public Sub SaveMasRecord()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 22 Then
             rsMas.Fields(Val(mControl.Tag)) = mControl
        End If
    Next
    rsMas.Fields(23) = Val(Me.txtExch_Rate) * Val(Me.txtNet_Amount)
    rsMas.Update
End Sub

'*********************************************************
Sub ClearMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 22 Then
             mControl = ""
        End If
    Next
End Sub


'*******************************************************
'************************************************************

Sub EnableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 22 Then
             mControl.Enabled = True
        End If
    Next
End Sub
'************************************************************

Sub DisableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 22 Then
             mControl.Enabled = False
        End If
    Next
End Sub

'***************************************************

Public Sub FillMas()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 22 Then
              If rsMas.Fields(Val(mControl.Tag)) <> "" Then
                    mControl = rsMas.Fields(Val(mControl.Tag))
               Else
                    mControl = " "
                End If
        End If
    Next

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        Me.txtEdate.SetFocus
    End If
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "new"
            
            new_Click
'*************************************************************************
         Case "save"
            save_Click
            
'*************************************************************************          '**********************************************
          Case "cancel"
            
            cancel_Click
'*******************************************************************************
          Case "last"
            
            last_Click
 '************************************************************
           
           Case "next"
                next_Click      'Call function to move next record
                
'*************************************************************************
          Case "first"
                First_Click         'Call subroutine to move first record
                              
'*************************************************************************
          Case "previous"
               previous_Click      'Call Subroutine to move Previous
               
'**************************************************************
            Case "find"
                If rsQry = False Then
                    MsgBox "First Press F7 or Query option from Search Menu", vbInformation
                Else
                    FindRecord      'Call subroutine to find a record
                    rsQry = False
                    Me.Toolbar1.Buttons(2).Enabled = True
                    Me.query.Enabled = False
                    Me.save.Enabled = True
                    Me.delete.Enabled = True
                End If
            Case "print"
                   frmPrint.Show
    End Select

End Sub


Private Sub txtB_Charges_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtOther_Charges.SetFocus
    End If
End Sub

Private Sub txtBill_Amount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtCommission.SetFocus
End If
End Sub

Private Sub txtBilldate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.Text1.SetFocus
    End If
End Sub


Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtBillDate.SetFocus
    End If
End Sub


Private Sub txtCity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMonth.SetFocus
    End If
End Sub

Private Sub txtCommission_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtB_Charges.SetFocus
    End If
End Sub


Private Sub txtEdate_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtShip_Date.SetFocus
    End If
        
End Sub

Private Sub txtExch_Rate_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtNet_Amount.SetFocus
    End If
    
End Sub

Private Sub txtInv_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtInvVal.SetFocus
    End If
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtInv_Date.SetFocus
    End If
End Sub

Private Sub txtInvVal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtName.SetFocus
    End If
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.txtSrNo.SetFocus
    End If
End Sub




Private Sub txtNet_Amount_GotFocus()
    txtNet_Amount.Text = Val(Me.txtBill_Amount) * Val(Me.txtExch_Rate)
End Sub

Private Sub txtNet_Amount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtReal_Date.SetFocus
    End If
    
End Sub

Private Sub txtOther_Charges_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExch_Rate.SetFocus
    End If
    
End Sub

Private Sub txtReal_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCity.SetFocus
    End If
End Sub

Private Sub txtShip_Bill_No_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtShip_BillDate.SetFocus
    End If
End Sub



Private Sub txtShip_BillDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtBill_Amount.SetFocus
    End If
End Sub

Private Sub txtShip_Date_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtTerms.SetFocus
    End If
End Sub

Private Sub txtsrDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtBillNo.SetFocus

    End If

End Sub
Private Sub txtSrNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.txtsrDate.SetFocus
    End If
End Sub


Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtShip_Bill_No.SetFocus
    End If
End Sub

Private Sub yes_Click()
    frmBank.Hide
End Sub
