VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPHMA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALPHA kNITTING (pvt) Ltd..(Packing List Form)"
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
   Icon            =   "frmPHMA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Visa Application Form :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5655
      Left            =   240
      TabIndex        =   38
      Top             =   480
      Width           =   8775
      Begin VB.TextBox txtExp_Name 
         DataField       =   "Exp_Name"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Tag             =   "2"
         Top             =   840
         Width           =   2085
      End
      Begin VB.TextBox txtBy 
         DataField       =   "By"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4605
         TabIndex        =   3
         Tag             =   "3"
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox txtShippedTo 
         DataField       =   "ShippedTo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6885
         TabIndex        =   4
         Tag             =   "4"
         Top             =   840
         Width           =   1740
      End
      Begin VB.TextBox txtDebitTo 
         DataField       =   "DebitTo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox txtVisNo 
         DataField       =   "VisNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4605
         TabIndex        =   6
         Tag             =   "6"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox txtExp_Id 
         DataField       =   "Exp_Id"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6885
         TabIndex        =   7
         Tag             =   "7"
         Top             =   1200
         Width           =   1740
      End
      Begin VB.TextBox txtAss_Name 
         DataField       =   "Ass_Name"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   8
         Tag             =   "8"
         Top             =   1560
         Width           =   2085
      End
      Begin VB.TextBox txtCategory 
         DataField       =   "Category"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4605
         TabIndex        =   9
         Tag             =   "9"
         Top             =   1560
         Width           =   1395
      End
      Begin VB.TextBox txtQty_Units 
         DataField       =   "Qty_Units"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6885
         TabIndex        =   10
         Tag             =   "10"
         Top             =   1560
         Width           =   1740
      End
      Begin VB.TextBox txtE_No 
         DataField       =   "E_No"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   11
         Tag             =   "11"
         Top             =   1920
         Width           =   2085
      End
      Begin VB.TextBox txtE_Date 
         DataField       =   "E_Date"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4605
         TabIndex        =   12
         Tag             =   "12"
         Top             =   1920
         Width           =   1395
      End
      Begin VB.TextBox txtValue 
         DataField       =   "Value"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6885
         TabIndex        =   13
         Tag             =   "13"
         Top             =   1920
         Width           =   1740
      End
      Begin VB.TextBox txtCF 
         DataField       =   "C&FDesc"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         TabIndex        =   14
         Tag             =   "14"
         Top             =   2400
         Width           =   3225
      End
      Begin VB.TextBox txtOther 
         DataField       =   "Other"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5445
         TabIndex        =   15
         Tag             =   "15"
         Top             =   2400
         Width           =   3180
      End
      Begin VB.TextBox txtExch_Rate 
         DataField       =   "Exch_Rate"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Tag             =   "16"
         Top             =   2760
         Width           =   1500
      End
      Begin VB.TextBox txtValue_Pak 
         DataField       =   "Value_Pak"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4365
         TabIndex        =   17
         Tag             =   "17"
         Top             =   2760
         Width           =   1740
      End
      Begin VB.TextBox txtDollar_Parity 
         DataField       =   "Dollar_Parity"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7485
         TabIndex        =   18
         Tag             =   "18"
         Top             =   2760
         Width           =   1140
      End
      Begin VB.TextBox txtCF1 
         DataField       =   "C&F_Value"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         TabIndex        =   19
         Tag             =   "19"
         Top             =   3120
         Width           =   1500
      End
      Begin VB.TextBox txtFreigat 
         DataField       =   "Freigat"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   20
         Tag             =   "20"
         Top             =   3120
         Width           =   1740
      End
      Begin VB.TextBox txtInsurance 
         DataField       =   "Insurance"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7485
         TabIndex        =   21
         Tag             =   "21"
         Top             =   3120
         Width           =   1140
      End
      Begin VB.TextBox txtFOB 
         DataField       =   "FOB"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   22
         Tag             =   "22"
         Top             =   3480
         Width           =   1500
      End
      Begin VB.TextBox txtUndertaking 
         DataField       =   "Undertaking"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4365
         TabIndex        =   23
         Tag             =   "23"
         Top             =   3480
         Width           =   1740
      End
      Begin VB.TextBox txtExplicNo 
         DataField       =   "ExplicNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7485
         TabIndex        =   24
         Tag             =   "24"
         Top             =   3480
         Width           =   1140
      End
      Begin VB.TextBox txtCoNo 
         DataField       =   "CoNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   25
         Tag             =   "25"
         Top             =   3840
         Width           =   1500
      End
      Begin VB.TextBox txtVisaNo 
         DataField       =   "VisaNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4365
         TabIndex        =   26
         Tag             =   "26"
         Top             =   3840
         Width           =   1740
      End
      Begin VB.TextBox txtInvNo 
         DataField       =   "InvNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   0
         Tag             =   "0"
         Top             =   360
         Width           =   2760
      End
      Begin VB.TextBox txtInvdate 
         DataField       =   "Invdate"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5445
         TabIndex        =   1
         Tag             =   "1"
         Top             =   360
         Width           =   3120
      End
      Begin VB.TextBox txtAWBNo 
         DataField       =   "AWBNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7485
         TabIndex        =   27
         Tag             =   "27"
         Top             =   3840
         Width           =   1155
      End
      Begin VB.TextBox txtAWBDate 
         DataField       =   "Date"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   28
         Tag             =   "28"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox txtAircraft 
         DataField       =   "Aircraft"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4365
         TabIndex        =   29
         Tag             =   "29"
         Top             =   4200
         Width           =   1740
      End
      Begin VB.TextBox txtFrigatNo 
         DataField       =   "FrigatNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7485
         TabIndex        =   30
         Tag             =   "30"
         Top             =   4200
         Width           =   1140
      End
      Begin VB.TextBox txtshipperBillNo 
         DataField       =   "shipperBillNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   31
         Tag             =   "31"
         Top             =   4560
         Width           =   1500
      End
      Begin VB.TextBox txtBilldate 
         DataField       =   "Billdate"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4365
         TabIndex        =   32
         Tag             =   "32"
         Top             =   4560
         Width           =   1740
      End
      Begin VB.TextBox txtMateNo 
         DataField       =   "MateNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   33
         Tag             =   "33"
         Top             =   4920
         Width           =   3225
      End
      Begin VB.TextBox txtMatedate 
         DataField       =   "Matedate"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6045
         TabIndex        =   34
         Tag             =   "34"
         Top             =   4920
         Width           =   2580
      End
      Begin VB.TextBox txtSailingdate 
         DataField       =   "Sailingdate"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   35
         Tag             =   "35"
         Top             =   5280
         Width           =   3225
      End
      Begin VB.TextBox txtCatNo 
         DataField       =   "CatNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6045
         TabIndex        =   36
         Tag             =   "36"
         Top             =   5280
         Width           =   2580
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exp_Name"
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
         Left            =   450
         TabIndex        =   75
         ToolTipText     =   "Exporter Name"
         Top             =   840
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "By"
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
         Left            =   4050
         TabIndex        =   74
         Top             =   840
         Width           =   195
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ship To"
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
         Left            =   6240
         TabIndex        =   73
         ToolTipText     =   "Shipment To"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DebitTo"
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
         Left            =   690
         TabIndex        =   72
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Visa No"
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
         Left            =   3765
         TabIndex        =   71
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exp Id"
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
         Left            =   6255
         TabIndex        =   70
         ToolTipText     =   "Exporter Id"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ass Name"
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
         Left            =   495
         TabIndex        =   69
         ToolTipText     =   "Association Name"
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Category"
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
         Left            =   3690
         TabIndex        =   68
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qty Units"
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
         Left            =   6135
         TabIndex        =   67
         ToolTipText     =   "Quantity In Units"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Form E No "
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
         Left            =   330
         TabIndex        =   66
         ToolTipText     =   "Form 'E' No."
         Top             =   1920
         Width           =   870
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
         Index           =   11
         Left            =   4110
         TabIndex        =   65
         ToolTipText     =   "Form 'E' Date"
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value:"
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
         Left            =   6285
         TabIndex        =   64
         ToolTipText     =   "C&F / CIF Value"
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&&F / CIF "
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
         Left            =   540
         TabIndex        =   63
         ToolTipText     =   "C&F / CIF Description"
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Other"
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
         Left            =   4890
         TabIndex        =   62
         ToolTipText     =   "Other Than US $"
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exch Rate"
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
         Left            =   615
         TabIndex        =   61
         ToolTipText     =   "Exchange Rate Against Pak Rupees. "
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value Pak"
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
         Left            =   3495
         TabIndex        =   60
         ToolTipText     =   "Value In Pak Rupee"
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dollar Parity"
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
         Left            =   6165
         TabIndex        =   59
         ToolTipText     =   "Dollar Rupee Parity"
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&&F / CIF "
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
         Left            =   570
         TabIndex        =   58
         ToolTipText     =   "C&F and CIF Value In US $"
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Freigat:"
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
         Left            =   3405
         TabIndex        =   57
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Insurance"
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
         Left            =   6165
         TabIndex        =   56
         ToolTipText     =   "Insurance In US $"
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FOB Value"
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
         Index           =   21
         Left            =   420
         TabIndex        =   55
         ToolTipText     =   "FOB Value In US $"
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Undertaking"
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
         Index           =   22
         Left            =   3210
         TabIndex        =   54
         Top             =   3480
         Width           =   1005
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exp license No"
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
         Index           =   23
         Left            =   6165
         TabIndex        =   53
         ToolTipText     =   "Export License Form No"
         Top             =   3480
         Width           =   1185
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Co No"
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
         Index           =   24
         Left            =   885
         TabIndex        =   52
         ToolTipText     =   "Certificate Of Origin Form No."
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Visa No"
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
         Index           =   25
         Left            =   3525
         TabIndex        =   51
         ToolTipText     =   "Visa Form No"
         Top             =   3840
         Width           =   615
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
         Index           =   26
         Left            =   480
         TabIndex        =   50
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoice Date"
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
         Index           =   27
         Left            =   4365
         TabIndex        =   49
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "AWB No"
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
         Index           =   28
         Left            =   6645
         TabIndex        =   48
         Top             =   3840
         Width           =   630
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "AWB Date"
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
         Index           =   29
         Left            =   465
         TabIndex        =   47
         Top             =   4200
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Air craft"
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
         Index           =   30
         Left            =   3405
         TabIndex        =   46
         ToolTipText     =   "Name Of Vessel Or AIRCRAFT"
         Top             =   4200
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flight No"
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
         Index           =   31
         Left            =   6540
         TabIndex        =   45
         Top             =   4200
         Width           =   705
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Shipping Bill No"
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
         Index           =   32
         Left            =   120
         TabIndex        =   44
         Top             =   4560
         Width           =   1260
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
         Index           =   33
         Left            =   3525
         TabIndex        =   43
         Top             =   4560
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mate No"
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
         Index           =   34
         Left            =   405
         TabIndex        =   42
         ToolTipText     =   "Mate Receipt No"
         Top             =   4920
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mate Date"
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
         Index           =   35
         Left            =   4965
         TabIndex        =   41
         Top             =   4920
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sailing Date"
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
         Index           =   36
         Left            =   285
         TabIndex        =   40
         Top             =   5280
         Width           =   945
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Category No"
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
         Index           =   37
         Left            =   4845
         TabIndex        =   39
         ToolTipText     =   "Attached With Category"
         Top             =   5280
         Width           =   1005
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
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
            Object.ToolTipText     =   "Next Record"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "previous"
            Object.ToolTipText     =   "Previous Record"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "last"
            Object.ToolTipText     =   "Last Record"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   "Find Record"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print Reocrd"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   240
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
            Picture         =   "frmPHMA.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":0AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":0BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":0D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":0EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":13F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":193A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":1A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":1C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPHMA.frx":1D7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuForms 
      Caption         =   "&Forms"
      Begin VB.Menu main 
         Caption         =   "&Main Form"
      End
      Begin VB.Menu Invoice 
         Caption         =   "&Invoice Form"
      End
      Begin VB.Menu export 
         Caption         =   "&Export License"
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
      Begin VB.Menu bank 
         Caption         =   "&Bank Form "
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
Attribute VB_Name = "frmPHMA"
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
    
    sSqlMas = "Select * from PHMATbl"
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
      rsQry = False
End Sub

Private Sub bank_Click()
    frmBank.Show
    Unload Me
    
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
    
    
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(6).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
        
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
Toolbar1.Buttons(9).Enabled = False
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(3).Enabled = True
End Sub

Sub FillDate()
    Me.txtBillDate = Date
    Me.txtAwbDate = Date
    Me.txtE_Date = Date
    Me.txtInvDate = Date
    Me.txtSailingdate = Date
    
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

Private Sub performa_Click()
    frmPerforma.Show
    Unload Me
    
End Sub

Private Sub phma_Click()
    frmPHMA.Show
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
    If mControl.Tag >= 0 And mControl.Tag <= 36 Then
             rsMas.Fields(Val(mControl.Tag)) = mControl
        End If
    Next
    rsMas.Update
End Sub

'*********************************************************
Sub ClearMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 36 Then
             mControl = ""
        End If
    Next
End Sub


'*******************************************************
'************************************************************

Sub EnableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 36 Then
             mControl.Enabled = True
        End If
    Next
End Sub
'************************************************************

Sub DisableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 36 Then
             mControl.Enabled = False
        End If
    Next
End Sub

'***************************************************

Public Sub FillMas()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 36 Then
              If rsMas.Fields(Val(mControl.Tag)) <> "" Then
                    mControl = rsMas.Fields(Val(mControl.Tag))
               Else
                    mControl = " "
                End If
        End If
    Next
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
                    Me.query.Enabled = False
                    Me.Toolbar1.Buttons(2).Enabled = True
                End If
            Case "print"
                frmPrint.Show
    End Select

End Sub

'**************************************************
'**************************************************

Private Sub txtAircraft_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFrigatNo.SetFocus
    End If
End Sub

Private Sub txtAss_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCategory.SetFocus
    End If
End Sub

Private Sub txtAwbDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAircraft.SetFocus
    End If
End Sub

Private Sub txtAwbNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAwbDate.SetFocus
    End If
        
End Sub

Private Sub txtBilldate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMateNo.SetFocus
    End If
End Sub

Private Sub txtBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtShippedTo.SetFocus
    End If
End Sub


Private Sub txtCategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtQty_Units.SetFocus
    End If
End Sub


Private Sub txtCF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtOther.SetFocus
    End If
End Sub

Private Sub txtCF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFreigat.SetFocus
    End If
End Sub

Private Sub txtCoNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtVisaNo.SetFocus
    End If
End Sub

Private Sub txtDebitTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtVisNo.SetFocus
    End If
End Sub

Private Sub txtDollar_Parity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCF1.SetFocus
    End If
End Sub

Private Sub txtE_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtValue.SetFocus
    End If
        
End Sub


Private Sub txtE_No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtE_Date.SetFocus
    End If

End Sub

Private Sub txtExch_Rate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtValue_Pak.SetFocus
    End If
End Sub

Private Sub txtExp_Id_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtAss_Name.SetFocus
    End If
    End Sub

Private Sub txtExp_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtBy.SetFocus
    End If
End Sub

Private Sub txtExplicNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCoNo.SetFocus
    
    End If
End Sub

Private Sub txtFOB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtUndertaking.SetFocus
    End If
End Sub

Private Sub txtFreigat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtInsurance.SetFocus
    End If
End Sub

Private Sub txtFrigatNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtshipperBillNo.SetFocus
    End If
End Sub

Private Sub txtInsurance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFOB.SetFocus
    End If
End Sub

Private Sub txtInvDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExp_Name.SetFocus
    End If
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtInvDate.SetFocus
    End If
End Sub



Private Sub txtMatedate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtSailingdate.SetFocus
    End If
End Sub

Private Sub txtMateNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Me.txtMatedate.SetFocus
    End If
End Sub

Private Sub txtOther_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExch_Rate.SetFocus
    End If
    
End Sub

Private Sub txtQty_Units_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtE_No.SetFocus
    End If
End Sub

Private Sub txtSailingdate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCatNo.SetFocus
    End If
End Sub

Private Sub txtShippedto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDebitTo.SetFocus
    End If
    
End Sub


Private Sub txtshipperBillNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtBillDate.SetFocus
    End If
End Sub

Private Sub txtUndertaking_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExplicNo.SetFocus
    End If
        
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCF.SetFocus
    End If
End Sub

Private Sub txtValue_Pak_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDollar_Parity.SetFocus
    End If
End Sub

Private Sub txtVisaNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAwbNo.SetFocus
    End If
End Sub

Private Sub txtVisNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExp_Id.SetFocus
    End If
End Sub
