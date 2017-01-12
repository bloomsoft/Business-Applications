VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPerforma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALPHA kNITTING (pvt) Ltd..(Performa Invoice  Form)"
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
   Icon            =   "frmPerforma.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Custom Performa Invoice"
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
      Height          =   5775
      Left            =   480
      TabIndex        =   28
      Top             =   480
      Width           =   8775
      Begin VB.TextBox txtAsloNotify 
         DataField       =   "Notify"
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
         Left            =   1395
         TabIndex        =   20
         Tag             =   "20"
         Top             =   3600
         Width           =   7215
      End
      Begin VB.CommandButton cmdEditNext 
         Caption         =   "Next >>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   59
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditUpdate 
         Caption         =   "Update Changes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   58
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   56
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   55
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtMarks 
         DataField       =   "Marks"
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
         Left            =   1440
         TabIndex        =   21
         Tag             =   "21"
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox txtNumber 
         DataField       =   "Number"
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
         Left            =   4320
         TabIndex        =   22
         Tag             =   "22"
         Top             =   4200
         Width           =   1620
      End
      Begin VB.TextBox txtCar_Qty 
         DataField       =   "Car_Qty"
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
         Left            =   1440
         TabIndex        =   23
         Tag             =   "23"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox txtPcs_Qty 
         DataField       =   "Pcs_Qty"
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
         Left            =   4320
         TabIndex        =   24
         Tag             =   "24"
         Top             =   4560
         Width           =   1620
      End
      Begin VB.TextBox txtGross_Weight 
         DataField       =   "Gross_Weight"
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
         Left            =   7440
         TabIndex        =   25
         Tag             =   "25"
         Top             =   4560
         Width           =   1140
      End
      Begin VB.TextBox txtMeasure 
         DataField       =   "Measure"
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
         Left            =   1440
         TabIndex        =   26
         Tag             =   "26"
         Top             =   4920
         Width           =   7095
      End
      Begin VB.TextBox txtAwbNo 
         DataField       =   "AwbNo"
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
         Left            =   1440
         TabIndex        =   2
         Tag             =   "2"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtAwbDate 
         DataField       =   "AwbDate"
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
         Left            =   4560
         TabIndex        =   3
         Tag             =   "3"
         Top             =   720
         Width           =   1560
      End
      Begin VB.TextBox txtOrigin 
         DataField       =   "Origin"
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
         Left            =   7080
         TabIndex        =   4
         Tag             =   "4"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtDestination 
         DataField       =   "Destination"
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
         Left            =   1440
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtRouting 
         DataField       =   "Routing"
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
         Left            =   4560
         TabIndex        =   6
         Tag             =   "6"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtFreight_Prepaid 
         DataField       =   "Freight_Prepaid"
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
         Left            =   1440
         TabIndex        =   7
         Tag             =   "7"
         Top             =   1440
         Width           =   2100
      End
      Begin VB.TextBox txtFreight_Collect 
         DataField       =   "Freight_Collect"
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
         Left            =   4560
         TabIndex        =   8
         Tag             =   "8"
         Top             =   1440
         Width           =   1620
      End
      Begin VB.TextBox txtDec_Cus 
         DataField       =   "Dec_Cus"
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
         Left            =   7200
         TabIndex        =   9
         Tag             =   "9"
         Top             =   1440
         Width           =   1435
      End
      Begin VB.TextBox txtDec_Carr 
         DataField       =   "Dec_Carr"
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
         Left            =   1440
         TabIndex        =   10
         Tag             =   "10"
         Top             =   1800
         Width           =   2100
      End
      Begin VB.TextBox txtE_No 
         DataField       =   "E_No"
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
         Left            =   4560
         TabIndex        =   11
         Tag             =   "11"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtE_Date 
         DataField       =   "E_Date"
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
         Left            =   7200
         TabIndex        =   12
         Tag             =   "12"
         Top             =   1800
         Width           =   1435
      End
      Begin VB.TextBox txtLcNo 
         DataField       =   "L/cNo"
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
         Left            =   1440
         TabIndex        =   13
         Tag             =   "13"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtLcDate 
         DataField       =   "L/cDate"
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
         Left            =   5880
         TabIndex        =   14
         Tag             =   "14"
         Top             =   2160
         Width           =   2745
      End
      Begin VB.TextBox txtExp_Reg 
         DataField       =   "Exp_Reg"
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
         Left            =   1440
         TabIndex        =   15
         Tag             =   "15"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtExp_Permit 
         DataField       =   "Exp_Permit"
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
         Left            =   5880
         TabIndex        =   16
         Tag             =   "16"
         Top             =   2520
         Width           =   2775
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
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Tag             =   "0"
         Top             =   360
         Width           =   2955
      End
      Begin VB.TextBox txtInvDate 
         DataField       =   "InvDate"
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
         Left            =   6000
         TabIndex        =   1
         Tag             =   "1"
         Top             =   360
         Width           =   2640
      End
      Begin VB.TextBox txtShipper 
         DataField       =   "Shipper"
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
         Left            =   1440
         TabIndex        =   17
         Tag             =   "17"
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txtConsignee 
         DataField       =   "Consignee"
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
         Left            =   5880
         TabIndex        =   18
         Tag             =   "18"
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtNotify 
         DataField       =   "Notify"
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
         Left            =   1440
         TabIndex        =   19
         Tag             =   "19"
         Top             =   3240
         Width           =   7215
      End
      Begin MSDataGridLib.DataGrid grdInvoice 
         Height          =   1695
         Left            =   0
         TabIndex        =   57
         Top             =   4080
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   2
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1500.095
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Also Notify"
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
         Left            =   360
         TabIndex        =   60
         Top             =   3600
         Width           =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8760
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Marks"
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
         Left            =   765
         TabIndex        =   54
         Top             =   4200
         Width           =   525
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number"
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
         Left            =   3645
         TabIndex        =   53
         Top             =   4200
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carton Qty"
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
         Left            =   480
         TabIndex        =   52
         ToolTipText     =   "Carton Quantity"
         Top             =   4560
         Width           =   870
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pcs Qty"
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
         Left            =   3570
         TabIndex        =   51
         ToolTipText     =   "Pcs Qunatity"
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Gross Weight"
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
         Left            =   6210
         TabIndex        =   50
         Top             =   4560
         Width           =   1125
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Measurement"
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
         Left            =   240
         TabIndex        =   49
         Top             =   4920
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "AwbNo"
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
         Left            =   645
         TabIndex        =   48
         ToolTipText     =   "Airway Bill No"
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Awb Date"
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
         Left            =   3600
         TabIndex        =   47
         ToolTipText     =   "Airway Bill Date"
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Origin:"
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
         Left            =   6240
         TabIndex        =   46
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Destination"
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
         Left            =   405
         TabIndex        =   45
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Routing"
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
         Left            =   3765
         TabIndex        =   44
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Freight Pre"
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
         Left            =   435
         TabIndex        =   43
         ToolTipText     =   "Freight prepaid"
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Collected"
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
         Left            =   3600
         TabIndex        =   42
         ToolTipText     =   "Freight Collected"
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dec Cust"
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
         Left            =   6270
         TabIndex        =   41
         ToolTipText     =   "Declare Custom"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dec Carraige"
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
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "Declared Carraige  "
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E_No"
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
         Left            =   3765
         TabIndex        =   39
         ToolTipText     =   "Form 'E' no."
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E_Date"
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
         Left            =   6405
         TabIndex        =   38
         ToolTipText     =   "Form 'E' Date"
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "L/c No"
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
         Left            =   600
         TabIndex        =   37
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "L/c Date"
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
         Left            =   5160
         TabIndex        =   36
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exp Reg"
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
         Left            =   480
         TabIndex        =   35
         ToolTipText     =   "Export Registraion No"
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exp Permit"
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
         Left            =   4890
         TabIndex        =   34
         ToolTipText     =   "Export Permit No"
         Top             =   2520
         Width           =   885
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
         Index           =   21
         Left            =   360
         TabIndex        =   33
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
         Index           =   22
         Left            =   4815
         TabIndex        =   32
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Shipper"
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
         Left            =   525
         TabIndex        =   31
         Top             =   2880
         Width           =   645
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Consignee"
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
         Left            =   4845
         TabIndex        =   30
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Notify"
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
         Left            =   600
         TabIndex        =   29
         Top             =   3240
         Width           =   465
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   27
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
            Object.ToolTipText     =   "Add new Record"
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
            Object.ToolTipText     =   "Move First"
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
            Object.ToolTipText     =   "Move Previous"
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
            Picture         =   "frmPerforma.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":0AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":0BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":0D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":0EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":13F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":193A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":1A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":1C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerforma.frx":1D7E
            Key             =   ""
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
         Caption         =   "&Bank Form"
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
      Begin VB.Menu edit 
         Caption         =   "&Edit"
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
Attribute VB_Name = "frmPerforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Object Variables to Create Instance of ADO

Dim rsMas As ADODB.Recordset
Dim rsDet As ADODB.Recordset
Dim rsEdit As ADODB.Recordset
Dim Cn As ADODB.Connection
Dim Cmd As ADODB.Command
Dim sSqlMas, sSqlDet, sTemp, sTemp1, sqlTemp As String
Dim bNew, bNewMas, rsDetail, btemp, rsQry As Boolean


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
    If rsMas.State = adStateOpen Then
        rsMas.Close
    End If
    sSqlMas = "Select * from PerformaMasTbl"
    rsMas.Open sSqlMas, Cn, adOpenKeyset, adLockPessimistic
End Sub

'****************************************************************
'Subroutine to open the Detail Recordset

Sub OpenDetConnection()
    
    Set Cn = New ADODB.Connection
    Set rsDet = New ADODB.Recordset
    With Cn
        .Provider = "MICROSOFT.JET.OLEDB.3.51"
        .ConnectionString = App.Path & "\Export.mdb"
        .Open
    End With
    
End Sub

'***************************************************************
'Subroutine to opent the Detail Recordset
Sub OpenDetRecordSet()
    If rsDet.State = adStateOpen Then
        rsDet.Close
    End If
    sSqlDet = "Select * from PerformaDetTbl"
    rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
    
End Sub
'****************************************************************
'Subroutine to open the Detail Recordset
Sub OpenEditConnection()
    
    Set Cn = New ADODB.Connection
    Set rsEdit = New ADODB.Recordset
    With Cn
        .Provider = "MICROSOFT.JET.OLEDB.3.51"
        .ConnectionString = App.Path & "\Export.mdb"
        .Open
    End With
    
End Sub
'************************************************************
  Sub DeleteTemp()
    
    
    If rsEdit.State = adStateOpen Then
        rsEdit.Close
    End If
        'Open Recordset to delete all record from the temp table
        rsEdit.Open "Select * from TempPerformaTbl", Cn, adOpenKeyset, adLockPessimistic
        rsEdit.MoveFirst
    While Not rsEdit.EOF
        rsEdit.delete adAffectCurrent
        rsEdit.MoveNext
    Wend
    btemp = False
 End Sub
'***************************************************************
'Subroutine to opent the Edit Recordset.
Sub OpenEditRecordSet()
    
    If rsEdit.State = adStateOpen Then
        rsEdit.Close
    End If
    
    sqlTemp = "Select * from TempPerformaTbl"
    rsEdit.Open sqlTemp, Cn, adOpenKeyset, adLockPessimistic
    
End Sub
'****************************************************
'*********************************************************
'*********************************************************
Public Sub ShowEdit()
    'method for saving the record in database
    'rsDet.Fields(0) = Me.txtInvNo.Text
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
              If rsEdit.Fields(Val(mControl.Tag) - Val(20)) <> "" Then
                    mControl = rsEdit.Fields(Val(mControl.Tag) - Val(20))
              Else
                 mControl = " "
                 'mControl.Locked = True
              End If
        End If
    Next
    
End Sub
 '************************************************
 Sub SwapEditData()
     Dim J
    rsDet.MoveFirst
    rsEdit.MoveFirst
        While Not rsEdit.EOF = True  'Swap values from the Temp Table to the Original Table
             For J = 0 To 6
                rsDet.Fields(J) = rsEdit.Fields(J)
             Next J
            rsDet.Update
            rsDet.MoveNext
            rsEdit.MoveNext
         Wend
 End Sub
'*******************************************************

Private Sub bank_Click()
    frmBank.Show
    Unload Me
    
End Sub

Private Sub cmdEditNext_Click()
          SaveEdit        'Subroutine to save edited data in Temp Table
          btemp = True
          rsEdit.MoveNext
        If rsEdit.EOF = True Then
            MsgBox "Updation Complete for the last  Record", vbInformation
            Me.cmdEditNext.Enabled = False
            
            rsEdit.MoveLast
                                
        Else
                       
            ShowEdit
            
        End If
    
End Sub

'Fill the Edited data into

Private Sub cmdEditUpdate_Click()
On Error GoTo 0
    'rsDet.MoveFirst
    'rsMas.MoveFirst
    UnLockRecord
    SaveMasRecord
    SwapEditData
    DeleteTemp      ' Delete Temportay Table from the database
    MsgBox "Records chage SuccessFully...", vbInformation
    Me.cmdEditUpdate.Enabled = False
    Me.cmdEditNext.Enabled = False
    
              

End Sub
'*********************************************************
Public Sub SaveEdit()
    'method for saving the record in database
    rsEdit.Fields(0) = Me.txtInvNo.Text
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
             rsEdit.Fields(Val(mControl.Tag) - Val(20)) = mControl
        End If
    Next
    rsEdit.Update
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
        Set Me.grdInvoice.DataSource = Nothing
        Me.grdInvoice.Visible = False
        VisibleDet
        EnableDet
     End If
    Me.delete.Enabled = False
End Sub
'*****************************************************
Private Sub edit_Click()
    Dim i, J
    
    If btemp = True Then
        DeleteTemp
    End If
    
    OpenEditConnection      'Open Edit Connection
    OpenEditRecordSet
    While Not rsDet.EOF = True
            rsEdit.AddNew
            For J = 0 To 6
                rsEdit.Fields(J) = rsDet.Fields(J)
             Next J
            rsEdit.Update
            rsDet.MoveNext
            btemp = True
         Wend
      
     
     sEdit = "Select * from TempPerformaTbl where InvNo = '" & Me.txtInvNo & "'"
     If rsEdit.State = adStateOpen Then
        rsEdit.Close
    End If
     rsEdit.Open sEdit, Cn, adOpenKeyset, adLockPessimistic
     rsEdit.MoveFirst
    Set Me.grdInvoice.DataSource = Nothing
     Me.grdInvoice.Visible = False
        
        VisibleDet
        EnableDet
        UnLockRecord
        ClearDet
        rsEdit.MoveFirst
        ShowEdit
                
    Me.new.Enabled = False
    Me.save.Enabled = False
    Me.cancel.Enabled = True
    Me.query.Enabled = False
    Me.cmdNext.Visible = False
    Me.CmdUpdate.Visible = False
    Me.cmdEditNext.Visible = True
    Me.cmdEditUpdate.Visible = True
    Me.cmdEditNext.Enabled = True
    Me.cmdEditUpdate.Enabled = True
    
    Me.mnufind.Enabled = False
    Me.edit.Enabled = False
    Me.grdInvoice.Visible = False
        
              Me.First.Enabled = False
              Me.next.Enabled = False
              Me.previous.Enabled = False
              Me.last.Enabled = False
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(8).Enabled = False
                
        
End Sub
'***************Edit Section Ends********************


'**************************************************************
' Subroutine to find a record on the given criteria

Sub FindRecord()
    If rsDet.State = adStateOpen Then
        rsDet.Close
    End If
        If rsMas.State = adStateOpen Then
            rsMas.Close
        End If
        'rsMas.MoveFirst
       sTemp = Me.txtInvNo.Text
       
       rsMas.Open "Select * From PerformaMasTbl", Cn, adOpenKeyset, adLockPessimistic
       If rsMas.RecordCount <> 0 Then
              
            rsMas.Find "InvNo ='" & sTemp & "'"
                If rsMas.EOF Then
                    MsgBox "Record does not exist", vbInformation
                    rsMas.MoveFirst
                    ClearMas
             'Me.grdInvoice.Visible = False
             'VisibleDet
              
                    Me.save.Enabled = False
                    Me.txtInvNo.SetFocus
         
                Else
                    FillMas
                    sSqlDet = "Select * from PerformaDetTbl where InvNo ='" & Me.txtInvNo & "'"
           
                    If rsDet.State = adStateOpen Then
                        rsDet.Close
                    End If
            
                    rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
                    rsDetail = True
                    unVisibleDet
       
                    Me.grdInvoice.Visible = True
                    Set Me.grdInvoice.DataSource = rsDet
            End If
        Else
            MsgBox "No Record found in the Database", vbInformation
        End If
      rsQry = False
End Sub
'************************************************************
'************************************************************
Private Sub cancel_Click()
    If bNew = True Then
        rsMas.CancelUpdate
        Me.new.Enabled = True
        
     ElseIf bNewMas = True Then
         rsMas.CancelUpdate
         rsDet.CancelUpdate
         Me.new.Enabled = True
    End If
      EnableMas
      EnableDet
      ClearMas
      ClearDet
      Me.txtInvNo.SetFocus
      Me.cancel.Enabled = False
      Me.query.Enabled = True
      Me.mnufind.Enabled = True
      Me.save.Enabled = False
      Me.mnufind.Enabled = False
      Me.new.Enabled = True
      Me.cmdEditNext.Visible = False
      Me.cmdEditUpdate.Visible = False
      Me.cmdNext.Visible = False
      Me.CmdUpdate.Visible = False
      
      ClearMas              'Clear Master Records
      
      ClearDet              'Clear Detail Records
      VisibleDet
      UnLockRecord
      
       Set Me.grdInvoice.DataSource = Nothing
      Me.grdInvoice.Visible = False
               
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
    Toolbar1.Buttons(3).Enabled = False
End Sub

Private Sub Certificate_Click()
    
    frmCertificate.Show
    Unload Me
End Sub
 
'****************************** '******************************
Private Sub cmdNext_Click()
      
    'Loop to enable and clear all detail fields
              Me.First.Enabled = True
              Me.next.Enabled = True
              Me.previous.Enabled = True
              Me.last.Enabled = True
        
        
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = True
    EnableDet
    ClearDet
    Me.txtMarks.SetFocus
    Me.cmdNext.Enabled = False
    Me.CmdUpdate.Enabled = True
    
End Sub


'************************************************************
Private Sub CmdUpdate_Click()
    If Me.txtMarks.Text = "" Then
        MsgBox "Incomplete data to save ", vbInformation
    Else
        
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
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = True
       
       rsDet.AddNew
        rsDet.Fields(0) = Me.txtInvNo.Text
        On Error GoTo Errorhandler
        SaveDetRecord
        DisableDet
        Me.CmdUpdate.Enabled = False
        Me.cmdNext.Enabled = True
        
        
    End If
Errorhandler:
        
        Select Case Err.Number
        
            Case -2147217887
                MsgBox "Record Already exist", vbInformation
                 rsMas.CancelUpdate
                 rsDet.CancelUpdate
                 
            Case -2147352571
                MsgBox "There is Invalid Data In some Fields,Record Cann't be saved ", vbInformation
                
                rsMas.CancelUpdate
                 rsDet.CancelUpdate
                 
        End Select
    
End Sub
Private Sub Combined_Click()
    frmGenSystem.Show
    Unload Me
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
Sub FILL()
    Me.txtCar_Qty = "Total Cartons ="
    Me.txtPcs_Qty = "Total Pcs ="
End Sub
'***************************************************************
Private Sub First_Click()
On Error GoTo 0
       If rsDet.State = adStateOpen Then       'Check if recordset is already open
            rsDet.Close
        End If
              rsMas.MoveFirst
              sSqlDet = "Select * from PerformaDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
              rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
              rsDetail = True
              EnableMas
              FillMas
              LockRecord
              unVisibleDet
              Me.grdInvoice.Visible = True
              Set Me.grdInvoice.DataSource = rsDet
              
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
End Sub
''*********************************************************
Private Sub Form_Load()
                    
'* Call subroutine to open the Master and Detail Connection
    
    OpenMasConnection
    OpenMasRecordSet
    OpenDetConnection
    OpenDetRecordSet
    rsDetail = True
    
    
    DisableMas
    DisableDet
    
    Me.new.Enabled = True
    Me.save.Enabled = False
    Me.cancel.Enabled = False
    
    Me.cmdNext.Visible = False
    Me.CmdUpdate.Visible = False
    Me.cmdEditNext.Visible = False
    Me.cmdEditUpdate.Visible = False
    Me.mnufind.Enabled = False
    Me.edit.Enabled = False
    Me.grdInvoice.Visible = False
    
    
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

'***************************************************

Private Sub last_Click()
On Error GoTo 0
   If rsDet.State = adStateOpen Then       'Check if recordset is already open
            rsDet.Close
        End If
                     
              rsMas.MoveLast
              sSqlDet = "Select * from PerformaDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
              rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
              rsDetail = True
              EnableMas
              FillMas
              unVisibleDet
              Me.grdInvoice.Visible = True
              Set Me.grdInvoice.DataSource = rsDet
              LockRecord            'Lock records
              
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
End Sub

Private Sub main_Click()
    frmMain.Show
    Unload Me
    
End Sub

Private Sub mnufind_Click()
    
    If Me.txtInvNo.Text = "" Then
        MsgBox "Enter Invoice No. to find Record", vbInformation
        
    Else
        FindRecord
        
    End If
    
    
     Me.new.Enabled = False
    Me.save.Enabled = False
    Me.cancel.Enabled = True
    Me.query.Enabled = True
    Me.delete.Enabled = True
    
    Me.cmdNext.Visible = False
    Me.CmdUpdate.Visible = False
    Me.cmdEditNext.Visible = False
    Me.cmdEditUpdate.Visible = False
    Me.edit.Enabled = True
    
    Me.mnufind.Enabled = False
    
    'Me.grdInvoice.Visible = False
        
              Me.First.Enabled = False
              Me.next.Enabled = False
              Me.previous.Enabled = False
              Me.last.Enabled = False
    
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(8).Enabled = False
    rsQry = False
End Sub


'***************************************************************
'****** Open the Recordset*********************
Private Sub new_Click()
    
    If bNew Then
       rsDet.delete adAffectCurrent
       Me.CmdUpdate.Visible = False
        
    End If
    If rsDet.State <> adStateOpen Then
        OpenDetRecordSet
    End If
    
    rsMas.AddNew
    rsDet.AddNew
    bNewMas = True
    
    ' To enable and clear all the controls
    
    EnableMas       'Enables Master Record
    VisibleDet      'Visible Detail fields
    EnableDet       'Enable Detail Fields
    ClearMas        'clear Master fields
    ClearDet        'Clear Detail Fields
    UnLockRecord    'Unlock Record
    FIllExisting    'To fill the default fields
    
    Me.new.Enabled = False
    Me.save.Enabled = True
    Me.cancel.Enabled = True
    Me.query.Enabled = False
    Me.cmdNext.Visible = False
    Me.CmdUpdate.Visible = False
    Me.cmdEditNext.Visible = False
    Me.cmdEditUpdate.Visible = False
    
    Me.mnufind.Enabled = False
    Me.edit.Enabled = False
    Me.grdInvoice.Visible = False
        
              Me.First.Enabled = False
              Me.next.Enabled = False
              Me.previous.Enabled = False
              Me.last.Enabled = False
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(8).Enabled = False
    FILL
    Me.txtInvNo.SetFocus
End Sub
Sub FillDate()
    
    
    Me.txtAwbDate = Date
    Me.txtE_Date = Date
    Me.txtInvDate = Date
    Me.txtLcDate = Date
    
End Sub

Sub FIllExisting()
    
    Me.txtOrigin.Text = "FAISALABAD"
    Me.txtShipper = "ALPHA KNITTING (PVT) LTD ,220 R.B ,PATHANWALA,JHANG ROAD ,FAISALABAD,PAKISTAN"
    Me.txtGross_Weight = "Kgs"
    
    
  
End Sub
Private Sub Packing_Click()
    frmPacking.Show
    Unload Me
End Sub

'***************************************************************
Private Sub next_Click()
    On Error GoTo 0
    If rsDet.State = adStateOpen Then
          rsDet.Close
     End If
               'On Error GoTo Errorhandler
    
          rsMas.MoveNext
          On Error GoTo Errorhandler
          If rsMas.EOF = True Then
            
            Toolbar1.Buttons(4).Enabled = True
            Toolbar1.Buttons(5).Enabled = False
            Toolbar1.Buttons(6).Enabled = True
            Toolbar1.Buttons(7).Enabled = True
            
            Me.First.Enabled = True
            Me.next.Enabled = False
            Me.previous.Enabled = True
            Me.last.Enabled = True
            MsgBox "You are at First Record", vbInformation
       
         Else
                    'rsMas.MoveNext
                    sSqlDet = "Select * from PerformaDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
                    rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
                                        
                    EnableMas
                    LockRecord          'lock records
                    FillMas
                    unVisibleDet
                    Me.grdInvoice.Visible = True
                    Set Me.grdInvoice.DataSource = rsDet
                    Toolbar1.Buttons(4).Enabled = True
              
              'Toolbar1.Buttons(5).Enabled = True
              Toolbar1.Buttons(6).Enabled = True
              Toolbar1.Buttons(7).Enabled = True
               
               'Menu options disabled and enabled
              Me.First.Enabled = True
              'Me.next.Enabled = True
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

'************************************************************
Private Sub previous_Click()
        On Error GoTo 0
            If rsDet.State = adStateOpen Then      'Check if recordset is open
                rsDet.Close
            End If
              
              LockRecord            'Lock records
              rsMas.MovePrevious
               On Error GoTo Errorhandler
                If rsMas.BOF = True Then
                    Toolbar1.Buttons(4).Enabled = False
                    Toolbar1.Buttons(5).Enabled = True
                    Toolbar1.Buttons(6).Enabled = False
                    Toolbar1.Buttons(7).Enabled = True
                    
                    Me.First.Enabled = False
                    Me.next.Enabled = True
                    Me.previous.Enabled = False
                    Me.last.Enabled = True
                    MsgBox "You are at the First Record", vbInformation
                    'rsMas.MoveFirst
                Else
                    
              sSqlDet = "Select * from PerformaDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
              rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
              rsDetail = True
              EnableMas             'Subroutine to enable master fields
              FillMas               'Subroutine to Fill master fields
              unVisibleDet
              Me.grdInvoice.Visible = True
              Set Me.grdInvoice.DataSource = rsDet
              Me.grdInvoice.Columns(0).Width = 0
              Toolbar1.Buttons(4).Enabled = True
              Toolbar1.Buttons(5).Enabled = True
              'Toolbar1.Buttons(6).Enabled = True
              Toolbar1.Buttons(7).Enabled = True
               
               'Menu options disabled and enabled
              Me.First.Enabled = True
              Me.next.Enabled = True
              'Me.previous.Enabled = True
              Me.last.Enabled = True
          End If
Errorhandler:
0
 End Sub

'************************************************************
Private Sub query_Click()
    'Call subroutine to clear and enable controls
    If rsDetail = True Then
        rsDet.CancelUpdate
        rsDet.Close
        rsDetail = False
    End If
    
    EnableMas
    ClearMas
    EnableDet
    ClearDet
    
    Set Me.grdInvoice.DataSource = Nothing
    Me.txtInvNo.SetFocus
    
     Me.new.Enabled = False
    Me.save.Enabled = False
    Me.cancel.Enabled = True
    Me.query.Enabled = False
    Me.cmdNext.Visible = False
    Me.CmdUpdate.Visible = False
    Me.cmdEditNext.Visible = False
    Me.cmdEditUpdate.Visible = False
    Me.edit.Enabled = False
    
    Me.mnufind.Enabled = True
    
    'Me.grdInvoice.Visible = False
        
              Me.First.Enabled = False
              Me.next.Enabled = False
              Me.previous.Enabled = False
              Me.last.Enabled = False
    
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(8).Enabled = True
    
    rsQry = True            'Query Variable
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
'******************************************************

Private Sub save_Click()
    
    If Me.txtInvNo.Text = "" Or Me.txtOrigin.Text = "" Or Me.txtConsignee.Text = "" Or Me.txtDestination.Text = "" Then
        MsgBox "Incomplete Data to Save", vbInformation
        Me.txtInvNo.SetFocus
   Else
    
'Call subroutine to save master and detail record
    
        On Error GoTo Errorhandler
        SaveMasRecord
        SaveDetRecord
    
 'Call subroutine to disable master and detail section
    DisableMas
    DisableDet
    
    
    Me.cmdNext.Visible = True
    Me.cmdNext.Enabled = True
    Me.CmdUpdate.Visible = True
    Me.CmdUpdate.Enabled = False
    Me.cmdEditNext.Visible = False
    Me.cmdEditUpdate.Visible = False
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
    End If

Errorhandler:
        
        Select Case Err.Number
        
            Case -2147217887
                MsgBox "Record Already exist", vbInformation
                
                 
             Case -2147352571
                MsgBox "There is Invalid Data In some Fields,Record Cann't be saved ", vbInformation
                               
               
        End Select
   
End Sub
Public Sub SaveMasRecord()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 20 Then
             rsMas.Fields(Val(mControl.Tag)) = mControl
        End If
    Next
    rsMas.Update
End Sub

'*********************************************************
Public Sub SaveDetRecord()
    'method for saving the record in database
    rsDet.Fields(0) = Me.txtInvNo.Text
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
             rsDet.Fields(Val(mControl.Tag) - Val(20)) = mControl
        End If
    Next
    rsDet.Update
End Sub
'*******************************************************
Sub ClearDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
        mControl = ""
    End If
    Next
End Sub

'************************************************************

Sub ClearMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 20 Then
             mControl = ""
        End If
    Next
End Sub
'*******************************************************
Sub EnableDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
             mControl.Enabled = True
        End If
    Next
End Sub

'*******************************************************
Sub DisableDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
             mControl.Enabled = False
        End If
    Next
End Sub

'************************************************************

Sub EnableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 20 Then
             mControl.Enabled = True
        End If
    Next
End Sub
'************************************************************

Sub DisableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 20 Then
             mControl.Enabled = False
        End If
    Next
End Sub
'***************************************************
Sub VisibleDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
             mControl.Visible = True
        End If
    Next
End Sub
'**********************************************
Sub unVisibleDet()
  
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
             mControl.Visible = False
        End If
    Next
End Sub
'****************************************************

Public Sub FillMas()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 20 Then
               If rsMas.Fields(Val(mControl.Tag)) <> "" Then
                    mControl = rsMas.Fields(Val(mControl.Tag))
                Else
                    mControl = " "
                End If
        End If
    Next
    
End Sub

'*********************************************************
Public Sub FillDet()
    'method for saving the record in database
    rsDet.Fields(0) = Me.txtInvNo.Text
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 21 And mControl.Tag <= 26 Then
             If rsDet.Fields(Val(mControl.Tag)) <> "" Then
              mControl = rsDet.Fields(Val(mControl.Tag) - Val(20))
        Else
            mControl = " "
            
        End If
        End If
    Next
    
End Sub
'*******************************************************
Sub LockRecord()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 20 Then
             mControl.Locked = True
        End If
    Next
End Sub

'*******************************************************
Sub UnLockRecord()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 20 Then
             mControl.Locked = False
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
                    Me.edit.Enabled = True
                    Me.mnufind.Enabled = False
                    Me.query.Enabled = True
                    Me.delete.Enabled = True
                    FindRecord      'Call subroutine to find a record
                    
                End If
            Case "print"
                frmPrint.Show
    End Select

End Sub


Private Sub txtAsloNotify_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtMarks.SetFocus
    End If
End Sub

Private Sub txtAwbDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtOrigin.SetFocus
    End If
End Sub

Private Sub txtAwbNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAwbDate.SetFocus
    End If
End Sub



Private Sub txtCar_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPcs_Qty.SetFocus
    End If
End Sub

Private Sub txtConsignee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtNotify.SetFocus
    End If
End Sub

Private Sub txtDec_Carr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtE_No.SetFocus
    End If
End Sub


Private Sub txtDec_Cus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDec_Carr.SetFocus
    End If
End Sub

Private Sub txtDestination_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtRouting.SetFocus
    End If
End Sub

Private Sub txtE_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtLcNo.SetFocus
    End If
End Sub

Private Sub txtE_No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtE_Date.SetFocus
    End If
End Sub

Private Sub txtExp_Permit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtShipper.SetFocus
    End If
End Sub

Private Sub txtExp_Reg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExp_Permit.SetFocus
    End If
End Sub

Private Sub txtFreight_Collect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDec_Cus.SetFocus
    End If
End Sub

Private Sub txtFreight_Prepaid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFreight_Collect.SetFocus
    End If
End Sub

Private Sub txtGross_Weight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.txtMeasure.SetFocus
    End If
End Sub



Private Sub txtInvDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAwbNo.SetFocus
    End If
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtInvDate.SetFocus
    End If
End Sub


Private Sub txtLcDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExp_Reg.SetFocus
    End If
End Sub


Private Sub txtLcNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtLcDate.SetFocus
    End If
End Sub


Private Sub txtMarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtNumber.SetFocus
    End If
End Sub


Private Sub txtNotify_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAsloNotify.SetFocus
    End If
End Sub


Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCar_Qty.SetFocus
    End If
End Sub


Private Sub txtOrigin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDestination.SetFocus
    End If
End Sub


Private Sub txtPcs_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtGross_Weight.SetFocus
    End If
End Sub


Private Sub txtRouting_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFreight_Prepaid.SetFocus
    End If
End Sub


Private Sub txtShipper_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtConsignee.SetFocus
    End If
End Sub
