VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCertificate 
   Caption         =   "ALPHA KNITTING (PVT) LTD..Certificate OF Origin)"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmCertificate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Certificate Of Origin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5895
      Left            =   600
      TabIndex        =   24
      Top             =   480
      Width           =   8055
      Begin VB.TextBox txtDate 
         DataField       =   "Destination"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Tag             =   "12"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtTotalPcs 
         DataField       =   "Qty_Desc"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   6960
         TabIndex        =   16
         Tag             =   "16"
         Top             =   3600
         Width           =   855
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
         Left            =   4080
         TabIndex        =   50
         Top             =   5400
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
         Left            =   6240
         TabIndex        =   49
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox txtPcs_Qty 
         DataField       =   "Car_Qty"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Tag             =   "20"
         Top             =   4680
         Width           =   2655
      End
      Begin VB.TextBox txtCar_Qty 
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
         Left            =   6360
         TabIndex        =   19
         Tag             =   "19"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   6480
         TabIndex        =   46
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >>"
         Height          =   375
         Left            =   4080
         TabIndex        =   45
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtInvNo 
         DataField       =   "Name"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Tag             =   "0"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtDescription 
         DataField       =   "Qty_Desc"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Tag             =   "21"
         Top             =   5040
         Width           =   6615
      End
      Begin VB.TextBox txtMarks 
         DataField       =   "Marks"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Tag             =   "17"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox txtNum 
         DataField       =   "Car_Qty"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   3600
         TabIndex        =   18
         Tag             =   "18"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txtFOB 
         DataField       =   "FOB"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   5040
         TabIndex        =   22
         Tag             =   "22"
         Top             =   4680
         Width           =   2745
      End
      Begin VB.TextBox txtQty_Desc 
         DataField       =   "Qty_Desc"
         DataMember      =   "Command2"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Tag             =   "15"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtSupp_Det 
         DataField       =   "Supp_Det"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Tag             =   "14"
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtTransport 
         DataField       =   "Transport"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Tag             =   "13"
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtPlace 
         DataField       =   "Place"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   4800
         TabIndex        =   11
         Tag             =   "11"
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox txtDestination 
         DataField       =   "Destination"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Tag             =   "10"
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox txtOrigin 
         DataField       =   "Origin"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   4800
         TabIndex        =   9
         Tag             =   "9"
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txtConCountry 
         DataField       =   "ConCountry"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Tag             =   "8"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtConAddress 
         DataField       =   "ConAddress"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Tag             =   "7"
         Top             =   2175
         Width           =   6615
      End
      Begin VB.TextBox txtConName 
         DataField       =   "ConName"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Tag             =   "6"
         Top             =   1800
         Width           =   6615
      End
      Begin VB.TextBox txtCatNo 
         DataField       =   "CatNo"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtQuotaYear 
         DataField       =   "QuotaYear"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Tag             =   "4"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtCountry 
         DataField       =   "Country"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1080
         Width           =   6615
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Tag             =   "2"
         Top             =   720
         Width           =   6615
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   285
         Left            =   4800
         TabIndex        =   1
         Tag             =   "1"
         Top             =   360
         Width           =   3015
      End
      Begin MSDataGridLib.DataGrid grdInvoice 
         Height          =   1695
         Left            =   0
         TabIndex        =   52
         Top             =   4080
         Width           =   7935
         _ExtentX        =   13996
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
         AutoSize        =   -1  'True
         Caption         =   "Total Pcs"
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
         Left            =   6120
         TabIndex        =   51
         ToolTipText     =   "Quantity Description"
         Top             =   3600
         Width           =   750
      End
      Begin VB.Label lblFieldLabel 
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
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Car Qty"
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
         Left            =   5520
         TabIndex        =   47
         ToolTipText     =   "Pieces Quantity"
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
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
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         Left            =   120
         TabIndex        =   43
         Top             =   5040
         Width           =   945
      End
      Begin VB.Label lblFieldLabel 
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
         Index           =   19
         Left            =   360
         TabIndex        =   42
         Top             =   4320
         Width           =   525
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Numbers"
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
         Left            =   2640
         TabIndex        =   41
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
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
         Index           =   16
         Left            =   4080
         TabIndex        =   40
         Top             =   4680
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
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
         Index           =   15
         Left            =   3480
         TabIndex        =   39
         ToolTipText     =   "Quantity Description"
         Top             =   3600
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Supp_Det"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   38
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Transp"
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
         Left            =   4080
         TabIndex        =   37
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label lblFieldLabel 
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
         Index           =   12
         Left            =   600
         TabIndex        =   36
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Place"
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
         Left            =   4200
         TabIndex        =   35
         Top             =   2880
         Width           =   435
      End
      Begin VB.Label lblFieldLabel 
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
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Origin"
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
         Left            =   4200
         TabIndex        =   33
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Country"
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
         Left            =   360
         TabIndex        =   32
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   360
         TabIndex        =   31
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Con_Name"
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
         TabIndex        =   30
         Top             =   1845
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cat No"
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
         Left            =   4080
         TabIndex        =   29
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "QuotaYear"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Country"
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
         Left            =   360
         TabIndex        =   27
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblFieldLabel 
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
         Left            =   3840
         TabIndex        =   25
         Top             =   360
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8040
         Y1              =   4080
         Y2              =   4080
      End
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
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "first"
            Object.ToolTipText     =   "Move First Record"
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
            Object.ToolTipText     =   "Move last Record"
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
            Picture         =   "frmCertificate.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":2CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":2E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":2F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":30B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":3212
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":3756
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":3C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":3DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":3F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertificate.frx":40DE
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
      Begin VB.Menu Packing 
         Caption         =   "Packing List Form"
      End
      Begin VB.Menu bank 
         Caption         =   "&Bank Form"
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
   Begin VB.Menu Action_menu 
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
   Begin VB.Menu mnusearch 
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
Attribute VB_Name = "frmCertificate"
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
Dim sSqlMas, sSqlDet, sTemp, sTemp1, sqlTemp, sEdit As String
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
    
    sSqlMas = "Select * from OriginTexMasTbl"
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
    
    sSqlDet = "Select * from OriginTexDetTbl"
    rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
    
End Sub

'****************************************************************************


Sub OpenEditConnection()
    
    Set Cn = New ADODB.Connection
    Set rsEdit = New ADODB.Recordset
    With Cn
        .Provider = "MICROSOFT.JET.OLEDB.3.51"
        .ConnectionString = App.Path & "\Export.mdb"
        .Open
    End With
    
End Sub
'***************************************************************
'Subroutine to opent the Edit Recordset.
Sub OpenEditRecordSet()
    
        If rsEdit.State = adStateOpen Then
            rsEdit.Close
        End If
        
    sqlTemp = "Select * from TempOriginTexTbl"
    rsEdit.Open sqlTemp, Cn, adOpenKeyset, adLockPessimistic
    
End Sub
'****************************************************
'*********************************************************
Public Sub ShowEdit()
    
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
              If rsEdit.Fields(Val(mControl.Tag) - Val(16)) <> "" Then
                    mControl = rsEdit.Fields(Val(mControl.Tag) - Val(16))
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
             For J = 0 To 5
                rsDet.Fields(J) = rsEdit.Fields(J)
             Next J
            rsDet.Update
            rsDet.MoveNext
            rsEdit.MoveNext
         Wend
 End Sub
'*******************************************************

Private Sub cmdEditNext_Click()
          
          SaveEdit        'Subroutine to save edited data in Temp Table
          btemp = True
          rsEdit.MoveNext
        If rsEdit.EOF = True Then
           ' MsgBox "Updation Complete for the last  Record", vbInformation
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
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
             rsEdit.Fields(Val(mControl.Tag) - Val(16)) = mControl
        End If
    Next
    rsEdit.Update
End Sub

'************************************************************
   Sub DeleteTemp()
    
    
    If rsEdit.State = adStateOpen Then
        rsEdit.Close
    End If
        'Open Recordset to delete all record from the temp table
        rsEdit.Open "Select * from TempOriginTexTbl", Cn, adOpenKeyset, adLockPessimistic
        rsEdit.MoveFirst
    While Not rsEdit.EOF
        rsEdit.delete adAffectCurrent
        rsEdit.MoveNext
    Wend
    btemp = False
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
    OpenEditRecordSet       'Open Recordset to open
        rsDet.MoveFirst
        While Not rsDet.EOF = True
            rsEdit.AddNew
            For J = 0 To 6
                rsEdit.Fields(J) = rsDet.Fields(J)
             Next J
            rsEdit.Update
            rsDet.MoveNext
            btemp = True
         Wend
         
         
     
     sEdit = "Select * from TempOriginTexTbl where InvNo = '" & Me.txtInvNo & "'"
        
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
    Me.delete.Enabled = False
    
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
       
       rsMas.Open "Select * From OriginTexMasTbl", Cn, adOpenKeyset, adLockPessimistic
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
                    sSqlDet = "Select * from OriginTexDetTbl where InvNo ='" & Me.txtInvNo & "'"
           
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
    If Me.txtMarks.Text = "" Or Me.txtCar_Qty = "" Or Me.txtFOB = "" Then
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

Private Sub CmdUpdate_KeyPress(KeyAscii As Integer)
    Me.cmdNext.SetFocus
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

'***************************************************************
Private Sub First_Click()
        On Error GoTo Errorhandler
        If rsDet.State = adStateOpen Then       'Check if recordset is already open
            rsDet.Close
        End If
             On errror GoTo Errorhandler
              rsMas.MoveFirst
              sSqlDet = "Select * from OriginTexDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
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
Errorhandler:
0
End Sub

'*********************************************************
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

Private Sub last_Click()
On Error GoTo 0
    If rsDet.State = adStateOpen Then     'Check if recordset is open
        
         rsDet.Close
    End If
                     
              rsMas.MoveLast
              sSqlDet = "Select * from OriginTexDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
              rsDet.Open sSqlDet, Cn, adOpenKeyset, adLockPessimistic
              rsDetail = True
              EnableMas
              FillMas
              unVisibleDet
              Me.grdInvoice.Visible = True
             
              
              Set Me.grdInvoice.DataSource = rsDet
              Me.grdInvoice.Columns(0).Visible = False
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

'*************************************************************
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
    Me.delete.Enabled = False
        
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
    Me.txtInvNo.SetFocus
End Sub

Sub FillDate()
  
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
                    sSqlDet = "Select * from OriginTexDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
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

'*************************************************************
Private Sub Packing_Click()
    frmPacking.Show
    Unload Me
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
        On Error GoTo Errorhandler
            If rsDet.State = adStateOpen Then      'Check if recordset is open
                rsDet.Close
            End If
              
              LockRecord            'Lock records
              On Error GoTo Errorhandler
              rsMas.MovePrevious
              
             
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
                    
              sSqlDet = "Select * from OriginTexDetTbl where InvNo ='" & rsMas.Fields(0) & "'"
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
    Me.delete.Enabled = True
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

'************************************************************
Sub FIllExisting()
    Me.txtName = " ALPHA KNITTING (PVT) LTD"
    Me.txtAddress = "220 R.B PATHANWALA ,JHANG ROAD,FAISALABAD"
    Me.txtCountry = "PAKISTAN"
    Me.txtOrigin = "PAKISTAN"
    Me.txtDate = Date
    Me.txtPlace = "FAISALABAD"
    
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
    
    If Me.txtInvNo = "" Or Me.txtName = "" Or Me.txtConName = "" Or Me.txtDate = "" Then
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
    Me.CmdUpdate.Enabled = True
    Me.cmdEditNext.Visible = False
    Me.cmdEditUpdate.Visible = False
    Me.query.Enabled = True
    Me.delete.Enabled = False
        
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

'****************************************************

Public Sub SaveMasRecord()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 16 Then
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
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
             rsDet.Fields(Val(mControl.Tag) - Val(16)) = mControl
        End If
    Next
    rsDet.Update
End Sub

'*******************************************************
Sub ClearDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
             mControl = ""
        End If
    Next
End Sub

'************************************************************

Sub ClearMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 16 Then
             mControl = ""
        End If
    Next
End Sub

'*******************************************************
Sub EnableDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
             mControl.Enabled = True
        End If
    Next
End Sub

'*******************************************************
Sub DisableDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
             mControl.Enabled = False
        End If
    Next
End Sub

'************************************************************
Sub EnableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 16 Then
             mControl.Enabled = True
        End If
    Next
End Sub

'************************************************************
Sub DisableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 16 Then
             mControl.Enabled = False
        End If
    Next
End Sub

'***************************************************
Sub VisibleDet()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
             mControl.Visible = True
        End If
    Next
End Sub

'**********************************************
Sub unVisibleDet()
  
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
             mControl.Visible = False
        End If
    Next
End Sub

'****************************************************

Public Sub FillMas()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 16 Then
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
    If mControl.Tag >= 17 And mControl.Tag <= 22 Then
        If rsDet.Fields(Val(mControl.Tag)) <> "" Then
              mControl = rsDet.Fields(Val(mControl.Tag) - Val(16))
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
    If mControl.Tag >= 0 And mControl.Tag <= 16 Then
             mControl.Locked = True
        End If
    Next
End Sub

'*******************************************************
Sub UnLockRecord()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 16 Then
             mControl.Locked = False
        End If
    Next
End Sub

'**********************************************************
Sub NumCheck()
    
    If KeyAscii < 47 Or KeyAscii > 57 Then
        MsgBox "Type Numeric entry in this field", vbInformation
    End If
    
End Sub
'*******************************************************
Public Sub Gridsetting()
    Me.grdInvoice.Columns(0).Width = 0
    
End Sub
Private Sub Invoice_Click()
    frmInvoice.Show
    Unload Me
End Sub

'**************************************************
'**************************************************

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


'*************************************************************************
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCountry.SetFocus
    End If
End Sub

Private Sub txtCar_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPcs_Qty.SetFocus
    End If
End Sub

Private Sub txtCatNo_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Me.txtConName.SetFocus
    End If
End Sub

Private Sub txtConAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.txtConCountry.SetFocus
    
    End If
End Sub

Private Sub txtConCountry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtOrigin.SetFocus
    End If
End Sub

Private Sub txtConName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtConAddress.SetFocus
    
    End If
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtQuotaYear.SetFocus
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtTransport.SetFocus
    End If
    
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFOB.SetFocus
    End If
End Sub

Private Sub txtDestination_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPlace.SetFocus
    End If
End Sub

Private Sub txtFOB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDescription.SetFocus
    End If
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.txtName.SetFocus
    End If
        
End Sub

Private Sub txtMarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtNum.SetFocus
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAddress.SetFocus
    End If
    
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCar_Qty.SetFocus
    End If
End Sub

Private Sub txtOrigin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDestination.SetFocus
    End If
End Sub

Private Sub txtPcsQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    End If
End Sub



Private Sub txtPcs_Qty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFOB.SetFocus
    End If
    
End Sub

Private Sub txtPlace_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDate.SetFocus
    End If
End Sub


Private Sub txtQty_Desc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Me.txtTotalPcs.SetFocus
    End If
End Sub

Private Sub txtQuotaYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCatNo.SetFocus
    End If
End Sub

Private Sub txtSupp_Det_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtQty_Desc.SetFocus
    End If
End Sub

Private Sub txtTotalPcs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMarks.SetFocus
    End If
End Sub

Private Sub txtTransport_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtSupp_Det.SetFocus
    End If
    
End Sub

Private Sub yes_Click()
    End
End Sub
