VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelease 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALPHA kNITTING (pvt) Ltd..(Release And Undertaking Form)"
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
   Icon            =   "frmRelease.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Release And Take :"
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
      Height          =   5055
      Left            =   1080
      TabIndex        =   14
      Top             =   600
      Width           =   6615
      Begin VB.TextBox txtDate 
         DataField       =   "Mr_Date"
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
         Left            =   1920
         TabIndex        =   12
         Tag             =   "12"
         Top             =   4560
         Width           =   2475
      End
      Begin VB.TextBox txtInvNo 
         DataField       =   "ExpidNo"
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
         Left            =   1920
         TabIndex        =   0
         Tag             =   "0"
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtInvDate 
         DataField       =   "S_No"
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
         Left            =   4440
         TabIndex        =   1
         Tag             =   "1"
         Top             =   360
         Width           =   1740
      End
      Begin VB.TextBox txtExp_Name 
         DataField       =   "Exp_Name"
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
         Left            =   1920
         TabIndex        =   2
         Tag             =   "2"
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtExpidNo 
         DataField       =   "ExpidNo"
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
         Left            =   1920
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1095
         Width           =   2460
      End
      Begin VB.TextBox txtS_No 
         DataField       =   "S_No"
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
         Left            =   1920
         TabIndex        =   4
         Tag             =   "4"
         Top             =   1440
         Width           =   2460
      End
      Begin VB.TextBox txtVisa_No 
         DataField       =   "Visa_No"
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
         Left            =   1920
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1860
         Width           =   2460
      End
      Begin VB.TextBox txtVisa_Date 
         DataField       =   "Visa_Date"
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
         Left            =   1920
         TabIndex        =   6
         Tag             =   "6"
         Top             =   2235
         Width           =   2460
      End
      Begin VB.TextBox txtCat 
         DataField       =   "Cat"
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
         Left            =   1920
         TabIndex        =   7
         Tag             =   "7"
         Top             =   2625
         Width           =   2460
      End
      Begin VB.TextBox txtQty 
         DataField       =   "Qty"
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
         Left            =   1920
         TabIndex        =   8
         Tag             =   "8"
         Top             =   3000
         Width           =   2460
      End
      Begin VB.TextBox txtSIB_No 
         DataField       =   "SIB_No"
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
         Left            =   1920
         TabIndex        =   9
         Tag             =   "9"
         Top             =   3375
         Width           =   2460
      End
      Begin VB.TextBox txtSIB_Date 
         DataField       =   "SIB_Date"
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
         Left            =   1920
         TabIndex        =   10
         Tag             =   "10"
         Top             =   3765
         Width           =   2460
      End
      Begin VB.TextBox txtMr_Date 
         DataField       =   "Mr_Date"
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
         Left            =   1920
         TabIndex        =   11
         Tag             =   "11"
         Top             =   4200
         Width           =   2475
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mr Date"
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
         Left            =   1080
         TabIndex        =   27
         Top             =   4560
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "InvDate"
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
         Left            =   3720
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "InvNo"
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
         Left            =   1080
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exp Name"
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
         Left            =   960
         TabIndex        =   24
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Expid No"
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
         Left            =   1080
         TabIndex        =   23
         Top             =   1140
         Width           =   690
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
         Left            =   1050
         TabIndex        =   22
         Top             =   1560
         Width           =   720
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
         Index           =   3
         Left            =   1170
         TabIndex        =   21
         Top             =   1905
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Visa Date"
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
         Left            =   960
         TabIndex        =   20
         Top             =   2280
         Width           =   765
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
         Index           =   5
         Left            =   960
         TabIndex        =   19
         Top             =   2640
         Width           =   750
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
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
         Left            =   1080
         TabIndex        =   18
         Top             =   3000
         Width           =   675
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SIB  No"
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
         Left            =   1080
         TabIndex        =   17
         Top             =   3360
         Width           =   555
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SIB Date"
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
         Left            =   960
         TabIndex        =   16
         Top             =   3810
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mr "
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
         Left            =   1440
         TabIndex        =   15
         Top             =   4200
         Width           =   270
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
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
            Key             =   "fint"
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
            Picture         =   "frmRelease.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":0AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":0BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":0D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":0EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":13F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":193A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":1A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":1C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelease.frx":1D7E
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
         Caption         =   "Cerificate Of Origin (Textile)"
      End
      Begin VB.Menu performa 
         Caption         =   "&Perform Invoice Form"
      End
      Begin VB.Menu phma 
         Caption         =   "&PHMA Form"
      End
      Begin VB.Menu release 
         Caption         =   "&Release And Undertaking"
      End
      Begin VB.Menu quota 
         Caption         =   "&Quota Transfer Form"
      End
      Begin VB.Menu bank 
         Caption         =   "&Bank Form"
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
Attribute VB_Name = "frmRelease"
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
    
    sSqlMas = "Select * from ReleaseTbl"
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
    Me.txtInvDate = Date
    
End Sub


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

'************************************************************
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
    If mControl.Tag >= 0 And mControl.Tag <= 12 Then
             rsMas.Fields(Val(mControl.Tag)) = mControl
        End If
    Next
    rsMas.Update
End Sub

'*********************************************************
Sub ClearMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 12 Then
             mControl = ""
        End If
    Next
End Sub


'*******************************************************
'************************************************************

Sub EnableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 12 Then
             mControl.Enabled = True
        End If
    Next
End Sub
'************************************************************

Sub DisableMas()
    
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 12 Then
             mControl.Enabled = False
        End If
    Next
End Sub

'***************************************************


Public Sub FillMas()
    'method for saving the record in database
    Dim mControl As Control
    For Each mControl In Me.Controls
    If mControl.Tag >= 0 And mControl.Tag <= 12 Then
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

'***************************************************
'***************************************************
Private Sub txtCat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtQty.SetFocus
    End If
End Sub

Private Sub txtExp_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtExpidNo.SetFocus
    End If
End Sub

Private Sub txtExpidNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtS_No.SetFocus
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

Private Sub txtMr_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.txtMr.SetFocus
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtSIB_No.SetFocus
    End If
End Sub

Private Sub txtS_No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtVisa_No.SetFocus
    End If
End Sub

Private Sub txtSIB_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMr_Date.SetFocus
    End If
End Sub

Private Sub txtSIB_No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtSIB_Date.SetFocus
    End If
End Sub

Private Sub txtVisa_Date_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCat.SetFocus
    End If
End Sub

Private Sub txtVisa_No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtVisa_Date.SetFocus
    End If
End Sub
