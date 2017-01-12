VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Bloomsoft Grey Conversion and Dying Manager"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00000080&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   8940
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dying Lots"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3390
         Picture         =   "MDIForm1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   105
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H80000009&
         Height          =   735
         Left            =   10800
         Picture         =   "MDIForm1.frx":0564
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "YARN ISSUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         Picture         =   "MDIForm1.frx":0971
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cloth Rec"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         Picture         =   "MDIForm1.frx":0EE6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Lot Packing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5460
         Picture         =   "MDIForm1.frx":13C4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   90
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H8000000E&
         Caption         =   "Loom Pay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7500
         Picture         =   "MDIForm1.frx":18A9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   105
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dying Pay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8805
         Picture         =   "MDIForm1.frx":1DEA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   105
         Width           =   1215
      End
   End
   Begin VB.Menu mnuCoding 
      Caption         =   "Data Entry"
      Begin VB.Menu mniClothCoding 
         Caption         =   "Cloth Quality Coding"
      End
      Begin VB.Menu mniDyingCoding 
         Caption         =   "Dying Coding"
      End
      Begin VB.Menu mnufactoryCoding 
         Caption         =   "Factory Coding"
      End
      Begin VB.Menu mnuyarncoding 
         Caption         =   "Yarn Coding"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuissueandrecieving 
      Caption         =   "Issue and Recieving"
      Begin VB.Menu mnuyarnissue 
         Caption         =   "Yarn Bags Issue"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnu_korarecieving 
         Caption         =   "Kora Recieving"
         Shortcut        =   ^K
      End
      Begin VB.Menu mniDyingLotsInformation 
         Caption         =   "Dying Lots Information"
      End
      Begin VB.Menu mnulotpacking 
         Caption         =   "Lot Complete Packing"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnu_voucher 
      Caption         =   "Vouchers Entry"
      Begin VB.Menu mnu_LoomPayment 
         Caption         =   "Looms Payment Voucher"
      End
      Begin VB.Menu mnudyingpayment 
         Caption         =   "Dying Payment Voucher"
      End
   End
   Begin VB.Menu mnu_Reports 
      Caption         =   "Reports"
      Begin VB.Menu mnuyarnreprts 
         Caption         =   "Yarn Issue"
         Begin VB.Menu mnu_yarn1 
            Caption         =   "Issue No's Wise "
         End
         Begin VB.Menu mnu_Yarn2 
            Caption         =   "Periodic Yarn Issue Notes"
         End
         Begin VB.Menu mnu_Yarn3 
            Caption         =   "Factory Wise Yarn Issues"
         End
      End
      Begin VB.Menu mnu_Cloth_Recievings 
         Caption         =   "Cloth Recieving"
         Begin VB.Menu mnu_Cloth1 
            Caption         =   "Reciept No Wise"
         End
         Begin VB.Menu mnu_Cloth2 
            Caption         =   "Periodic Cloth Reciepts "
         End
         Begin VB.Menu mnu_cloth3 
            Caption         =   "Factory Wise Cloth Recieving"
         End
         Begin VB.Menu mnu_Cloth4 
            Caption         =   "Dying Wise Cloth Issues"
         End
      End
      Begin VB.Menu mnu_packingreports 
         Caption         =   "Lot Packings"
         Begin VB.Menu mnu_packing1 
            Caption         =   "Voucher No Wise"
         End
         Begin VB.Menu mnu_Packing2 
            Caption         =   "Periodic Lot Packing Voucher"
         End
         Begin VB.Menu mnu_Packing3 
            Caption         =   "Dying Wise Lot Recieving"
         End
      End
      Begin VB.Menu mnu_LooomPayments 
         Caption         =   "Loom Payments"
         Begin VB.Menu mnu_LoomPayment1 
            Caption         =   "Voucher No Wise"
         End
         Begin VB.Menu mnu_loomPayment2 
            Caption         =   "Periodic Loom Payment Voucher "
         End
         Begin VB.Menu mnu_LoomPayment3 
            Caption         =   "Factory Wise Payment Vouchers"
         End
      End
      Begin VB.Menu mnu_DyingPayments 
         Caption         =   "Dying Payments"
         Begin VB.Menu mnu_dyingPayment1 
            Caption         =   "Voucher No Wise"
         End
         Begin VB.Menu mnu_DyingPayment2 
            Caption         =   "Periodic Dying Payment Voucher"
         End
         Begin VB.Menu mnu_DyingPayment3 
            Caption         =   "Dying Wise Payment Vouchers"
         End
      End
      Begin VB.Menu mniFactoryPaymentIssueRec 
         Caption         =   "Looms Factory Payment, Issue And Receipts"
         Shortcut        =   ^L
      End
      Begin VB.Menu mniLoomPaymentsMadetoLoomFactories 
         Caption         =   "Payments Details Made to Loom Factories"
      End
      Begin VB.Menu mniDyingReport 
         Caption         =   "Dying Report"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuUtiities 
      Caption         =   "Utilities"
      Begin VB.Menu mniDeleteAllData 
         Caption         =   "Delete All Data"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blm As New bloom1
Private Sub Command1_Click()
YarnIssue.Show
End Sub

Private Sub Command2_Click()
ClothRec.Show
End Sub

Private Sub Command3_Click()
LotPacking.Show
End Sub

Private Sub Command4_Click()
vouloompay.Show
End Sub

Private Sub Command5_Click()
voudyingpay.Show
End Sub

Private Sub Command6_Click()
    End
End Sub

Private Sub Command7_Click()
mniDyingLotsInformation_Click
End Sub

Private Sub mniClothCoding_Click()
Cloths.Show
End Sub

Private Sub mniDeleteAllData_Click()
Dim R As VbMsgBoxResult
R = MsgBox("Do You Realy Want to Delete the All Files Data", vbYesNo)
If R = vbYes Then
Dim DB As Database
Dim RS As Recordset
Dim Ssql As String
Dim J As String
J = InputBox("Please Enter Deletion Password")
If J = "QASIM" Then
Set DB = OpenDatabase(blm.pathMain)

Ssql = "Delete from ClothRec"
DB.Execute (Ssql)

'Ssql = "Delete from Cloths"
'DB.Execute (Ssql)

'Ssql = "Delete from DyingChart"
'DB.Execute (Ssql)

'Ssql = "Delete from FactoryChart"
'DB.Execute (Ssql)

Ssql = "Delete from Packing"
DB.Execute (Ssql)

Ssql = "Delete from PaymentDying"
DB.Execute (Ssql)

Ssql = "Delete from PaymentLoom"
DB.Execute (Ssql)

Ssql = "Delete from YarnIssue"
DB.Execute (Ssql)

Ssql = "Delete from Yarns"
DB.Execute (Ssql)

DB.Close
MsgBox "All Files Data has Been Deleted!"
End If
End If
End Sub

Private Sub mniDyingCoding_Click()
Dyingcoding.Show
End Sub

Private Sub mniDyingLotsInformation_Click()
frmLots.Show
End Sub

Private Sub mniDyingReport_Click()
Load vour5
vour5.Text5.Text = 1
vour5.Caption = "Dying Report"
vour5.Check1.Visible = True
vour5.Show
End Sub

Private Sub mniFactoryPaymentIssueRec_Click()
Load vour2
vour2.Caption = mniFactoryPaymentIssueRec.Caption
vour2.Text5.Text = 4
vour2.Show
End Sub

Private Sub mniLoomPaymentsMadetoLoomFactories_Click()
Load vour2
vour2.Caption = mniLoomPaymentsMadetoLoomFactories.Caption
vour2.Text5.Text = 5
vour2.Show

End Sub

Private Sub mnu_Cloth1_Click()
Load vour1
With vour1
    .Text3.Text = 2
    .Caption = "Cloth Recieving Note Preview "
    .Show
End With
End Sub

Private Sub mnu_Cloth2_Click()
Load vour3
With vour3
    .Text1.Text = 2
    .Caption = "Periodic Cloth Reciepts"
    .Show
End With
End Sub

Private Sub mnu_cloth3_Click()
Load vour2
With vour2
    .Text5.Text = 2
    .Caption = "Between Dates Cloth Recieving Factory Wise "
    .Show
End With
End Sub

Private Sub mnu_Cloth4_Click()
Load vour4
With vour4
    .Text5.Text = 1
    .Caption = "Between Dates Cloth Issue Dying Wise "
    .Show
End With
End Sub

Private Sub mnu_dyingPayment1_Click()
Load vour1
With vour1
    .Text3.Text = 5
    .Caption = "Dying Payment Voucher Preview "
    .Show
End With
End Sub

Private Sub mnu_DyingPayment2_Click()
Load vour3
With vour3
    .Text1.Text = 5
    .Caption = "Periodic Dying Payments"
    .Show
End With
End Sub

Private Sub mnu_DyingPayment3_Click()
Load vour4
With vour4
    .Text5.Text = 3
    .Caption = "Between Dates Payment Vouchers Dying Wise "
    .Show
End With
End Sub

Private Sub mnu_korarecieving_Click()
ClothRec.Show
End Sub

Private Sub mnu_loompayment_Click()
vouloompay.Show
End Sub

Private Sub mnu_LoomPayment1_Click()
Load vour1
With vour1
    .Text3.Text = 4
    .Caption = "Loom Payment Voucher Preview "
    .Show
End With
End Sub

Private Sub mnu_loomPayment2_Click()
Load vour3
With vour3
    .Text1.Text = 4
    .Caption = "Periodic Loom Payments"
    .Show
End With
End Sub

Private Sub mnu_LoomPayment3_Click()
Load vour2
With vour2
    .Text5.Text = 3
    .Caption = "Between Dates Loom Payments Factory Wise "
    .Show
End With
End Sub

Private Sub mnu_packing1_Click()
Load vour1
With vour1
    .Text3.Text = 3
    .Caption = "Lot Packing Note Preview "
    .Show
End With
End Sub

Private Sub mnu_Packing2_Click()
Load vour3
With vour3
    .Text1.Text = 3
    .Caption = "Periodic Lot Packings"
    .Show
End With
End Sub

Private Sub mnu_Packing3_Click()
Load vour4
With vour4
    .Text5.Text = 2
    .Caption = "Between Dates Lot Packing Dying Wise "
    .Show
End With
End Sub

Private Sub mnu_yarn1_Click()
Load vour1
With vour1
    .Text3.Text = 1
    .Caption = "Yarn Issue Note Preview "
    .Show
End With
End Sub

Private Sub mnu_Yarn2_Click()
Load vour3
With vour3
    .Text1.Text = 1
    .Caption = "Periodic Yarn ISSUES"
    '.Frame2.Visible = False
    .Show
End With
End Sub

Private Sub mnu_Yarn3_Click()
Load vour2
With vour2
    .Text5.Text = 1
    .Caption = "Between Dates Yarn Issue Factory Wise "
    .Show
End With
End Sub

Private Sub mnudyingpayment_Click()
voudyingpay.Show
End Sub

Private Sub mnufactoryCoding_Click()
FactoryCoding.Show
End Sub

Private Sub mnulotpacking_Click()
LotPacking.Show
End Sub

Private Sub mnuyarncoding_Click()
Yarns.Show
End Sub

Private Sub mnuyarnissue_Click()
YarnIssue.Show
End Sub
