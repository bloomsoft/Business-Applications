VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BloomSoft Accounts & Inventory Management System"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Accounts Ledger"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Aging-Accounts ledger"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Complete Trial Balance"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Closing Stocks of All Items"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Daily Cost Report (Sub-Head Wise)"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Daily Cost Report (Item Wise)"
            Object.Tag             =   ""
            ImageIndex      =   4
            Object.Width           =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Cost Report - 2"
            Object.Tag             =   ""
            ImageIndex      =   5
            Object.Width           =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Day Book"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Banks Position"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Parties Ledger"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1740
      Top             =   1710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   4620
      TabIndex        =   2
      Top             =   390
      Visible         =   0   'False
      Width           =   4680
      Begin VB.Image Image1 
         Height          =   6000
         Left            =   480
         Picture         =   "MDIForm1.frx":27A2
         Top             =   120
         Width           =   7995
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   1
      Top             =   3030
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   291
      _Version        =   327682
      Appearance      =   1
   End
   Begin Crystal.CrystalReport R1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":17827
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":17A01
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":17BDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":183F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":18C0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":18DE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":18FC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":197DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":199B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":19B91
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCoding 
      Caption         =   "Coding"
      Begin VB.Menu mni_chart_ac_setup 
         Caption         =   "Chart of A/c Setup"
         Begin VB.Menu mni_ac_heads 
            Caption         =   "A/c Heads Coding"
         End
         Begin VB.Menu mnu_subhead 
            Caption         =   "A/c Sub Heads Coding"
         End
         Begin VB.Menu mnuaccode 
            Caption         =   "A/c Coding"
         End
      End
      Begin VB.Menu mniItemsCoding 
         Caption         =   "Items Coding"
         Begin VB.Menu mniItemGroups 
            Caption         =   "Item Groups"
         End
         Begin VB.Menu mniSubGroupsCoding 
            Caption         =   "Item Sub Groups"
         End
         Begin VB.Menu mni_Item_Code 
            Caption         =   "Items Coding"
         End
      End
      Begin VB.Menu sepfinance 
         Caption         =   "-"
      End
      Begin VB.Menu mniProfitLossSheet 
         Caption         =   "Profit / Loss Sheet"
      End
      Begin VB.Menu mniBalanceSheet 
         Caption         =   "Balance Sheet"
      End
   End
   Begin VB.Menu mni_Data_ent 
      Caption         =   "&Data Entry"
      Begin VB.Menu mnu_vou_ent 
         Caption         =   "Cash, Bank and Journal Vouchers"
      End
      Begin VB.Menu mni_P_Vou 
         Caption         =   "Purchase Vouchers"
      End
      Begin VB.Menu mni_Sale_Vou 
         Caption         =   "Sales Vouchers"
      End
      Begin VB.Menu mniPurchaseReturnVoucher 
         Caption         =   "Purchase Return Vouchers"
      End
      Begin VB.Menu mniSalesReturnVoucher 
         Caption         =   "Sales Return Vouchers"
      End
   End
   Begin VB.Menu mnuJobs 
      Caption         =   "Jobs Definitions"
      Begin VB.Menu mniPurchaseJobDefinition 
         Caption         =   "Purchase Job Definitions"
      End
      Begin VB.Menu mniSaleJobDefinition 
         Caption         =   "Sale Job Definitions"
      End
   End
   Begin VB.Menu mnuCostSheets 
      Caption         =   "Cost Sheets"
      Begin VB.Menu mniCosSheetEntries 
         Caption         =   "Daily Cost Sheet Entries"
         Begin VB.Menu mniDailyIssueEntry 
            Caption         =   "Daily Issue Voucher (Item Wise)"
         End
         Begin VB.Menu mniDailyIssueVoucherSubHead 
            Caption         =   "Daily Issue Voucher (Sub Head)"
         End
         Begin VB.Menu mniDailyProductionVoucher 
            Caption         =   "Daily Production Vouchers"
         End
         Begin VB.Menu mniDailyExpences 
            Caption         =   "Daily Expenses Vouchers"
         End
         Begin VB.Menu mniDailyArrivals 
            Caption         =   "Daily Arrivals"
         End
         Begin VB.Menu mniDailyDistpaches 
            Caption         =   "Daily Dispatches"
         End
      End
      Begin VB.Menu mniManualCostSheet 
         Caption         =   "Cost Sheet - 2"
      End
   End
   Begin VB.Menu mnuPayroll 
      Caption         =   "Payroll"
      Begin VB.Menu mniEmpAtdEntry 
         Caption         =   "Employees Attendance"
      End
      Begin VB.Menu mniEmployeeOverTime 
         Caption         =   "Employees Over Time"
      End
      Begin VB.Menu mniEmpAdvancesEntry 
         Caption         =   "Employees Advances"
      End
      Begin VB.Menu mniSalariesVoucher 
         Caption         =   "Monthly Salaries Vouchers"
      End
   End
   Begin VB.Menu mni_rep 
      Caption         =   "&Reports"
      Begin VB.Menu mnuchart 
         Caption         =   "Charts"
         Begin VB.Menu mni_chart_ac_rep 
            Caption         =   "Chart of A/c"
         End
         Begin VB.Menu mniItemsChart 
            Caption         =   "Items Chart"
         End
         Begin VB.Menu mni_Add 
            Caption         =   "Address and Phone List"
         End
      End
      Begin VB.Menu mniDailyCostReports 
         Caption         =   "Daily Cost Reports"
         Begin VB.Menu mniDailyReports 
            Caption         =   "Daily Reports"
            Begin VB.Menu mniDailyExpencesReps 
               Caption         =   "Daily Expenses"
            End
            Begin VB.Menu mniDailyArrivalsRep 
               Caption         =   "Daily Arrivals"
            End
            Begin VB.Menu mniDailyDispatchesReps 
               Caption         =   "Daily Dispatches"
            End
            Begin VB.Menu mnudailycombineinformation 
               Caption         =   "Daily Combine Information"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mniProductionReports 
            Caption         =   "Production"
            Begin VB.Menu mniDailyIssue 
               Caption         =   "Daily Issue (Item Wise)"
            End
            Begin VB.Menu mniDailyIssueReportSubHeadWise 
               Caption         =   "Daily Issue (Sub-Head Wise)"
            End
            Begin VB.Menu mniitemwiseissueregister 
               Caption         =   "Single Item Wise Issue Register"
            End
            Begin VB.Menu mniDailyItemWiseProduction 
               Caption         =   "Daily Item Wise Production"
            End
            Begin VB.Menu mniProductionRegister 
               Caption         =   "Daily Production Summary"
            End
            Begin VB.Menu mniMonthlyPurchasesRegister 
               Caption         =   "Monthly Purchases Register"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mniCostAndStockReportSH 
            Caption         =   "Daily Cost Report (Sub Head Wise)"
         End
         Begin VB.Menu mniCostAndStockReport 
            Caption         =   "Daily Cost Report (Item Wise)"
         End
      End
      Begin VB.Menu mni_ledger_rep 
         Caption         =   "Journal Ledger"
         Begin VB.Menu mni_ac_ledger 
            Caption         =   "Accounts Ledger"
            Visible         =   0   'False
         End
         Begin VB.Menu mni_ref_Ledger 
            Caption         =   "Ref. No. Wise Ledger"
            Visible         =   0   'False
         End
         Begin VB.Menu mni_comp_ac_ledger 
            Caption         =   "Accounts Ledger"
         End
         Begin VB.Menu mniAgingLedgerofAnAccount 
            Caption         =   "Aging-Accounts Ledger"
         End
         Begin VB.Menu mni_comb_ledger 
            Caption         =   "All Ledgers"
         End
      End
      Begin VB.Menu mniPurchaseJobReports 
         Caption         =   "Jobs"
         Begin VB.Menu mniPurchaseJobLedger 
            Caption         =   "Purchase Job Ledger"
         End
         Begin VB.Menu mniPurchaseJobSummary 
            Caption         =   "Purchase Job Summary"
            Visible         =   0   'False
         End
         Begin VB.Menu mniSaleJobLedger 
            Caption         =   "Sale Job Ledger"
         End
      End
      Begin VB.Menu mniPurchaseReps 
         Caption         =   "Purchases"
         Begin VB.Menu mniPurchasesPartyWise 
            Caption         =   "Purchases Party Wise"
         End
         Begin VB.Menu mniPurchaseItemWise 
            Caption         =   "Purchases Item Wise"
         End
         Begin VB.Menu mniPurchasesPartyItemWise 
            Caption         =   "Purchases Party and Item Wise"
         End
         Begin VB.Menu mniPurchasesSummary 
            Caption         =   "Purchases Summary"
         End
         Begin VB.Menu mniPBBook 
            Caption         =   "Over All Purchases"
         End
         Begin VB.Menu mniOverAllGSTInvoicePV 
            Caption         =   "Over All GST Invoices (PV) Party Wise"
         End
      End
      Begin VB.Menu mniSalesReps 
         Caption         =   "Sales"
         Begin VB.Menu mniPartyWiseSalesRep 
            Caption         =   "Party Wise Sales"
         End
         Begin VB.Menu mniItemWiseSales 
            Caption         =   "Item Wise Sales"
         End
         Begin VB.Menu mniPartyItemWiseSales 
            Caption         =   "Party and Item Wise Sales"
         End
         Begin VB.Menu mniSalesSummary 
            Caption         =   "Sales Summary"
         End
         Begin VB.Menu mniSalesBook 
            Caption         =   "Over All Sales"
         End
         Begin VB.Menu mniOverAllGSTInvoicesSVPartyWise 
            Caption         =   "Over All GST Invoices (SV) Party Wise"
         End
      End
      Begin VB.Menu mniInventoryAndStockReps 
         Caption         =   "Inventory and Stocks"
         Begin VB.Menu mniItemWiseLedgerNew 
            Caption         =   "Item Wise Ledger"
         End
         Begin VB.Menu mniClosingStocks 
            Caption         =   "Closing Stocks of All Items"
         End
         Begin VB.Menu mniClosingStocksofSelectiveHeads 
            Caption         =   "Closing Stocks of Selective Heads"
         End
         Begin VB.Menu mniClosingStocksofSelectiveSubHeads 
            Caption         =   "Closing Stocks of Selective Sub-Heads"
         End
      End
      Begin VB.Menu mniPeriodicVouchers 
         Caption         =   "Classified Vouchers"
         Begin VB.Menu mniJournalVoucherPeriod 
            Caption         =   "Journal Vouchers"
         End
         Begin VB.Menu mniCashVouchersPeriodic 
            Caption         =   "Cash Vouchers"
         End
         Begin VB.Menu mniBankVouchers 
            Caption         =   "Bank Vouchers"
         End
         Begin VB.Menu sepperiodic 
            Caption         =   "-"
         End
         Begin VB.Menu mniPurchaseVouchersPeriodic 
            Caption         =   "Purchase Vouchers"
         End
         Begin VB.Menu mniSaleVouchersPeriodic 
            Caption         =   "Sales Vouchers"
         End
         Begin VB.Menu mniPurchaseReturnPeriodic 
            Caption         =   "Purchase Return Vouchers"
         End
         Begin VB.Menu mniSalesReturnVouchersPeriodic 
            Caption         =   "Sales Return Vouchers"
         End
         Begin VB.Menu sepabc 
            Caption         =   "-"
         End
         Begin VB.Menu mniIssueVouchersPeriodic 
            Caption         =   "Issue Vouchers (Item Wise)"
         End
         Begin VB.Menu mniDailyIssueHeadWiseRep 
            Caption         =   "Issue Vouchers (Sub Head Wise)"
         End
         Begin VB.Menu mniProductionVouchers 
            Caption         =   "Production Vouchers"
         End
         Begin VB.Menu mniStockTransferVouchers 
            Caption         =   "Stock Transfer Vouchers"
         End
         Begin VB.Menu sepissue 
            Caption         =   "-"
         End
         Begin VB.Menu mniAdvanceVouchersPeriodic 
            Caption         =   "Advances Vouchers"
         End
         Begin VB.Menu mniSalariesVoucherPeriodic 
            Caption         =   "Salaries Vouchers"
         End
      End
      Begin VB.Menu mni_vou_print 
         Caption         =   "Vouchers Printing"
         Visible         =   0   'False
         Begin VB.Menu mni_jou_vou_rep 
            Caption         =   "Journal Vouchers"
         End
         Begin VB.Menu mni_bank_vou_print 
            Caption         =   "Bank Vouchers"
         End
         Begin VB.Menu mni_cash_vou_print 
            Caption         =   "Cash Vouchers"
         End
         Begin VB.Menu mni_SJV 
            Caption         =   "Sale Journal Vouchers"
         End
         Begin VB.Menu mni_PJV 
            Caption         =   "Purchase Journal Vouchers"
         End
         Begin VB.Menu mni_CJV 
            Caption         =   "Consumption Voucher"
            Visible         =   0   'False
         End
         Begin VB.Menu mniPurchaseJobDefine 
            Caption         =   "Purchase Job Definitions"
         End
         Begin VB.Menu mniSaleJobDefine 
            Caption         =   "Sale Job Definitions"
         End
      End
      Begin VB.Menu mni_books 
         Caption         =   "Books"
         Visible         =   0   'False
         Begin VB.Menu mni_journal_Book 
            Caption         =   "Journal Book"
         End
         Begin VB.Menu mni_Bank_book 
            Caption         =   "Bank Book"
         End
         Begin VB.Menu mni_cash_book 
            Caption         =   "Cash Book"
         End
      End
      Begin VB.Menu mni_pay_lst 
         Caption         =   "Payables List"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_rec_lst 
         Caption         =   "Receivable List"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_trial 
         Caption         =   "Trial Balance"
         Begin VB.Menu mni_open_trial 
            Caption         =   "Opening Trial Balance"
         End
         Begin VB.Menu mni_simple_trial 
            Caption         =   "Simple Trial Balance (Without Groups)"
            Visible         =   0   'False
         End
         Begin VB.Menu mni_dates_trial 
            Caption         =   "Complete Trial Balance"
         End
         Begin VB.Menu mni_group_trial 
            Caption         =   "Groups Trial Balance"
            Visible         =   0   'False
         End
         Begin VB.Menu mni_Sub_H_Totals_Tria 
            Caption         =   "Groups Trail Balance (Sub Head Totals)"
            Visible         =   0   'False
         End
         Begin VB.Menu mniSelHeadsTrialBalance 
            Caption         =   "Selective Heads Trial Balance"
         End
         Begin VB.Menu mniSelectiveSubHeadsTrialBalance 
            Caption         =   "Selective Sub Heads Trial Balance"
         End
      End
      Begin VB.Menu mni_recable 
         Caption         =   "Recievables"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_payable 
         Caption         =   "Payables"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_Stock 
         Caption         =   "Stocks Over All to An End Date"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_Period_Stock 
         Caption         =   "Periodic Stock Statement"
         Visible         =   0   'False
      End
      Begin VB.Menu mniDayBook 
         Caption         =   "Day Book"
      End
      Begin VB.Menu mni_CashBook 
         Caption         =   "Cash Flow Statement"
      End
      Begin VB.Menu mniPayrollReps 
         Caption         =   "Payroll"
         Begin VB.Menu mnuempchart 
            Caption         =   "Chart of Employees"
         End
         Begin VB.Menu mniEmpAtdregisterRep 
            Caption         =   "Employees Attendance Register"
         End
         Begin VB.Menu mniDailyOverTimeSheet 
            Caption         =   "Over Time Sheet"
         End
         Begin VB.Menu mniEmployeesAdvacensRep 
            Caption         =   "Employees Advances"
            Visible         =   0   'False
         End
         Begin VB.Menu mniAdvancesBook 
            Caption         =   "Advances Book"
         End
         Begin VB.Menu mniSalariesSheet 
            Caption         =   "Monthly Salaries Sheet"
         End
      End
      Begin VB.Menu mniCostReports 
         Caption         =   "Cost Reports"
         Visible         =   0   'False
         Begin VB.Menu mniNotestotheAccounts 
            Caption         =   "Notes to the Accounts"
         End
         Begin VB.Menu mniPeriodicCostSheet 
            Caption         =   "Periodic Cost Sheet"
         End
      End
      Begin VB.Menu mni_Finance_State 
         Caption         =   "Financial Statements"
         Visible         =   0   'False
      End
      Begin VB.Menu mniAgingStatement 
         Caption         =   "Aging Statement"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mni_user 
      Caption         =   "Settings"
      Begin VB.Menu mniAlarmSettings 
         Caption         =   "Alarm Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_Set_OrgName 
         Caption         =   "Set or Change Organization Name"
      End
      Begin VB.Menu mni_create_change 
         Caption         =   "Create or Change User Info"
      End
      Begin VB.Menu mniDeleteAllData 
         Caption         =   "Delete All the Data"
      End
      Begin VB.Menu mniChangeWallPaper 
         Caption         =   "Change Wallpaper"
      End
      Begin VB.Menu mniYearClosing 
         Caption         =   "Year Closing"
      End
   End
   Begin VB.Menu mnuAboutus 
      Caption         =   "About Us"
   End
   Begin VB.Menu mni_quit 
      Caption         =   "&Quit"
      Begin VB.Menu mni_exit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mni_cancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blmr As New bloom_r
Private Sub SaveWallPaper(W As String)
On Error GoTo ErrHand
Dim FS As New FileSystemObject
Dim TR As TextStream
Set TR = FS.CreateTextFile(App.path & "\WallPaper.txt", True)
    TR.Write W
TR.Close
Set FS = Nothing
Image1.Picture = LoadPicture(W)
DoEvents
SetWallPaperSettings
Me.WindowState = 1
DoEvents
Me.WindowState = 2
Exit Sub
ErrHand:
MsgBox "There is an Error in Wall Paper Settings " & Err.Description

End Sub

Private Sub ShowWallPaper()
On Error GoTo ErrHand
Dim FS As New FileSystemObject
Dim TR As TextStream
Set TR = FS.OpenTextFile(App.path & "\WallPaper.txt", ForReading)
    CurrentWallPaper = TR.ReadAll
TR.Close
Set FS = Nothing
Image1.Picture = LoadPicture(CurrentWallPaper)
Exit Sub
ErrHand:
    CurrentWallPaper = App.path & "\AllahAkbar.jpg"
    Image1.Picture = LoadPicture(CurrentWallPaper)
    
End Sub

Private Sub SetWallPaperSettings()
    Picture2.AutoRedraw = True
    Picture2.Cls
    Picture2.Height = Me.ScaleHeight + Picture2.Height - Picture2.ScaleHeight


    Select Case Mode
        Case 0 'centered
        Picture2.PaintPicture Image1.Picture, (Me.ScaleWidth - Image1.Width) / 2, (Me.ScaleHeight - Image1.Height) / 2
        Case 1 'stretched
        If Me.WindowState <> 1 Then Picture2.PaintPicture Image1.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
        Case 2 'tiled
        Dim I As Integer, j As Integer


        For I = 0 To Screen.Height Step Image1.Height


            For j = 0 To Screen.Width Step Image1.Width
                Picture2.PaintPicture Image1, j, I, Image1.Width, Image1.Height
            Next
        Next
    End Select
Me.Picture = Picture2.Image
Picture2.AutoRedraw = False

End Sub

Private Sub MDIForm_Load()
ShowWallPaper
GetDates
Load frmAlarmer
End Sub

Private Sub MDIForm_Resize()
SetWallPaperSettings
End Sub

Private Sub mni_ac_heads_Click()
Load head1
head1.Show
End Sub

Private Sub mni_ac_ledger_Click()
Load ledger
ledger.Text1.Text = 1
ledger.Show

End Sub

Private Sub mni_Add_Click()
Load ac1
ac1.Text1.Text = 3
ac1.Caption = "Addresses List"
ac1.Show

End Sub

Private Sub mni_Bank_book_Click()
Load book
book.Text1.Text = 2
book.Show

End Sub

Private Sub mni_bank_vou_print_Click()
Load vour
vour.Text2.Text = 2
vour.Show
End Sub

Private Sub mni_cash_book_Click()
Load book
book.Text1.Text = 3
book.Show

End Sub

Private Sub mni_cash_vou_print_Click()
Load vour
vour.Text2.Text = 3
vour.Show

End Sub

Private Sub mni_CashBook_Click()
Load ledger
ledger.Text1.Text = 4
ledger.Label1.Visible = True
ledger.Label2.Visible = True
ledger.date1.Visible = True
ledger.Date2.Visible = True

ledger.Caption = "Cash Book"
ledger.Show

End Sub

Private Sub mni_chart_ac_rep_Click()
Load ac1
ac1.Text1.Text = 1
ac1.Show
End Sub

Private Sub mni_CJV_Click()
Load vour
vour.Text2.Text = 6
vour.Caption = "Consumption Voucher Print"
vour.Show
End Sub

Private Sub mni_comb_ledger_Click()
Load ledger
ledger.Text1.Text = 3
ledger.Label1.Visible = True
ledger.Label2.Visible = True
ledger.Label3.Visible = False
ledger.Text3.Visible = False
ledger.Label5.Visible = False
ledger.Text2.Visible = False
ledger.date1.Visible = True
ledger.Date2.Visible = True
ledger.Show

End Sub

Private Sub mni_comp_ac_ledger_Click()
Load ledger
ledger.Text1.Text = 2
ledger.Label1.Visible = True
ledger.Label2.Visible = True
ledger.date1.Visible = True
ledger.Date2.Visible = True
ledger.Show

End Sub

Private Sub mni_Consume_Vou_Click()
Load vou1C
vou1C.Show
End Sub

Private Sub mni_create_change_Click()
Load login
login.Show
End Sub

Private Sub mni_dates_trial_Click()
Load trial
trial.Caption = "Trial Balance With Groups (Between Dates)"
trial.Date2.Visible = True
trial.Label3.Visible = True
trial.Text1.Text = 7
trial.Show
End Sub

Private Sub mni_exit_Click()
End
End Sub

Private Sub mni_Finance_State_Click()
Load trial
trial.Text1.Text = 10
trial.Caption = "Financial Statements"
trial.Show
End Sub

Private Sub mni_group_trial_Click()
Load trial
trial.Caption = "Groups Trial Balance"
trial.Text1.Text = 1
trial.Visible = True
trial.Show
End Sub

Private Sub mni_itm_cod_Click()
Load item1
item1.Show
End Sub

Private Sub mni_itm_lst_Click()
Load ac1
ac1.Text1.Text = 2
ac1.Show
End Sub

Private Sub mni_Item_Code_Click()
Load Items
Items.Show
End Sub

Private Sub mni_Item_Ledger_Click()
Load itmledger
itmledger.Text1.Text = 1
itmledger.Show
End Sub

Private Sub mni_jou_vou_rep_Click()
Load vour
vour.Text2.Text = 1
vour.Show
End Sub

Private Sub mni_journal_Book_Click()
Load book
book.Text1.Text = 1
book.Show
End Sub

Private Sub mni_New_App_Create_Click()
newapp.Show
End Sub

Private Sub mni_open_trial_Click()
Load trial
trial.Caption = "Opening Trial Balance"
trial.Text1.Text = 5
trial.Label2.Visible = False
trial.date1.Visible = False
trial.Show
End Sub

Private Sub mni_p_l_Notes_Click()
Load Setup1
Setup1.Caption = "Profit & Loss Statements"
Setup1.Show
End Sub

Private Sub mni_P_Vou_Click()
Load vou1P
vou1P.Show
End Sub

Private Sub mni_pay_lst_Click()
Load trial
trial.Text1.Text = 3
trial.Caption = "Payables Report"
trial.Show
End Sub

Private Sub mni_purc_jou_vou_Click()
Load purc
purc.Show
End Sub

Private Sub mni_purchase_book_Click()
Load book
book.Text1.Text = 5
book.Show

End Sub

Private Sub mni_purchase_vou_print_Click()
Load vour
vour.Text2.Text = 5
vour.Show

End Sub

Private Sub mni_payable_Click()
Load trial
trial.Text1.Text = 9
trial.Caption = "Payables Report"
trial.Show

End Sub

Private Sub mni_Period_Stock_Click()
Load trial
trial.Caption = "Periodic Stock Statement"
trial.Text1.Text = 11
trial.Label3.Visible = True
trial.Date2.Visible = True
trial.Show

End Sub

Private Sub mni_PJV_Click()
Load vour
vour.Text2.Text = 5
vour.Caption = "Purchase Journal Voucher Print"
vour.Show
End Sub

Private Sub mni_rec_lst_Click()
Load trial
trial.Text1.Text = 4
trial.Caption = "Recievables Report"
trial.Show
End Sub

Private Sub mni_sale_book_Click()
Load book
book.Text1.Text = 4
book.Show

End Sub

Private Sub mni_sale_jou_vou_Click()
Load sale
sale.Show
End Sub

Private Sub mni_sale_vou_print_Click()
Load vour
vour.Text2.Text = 4
vour.Show

End Sub

Private Sub mni_recable_Click()
Load trial
trial.Text1.Text = 8
trial.Caption = "Recievables Report"
trial.Show
End Sub

Private Sub mni_ref_Ledger_Click()
ICTLedger.Show
End Sub

Private Sub mni_Sale_Vou_Click()
Load vou1S
vou1S.Show
End Sub

Private Sub mni_Set_OrgName_Click()
orgname.Show
End Sub

Private Sub mni_simple_trial_Click()
Load trial
trial.Caption = "Simple Trial Balance (Without Groups)"
trial.Text1.Text = 6
trial.Show

End Sub

Private Sub mni_SJV_Click()
Load vour
vour.Text2.Text = 4
vour.Caption = "Sale Journal Voucher Print"
vour.Show

End Sub

Private Sub mni_Stock_Click()
Load Stocks
Stocks.Show
End Sub

Private Sub mni_Sub_H_Totals_Tria_Click()
Load trial
trial.Caption = "Groups Trial Balance"
trial.Text1.Text = 1
trial.Check1.Visible = True
trial.Show

End Sub

Private Sub mniAdvancesBook_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 14
PeriodicRep.Frame2.Visible = True
PeriodicRep.Caption = "Periodic Advance Book"
PeriodicRep.Show

End Sub

Private Sub mniAdvanceVouchersPeriodic_Click()
SelectedVType = 11
SelectedReportTitle = "Advances Vouchers"
DoPeriodicRep "Advances Vouchers"
End Sub

Private Sub mniAgingLedgerofAnAccount_Click()
Load ledger
ledger.Text1.Text = 7
ledger.Label1.Visible = True
ledger.Label2.Visible = True
ledger.date1.Visible = True
ledger.Date2.Visible = True
ledger.Label6.Visible = True
ledger.Text4.Visible = True
ledger.Show

End Sub

Private Sub mniAgingStatement_Click()
Load Stocks
Stocks.Text1.Text = 4
Stocks.Caption = "Aging Statement"
Stocks.Show

End Sub

Private Sub mniAlarmSettings_Click()
frmAlarm.Show
End Sub

Private Sub mniBalanceSheet_Click()
frmManualBS.Show
End Sub

Private Sub mniBankVouchers_Click()
SelectedReportTitle = "Bank Vouchers"
SelectedVType = 2
DoPeriodicRep "Bank Vouchers"
End Sub

Private Sub mniCashVouchersPeriodic_Click()
SelectedReportTitle = "Cash Vouchers"
SelectedVType = 3
DoPeriodicRep "Cash Vouchers"
End Sub

Private Sub mniChangeWallPaper_Click()
With CommonDialog1
    .Filter = "All Images (BMP,JPG,GIF)|*.bmp;*.gif;*.jpg"
    .FilterIndex = 1
    .ShowOpen
    If .FileName <> "" Then
        SaveWallPaper .FileName
        
    End If
End With

End Sub

Private Sub mniClosingStocks_Click()
Load trial
trial.Text1.Text = 14
trial.Caption = "Closing Stocks of All Items"
trial.Date2.Visible = True
trial.Label3.Visible = True

trial.Show

End Sub

Private Sub mniClosingStocksofSelectiveHeads_Click()
Load SelectiveITHeads
SelectiveITHeads.Caption = "Selective Heads Stocks"
SelectiveITHeads.Text1.Text = 1
SelectiveITHeads.Label3.Visible = True
SelectiveITHeads.DTPicker2.Visible = True

SelectiveITHeads.Show

End Sub

Private Sub mniClosingStocksofSelectiveSubHeads_Click()
Load SelectiveITSubHeads
SelectiveITSubHeads.Caption = "Selective Sub-Head Stocks"
SelectiveITSubHeads.Text1.Text = 1
SelectiveITSubHeads.Label4.Visible = True
SelectiveITSubHeads.DTPicker2.Visible = True

SelectiveITSubHeads.Show

End Sub

Private Sub mniCostAndStockReport_Click()
Load frmCostsheetrep
frmCostsheetrep.Text2.Text = 1
frmCostsheetrep.Caption = "Cost And Stock Report"
frmCostsheetrep.Show
End Sub

Private Sub mniCostAndStockReportSH_Click()
Load frmCostsheetrep
frmCostsheetrep.Text2.Text = 2
frmCostsheetrep.Caption = "Cost And Stock Report(Sub Head Wise)"
frmCostsheetrep.Show

End Sub

Private Sub mniDailyArrivals_Click()
Arrivals.Show
End Sub

Private Sub mniDailyArrivalsRep_Click()
'Load book
'book.Text1.Text = 7
'book.Frame2.Visible = True
'book.Show
Load PeriodicRep
PeriodicRep.Text1.Text = 19
PeriodicRep.Frame1.Visible = True
PeriodicRep.Show

End Sub

Private Sub mniDailyDispatchesReps_Click()
'Load book
'book.Text1.Text = 8
'book.Frame2.Visible = True
'book.Show
Load PeriodicRep
PeriodicRep.Text1.Text = 20
PeriodicRep.Frame1.Visible = True
PeriodicRep.Show

End Sub

Private Sub mniDailyDistpaches_Click()
Dispatches.Show
End Sub

Private Sub mniDailyExpences_Click()
Expences.Show
End Sub

Private Sub mniDailyExpencesReps_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 17
PeriodicRep.Caption = "Daily Expences"
PeriodicRep.Frame1.Visible = True
PeriodicRep.Show

End Sub

Private Sub mniDailyissue_Click()
'Load book
'book.Text1.Text = 6
'book.Frame2.Visible = True
'book.Caption = "Daily Issue Report"
'book.Show
Load PeriodicRep
PeriodicRep.Text1.Text = 18
PeriodicRep.Frame1.Visible = True
PeriodicRep.Caption = "Daily Issue Report (Item Wise)"
PeriodicRep.Show

End Sub

Private Sub mniDailyIssueEntry_Click()
Issue.Show
End Sub

Private Sub mniDailyIssueHeadWiseRep_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 16
PeriodicRep.Frame1.Visible = True
PeriodicRep.Show

End Sub

Private Sub mniDailyIssueReportSubHeadWise_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 16
PeriodicRep.Frame1.Visible = True
PeriodicRep.Caption = "Daily Issue Report (Sub Head Wise)"
PeriodicRep.Show

End Sub

Private Sub mniDailyIssueVoucherSubHead_Click()
IssueSH.Show
End Sub

Private Sub mniDailyItemWiseProduction_Click()
Load itmledger
itmledger.Text1.Text = 6
itmledger.Label5.Visible = True
itmledger.Text4.Visible = True
itmledger.Show

End Sub

Private Sub mniDailyOverTimeSheet_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 10
PeriodicRep.Frame2.Visible = True
PeriodicRep.Caption = "Periodic Over Time / Short Time"
PeriodicRep.Show

End Sub

Private Sub mniDailyProductionVoucher_Click()
Production.Show
End Sub

Private Sub mniDayBook_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 11
PeriodicRep.Caption = "Day Book"
PeriodicRep.Show

End Sub

Private Sub mniDeleteAllData_Click()
Dim Ssql As String
Dim DBM As Database

Dim Result As VbMsgBoxResult

Result = MsgBox("Do You Realy Want to Delete All the Data", vbYesNo)
If Result = vbYes Then
    Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
    Ssql = "Delete from VouMST"
    DBM.Execute Ssql
    Ssql = "Delete from VouDTL"
    DBM.Execute Ssql
    Ssql = "Delete from SalVoucher"
    DBM.Execute Ssql
    Ssql = "Delete from Sales"
    DBM.Execute Ssql
    Ssql = "Delete from Purchase"
    DBM.Execute Ssql
    Ssql = "Delete from SaleJob"
    DBM.Execute Ssql
    Ssql = "Delete from PurJob"
    DBM.Execute Ssql
    'Ssql = "Delete from Acchart"
    'DBM.Execute Ssql
    Ssql = "Delete from Consume"
    DBM.Execute Ssql
    Ssql = "Delete from EmpATD"
    DBM.Execute Ssql
    'Ssql = "Delete from Heads"
    'DBM.Execute Ssql
    Ssql = "Delete from Issue"
    DBM.Execute Ssql
    'Ssql = "Delete from Items"
    'DBM.Execute Ssql
    Ssql = "Delete from Production"
    DBM.Execute Ssql
    
    Ssql = "Delete from OverTime"
    DBM.Execute Ssql
    
    DBM.Close
End If

End Sub

Private Sub mniEmpAdvancesEntry_Click()
EmpAdv.Show
End Sub

Private Sub mniEmpAtdEntry_Click()
EmpAtd.Show
End Sub

Private Sub mniEmpAtdregisterRep_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 3
PeriodicRep.Caption = "Employees Attendance Register"
PeriodicRep.Show
End Sub

Private Sub mniEmployeeOverTime_Click()
EmpOverTime.Show
End Sub

Private Sub mniEmployeesAdvacensRep_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 5
PeriodicRep.Caption = "Employee Advances Book"
PeriodicRep.Show

End Sub

Private Sub mniIssueVouchersPeriodic_Click()
SelectedVType = 20
DoPeriodicRep "Issue Vouchers"
End Sub

Private Sub mniItemGroups_Click()
Groups1.Show
End Sub

Private Sub mniItemsChart_Click()
Load ac1
ac1.Text1.Text = 2
ac1.Show

End Sub

Private Sub mniitemwiseissueregister_Click()
Load itmledger
itmledger.Text1.Text = 2
itmledger.Text4.Visible = True
itmledger.Label5.Visible = True
itmledger.Caption = "Item Wise Monthly Issue Register"
itmledger.Show

End Sub

Private Sub mniItemWiseLedgerNew_Click()
Load itmledger
itmledger.Text1.Text = 5
itmledger.Show

End Sub

Private Sub mniItemWiseSales_Click()
Load itmledger
itmledger.Text1.Text = 4
itmledger.Show

End Sub

Private Sub mniJournalVoucherPeriod_Click()
SelectedReportTitle = "Journal Vouchers"
SelectedVType = 1
DoPeriodicRep "Journal Vouchers"
End Sub
Private Sub DoPeriodicRep(S As String)
Load PeriodicRep
PeriodicRep.Text1.Text = 15
PeriodicRep.Caption = S
PeriodicRep.Show

End Sub

Private Sub mniManualCostSheet_Click()
frmManualCostRep.Show
End Sub

Private Sub mniMillsEntry_Click()
Mill.Show
End Sub

Private Sub mniMonthlyPurchasesRegister_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 7
PeriodicRep.Caption = "Monthly Purchases Register"
PeriodicRep.Show

End Sub

Private Sub mniNotestotheAccounts_Click()
Load trial
trial.Caption = "Notes to the Accounts"
trial.Text1.Text = 12
trial.Show
End Sub

Private Sub mniOverAllGSTInvoicePV_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 12
PeriodicRep.Caption = "Over All GST Invoices (PV)"
PeriodicRep.Show

End Sub

Private Sub mniOverAllGSTInvoicesSVPartyWise_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 13
PeriodicRep.Caption = "Over All GST Invoices (SV)"
PeriodicRep.Show

End Sub

Private Sub mniPartyItemWiseSales_Click()
Load PartyItemledger
PartyItemledger.Label1.Visible = True
PartyItemledger.Label2.Visible = True
PartyItemledger.date1.Visible = True
PartyItemledger.Date2.Visible = True
PartyItemledger.Text1.Text = 2
PartyItemledger.Caption = "Party Item Wise Sales"
PartyItemledger.Show

End Sub

Private Sub mniPartyWiseSalesRep_Click()
Load ledger
ledger.Label1.Visible = True
ledger.Label2.Visible = True
ledger.date1.Visible = True
ledger.Date2.Visible = True
ledger.Check1.Visible = True
ledger.Check2.Visible = True
ledger.Text5.Visible = True
ledger.Text1.Text = 6
ledger.Caption = "Party Wise Sales"
ledger.Show

End Sub

Private Sub mniPBBook_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 1
PeriodicRep.Caption = "Periodic Purchase Ledger"
PeriodicRep.Show
End Sub

Private Sub mniPeriodicCostSheet_Click()
Load trial
trial.Caption = "Cost Sheet (Periodic)"
trial.Date2.Visible = True
trial.Label3.Visible = True
trial.Text1.Text = 13
trial.Show

End Sub

Private Sub mniProductionRegister_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 21
PeriodicRep.Caption = "Monthly Production Register"
PeriodicRep.Frame1.Visible = True
PeriodicRep.Show

End Sub

Private Sub mniProductionVouchers_Click()
SelectedVType = 21
DoPeriodicRep "Production Vouchers"
End Sub

Private Sub mniProfitLossSheet_Click()
frmManualPL.Show
End Sub

Private Sub mniPurchaseItemWise_Click()
Load itmledger
itmledger.Text1.Text = 3
itmledger.Show

End Sub

Private Sub mniPurchaseJobDefine_Click()
Load vour
vour.Text2.Text = 8
vour.Text3.Visible = False
vour.Label2.Visible = False
vour.Caption = "Purchase Job Print"
vour.Show

End Sub

Private Sub mniPurchaseJobDefinition_Click()
PurJob.Show
End Sub

Private Sub mniPurchaseJobLedger_Click()
Load ContractRep
ContractRep.Text1.Text = 1
ContractRep.Caption = "Purchase Job Ledger"
ContractRep.Show
End Sub

Private Sub mniPurchaseJobSummary_Click()
'Load Stocks
'Stocks.Text1.Text = 2
'Stocks.Caption = "Purchase Jobs Summary"
'Stocks.Show
Load ac1
ac1.Text1.Text = 4
ac1.Caption = "Purchase Job Summary"
ac1.Show
End Sub

Private Sub mniPurchaseReturnPeriodic_Click()
SelectedReportTitle = "Purchase Return Vouchers"
SelectedVType = 16
DoPeriodicRep "Purchase Return Vouchers"
End Sub

Private Sub mniPurchaseReturnVoucher_Click()
vou1PR.Show
End Sub

Private Sub mniPurchasesPartyItemWise_Click()

Load PartyItemledger
PartyItemledger.Label1.Visible = True
PartyItemledger.Label2.Visible = True
PartyItemledger.date1.Visible = True
PartyItemledger.Date2.Visible = True

PartyItemledger.Text1.Text = 1
PartyItemledger.Caption = "Party Item Wise Purchases"
PartyItemledger.Show

End Sub

Private Sub mniPurchasesPartyWise_Click()
Load ledger
ledger.Label1.Visible = True
ledger.Label2.Visible = True
ledger.date1.Visible = True
ledger.Date2.Visible = True
ledger.Check1.Visible = True
ledger.Check2.Visible = True
ledger.Text5.Visible = True
ledger.Text1.Text = 5
ledger.Caption = "Party Wise Purchases"
ledger.Show

End Sub

Private Sub mniPurchasesSummary_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 8
PeriodicRep.Caption = "Purchases Summary"
PeriodicRep.Show

End Sub

Private Sub mniPurchaseVouchersPeriodic_Click()
SelectedVType = 4
DoPeriodicRep "Purchase Vouchers"
End Sub

Private Sub mniSalariesSheet_Click()
'Load PeriodicRep
'PeriodicRep.Text1.Text = 4
'PeriodicRep.Caption = "Monhtly Salary Sheet"
'PeriodicRep.Show
Load SelectiveSubHeads
SelectiveSubHeads.Label4.Visible = True
SelectiveSubHeads.DTPicker2.Visible = True
SelectiveSubHeads.Caption = "Monhtly Salary Sheet"
SelectiveSubHeads.Text1.Text = 2
SelectiveSubHeads.Show

End Sub

Private Sub mniSalariesVoucher_Click()
SalVoucher.Show
End Sub

Private Sub mniSalariesVoucherPeriodic_Click()
SelectedVType = 15
DoPeriodicRep "Salaries Vouchers"
End Sub

Private Sub mniSaleJobDefine_Click()
Load vour
vour.Text2.Text = 7
vour.Text3.Visible = False
vour.Label2.Visible = False
vour.Caption = "Sale Job Print"
vour.Show

End Sub

Private Sub mniSaleJobDefinition_Click()
SaleJob.Show
End Sub

Private Sub mniSaleJobLedger_Click()
Load ContractRep
ContractRep.Text1.Text = 2
ContractRep.Caption = "Sale Job Ledger"
ContractRep.Show

End Sub

Private Sub mniSaleJobSummary_Click()
Load Stocks
Stocks.Text1.Text = 3
Stocks.Caption = "Sale Jobs Summary"
Stocks.Show

End Sub

Private Sub mniSalesBook_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 2
PeriodicRep.Caption = "Periodic Sales Ledger"
PeriodicRep.Show

End Sub

Private Sub mniSalesReturnVoucher_Click()
vou1SR.Show
End Sub

Private Sub mniSalesReturnVouchersPeriodic_Click()
SelectedVType = 17
SelectedReportTitle = "Sales Return Vouchers"
DoPeriodicRep "Sales Return Vouchers"
End Sub

Private Sub mniSalesSummary_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 9
PeriodicRep.Caption = "Sales Summary"
PeriodicRep.Show

End Sub

Private Sub mniSaleVouchersPeriodic_Click()
SelectedVType = 5
DoPeriodicRep "Sales Vouchers"
End Sub

Private Sub mniSelectiveSubHeadsTrialBalance_Click()
Load SelectiveSubHeads
SelectiveSubHeads.Label4.Visible = True
SelectiveSubHeads.DTPicker2.Visible = True
SelectiveSubHeads.Caption = "Selective Sub-Head Trial Balance"
SelectiveSubHeads.Text1.Text = 1
SelectiveSubHeads.Show

End Sub

Private Sub mniSelHeadsTrialBalance_Click()
Load SelectiveHeads
SelectiveHeads.Caption = "Selective Head Trial Balance"
SelectiveHeads.Label3.Visible = True
SelectiveHeads.DTPicker2.Visible = True
SelectiveHeads.Text1.Text = 1
SelectiveHeads.Show
End Sub

Private Sub mniStockTransferVouchers_Click()
SelectedVType = 18
DoPeriodicRep "Stock Transfer Vouchers"
End Sub

Private Sub mniSubGroupsCoding_Click()
SubGroups1.Show
End Sub

Private Sub mniYearClosing_Click()
YearClose.Show
End Sub

Private Sub mnu_subhead_Click()
Load sub1
sub1.Show

End Sub

Private Sub mnu_vou_ent_Click()
Load vou1
vou1.Show

End Sub

Private Sub mnuAboutus_Click()
frmAbout.Show
End Sub

Private Sub mnuaccode_Click()
Load acchart1
acchart1.Show
End Sub

Private Sub mnudailycombineinformation_Click()
Load book
book.Text1.Text = 10
book.Show
End Sub

Private Sub mnuempchart_Click()
Load ac1
ac1.Text1.Text = 5
ac1.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
    Case 1
    mni_comp_ac_ledger_Click
    Case 2
    mniAgingLedgerofAnAccount_Click
    Case 3
    mni_dates_trial_Click
    Case 4
    mniClosingStocks_Click
    Case 5
    mniCostAndStockReportSH_Click
    Case 6
    mniCostAndStockReport_Click
    Case 7
    mniManualCostSheet_Click
    Case 8
    mniDayBook_Click
    Case 9
        Blmr.trial2 FStartDate, Date, ProgressBar1, , "36002,"
        R1.ReportFileName = App.path & "\Trial3.rpt"
        R1.DataFiles(0) = App.path & "\Book.mdb"
        R1.WindowTitle = "Closing Balances of All Accounts in a Sub-Head"
        R1.Action = 1
    Case 10
        Blmr.trial2 FStartDate, Date, ProgressBar1, "31,"
        R1.ReportFileName = App.path & "\Trial3.rpt"
        R1.DataFiles(0) = App.path & "\Book.mdb"
        R1.WindowTitle = "Closing Balances of All Accounts in a Head"
        R1.Action = 1
    
End Select
End Sub
