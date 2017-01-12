VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BloomSoft Financial Management System"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
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
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Heads Coding"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Sub Heads Coding"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Accounts Coding"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Voucher Entry"
            Object.Tag             =   ""
            ImageIndex      =   4
            Object.Width           =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Journal Voucher Print"
            Object.Tag             =   ""
            ImageIndex      =   5
            Object.Width           =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Journal Book Print"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Account Ledger"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Trial Balance"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
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
            Picture         =   "MDIForm1.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":297C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3370
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3D64
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4758
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4932
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4B0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mni_Data_ent 
      Caption         =   "&Data Entry"
      Begin VB.Menu mni_chart_ac_setup 
         Caption         =   "Chart of A/c Setup"
         Begin VB.Menu mni_ac_heads 
            Caption         =   "A/c Heads Coding"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnu_subhead 
            Caption         =   "A/c Sub Heads Coding"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuaccode 
            Caption         =   "A/c Coding"
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu mni_Item_Code 
         Caption         =   "Items Coding"
      End
      Begin VB.Menu mniSaleJobDefinition 
         Caption         =   "Sale Job Definition"
      End
      Begin VB.Menu mnu_vou_ent 
         Caption         =   "Voucher Entry"
         Shortcut        =   ^V
      End
      Begin VB.Menu mni_P_Vou 
         Caption         =   "Purchase Voucher"
      End
      Begin VB.Menu mni_Sale_Vou 
         Caption         =   "Sales Voucher"
      End
      Begin VB.Menu mni_Consume_Vou 
         Caption         =   "Consumption Voucher"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mni_rep 
      Caption         =   "&Reports"
      Begin VB.Menu mni_chart_ac_rep 
         Caption         =   "Chart of A/c"
         Shortcut        =   ^C
      End
      Begin VB.Menu mni_Add 
         Caption         =   "Address and Phone List"
      End
      Begin VB.Menu mni_vou_print 
         Caption         =   "Voucher Printing"
         Begin VB.Menu mni_jou_vou_rep 
            Caption         =   "Journal Voucher"
            Shortcut        =   ^J
         End
         Begin VB.Menu mni_bank_vou_print 
            Caption         =   "Bank Voucher"
            Shortcut        =   ^B
         End
         Begin VB.Menu mni_cash_vou_print 
            Caption         =   "Cash Voucher"
            Shortcut        =   ^U
         End
         Begin VB.Menu mni_SJV 
            Caption         =   "Sale Journal Voucher (Invoice)"
         End
         Begin VB.Menu mni_PJV 
            Caption         =   "Purchase Journal Voucher"
         End
         Begin VB.Menu mni_CJV 
            Caption         =   "Consumption Voucher"
            Visible         =   0   'False
         End
         Begin VB.Menu mniSaleJobDefine 
            Caption         =   "Sale Job Definition"
         End
      End
      Begin VB.Menu mni_books 
         Caption         =   "Books"
         Begin VB.Menu mni_journal_Book 
            Caption         =   "Journal Book"
            Shortcut        =   ^D
         End
         Begin VB.Menu mni_Bank_book 
            Caption         =   "Bank Book"
            Shortcut        =   ^E
         End
         Begin VB.Menu mni_cash_book 
            Caption         =   "Cash Book"
            Shortcut        =   ^F
         End
         Begin VB.Menu mniPBBook 
            Caption         =   "Purchase Book"
         End
         Begin VB.Menu mniSalesBook 
            Caption         =   "Sales Book W/O GST"
         End
         Begin VB.Menu SalesBookWithGST 
            Caption         =   "Sales Book With GST"
         End
      End
      Begin VB.Menu mni_CashBook 
         Caption         =   "Cash Flow"
      End
      Begin VB.Menu mni_ledger_rep 
         Caption         =   "Journal Ledger"
         Begin VB.Menu mni_ac_ledger 
            Caption         =   "Account Ledger"
            Shortcut        =   ^I
         End
         Begin VB.Menu mni_ref_Ledger 
            Caption         =   "Ref. No. Wise Ledger"
         End
         Begin VB.Menu mni_comp_ac_ledger 
            Caption         =   "Complete Account Ledger (Periodic)"
            Shortcut        =   ^K
         End
         Begin VB.Menu mni_comb_ledger 
            Caption         =   "Combine Ledger"
            Shortcut        =   ^M
         End
         Begin VB.Menu mni_Item_Ledger 
            Caption         =   "Item Wise Ledger (Periodic)"
         End
      End
      Begin VB.Menu mni_pay_lst 
         Caption         =   "Payables List"
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mni_rec_lst 
         Caption         =   "Receivable List"
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mni_trial 
         Caption         =   "Trial Balance"
         Begin VB.Menu mni_open_trial 
            Caption         =   "Opening Trial Balance"
            Shortcut        =   ^O
         End
         Begin VB.Menu mni_simple_trial 
            Caption         =   "Simple Trial Balance (Without Groups)"
            Shortcut        =   ^Q
         End
         Begin VB.Menu mni_dates_trial 
            Caption         =   "Trial Balance Between Dates (With Groups)"
            Shortcut        =   ^T
         End
         Begin VB.Menu mni_group_trial 
            Caption         =   "Groups Trial Balance"
            Shortcut        =   ^G
         End
         Begin VB.Menu mni_Sub_H_Totals_Tria 
            Caption         =   "Groups Trail Balance (Sub Head Totals)"
         End
      End
      Begin VB.Menu mni_recable 
         Caption         =   "Recievables"
         Shortcut        =   ^W
         Visible         =   0   'False
      End
      Begin VB.Menu mni_payable 
         Caption         =   "Payables"
         Shortcut        =   ^Y
         Visible         =   0   'False
      End
      Begin VB.Menu mni_Finance_State 
         Caption         =   "Financial Statements"
      End
      Begin VB.Menu mni_Stock 
         Caption         =   "Stocks Over All to An End Date"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_Period_Stock 
         Caption         =   "Periodic Stock Statement"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mni_user 
      Caption         =   "Settings"
      Begin VB.Menu mni_Set_OrgName 
         Caption         =   "Set or Change Organization Name"
      End
      Begin VB.Menu mni_create_change 
         Caption         =   "Create or Change User Info"
      End
      Begin VB.Menu mni_New_App_Create 
         Caption         =   "New Application Creater"
      End
      Begin VB.Menu mniChangeOpenBalDate 
         Caption         =   "Change Opening Balances Date"
      End
      Begin VB.Menu mniDeleteAllOpeningbalances 
         Caption         =   "Delete All Opening Balances"
      End
      Begin VB.Menu mniEmptytheWholeDatabase 
         Caption         =   "Empty the Whole Database"
      End
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
ledger.Date1.Visible = True
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
ledger.Combo1.Visible = False
ledger.Date1.Visible = True
ledger.Date2.Visible = True
ledger.Show

End Sub

Private Sub mni_comp_ac_ledger_Click()
Load ledger
ledger.Text1.Text = 2
ledger.Label1.Visible = True
ledger.Label2.Visible = True
ledger.Date1.Visible = True
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
trial.Date1.Visible = False
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
vour.Text3.Visible = False
vour.Label2.Visible = False
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

Private Sub mniChangeOpenBalDate_Click()
OpenDateChange.Show
End Sub

Private Sub mniDeleteAllOpeningbalances_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete All Opening Balances", vbYesNo)
If Result = vbYes Then
    Dim DB As Database
    Dim Ssql As String
    Set DB = OpenDatabase(App.path & "\Bloom.mdb")
    Ssql = "Update Acchart Set Debit=0,Credit=0"
    DB.Execute Ssql
    Ssql = "Delete from vouDTL where V_type=10"
    DB.Execute Ssql
    DB.Close
    MsgBox "All Opening Balances Has been Deleted"
    
End If

End Sub

Private Sub mniEmptytheWholeDatabase_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Delete All Data", vbYesNo)
If Result = vbYes Then
    Dim DB As Database
    Dim Ssql As String
    Set DB = OpenDatabase(App.path & "\Bloom.mdb")
    Ssql = "Delete from Acchart where Name <> 'CASH IN HAND'"
    DB.Execute Ssql
    Ssql = "Delete from vouMST"
    DB.Execute Ssql
    Ssql = "Delete from vouDTL"
    DB.Execute Ssql
    Ssql = "Delete from Purchase"
    DB.Execute Ssql
    Ssql = "Delete from Sales"
    DB.Execute Ssql
    Ssql = "Delete from SaleJob"
    DB.Execute Ssql
    
    DB.Close
    MsgBox "All Data Has been Deleted"
End If

End Sub

Private Sub mniPBBook_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 1
PeriodicRep.Caption = "Periodic Purchase Ledger"
PeriodicRep.Show
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

Private Sub mniSalesBook_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 2
PeriodicRep.Caption = "Periodic Sales Book W/O GST"
PeriodicRep.Show

End Sub

Private Sub mnu_subhead_Click()
Load sub1
sub1.Show

End Sub

Private Sub mnu_vou_ent_Click()
Load vou1
vou1.Show

End Sub

Private Sub mnuaccode_Click()
Load acchart1
acchart1.Show
End Sub

Private Sub SalesBookWithGST_Click()
Load PeriodicRep
PeriodicRep.Text1.Text = 3
PeriodicRep.Caption = "Periodic Sales Book With GST"
PeriodicRep.Show

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
    Case 1
    'mni_ac_heads_Click
    Case 2
    mnu_subhead_Click
    Case 3
    mnuaccode_Click
    Case 4
    mnu_vou_ent_Click
    Case 5
    mni_jou_vou_rep_Click
    Case 6
    mni_journal_Book_Click
    Case 7
    mni_comp_ac_ledger_Click
    Case 8
    mni_group_trial_Click
    Case 9
    mni_rec_lst_Click
    Case 10
    mni_pay_lst_Click
    
End Select
End Sub
