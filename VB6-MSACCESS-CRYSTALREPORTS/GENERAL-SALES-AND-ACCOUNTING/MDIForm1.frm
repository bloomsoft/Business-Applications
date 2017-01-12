VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BloomSoft Point of Sale"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Create or Manage Cities"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Create or Manage Ledger Accounts"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Create or Manage Item Groups"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Create or Manage Items"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Purchases"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Sales"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cash"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   960
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":1736
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":20B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":33A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3D1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&Data Entry"
      Begin VB.Menu mni_City_Code 
         Caption         =   "Cities Information (Coding)"
      End
      Begin VB.Menu mniWareHousent 
         Caption         =   "WareHouse Information"
      End
      Begin VB.Menu mni_Parties_Code 
         Caption         =   "Parties Information (Coding)"
      End
      Begin VB.Menu mni_broker_code 
         Caption         =   "Brokers Information"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Items_Coding 
         Caption         =   "Items Groups Information (Coding)"
      End
      Begin VB.Menu mni_Item_Code 
         Caption         =   "Items Information (Coding)"
      End
      Begin VB.Menu b 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuvoucherentry 
      Caption         =   "&Vouchers Entry"
      Begin VB.Menu mni_Purs_Inward 
         Caption         =   "Purchase or Inward"
      End
      Begin VB.Menu mni_Sale_Invoice 
         Caption         =   "Sale Invoice"
      End
      Begin VB.Menu mnupurchasejob 
         Caption         =   "Purchase Job"
      End
      Begin VB.Menu mnusaleagreement 
         Caption         =   "Sale Job"
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Pay_rec_Entry 
         Caption         =   "Cash Payments and Reciepts"
      End
      Begin VB.Menu f 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mni_Quit 
         Caption         =   "Quit"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_Reports 
      Caption         =   "&Reports"
      Begin VB.Menu Mni_lists 
         Caption         =   "Lists"
         Begin VB.Menu mni_Accounts_List 
            Caption         =   "Accounts List Cities Wise"
         End
         Begin VB.Menu mni_Cities 
            Caption         =   "Cities List"
         End
         Begin VB.Menu mniWareHouseList 
            Caption         =   "WareHouse List"
         End
         Begin VB.Menu mni_Accountsonecity 
            Caption         =   "Accounts of One City"
         End
         Begin VB.Menu we 
            Caption         =   "-"
         End
         Begin VB.Menu mni_Items_Grousp 
            Caption         =   "Items Groups List"
         End
         Begin VB.Menu itemslistgroupswise 
            Caption         =   "Items List Groups Wise"
         End
         Begin VB.Menu mni_ItemofOneGroup 
            Caption         =   "Items of One Group"
         End
      End
      Begin VB.Menu qwd 
         Caption         =   "-"
      End
      Begin VB.Menu mni_PJobs 
         Caption         =   "Purchase Job"
         Begin VB.Menu mni_pjob_contnowise 
            Caption         =   "Purchase Contract No.Wise"
         End
         Begin VB.Menu mni_pjob_partywisecont 
            Caption         =   "Purchase Contract Party Wise "
         End
         Begin VB.Menu mni_pjob_itemwisecont 
            Caption         =   "Purchase Contract Item Wise"
         End
         Begin VB.Menu mniShortfallStatementofAllContracts 
            Caption         =   "Short Fall Statement of All Contracts"
         End
      End
      Begin VB.Menu mni_SJobs 
         Caption         =   "Sale Jobs "
         Begin VB.Menu mni_sjob_contnowise 
            Caption         =   "Sale Contract No.Wise"
         End
         Begin VB.Menu mni_sjob_partywisecont 
            Caption         =   "Sale Contract Party Wise "
         End
         Begin VB.Menu mni_sjob_itemwisecont 
            Caption         =   "Sale Contract Item Wise"
         End
         Begin VB.Menu mniShortfallStatementofAllContractsSales 
            Caption         =   "Short Fall Statement of All Contracts"
         End
      End
      Begin VB.Menu qw 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Purc_Inward 
         Caption         =   "Purchase or Inwards"
         Begin VB.Menu mni_Purchase_Inwards_BetweenDates 
            Caption         =   "Purchase or Inwards Between Dates"
         End
         Begin VB.Menu mni_Purc_Inward_Account_Wise_Period 
            Caption         =   "Purchase or Inwards Account Wise Periodic"
         End
         Begin VB.Menu mni_Purchase_Inward_West_Item 
            Caption         =   "Purchase or Inward Item Wise Periodic"
         End
      End
      Begin VB.Menu mni_Sales 
         Caption         =   "Sales"
         Begin VB.Menu mni_SaleInvoiceWest 
            Caption         =   "Sales Invoice"
         End
         Begin VB.Menu mni_Sales_MultiWestNumbers 
            Caption         =   "Sales Invoices Multiple Numbers Wise"
            Visible         =   0   'False
         End
         Begin VB.Menu mnimultiSalesWestInv 
            Caption         =   "Date Wise Multiple Sales Invoices"
            Visible         =   0   'False
         End
         Begin VB.Menu mni_Sales_BetweendatesWest 
            Caption         =   "Sales Between Dates"
         End
         Begin VB.Menu mni_Sales_Between_Dates_West 
            Caption         =   "Sales Item Wise Between Dates"
         End
         Begin VB.Menu mni_Sales_AccountWisePeriodic 
            Caption         =   "Sales Account Wise Periodic"
         End
         Begin VB.Menu mni_DailySalesOverAll 
            Caption         =   "Daily Sales OverAll"
         End
         Begin VB.Menu mni_PartySaleItem 
            Caption         =   "Party Wise Item Sales Periodic"
         End
      End
      Begin VB.Menu alpha1 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Ac_Ledger_Dates 
         Caption         =   "Account Wise Ledger (Between Dates)"
      End
      Begin VB.Menu alpha2 
         Caption         =   "-"
      End
      Begin VB.Menu mniDayBook 
         Caption         =   "Day Book"
      End
      Begin VB.Menu mni_Trial_Balance 
         Caption         =   "Trial Balance (to An End Date)"
      End
      Begin VB.Menu alpha3 
         Caption         =   "-"
      End
      Begin VB.Menu mni_ItemLedger 
         Caption         =   "Item Ledger Periodic"
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Stocks_All_Item_End_date 
         Caption         =   "Stocks of All Items"
      End
      Begin VB.Menu mni_Stock_All_Values 
         Caption         =   "Stocks of All Items With Values"
      End
      Begin VB.Menu mni_Stock_Betweendates 
         Caption         =   "Stocks Details (Between Dates)"
      End
      Begin VB.Menu SepRecAbles 
         Caption         =   "-"
      End
      Begin VB.Menu mniRecablesList 
         Caption         =   "Receiveables List"
      End
   End
   Begin VB.Menu mni_Setup 
      Caption         =   "Setup"
      Begin VB.Menu mniUserManage 
         Caption         =   "User Management"
      End
      Begin VB.Menu mniBackupData 
         Caption         =   "Backup"
      End
      Begin VB.Menu mniRestore 
         Caption         =   "Restore"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub itemslistgroupswise_Click()
Load Lists
With Lists
    .Text1.Text = 5
    .Label1.Visible = False
    .Combo1.Visible = False
    .Caption = "Items List Groups Wise"
    .Show
End With
End Sub

Private Sub mni_Ac_Ledger_Dates_Click()
Ledger.Caption = "Account Wise Ledger Between Dates"
Ledger.Text3.Text = 1
Ledger.Show
End Sub

Private Sub mni_Accounts_List_Click()
Load Lists
With Lists
    .Text1.Text = 1
    .Label1.Visible = False
    .Combo1.Visible = False
    .Caption = "Accounts List Cities Wise"
    .Show
End With
End Sub

Private Sub mni_Accountsonecity_Click()
Load Lists
With Lists
    .Text1.Text = 3
    .Caption = "Accounts List of One City"
    .Show
End With

End Sub

Private Sub mni_broker_code_Click()
Brokers.Show
End Sub

Private Sub mni_Cities_Click()
Load Lists
With Lists
    .Text1.Text = 2
    .Label1.Visible = False
    .Combo1.Visible = False
    .Caption = "Cities List"
    .Show
End With

End Sub

Private Sub mni_City_Code_Click()
City.Show
End Sub

Private Sub mni_DailySalesOverAll_Click()
Load notes
With notes
    .Text3.Text = 151
    .Caption = "Daily Sales"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_DailySalesOverAllTowels_Click()
Load notes
With notes
    .Text3.Text = 153
    .Caption = "Daily Sales (Towels)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_DailySalesSocks_Click()
Load notes
With notes
    .Text3.Text = 152
    .Caption = "Daily Sales (Socks)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_Item_Code_Click()
Item1.Show
End Sub

Private Sub mni_ItemLedger_Click()
Ledger.Caption = "Item Ledger Periodic"
Ledger.Text3.Text = 2
Ledger.Frame4.Visible = True
Ledger.Show

End Sub

Private Sub mni_ItemofOneGroup_Click()
Load Lists
With Lists
    .Text1.Text = 6
    .Caption = "Items of One Group"
    .Show
End With

End Sub

Private Sub mni_Items_Coding_Click()
Groups1.Show
End Sub

Private Sub mni_Items_Grousp_Click()
Load Lists
With Lists
    .Text1.Text = 4
    .Label1.Visible = False
    .Combo1.Visible = False
    .Caption = "Item Groups List"
    .Show
End With

End Sub

Private Sub mni_ItemWiesSalesPeriodic_Click()
Load notes
With notes
    .Text3.Text = 55
    .Caption = "Between Dates Item Wise Sales (Socks)"
    .Frame3.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_MultilsalesInvTowels_Click()
Load notes
With notes
    .Text3.Text = 62
    .Caption = "Between Dates Sales Invoices (Towels)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_MultiSalesInvSocks_Click()
Load notes
With notes
    .Text3.Text = 61
    .Caption = "Between Dates Sales Invoices (Socks)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_Parties_Code_Click()
Party.Show
End Sub

Private Sub mni_PartySaleItem_Click()
Load PItem
PItem.Text5.Text = 1
PItem.Caption = "Party Wise Item Sale"
PItem.Show
End Sub

Private Sub mni_partySalesItemSocks_Click()
Load PItem
PItem.Text5.Text = 2
PItem.Caption = "Party Wise Item Sale"
PItem.Show

End Sub

Private Sub mni_PartySalesItemTowels_Click()
Load PItem
PItem.Text5.Text = 3
PItem.Caption = "Party Wise Item Sale"
PItem.Show

End Sub

Private Sub mni_Pay_rec_Entry_Click()
Load vou_1
vou_1.Show
End Sub

Private Sub mni_pjob_contnowise_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 8
    .Caption = "Purchase Contract Print "
    .Label4.Visible = False
    .Combo3.Visible = False
    .Label3.Visible = False
    .Label1.Caption = "Contract#"
    .Show
End With
End Sub

Private Sub mni_pjob_itemwisecont_Click()
Load notes
With notes
    .Text3.Text = 202
    .Caption = "Purchases Contract Item Wise "
    .Frame3.Visible = True
    .Show
    
End With
End Sub

Private Sub mni_pjob_partywisecont_Click()
Load notes
With notes
    .Text3.Text = 200
    .Caption = "Purchases Contract Party Wise"
    .Frame2.Visible = True
    .Show
    
End With
End Sub

Private Sub mni_Purc_Inward_Account_Wise_Period_Click()
Load notes
With notes
    .Text3.Text = 2
    .Caption = "Between Dates Purchases Account Wise"
    .Frame2.Visible = True
    .Show
    
End With
End Sub

Private Sub mni_Purc_Inward_Item_Wise_Periodic_Towels_Click()
Load notes
With notes
    .Text3.Text = 9
    .Caption = "Between Dates Purchases Item Wise (Towels)"
    .Frame3.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_Purc_Inwards_Item_Wise_Socks_Click()
Load notes
With notes
    .Text3.Text = 6
    .Caption = "Between Dates Purchases Item Wise (Socks)"
    .Frame3.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_Purc_UInwards_Purc_Socks_Dates_Click()
Load notes
With notes
    .Text3.Text = 4
    .Caption = "Between Dates Purchases (Socks)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_Purchase_Inward_West_Item_Click()
Load notes
With notes
    .Text3.Text = 3
    .Caption = "Between Dates Purchases Item Wise "
    .Frame3.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_Purchase_Inwards_BetweenDates_Click()
Load notes
With notes
    .Text3.Text = 1
    .Caption = "Between Dates Purchases "
    .Frame2.Visible = False
    .Show
    
End With
End Sub

Private Sub mni_Purchases_Socks_Periodic_Click()
Load notes
With notes
    .Text3.Text = 5
    .Caption = "Between Dates Purchases Account Wise(Socks)"
    .Frame2.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_PurInwardTowelsAccountwise_Click()
Load notes
With notes
    .Text3.Text = 8
    .Caption = "Between Dates Purchases Account Wise(Towels)"
    .Frame2.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_PurInwardTowelsPeriodic_Click()
Load notes
With notes
    .Text3.Text = 7
    .Caption = "Between Dates Purchases (Towels)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_Purs_Inward_Click()
Load In1
In1.Text2.Text = 1
In1.Show
End Sub

Private Sub mni_Quit_Click()
 
End
End Sub

Private Sub mni_Sale_Invoice_Click()
Load Inv1
Inv1.Text20.Text = 1
Inv1.Show
End Sub

Private Sub mni_SaleInvoiceWest_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 1
    .Caption = "Sales Invoice Print "
    .Show
End With
End Sub

Private Sub mni_Sales_AccountWisePeriodic_Click()
Load notes
With notes
    .Text3.Text = 53
    .Caption = "Between Dates Sales Account Wise "
    .Frame2.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_Sales_Between_Dates_West_Click()
Load notes
With notes
    .Text3.Text = 52
    .Caption = "Between Dates Item Wise Sales "
    .Frame3.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_Sales_BetweendatesWest_Click()
Load notes
With notes
    .Text3.Text = 51
    .Caption = "Between Dates Sales "
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_Sales_InvSocksMultiNumbers_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 5
    .Caption = "Sales Invoice Print Multiple (Socks)"
    .Frame2.Visible = True
    .Show
End With
End Sub

Private Sub mni_Sales_MultiWestNumbers_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 4
    .Caption = "Sales Invoice Print Multiple "
    .Frame2.Visible = True
    .Show
End With

End Sub

Private Sub mni_Sales_PeriodicSocks_Click()
Load notes
With notes
    .Text3.Text = 54
    .Caption = "Between Dates Sales (Socks)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_Sales_PeriodicTowels_Click()
Load notes
With notes
    .Text3.Text = 57
    .Caption = "Between Dates Sales (Towels)"
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mni_Sales_Socks_Invoice_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 2
    .Caption = "Sales Invoice Print (Socks)"
    .Show
End With

End Sub

Private Sub mni_sales_Towels_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 3
    .Caption = "Sales Invoice Print (Towels)"
    .Show
End With

End Sub

Private Sub mni_SalesItemWisePeriodic_Click()
Load notes
With notes
    .Text3.Text = 58
    .Caption = "Between Dates Item Wise Sales (Towels)"
    .Frame3.Visible = True
    .Show
    
End With

End Sub

Private Sub mni_Sock_sale_entry_Click()
Load Inv1
Inv1.Text20.Text = 2
Inv1.Caption = "Socks Sale Invoice"
Inv1.Show

End Sub

Private Sub mni_Socks_Purs_Entry_Click()
Load In1
In1.Text2.Text = 2
In1.Caption = "Socks Purchase or Inward"
In1.Show
End Sub

Private Sub mni_sjob_contnowise_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 9
    .Caption = "Sale Contract Print "
    .Label1.Caption = "Contract#"
    .Label4.Visible = False
    .Label2.Visible = False
    .Combo3.Visible = False
    .Show
End With
End Sub

Private Sub mni_sjob_itemwisecont_Click()
Load notes
With notes
    .Text3.Text = 203
    .Caption = "Sale Contract Item Wise "
    .Frame3.Visible = True
    .Show
    
End With
End Sub

Private Sub mni_sjob_partywisecont_Click()
Load notes
With notes
    .Text3.Text = 201
    .Caption = "Sale Contract Party Wise"
    .Frame2.Visible = True
    .Show
    
End With
End Sub

Private Sub mni_Stock_All_Values_Click()
Load Stock1
Stock1.Caption = "Stocks to An End Date with Values"
Stock1.Text1.Text = 4
Stock1.Show

End Sub

Private Sub mni_Stock_Betweendates_Click()
With Stock1
.Text1.Text = 2
.Label1.Caption = "From"
.Label2.Visible = True
.DTPicker2.Visible = True
.Show
End With

End Sub

Private Sub mni_Stocks_All_Item_End_date_Click()
Load Stock1
Stock1.Caption = "Stocks to An End Date"
Stock1.Text1.Text = 1
Stock1.Show

End Sub

Private Sub mni_Towel_Sale_Entry_Click()
Load Inv1
Inv1.Text20.Text = 3
Inv1.Caption = "Towels Sale Invoice"
Inv1.Show

End Sub

Private Sub mni_Towels_Purc_Click()
Load In1
In1.Text2.Text = 3
In1.Caption = "Towels Purchase or Inward"
In1.Show
End Sub

Private Sub mni_Trial_Balance_Click()
Stock1.Text1.Text = 3
Stock1.Caption = "Trial Balance to An End Date"
Stock1.Show
End Sub

Private Sub mni_West_purc_Entry_Click()
Load In1
In1.Text2.Text = 1
In1.Caption = "Vest Purchase or Inward"
In1.Show
End Sub

Private Sub mni_West_sale_entry_Click()
Load Inv1
Inv1.Text20.Text = 1
Inv1.Caption = "Vest Sale Invoice"
Inv1.Show
End Sub

Private Sub mniBackupData_Click()
Dim FS As New FileSystemObject
With CD1
    .Filter = "Database File|*.mdb"
    .FilterIndex = 1
    .FileName = "BLOOM.MDB"
    .DefaultExt = "MDB"
    .ShowSave
    If .FileName <> "" Then
        FS.CopyFile App.Path & "\Bloom.Mdb", .FileName, True
        MsgBox "Data Backup Taken"
    End If
End With
End Sub

Private Sub mniDayBook_Click()
DayRep.Show
End Sub

Private Sub mnimultiSalesWestInv_Click()
Load notes
With notes
    .Text3.Text = 60
    .Caption = "Between Dates Sales Invoices "
    .Frame2.Visible = False
    .Show
    
End With

End Sub

Private Sub mniRecablesList_Click()
Stock1.Text1.Text = 5
Stock1.Caption = "Receiveables to An End Date"
Stock1.Show

End Sub

Private Sub mniRestore_Click()
Dim FS As New FileSystemObject
With CD1
    .Filter = "Database File|*.mdb"
    .FilterIndex = 1
    .FileName = "BLOOM.MDB"
    
    .ShowOpen
    If .FileName <> "" Then
        FS.CopyFile .FileName, App.Path & "\Bloom.mdb", True
        MsgBox "Data Restored"
    End If
End With

End Sub

Private Sub mniSaleInvMultiTowelsNumbers_Click()
Load InvPrint
With InvPrint
    .Text2.Text = 6
    .Caption = "Sales Invoice Print Multiple (Towels)"
    .Frame2.Visible = True
    .Show
End With
End Sub

Private Sub mnisalesaccountPeriodicSocks_Click()
Load notes
With notes
    .Text3.Text = 56
    .Caption = "Between Dates Account Wise Sales (Socks)"
    .Frame2.Visible = True
    .Show
    
End With

End Sub

Private Sub mniSalesAccountWisePeriodic_Click()
Load notes
With notes
    .Text3.Text = 59
    .Caption = "Between Dates Account Wise Sales (Towels)"
    .Frame2.Visible = True
    .Show
    
End With

End Sub

Private Sub mniShortfallStatementofAllContracts_Click()
Stock1.Text1.Text = 6
Stock1.Caption = "ShortFall Statement of All Purchase Contracts"
Stock1.Show

End Sub

Private Sub mniShortfallStatementofAllContractsSales_Click()
Stock1.Text1.Text = 7
Stock1.Caption = "ShortFall Statement of All Sale Contracts"
Stock1.Show

End Sub

Private Sub mniUserManage_Click()
login.Show
End Sub

Private Sub mniWareHouseList_Click()
Load Lists
With Lists
    .Text1.Text = 7
    .Label1.Visible = False
    .Combo1.Visible = False
    .Caption = "WareHouses List"
    .Show
End With

End Sub

Private Sub mniWareHousent_Click()
WareHouse.Show
End Sub

Private Sub mnupurchasejob_Click()
PContract.Show
End Sub

Private Sub mnusaleagreement_Click()
    SContract.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
    Case 1
        mni_City_Code_Click
    Case 2
        mni_Parties_Code_Click
    Case 3
        mni_Items_Coding_Click
    Case 4
        mni_Item_Code_Click
    Case 5
        mni_Purs_Inward_Click
    Case 6
        mni_Sale_Invoice_Click
    Case 7
        mni_Pay_rec_Entry_Click
    
End Select

End Sub
