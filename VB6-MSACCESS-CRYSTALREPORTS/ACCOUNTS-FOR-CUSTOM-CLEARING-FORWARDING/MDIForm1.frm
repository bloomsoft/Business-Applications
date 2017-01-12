VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BloomSoft Custom Clearing Accounts"
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
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Create or Manage Ledger Accounts"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Create or Manage Items"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
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
      Caption         =   "&File"
      Begin VB.Menu mni_Parties_Code 
         Caption         =   "Parties And Accounts Information (Coding)"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Item_Code 
         Caption         =   "Clearning Expences Information (Coding)"
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu mniShippingDocumentsDefine 
         Caption         =   "Shipping Documents Definition"
      End
      Begin VB.Menu mni_Pay_rec_Entry 
         Caption         =   "Expence And Cash Payments / Reciepts"
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnu_Reports 
      Caption         =   "&Reports"
      Begin VB.Menu Mni_lists 
         Caption         =   "Lists"
         Begin VB.Menu mni_Accounts_List 
            Caption         =   "Parties And Accounts List"
         End
         Begin VB.Menu we 
            Caption         =   "-"
         End
         Begin VB.Menu itemslistgroupswise 
            Caption         =   "Clearing Expences List"
         End
      End
      Begin VB.Menu qwd 
         Caption         =   "-"
      End
      Begin VB.Menu mniDayBook 
         Caption         =   "Day Book"
      End
      Begin VB.Menu mniClearingCaseInvoice 
         Caption         =   "Clearing Case Invoice"
      End
      Begin VB.Menu mniClearingCaseGSTInv 
         Caption         =   "Clearing Case GST Invoice"
      End
      Begin VB.Menu mni_Ac_Ledger_Dates 
         Caption         =   "Account Wise Ledger (Between Dates)"
      End
      Begin VB.Menu mniClearCaseLedger 
         Caption         =   "Clearing Case Ledger"
      End
      Begin VB.Menu Sep14 
         Caption         =   "-"
      End
      Begin VB.Menu mni_Trial_Balance 
         Caption         =   "Trial Balance (to An End Date)"
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
      Begin VB.Menu mniDeleteAllData 
         Caption         =   "Delete All Data"
      End
      Begin VB.Menu mniAutoBackupPathSettings 
         Caption         =   "Auto Backup Path Settings"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cport As String
Dim Counter As Integer
Private Sub GetPortNumber()
Open App.Path & "\CommPort.txt" For Input As #1
    Cport = Input(1, LOF(1))
Close #1

End Sub
Private Sub GetBackupPath()
Dim db As Database
Dim ssql As String
Dim tb As Recordset

Set db = OpenDatabase(App.Path & "\User.mdb")
ssql = "Select * from Backup"
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    BackupPath = tb.Fields("Path").Value & ""
Else
    MsgBox "No Auto Backup Path has been Set Please Set it From Setup Menu"
End If
tb.Close
db.Close
End Sub
Private Sub TakeBackup()
Dim FS As New FileSystemObject
Dim NP As String
If Len(BackupPath) > 3 Then
    NP = BackupPath & "\" & Format(Date, "MMMMddyyyy")
Else
    NP = BackupPath & Format(Date, "MMMMddyyyy")
End If
If FS.FolderExists(NP) Then
    FS.CopyFile App.Path & "\Bloom.mdb", NP & "\Bloom.mdb", True
Else
    FS.CreateFolder NP
    FS.CopyFile App.Path & "\Bloom.mdb", NP & "\Bloom.mdb", True
End If
DoEvents
End Sub
Private Sub itemslistgroupswise_Click()
Load Lists
With Lists
    .Text1.Text = 2
    .Label1.Visible = False
    .Combo1.Visible = False
    .Caption = "Clearing Expences List"
    .Show
End With
End Sub

Private Sub MDIForm_Load()
GetBackupPath

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
TakeBackup
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
'Groups1.Show
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
End Sub

Private Sub mni_Sales_MultiWestNumbers_Click()

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

Private Sub mniAutoBackupPathSettings_Click()
AutoBackup.Show
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

Private Sub mniCityWiseBalances_Click()
CityBal.Show
End Sub

Private Sub mniClearCaseLedger_Click()
Load InvPrint
InvPrint.Text2.Text = 2
InvPrint.Caption = "Clearing Case Ledger"
InvPrint.Show

End Sub

Private Sub mniClearingCaseGSTInv_Click()
STInvoice.Show
End Sub

Private Sub mniClearingCaseInvoice_Click()
Load InvPrint
InvPrint.Text2.Text = 1
InvPrint.Caption = "Clearing Case Invoice"
InvPrint.Show
End Sub

Private Sub mniDayBook_Click()
Stock1.Text1.Text = 5
Stock1.Caption = "Day Book"
Stock1.Show
End Sub

Private Sub mniDeleteAllData_Click()
Dim Result As VbMsgBoxResult

Result = MsgBox("Do You Realy Want to Delete All Data", vbYesNo)
If Result = vbYes Then
    Dim ssql As String
    Dim db As Database
    
    Set db = OpenDatabase(App.Path & "\Bloom.mdb")
    ssql = "Delete from Parties"
    db.Execute ssql
    ssql = "Delete from Voucher"
    db.Execute ssql
    ssql = "Delete from Pre_Cash"
    db.Execute ssql
    ssql = "Delete from Docs"
    db.Execute ssql
    ssql = "Delete from Invoices"
    db.Execute ssql
    ssql = "Delete from GSTInvoice"
    db.Execute ssql
    db.Close
    MsgBox "All Data Deleted"
End If
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

Private Sub mniShippingDocumentsDefine_Click()
newDocs.Show
End Sub

Private Sub mniUserManage_Click()
login.Show
End Sub

Private Sub Timer1_Timer()
Dim s As String
MSComm1.Output = "zakria"
DoEvents
End Sub

Private Sub Timer2_Timer()
If MSComm1.InBufferCount > 0 Then
    s = MSComm1.Input
    If Len(s) > 0 Then
        Counter = 0
    Else
     '   Counter = Counter + 1
    '    MsgBox "Please Attach the Identity Device on the Comm port ", , Format(Now, "sshhmm")
    End If
Else
   '     Counter = Counter + 1
  '      MsgBox "Please Attach the Identity Device on the Comm port ", , Format(Now, "sshhmm")
End If

If Counter > 12 Then
 '   MsgBox "Sorry You did not Attach the Identity Device on Comm Port So Software will Now Close"
'    End
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
    Case 1
        mni_Parties_Code_Click
    Case 2
        mni_Item_Code_Click
        
    Case 3
        mni_Pay_rec_Entry_Click
'    Case 4
'        mni_Item_Code_Click
'    Case 5
'        mni_Purs_Inward_Click
'    Case 6
'        mni_Sale_Invoice_Click
'    Case 7
        
    
End Select

End Sub
