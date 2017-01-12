VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BloomSoft Knitting Unit Manager"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   635
      ButtonWidth     =   2566
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Codings"
            Object.ToolTipText     =   "Drop Down to Select Codings Options"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Acchart"
                  Text            =   "Parties Coding"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ClothCode"
                  Text            =   "Cloth Coding"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "YarnCode"
                  Text            =   "Yarns Coding"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "M_Code"
                  Text            =   "Machine Coding"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ECode"
                  Text            =   "Employees Coding"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contracts"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Pknitt"
                  Text            =   "Purchase Knitting"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SKnitt"
                  Text            =   "Sale Knitting"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDye"
                  Text            =   "Purchase Dyeing"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inwards"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "InYarn"
                  Text            =   "Inward Yarn"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "incloth"
                  Text            =   "Inward Cloth"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "indyecloth"
                  Text            =   "Inward Dyed Cloth"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Outwards"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "clothout"
                  Text            =   "Cloth Outward"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Yarnout"
                  Text            =   "Yarn Outward"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ClothDyeout"
                  Text            =   "Dye Cloth Outward"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Listings"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PList"
                  Text            =   "Parties"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ClothL"
                  Text            =   "Cloths"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "YarnsL"
                  Text            =   "Yarns"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ML"
                  Text            =   "Machines"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EMPL"
                  Text            =   "Employees"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contract Print"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PKnittP"
                  Text            =   "Purchase Knitting"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SKnittP"
                  Text            =   "Sale Knitting Print"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDYEP"
                  Text            =   "Purchase Dye Print"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "G.Pass Print"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "YIN"
                  Text            =   "Yarn Inwards"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Clothin"
                  Text            =   "Cloth Inwards"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DCLOTHIN"
                  Text            =   "Dye Cloth Inwards"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "YOUT"
                  Text            =   "Yarn Outwards"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "COUt"
                  Text            =   "Cloth Outwards"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dyeclothout"
                  Text            =   "Dye Cloth Outwards"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Yarn InWard"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cloth Outward"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Yarn Outward"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cloth Inward"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   108
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0632
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":16C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2016
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2332
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":264E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":296A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":32BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":35DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":38F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":424A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4566
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4882
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":51D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":54F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":580E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5E46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_Data_Entry 
      Caption         =   "Data Entry"
      Begin VB.Menu mni_parties_Code 
         Caption         =   "Parties Coding"
      End
      Begin VB.Menu mni_cloth_code 
         Caption         =   "&Cloth Coding"
      End
      Begin VB.Menu mni_Yarn_Code 
         Caption         =   "&Yarn Coding"
      End
      Begin VB.Menu mni_machine_Code 
         Caption         =   "Machines Coding"
      End
      Begin VB.Menu mni_employee_code 
         Caption         =   "Employees Coding"
      End
      Begin VB.Menu mni_contracts 
         Caption         =   "Contracts"
         Begin VB.Menu mni_knit_purc_contract 
            Caption         =   "Knitting Purchase Contract"
         End
         Begin VB.Menu mni_knit_sale_contract 
            Caption         =   "Knitting Sale Contract"
         End
         Begin VB.Menu mni_dye_purc_contract 
            Caption         =   "Dyeing Purchase Contract"
         End
      End
      Begin VB.Menu mni_inwards_entry 
         Caption         =   "Inwards"
         Begin VB.Menu mni_ran_inward 
            Caption         =   "Yarn for Knitting InWard"
            Shortcut        =   ^I
         End
         Begin VB.Menu mni_cloth_Inward 
            Caption         =   "Cloth After Knitting InWard"
         End
         Begin VB.Menu mni_cloth_dye_InWard 
            Caption         =   "Cloth After Dyeing InWard"
         End
      End
      Begin VB.Menu mni_outwards_entry 
         Caption         =   "Outwards"
         Begin VB.Menu mni_cloth_outward 
            Caption         =   "Cloth After Knitting OutWard"
            Shortcut        =   ^O
         End
         Begin VB.Menu mni_yarn_outward 
            Caption         =   "Yarn for Knitting OutWard"
         End
         Begin VB.Menu mni_cloth_dye_outward 
            Caption         =   "Cloth For Dyeing OutWard"
         End
      End
      Begin VB.Menu stp2 
         Caption         =   "-"
      End
      Begin VB.Menu mniProductionEntry 
         Caption         =   "Production"
         Begin VB.Menu mniYarnIssuetoMachine 
            Caption         =   "Yarn Issued to Machine"
         End
         Begin VB.Menu mniFabricRecFromMachineEntry 
            Caption         =   "Fabric Received from Machine"
         End
         Begin VB.Menu ins_def 
            Caption         =   "Inspection Deffination"
         End
      End
      Begin VB.Menu mniNeedlesandSinkersEntry 
         Caption         =   "Needles and Sinkers"
         Begin VB.Menu mniNeedlesandSinkersDefinition 
            Caption         =   "Needles and Sinkers Definition"
         End
         Begin VB.Menu need_sin_in 
            Caption         =   "Needles and Sinkers InWard"
         End
         Begin VB.Menu need_sin_out 
            Caption         =   "Needles and Sinkers OutWard"
         End
      End
   End
   Begin VB.Menu mni_Reports 
      Caption         =   "Reports"
      Begin VB.Menu mni_party_rep 
         Caption         =   "Parties List"
      End
      Begin VB.Menu mni_CLoth_rep 
         Caption         =   "Cloth List"
      End
      Begin VB.Menu mni_yarn_list_rep 
         Caption         =   "Yarn List"
      End
      Begin VB.Menu mni_Machines_List_Rep 
         Caption         =   "Machines List"
      End
      Begin VB.Menu mni_Emp_List_Rep 
         Caption         =   "Employees List"
      End
      Begin VB.Menu p_list_rep 
         Caption         =   "Parts (Needles and Sinkers) List"
      End
      Begin VB.Menu mni_Cont_Print 
         Caption         =   "Contracts Printing"
         Begin VB.Menu mni_Pur_Knitt_Con_Print 
            Caption         =   "Purchase Knitting Contract"
         End
         Begin VB.Menu mni_Sale_Knitt_Cont_Print 
            Caption         =   "Sale Knitting Contract"
         End
         Begin VB.Menu mni_Pur_Dye_Cont_Print 
            Caption         =   "Purchase Dyeing Contract"
         End
         Begin VB.Menu mni_Cont_Details 
            Caption         =   "Contracts Details"
         End
      End
      Begin VB.Menu mni_Gate_Passes 
         Caption         =   "Gate Passes"
         Begin VB.Menu mni_Inward_Yarn_Rec 
            Caption         =   "Inward for Knitting Sale Contract (Yarn Recvd)"
         End
         Begin VB.Menu mni_inward_cloth_rec 
            Caption         =   "Inward for Knitting Purchase Contract (Cloth Recvd)"
         End
         Begin VB.Menu mni_cloth_rec_after_dye_rep 
            Caption         =   "Inward for Dyeing Purchase Contract (Cloth Recvd After Dyeing)"
         End
         Begin VB.Menu mni_outward_cloth_sent_knit 
            Caption         =   "Outward for Knitting Sale Contract (Cloth Sent)"
         End
         Begin VB.Menu mni_outward_yarn_sent_knit 
            Caption         =   "Outward for Knitting Purchase Contract (Yarn Sent)"
         End
         Begin VB.Menu mni_outward_cloth_sent_dye 
            Caption         =   "Outward for Dyeing Purchase Contract (Cloth Sent for Dyeing)"
         End
      End
      Begin VB.Menu mni_date_wise_in_out 
         Caption         =   "Date Wise Inwards && OutWards"
         Begin VB.Menu mni_inward_d_yarn_rec 
            Caption         =   "Inwards for Knitting Sale Contracts (Yarns Recvd.)"
         End
         Begin VB.Menu mni_inward_d_yarn_sent 
            Caption         =   "Inwards for Knitting Purchase Contracts (Cloths Recvd.)"
         End
         Begin VB.Menu mni_inward_cloth_d_rec 
            Caption         =   "Inwards for Dyeing Purchase Contracts (Cloths Recvd.)"
         End
         Begin VB.Menu mni_outward_d_cloth_sent 
            Caption         =   "Outwards for Sale Knitting Contracts (Cloths Sent)"
         End
         Begin VB.Menu mni_outward_yarn_sent_d 
            Caption         =   "Outwards for Purchase Knitting Contracts (Yarns Sent)"
         End
         Begin VB.Menu mni_outward_d_cloth_sent_dye 
            Caption         =   "Outwards for Purchase Dyeing Contracts (Cloths Sent)"
         End
      End
      Begin VB.Menu mni_cont_detail_summ 
         Caption         =   "Contracts Detail Summaries"
         Begin VB.Menu mni_purc_knitt_detals 
            Caption         =   "Purchase Knitting Contract Details"
         End
         Begin VB.Menu mni_sale_knitt_cont_details 
            Caption         =   "Sale Knitting Contract Details"
         End
         Begin VB.Menu mni_pur_dye_cont_details 
            Caption         =   "Purchase Dyeing Contract Details"
         End
      End
      Begin VB.Menu mni_party_con_summ_details 
         Caption         =   "Party Wise Contract Summaries"
         Begin VB.Menu mni_Party_Purc_knitt_cont_party 
            Caption         =   "Purchase Knitting Contracts"
         End
         Begin VB.Menu mni_sale_knit_cont_details 
            Caption         =   "Sale Knitting Contracts"
         End
         Begin VB.Menu mni_Party_Purc_Dye_Cont 
            Caption         =   "Purchase Dyeing Contracts"
         End
      End
      Begin VB.Menu mni_invt_led 
         Caption         =   "Inventory Ledger"
         Begin VB.Menu mni_p_inv_led 
            Caption         =   "Party Wise Inventory Ledger"
         End
         Begin VB.Menu yarn_inv_led 
            Caption         =   "Yarn Wise Inventory Ledger"
         End
         Begin VB.Menu fab_inv_led 
            Caption         =   "Quality Wise Inventory Ledger"
         End
      End
      Begin VB.Menu mni_inward 
         Caption         =   "Inwards Only"
         Begin VB.Menu min_p_inward 
            Caption         =   "Party Wise Inwards"
         End
         Begin VB.Menu y_inward 
            Caption         =   "Yarn Wise Inwards"
         End
         Begin VB.Menu q_inward 
            Caption         =   "Quality Wise Inwards"
         End
      End
      Begin VB.Menu mni_outward 
         Caption         =   "Outwards Only"
         Begin VB.Menu mni_p_out 
            Caption         =   "Party Wise Outwards"
         End
         Begin VB.Menu y_outward 
            Caption         =   "Yarn Wise Outwards"
         End
         Begin VB.Menu q_outward 
            Caption         =   "Quality Wise Outwards"
         End
      End
      Begin VB.Menu close_stock 
         Caption         =   "Trial Closing Stock"
      End
      Begin VB.Menu STP1 
         Caption         =   "-"
      End
      Begin VB.Menu produc 
         Caption         =   "Production"
         Begin VB.Menu yarn_issue_mac 
            Caption         =   "Yarn Issued To Machine"
            Begin VB.Menu y_issue_mac 
               Caption         =   "Yarn Issued To Machine No Wise"
            End
            Begin VB.Menu y_issue_mac_dates 
               Caption         =   "Yarn Issued To Machine Between Dates"
            End
         End
         Begin VB.Menu fabric_rec_mac 
            Caption         =   "Fabric Received From Machine"
            Begin VB.Menu fab_rec_mat 
               Caption         =   "Fabric Received From Machine No Wise"
            End
            Begin VB.Menu fab_rec_mat_dates 
               Caption         =   "Fabric Received From Machine Between Dates"
            End
         End
      End
      Begin VB.Menu needl_in_rep 
         Caption         =   "Needles and Sinkers InWard"
         Begin VB.Menu need_sink_inno 
            Caption         =   "Needles and Sinkers InWard No Wise"
         End
         Begin VB.Menu need_sink_in_dates 
            Caption         =   "Needles and Sinkers InWard Between Dates"
         End
      End
      Begin VB.Menu needl_out_rep 
         Caption         =   "Needles and Sinkers OutWard"
         Begin VB.Menu need_sink_outno 
            Caption         =   "Needles and Sinkers OutWard No Wise"
         End
         Begin VB.Menu need_sink_out_dates 
            Caption         =   "Needles and Sinkers OutWard Between Dates"
         End
      End
   End
   Begin VB.Menu mnu_quit 
      Caption         =   "&Quit"
      Begin VB.Menu mni_Cancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu backup 
         Caption         =   "Backup"
      End
      Begin VB.Menu mni_Exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub backup_Click()
Dim S As String
'S = "d:\orant\bin\exp80 " & BLOOMNEW & "/mlb File =" & App.path & "\Backups\BU" & Format(Date, "ddMMMyyyy") & ".dmp"
S = ""
Shell S, vbNormalNoFocus
End Sub

Private Sub close_stock_Click()
Load vou_r
vou_r.Text1.Text = 16
vou_r.date2.Visible = False
vou_r.Label3.Visible = False
vou_r.Label2.Caption = "Up To"
vou_r.Caption = "Trial Closing Stock"
vou_r.Show
End Sub

Private Sub fab_inv_led_Click()
Load vou_r
vou_r.Text1.Text = 11
vou_r.Frame4.Visible = True
vou_r.Caption = "Quality Wise In/Out Inventory Ledger"
vou_r.Show

End Sub

Private Sub fab_rec_mat_Click()
Load vour
vour.Caption = "Fabric Received From Machine"
vour.Text2.Text = 67
vour.Label1.Caption = "Receive #"
vour.Show

End Sub

Private Sub fab_rec_mat_dates_Click()
Load vou_r
vou_r.Caption = "Fabric Received From Machines"
vou_r.Text1.Text = 20
vou_r.Show

End Sub

Private Sub ins_def_Click()
Load Ins
Ins.Show
End Sub

Private Sub MDIForm_Load()
Main
End Sub

Private Sub min_p_inward_Click()
Load vou_r
vou_r.Text1.Text = 8
vou_r.Frame3.Visible = True
vou_r.Caption = "Party Wise Inwards"
vou_r.Show
End Sub

Private Sub mni_cloth_code_Click()
Load Item2
Item2.Show
End Sub

Private Sub mni_cloth_dye_InWard_Click()
Load out4
out4.Show
End Sub

Private Sub mni_cloth_dye_outward_Click()
Load kachi3
kachi3.Show
End Sub

Private Sub mni_cloth_Inward_Click()
Load InwardKPC
InwardKPC.Show
End Sub

Private Sub mni_cloth_outward_Click()
OutWarrdSKC.Show

End Sub

Private Sub mni_cloth_rec_after_dye_rep_Click()
Load vour
vour.Caption = "Inward for Dyeing Purchase Contract"
vour.Text2.Text = 3
vour.Label1.Caption = "Inward #"
vour.Show

End Sub

Private Sub mni_CLoth_rep_Click()
Load ac1
ac1.Caption = "Cloths Listings"
ac1.Text1.Text = 2
ac1.Show
End Sub

Private Sub mni_Cont_Details_Click()
Load ac1
ac1.Caption = "Contracts Details"
ac1.Text1.Text = 6
ac1.Show
End Sub

Private Sub mni_dye_purc_contract_Click()
Load cont_p_c
cont_p_c.Show
End Sub

Private Sub mni_Emp_List_Rep_Click()
Load ac1
ac1.Caption = "Employees Listings"
ac1.Text1.Text = 5
ac1.Show

End Sub

Private Sub mni_employee_code_Click()
Load emp1
emp1.Show
End Sub

Private Sub mni_Exit_Click()
End
End Sub

Private Sub mni_inward_cloth_d_rec_Click()
Load vou_r
vou_r.Caption = "Inwards for Purchase Dyeing Contracts.."
vou_r.Text1.Text = 3
vou_r.Show

End Sub

Private Sub mni_inward_cloth_rec_Click()
Load vour
vour.Caption = "Inward for Knitting Purchase Contract"
vour.Text2.Text = 2
vour.Label1.Caption = "Inward #"
vour.Show

End Sub

Private Sub mni_inward_d_yarn_rec_Click()
Load vou_r
vou_r.Caption = "Inwards for Sale Knitting Contracts.."
vou_r.Text1.Text = 1
vou_r.Show
End Sub

Private Sub mni_inward_d_yarn_sent_Click()
Load vou_r
vou_r.Caption = "Inwards for Purchase Knitting Contracts.."
vou_r.Text1.Text = 2
vou_r.Show

End Sub

Private Sub mni_Inward_Yarn_Rec_Click()
Load vour
vour.Caption = "Inward for Knitting Sale Contract"
vour.Text2.Text = 1
vour.Label1.Caption = "Inward #"
vour.Show
End Sub

Private Sub mni_knit_purc_contract_Click()
Load cont_PNN
cont_PNN.Show
End Sub

Private Sub mni_knit_sale_contract_Click()
Load cont_S
cont_S.Show
End Sub

Private Sub mni_machine_Code_Click()
Load machine
machine.Show
End Sub

Private Sub mni_Machines_List_Rep_Click()
Load ac1
ac1.Caption = "Machines Listings"
ac1.Text1.Text = 4
ac1.Show

End Sub

Private Sub mni_outward_cloth_sent_dye_Click()
Load vour
vour.Caption = "Outward for Dyeing Purchase Contract"
vour.Text2.Text = 6
vour.Label1.Caption = "Outward #"
vour.Show

End Sub

Private Sub mni_outward_cloth_sent_knit_Click()
Load vour
vour.Caption = "Outward for Knitting Sale Contract"
vour.Text2.Text = 4
vour.Frame2.Visible = True
vour.Label1.Caption = "Outward #"

vour.Show

End Sub

Private Sub mni_outward_d_cloth_sent_Click()
Load vou_r
vou_r.Caption = "Outwards for Sale Knitting Contracts.."
vou_r.Text1.Text = 4
vou_r.Show
End Sub

Private Sub mni_outward_d_cloth_sent_dye_Click()
Load vou_r
vou_r.Caption = "Outwards for Purchase Dyeing Contracts.."
vou_r.Text1.Text = 6
vou_r.Show
End Sub

Private Sub mni_outward_yarn_sent_d_Click()
Load vou_r
vou_r.Caption = "Outwards for Purchase Knitting Contracts.."
vou_r.Text1.Text = 5
vou_r.Show
End Sub

Private Sub mni_outward_yarn_sent_knit_Click()
Load vour
vour.Caption = "Outward for Knitting Purchase Contract"
vour.Text2.Text = 5
vour.Label1.Caption = "Outward #"
vour.Show

End Sub

Private Sub mni_p_inv_led_Click()
Load vou_r
vou_r.Text1.Text = 7
vou_r.Frame3.Visible = True
vou_r.Caption = "Party Wise In/Out Inventory Ledger"
vou_r.Show
End Sub

Private Sub mni_p_out_Click()
Load vou_r
vou_r.Text1.Text = 9
vou_r.Frame3.Visible = True
vou_r.Caption = "Party Wise Outwards"
vou_r.Show
End Sub

Private Sub mni_parties_Code_Click()
Load acchart1
acchart1.Show
End Sub

Private Sub mni_Party_Purc_Dye_Cont_Click()
Load vour
vour.Caption = "Party Wise Purchase Dyeing Contracts"
'vour.Label1.Caption = "Contract #"
vour.Frame1.Visible = False
vour.Frame3.Visible = True
vour.Text2.Text = 56
vour.Show

End Sub

Private Sub mni_Party_Purc_knitt_cont_party_Click()
Load vour
vour.Caption = "Party Wise Purchase Knitting Contracts"
'vour.Label1.Caption = "Contract #"
vour.Frame1.Visible = False
vour.Frame3.Visible = True
vour.Text2.Text = 54
vour.Show

End Sub

Private Sub mni_party_rep_Click()
Load ac1
ac1.Caption = "Parties Listings"
ac1.Text1.Text = 1
ac1.Show

End Sub

Private Sub mni_pur_dye_cont_details_Click()
Load vour
vour.Caption = "Purchase Dyeing Contract Summary"
vour.Label1.Caption = "Contract #"
vour.Text2.Text = 53
vour.Show

End Sub

Private Sub mni_Pur_Dye_Cont_Print_Click()
Load vour
vour.Caption = "Purchase Dyeing Contract"
vour.Label1.Caption = "Contract #"
vour.Text2.Text = 63
vour.Show

End Sub

Private Sub mni_Pur_Knitt_Con_Print_Click()
Load vour
vour.Caption = "Purchase Knitting Contract"
vour.Label1.Caption = "Contract #"
vour.Text2.Text = 61
vour.Show

End Sub

Private Sub mni_purc_knitt_detals_Click()
Load vour
vour.Caption = "Purchase Knitting Contract Summary"
vour.Label1.Caption = "Contract #"
vour.Text2.Text = 51
vour.Show

End Sub

Private Sub mni_ran_inward_Click()
Load kachi
kachi.Show
End Sub

Private Sub mni_sale_knit_cont_details_Click()
Load vour
vour.Caption = "Party Wise Sale Knitting Contracts"
vour.Frame1.Visible = False
vour.Frame3.Visible = True
vour.Text2.Text = 55
vour.Show

End Sub

Private Sub mni_sale_knitt_cont_details_Click()
Load vour
vour.Caption = "Sale Knitting Contract Summary"
vour.Label1.Caption = "Contract #"
vour.Text2.Text = 52
vour.Show

End Sub

Private Sub mni_Sale_Knitt_Cont_Print_Click()
Load vour
vour.Caption = "Sale Knitting Contract"
vour.Label1.Caption = "Contract #"
vour.Text2.Text = 62
vour.Show

End Sub

Private Sub mni_Yarn_Code_Click()
Load Item1
Item1.Show
End Sub

Private Sub mni_yarn_list_rep_Click()
Load ac1
ac1.Caption = "Yarn Listings"
ac1.Text1.Text = 3
ac1.Show

End Sub

Private Sub mni_yarn_outward_Click()
OutwardPKC.Show
End Sub

Private Sub mniFabricRecFromMachineEntry_Click()
FabricRcvdFromMachine.Show
End Sub

Private Sub mniNeedlesandSinkersDefinition_Click()
Needles.Show
End Sub

Private Sub mniYarnIssuetoMachine_Click()
YarnIssueToMachine.Show
End Sub

Private Sub need_sin_in_Click()
Load NeedleIn
NeedleIn.Show
End Sub

Private Sub need_sin_out_Click()
Load NeedleOut
NeedleOut.Show
End Sub

Private Sub need_sink_in_dates_Click()
Load vou_r
vou_r.Caption = "Needles and Sinkers InWard"
vou_r.Text1.Text = 17
vou_r.Show

End Sub

Private Sub need_sink_inno_Click()
Load vour
vour.Caption = "Needles and Sinkers Inward"
vour.Text2.Text = 64
vour.Label1.Caption = "Inward #"
vour.Show

End Sub

Private Sub need_sink_out_dates_Click()
Load vou_r
vou_r.Caption = "Needles and Sinkers OutWard"
vou_r.Text1.Text = 18
vou_r.Show

End Sub

Private Sub need_sink_outno_Click()
Load vour
vour.Caption = "Needles and Sinkers Outward"
vour.Text2.Text = 65
vour.Label1.Caption = "Outward #"
vour.Show

End Sub

Private Sub p_list_rep_Click()
Load ac1
ac1.Caption = "Parts List"
ac1.Text1.Text = 7
ac1.Show

End Sub

Private Sub q_inward_Click()
Load vou_r
vou_r.Text1.Text = 13
vou_r.Frame4.Visible = True
vou_r.Caption = "Quality Wise Inwards"
vou_r.Show

End Sub

Private Sub q_outward_Click()
Load vou_r
vou_r.Text1.Text = 15
vou_r.Frame4.Visible = True
vou_r.Caption = "Quality Wise Outwards"
vou_r.Show

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
            mni_parties_Code_Click
    Case 2
           mni_knit_purc_contract_Click
     Case 11
           mni_ran_inward_Click
   Case 14
            mni_cloth_Inward_Click
    Case 12
           mni_cloth_outward_Click
        Case 13
           mni_yarn_outward_Click
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Parent.Index = 1 Then
    Select Case ButtonMenu.Index
        Case 1
            mni_parties_Code_Click
        Case 2
            mni_cloth_code_Click
        Case 3
            mni_Yarn_Code_Click
        Case 4
            mni_machine_Code_Click
        Case 5
            mni_employee_code_Click
    End Select
End If
            
If ButtonMenu.Parent.Index = 2 Then
    Select Case ButtonMenu.Index
        Case 1
           mni_knit_purc_contract_Click
        Case 2
            mni_knit_sale_contract_Click
        Case 3
            mni_dye_purc_contract_Click
    End Select
End If
            
If ButtonMenu.Parent.Index = 3 Then
    Select Case ButtonMenu.Index
        Case 1
           mni_ran_inward_Click
        Case 2
            mni_cloth_Inward_Click
        Case 3
            mni_cloth_dye_InWard_Click
    End Select
End If
            
If ButtonMenu.Parent.Index = 4 Then
    Select Case ButtonMenu.Index
        Case 1
           mni_cloth_outward_Click
        Case 2
           mni_yarn_outward_Click
        Case 3
           mni_cloth_dye_outward_Click
    End Select
End If

If ButtonMenu.Parent.Index = 6 Then
    Select Case ButtonMenu.Index
        Case 1
           mni_party_rep_Click
        Case 2
           mni_CLoth_rep_Click
        Case 3
           mni_yarn_list_rep_Click
        Case 4
           mni_Machines_List_Rep_Click
        Case 5
            mni_Emp_List_Rep_Click
    End Select
End If

If ButtonMenu.Parent.Index = 8 Then
    Select Case ButtonMenu.Index
        Case 1
           mni_Pur_Knitt_Con_Print_Click
        Case 2
          mni_Sale_Knitt_Cont_Print_Click
        Case 3
          mni_Pur_Dye_Cont_Print_Click
        
    End Select
End If

If ButtonMenu.Parent.Index = 10 Then
    Select Case ButtonMenu.Index
        Case 1
           mni_Inward_Yarn_Rec_Click
        Case 2
          mni_inward_cloth_rec_Click
        Case 3
         mni_cloth_rec_after_dye_rep_Click
        Case 4
         mni_outward_yarn_sent_knit_Click
        Case 5
         mni_outward_cloth_sent_knit_Click
        Case 6
         mni_outward_cloth_sent_dye_Click
    End Select
End If

End Sub

Private Sub y_inward_Click()
Load vou_r
vou_r.Text1.Text = 12
vou_r.Frame2.Visible = True
vou_r.Caption = "Yarn Wise Inwards"
vou_r.Show

End Sub

Private Sub y_issue_mac_Click()
Load vour
vour.Caption = "Yarn Issued To Machine"
vour.Text2.Text = 66
vour.Label1.Caption = "Issue #"
vour.Show

End Sub

Private Sub y_issue_mac_dates_Click()
Load vou_r
vou_r.Caption = "Yarn Issued To Machines"
vou_r.Text1.Text = 19
vou_r.Show

End Sub

Private Sub y_outward_Click()
Load vou_r
vou_r.Text1.Text = 14
vou_r.Frame2.Visible = True
vou_r.Caption = "Yarn Wise In/Out Inventory Ledger"
vou_r.Show

End Sub

Private Sub yarn_inv_led_Click()
Load vou_r
vou_r.Text1.Text = 10
vou_r.Frame2.Visible = True
vou_r.Caption = "Yarn Wise In/Out Inventory Ledger"
vou_r.Show

End Sub
