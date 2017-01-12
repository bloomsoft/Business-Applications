VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form notes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "(F3) to Search"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   390
   End
   Begin Crystal.CrystalReport r1 
      Left            =   2040
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   1095
      Left            =   2640
      Picture         =   "notes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Prev"
      Height          =   1095
      Left            =   600
      Picture         =   "notes.frx":0561
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "(F2) to Search"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Account Title"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Account ID"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60358659
         CurrentDate     =   37422
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   60358659
         CurrentDate     =   37422
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Blm1 As New bloom1
Private Blmr As New Bloom_r

Private Sub Command1_Click()
Dim f As String
Screen.MousePointer = vbHourglass
    If Val(Text3.Text) = 1 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\PurchaseW.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    If Val(Text3.Text) = 2 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Inward_vw_f.party} = " & Val(Text1.Text)
        r1.ReportFileName = App.Path & "\PurchaseParty.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    If Val(Text3.Text) = 3 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Inward_vw_f.item} = " & Val(Text4.Text)
        r1.ReportFileName = App.Path & "\PurchaseItems.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 4 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\PurchaseS.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 5 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Inward_vw_f.party} = " & Val(Text1.Text)
        r1.ReportFileName = App.Path & "\PurchasePartyS.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 6 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Inward_vw_f.item} = " & Val(Text4.Text)
        r1.ReportFileName = App.Path & "\PurchaseItemsS.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 7 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.ReportFileName = App.Path & "\PurchaseT.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 8 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Inward_vw_f.party} = " & Val(Text1.Text)
        r1.ReportFileName = App.Path & "\PurchasePartyT.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 9 Then
        f = "{inward_vw_f.a.a.v_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {Inward_vw_f.item} = " & Val(Text4.Text)
        r1.ReportFileName = App.Path & "\PurchaseItemsT.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 51 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SaleW.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 52 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {sale_vw_Final.Item} = " & Val(Text4.Text)
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SaleItem.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 53 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {sale_vw_Final.Party} = " & Val(Text1.Text)
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SaleParty.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    If Val(Text3.Text) = 54 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SaleS.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 55 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {sale_vw_Final.Item} = " & Val(Text4.Text)
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SaleItemS.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 56 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {sale_vw_Final.Party} = " & Val(Text1.Text)
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SalePartyS.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

If Val(Text3.Text) = 57 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SaleT.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 58 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {sale_vw_Final.Item} = " & Val(Text4.Text)
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SaleItemT.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 59 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {sale_vw_Final.Party} = " & Val(Text1.Text)
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\SalePartyT.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

    If Val(Text3.Text) = 60 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\invoice.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 61 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\invoiceSocks.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 62 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\invoiceTowels.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If

If Val(Text3.Text) = 151 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\TotalSales.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
End If

If Val(Text3.Text) = 152 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\TotalSalesS.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
End If

If Val(Text3.Text) = 153 Then
        f = "{Sale_vw_final.inv_date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportFileName = App.Path & "\TotalSalesT.rpt"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
End If

    If Val(Text3.Text) = 200 Then
        f = "{PContractVW.Cont_Date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {PContractVW.SellerCode} = " & Val(Text1.Text)
        r1.ReportFileName = App.Path & "\PContract.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 201 Then
        f = "{SContractVW.Cont_Date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {SContractVW.SellerCode} = " & Val(Text1.Text)
        r1.ReportFileName = App.Path & "\SContract.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 202 Then
        f = "{PContractVW.Cont_Date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {PContractVW.ClothCode} = " & Val(Text4.Text)
        r1.ReportFileName = App.Path & "\PContract.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
    
    If Val(Text3.Text) = 203 Then
        f = "{SContractVW.Cont_Date} in Date(" & Format(DTPicker1.Value, "yyyy,MM,dd") & ")"
        f = f & " To Date(" & Format(DTPicker2.Value, "yyyy,MM,dd") & ")"
        f = f & " and {SContractVW.ClothCode} = " & Val(Text4.Text)
        r1.ReportFileName = App.Path & "\SContract.rpt"
        r1.DataFiles(0) = App.Path & "\Bloom.mdb"
        r1.ReportTitle = "From : " & Format(DTPicker1.Value, "dd/MM/yyyy") & " To : " & Format(DTPicker2.Value, "dd/MM/yyyy")
        r1.SelectionFormula = f
        r1.WindowTop = 0
        r1.WindowLeft = 0
        r1.WindowState = crptMaximized
        r1.Action = 1
        r1.PageZoom 100
    End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    Search2.Text3.Text = 6
    Search2.Show
End If
If KeyCode = vbKeyF3 Then
    Search1.Text3.Text = 6
    Search1.Show
End If

End Sub

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Val(Text1.Text) > 0 Then
    Text2.Text = Blm1.party1(Val(Text1.Text))
    If Text2.Text = "NOT" Then
        MsgBox "Invalid Account Code"
        
    End If
    
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Val(Text4.Text) > 0 Then
    Text5.Text = Blm1.Item1(Val(Text4.Text))
    If Text5.Text = "NOT" Then
        MsgBox "Invalid Item Code"
        
    End If
    
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Exit Sub
Else
    If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    Else
        KeyAscii = 0
        Beep
    End If
End If
End Sub
