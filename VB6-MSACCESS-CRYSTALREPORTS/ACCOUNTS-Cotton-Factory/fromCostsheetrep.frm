VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCostsheetrep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cost Sheet & Stock Reports"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3510
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   810
      TabIndex        =   7
      Top             =   690
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1260
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1260
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   56098819
      CurrentDate     =   39206
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   780
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "01/MM/yyyy"
      Format          =   56098819
      CurrentDate     =   39206
   End
   Begin Crystal.CrystalReport R1 
      Left            =   1920
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Ref.No."
      Height          =   255
      Left            =   210
      TabIndex        =   6
      Top             =   690
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCostsheetrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Private Sub GenData()
Dim TBT As Recordset
Dim DBT As Database
Dim DBM As Database
Dim TBM As Recordset
Dim S As String

Dim ThisDate As Date
Dim d As Integer
Dim Diff As Integer

Diff = Abs(DateDiff("d", DTPicker1.Value, DTPicker2.Value)) + 1
'cntl.Max = Abs(DateDiff("d", DTPicker1.Value, DTPicker2.Value)) + 1
'MsgBox "Test"
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")
S = "Delete from CostSheetRep"
DBT.Execute S

Set TBT = DBT.OpenRecordset("CostSheetRep", dbOpenTable)

For d = 1 To Diff
    ThisDate = DateAdd("d", d - 1, DTPicker1.Value)
    TBT.AddNew
    TBT.Fields("VDate").Value = ThisDate
    TBT.Fields("RefNo").Value = Val(Text1.Text)
        S = "Select * from Expences where EDate=#" & ThisDate & "#"
        S = S & " and RefNo = " & Val(Text1.Text)
        Set TBM = DBM.OpenRecordset(S)
        If Not TBM.EOF Then
            TBT.Fields("Electric").Value = TBM.Fields("ElectricAmount").Value & ""
            TBT.Fields("Maintain").Value = TBM.Fields("Maintain").Value & ""
            TBT.Fields("Salaries").Value = Val(TBM.Fields("Salaries").Value & "") + Val(TBM.Fields("Misc").Value & "")
            TBT.Fields("Contractor").Value = TBM.Fields("Contractor").Value & ""
            
        Else
            TBT.Fields("Electric").Value = 0
            TBT.Fields("Maintain").Value = 0
            TBT.Fields("Salaries").Value = 0
            TBT.Fields("Contractor").Value = 0
        End If
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,sum(Amount) as A from Issue where V_Date=#" & ThisDate & "#"
        S = S & " and RefNo = " & Val(Text1.Text)
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("Q").Value) Then
            TBT.Fields("IssueWT").Value = TBM.Fields("Q").Value & ""
        Else
            TBT.Fields("IssueWT").Value = 0
        End If
        If Not IsNull(TBM.Fields("A").Value) Then
            TBT.Fields("MaterialCost").Value = TBM.Fields("A").Value & ""
        Else
            TBT.Fields("MaterialCost").Value = 0
        End If
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,Avg(rate) as R,sum(Rate * Qty) as A from Production where V_Date=#" & ThisDate & "#"
        S = S & " and RefNo = " & Val(Text1.Text)
        S = S & " and ItemCode in (Select Distinct Code from Items where IType=1)"
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("Q").Value) Then
            TBT.Fields("ProductionWT").Value = TBM.Fields("Q").Value & ""
        Else
            TBT.Fields("ProductionWT").Value = 0
        End If
        If Not IsNull(TBM.Fields("R").Value) Then
            TBT.Fields("ProductionRate").Value = TBM.Fields("A").Value / TBM.Fields("Q").Value ' & ""
        Else
            TBT.Fields("ProductionRate").Value = 0
        End If
        If Not IsNull(TBM.Fields("B").Value) Then
            TBT.Fields("ProductionBales").Value = TBM.Fields("B").Value & ""
        Else
            TBT.Fields("ProductionBales").Value = 0
        End If
        
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,Avg(rate) as R,sum(Rate * Qty) as A from Sales where V_Date=#" & ThisDate & "#"
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("B").Value) Then
            TBT.Fields("SaleBales").Value = TBM.Fields("B").Value & ""
        Else
            TBT.Fields("SaleBales").Value = 0
        End If
                
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,Avg(rate) as R,sum(Rate * Qty) as A from SalesReturn where V_Date=#" & ThisDate & "#"
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("B").Value) Then
            TBT.Fields("ProductionBales").Value = TBT.Fields("ProductionBales").Value + TBM.Fields("B").Value & ""
        End If
        TBM.Close
        
        
        
     TBT.Update
Next d
TBT.Close
DBT.Close
DBM.Close
End Sub

Private Sub GenDataSH()
Dim TBT As Recordset
Dim DBT As Database
Dim DBM As Database
Dim TBM As Recordset
Dim S As String

Dim ThisDate As Date
Dim d As Integer
Dim Diff As Integer

Diff = Abs(DateDiff("d", DTPicker1.Value, DTPicker2.Value)) + 1
'cntl.Max = Abs(DateDiff("d", DTPicker1.Value, DTPicker2.Value)) + 1
'MsgBox "Test"
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")
S = "Delete from CostSheetRep"
DBT.Execute S

Set TBT = DBT.OpenRecordset("CostSheetRep", dbOpenTable)
'Mgsbox Diff
For d = 1 To Diff
    ThisDate = DateAdd("d", d - 1, DTPicker1.Value)
    TBT.AddNew
    TBT.Fields("VDate").Value = ThisDate
    TBT.Fields("RefNo").Value = Val(Text1.Text)
        S = "Select * from Expences where EDate=#" & ThisDate & "#"
        S = S & " and RefNo = " & Val(Text1.Text)
        Set TBM = DBM.OpenRecordset(S)
        If Not TBM.EOF Then
            TBT.Fields("Electric").Value = TBM.Fields("ElectricAmount").Value & ""
            TBT.Fields("Maintain").Value = TBM.Fields("Maintain").Value & ""
            TBT.Fields("Salaries").Value = Val(TBM.Fields("Salaries").Value & "") + Val(TBM.Fields("Misc").Value & "")
            TBT.Fields("Contractor").Value = TBM.Fields("Contractor").Value & ""
        Else
            TBT.Fields("Electric").Value = 0
            TBT.Fields("Maintain").Value = 0
            TBT.Fields("Salaries").Value = 0
            TBT.Fields("Contractor").Value = 0
        End If
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,sum(Amount) as A from IssueSH where V_Date=#" & ThisDate & "#"
        S = S & " and RefNo = " & Val(Text1.Text)
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("Q").Value) Then
            TBT.Fields("IssueWT").Value = TBM.Fields("Q").Value & ""
        Else
            TBT.Fields("IssueWT").Value = 0
        End If
        If Not IsNull(TBM.Fields("A").Value) Then
            TBT.Fields("MaterialCost").Value = TBM.Fields("A").Value & ""
        Else
            TBT.Fields("MaterialCost").Value = 0
        End If
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,(sum(Rate * Qty)/Sum(Qty)) as R,sum(Rate * Qty) as A from Production where V_Date=#" & ThisDate & "#"
        S = S & " and RefNo = " & Val(Text1.Text)
        S = S & " and ItemCode in (Select Distinct Code from Items where IType=1)"
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("Q").Value) Then
            TBT.Fields("ProductionWT").Value = TBM.Fields("Q").Value & ""
        Else
            TBT.Fields("ProductionWT").Value = 0
        End If
        If Not IsNull(TBM.Fields("R").Value) Then
            TBT.Fields("ProductionRate").Value = TBM.Fields("R").Value & ""
        Else
            TBT.Fields("ProductionRate").Value = 0
        End If
        If Not IsNull(TBM.Fields("B").Value) Then
            TBT.Fields("ProductionBales").Value = TBM.Fields("B").Value & ""
        Else
            TBT.Fields("ProductionBales").Value = 0
        End If
        
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,Avg(rate) as R,sum(Rate * Qty) as A from Sales where V_Date=#" & ThisDate & "#"
'        MsgBox s
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("B").Value) Then
            TBT.Fields("SaleBales").Value = TBM.Fields("B").Value & ""
        Else
            TBT.Fields("SaleBales").Value = 0
        End If
                
        TBM.Close
        
        S = "Select Sum(Qty) as Q,Sum(Bales) as B,Avg(rate) as R,sum(Rate * Qty) as A from SalesReturn where V_Date=#" & ThisDate & "#"
        Set TBM = DBM.OpenRecordset(S)
        If Not IsNull(TBM.Fields("B").Value) Then
            TBT.Fields("ProductionBales").Value = TBT.Fields("ProductionBales").Value + TBM.Fields("B").Value & ""
        End If
        TBM.Close
        
        
        
     TBT.Update
Next d
TBT.Close
DBT.Close
DBM.Close
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Val(Text2.Text) = 1 Then
    GenData
    R1.ReportTitle = "(Item Wise) " & vbCrLf & "From : " & Format(DTPicker1.Value, "dd-MM-yyyy") & " To : " & Format(DTPicker2.Value, "dd-MM-yyyy")
End If
If Val(Text2.Text) = 2 Then
    GenDataSH
    R1.ReportTitle = "(Sub Head Wise) " & vbCrLf & "From : " & Format(DTPicker1.Value, "dd-MM-yyyy") & " To : " & Format(DTPicker2.Value, "dd-MM-yyyy")
End If
'MsgBox "Test1"
DoEvents
R1.ReportFileName = App.path & "\" & "CostSheetAndStock.rpt"
R1.DataFiles(0) = App.path & "\Book.mdb"

R1.WindowTop = 0
R1.WindowLeft = 0
R1.WindowState = crptMaximized
R1.Action = 1

Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub DTPicker1_LostFocus()
DTPicker1.Value = CDate("01/" & MonthName(DTPicker1.Month, True) & "/" & DTPicker1.Year)
End Sub

Private Sub Form_Load()
DTPicker1.Value = FStartDate
DTPicker2.Value = Date
End Sub
