VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bloom_r"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim t_path As String
Dim m_path As String
Private Blm As bloom1
Public Sub DailyIssueHeadWise(SDate As Date, EDate As Date, Optional VNo As Long, Optional RefNo As Long)
Dim db_t As DAO.Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim WPB As CottonIssue
Dim Ssql As String
Dim Db_M As Database
Set Db_M = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set db_t = OpenDatabase(App.path & "\Book.mdb")
Ssql = "delete from DailyIssue"
db_t.Execute Ssql

Set tb_t = db_t.OpenRecordset("DailyIssue", dbOpenTable)

Ssql = "Select * from IssueSH where V_date Between #" & SDate & "# and #" & EDate & "#"
If VNo > 0 Then
    Ssql = Ssql & " and VNo=" & VNo
End If
If RefNo > 0 Then
    Ssql = Ssql & " and RefNo=" & RefNo
End If
'Ssql = Ssql & " Group By ItemCode"

'MsgBox Ssql

Set TB = Db_M.OpenRecordset(Ssql)

If Not TB.EOF Then
    Do While Not TB.EOF
    'WPB = bloom.AvgRateAndWeightHeadWise(EDate, TB.Fields("HCode").Value)
    tb_t.AddNew
        tb_t.Fields("EDate").Value = TB.Fields("V_Date").Value
        tb_t.Fields("VNo").Value = TB.Fields("VNo").Value
        tb_t.Fields("RefNo").Value = TB.Fields("RefNo").Value
        tb_t.Fields("HCode").Value = TB.Fields("ItemCode").Value
        tb_t.Fields("HName").Value = Blm.SubGroupName(TB.Fields("ItemCode").Value)
        tb_t.Fields("Bales").Value = TB.Fields("Bales").Value
        tb_t.Fields("Rate").Value = TB.Fields("Rate").Value
        tb_t.Fields("WTPerBale").Value = TB.Fields("AvgBaleWT").Value
        tb_t.Fields("Weight").Value = TB.Fields("Qty").Value
        tb_t.Fields("Pcent").Value = (TB.Fields("Percent").Value)
        tb_t.Fields("Amount").Value = TB.Fields("Amount").Value
    tb_t.Update
    TB.MoveNext
    Loop
End If
TB.Close
tb_t.Close
Db_M.Close
db_t.Close

End Sub
Public Sub DailyProduction(SDate As Date, EDate As Date, Optional VNo As Long, Optional RefNo As Long, Optional ItemCode As Long)
Dim db_t As DAO.Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim WPB As CottonIssue
Dim Ssql As String
Dim Db_M As Database
Set Db_M = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set db_t = OpenDatabase(App.path & "\Book.mdb")
Ssql = "delete from DailyIssue"
db_t.Execute Ssql

Set tb_t = db_t.OpenRecordset("DailyIssue", dbOpenTable)

Ssql = "Select * from Production where V_date Between #" & SDate & "# and #" & EDate & "#"
If VNo > 0 Then
    Ssql = Ssql & " and VNo=" & VNo
End If
If RefNo > 0 Then
    Ssql = Ssql & " and RefNo=" & RefNo
End If
If ItemCode > 0 Then
    Ssql = Ssql & " and ItemCode=" & ItemCode
End If

'Ssql = Ssql & " Group By ItemCode"

'MsgBox Ssql

Set TB = Db_M.OpenRecordset(Ssql)

If Not TB.EOF Then
    Do While Not TB.EOF
    'WPB = bloom.AvgRateAndWeightHeadWise(EDate, TB.Fields("HCode").Value)
    tb_t.AddNew
        tb_t.Fields("EDate").Value = TB.Fields("V_Date").Value
        tb_t.Fields("VNo").Value = TB.Fields("VNo").Value
        tb_t.Fields("RefNo").Value = TB.Fields("RefNo").Value
        tb_t.Fields("HCode").Value = TB.Fields("ItemCode").Value
        tb_t.Fields("HName").Value = Blm.item1(TB.Fields("ItemCode").Value)
        tb_t.Fields("Bales").Value = TB.Fields("Bales").Value
        tb_t.Fields("Rate").Value = TB.Fields("Rate").Value
        'tb_t.Fields("WTPerBale").Value = TB.Fields("AvgBaleWT").Value
        tb_t.Fields("Weight").Value = TB.Fields("Qty").Value
        'tb_t.Fields("Pcent").Value = (TB.Fields("Percent").Value)
        tb_t.Fields("Amount").Value = TB.Fields("Rate").Value * TB.Fields("Qty").Value
        tb_t.Fields("Description").Value = Blm.party1(TB.Fields("CrCode").Value)
    tb_t.Update
    TB.MoveNext
    Loop
End If
TB.Close
tb_t.Close
Db_M.Close
db_t.Close

End Sub

Public Sub Aging(V_Date As Date, CNTL As Control, Cntl2 As Control)
Dim db_t As DAO.Database
Dim test As Integer
Dim TB As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim tb_s As Recordset
Dim t_cred As Double, T_Deb As Double
Dim Ssql As String
Dim tbtrial As Recordset, I As Double
Dim TB_C As Recordset
Dim Db_M As Database
Set Db_M = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set db_t = OpenDatabase(App.path & "\Book.mdb")
Ssql = "delete from aging"
db_t.Execute Ssql
CNTL.Value = 0
Set tb_t = db_t.OpenRecordset("Aging", dbOpenTable)
Ssql = "SELECT * FROM ACCHART ORDER BY CODE"
Set TB_C = Db_M.OpenRecordset(Ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
CNTL.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    
    Cntl2.SimpleText = TB_C.Fields("Name").Value
        Ssql = "SELECT Val(SUM(CREDIT)) AS DEB FROM VOUDTL WHERE PARTY = " & TB_C.Fields("CODE").Value
        Set TB = Db_M.OpenRecordset(Ssql)
        If Not TB.EOF Then
        If Not IsNull(TB.Fields("DEB").Value) Then
            t_cred = TB.Fields("DEB").Value
        End If
        End If
'        MsgBox TB_C.Fields("Name").Value & " Balance " & t_cred
        TB.Close
        
        Cntl2.SimpleText = Cntl2.SimpleText & " Total Due Balance = " & t_cred
     'TOTAL Sales ITEMS ADJUSTMENTS
     
        

        Ssql = "SELECT V_Date,V_no,Sum(Amount) as Total FROM Sales WHERE Seller = " & TB_C.Fields("CODE").Value
        Ssql = Ssql & " Group BY V_DATE,V_no Order by V_Date,V_no"
        Set TB = Db_M.OpenRecordset(Ssql)
        
        If Not TB.EOF Then
            TB.MoveLast
'            MsgBox tb.Fields("V_Date").Value & " " & tb.Fields("Ent_no").Value
            TB.MoveFirst
            Do While Not TB.EOF
                I = I + 1
                If t_cred > 0 Then
                    If Not IsNull(TB.Fields("Total").Value) Then
        '                Debug.Print TB.Fields("V_date").Value & " " & TB.Fields("Ent_no").Value & " Bill Amount " & TB.Fields("Total").Value & " = " & t_cred
                        'If tb.Fields("Ent_no").Value = 74 Then
                        
                        '    MsgBox tb.Fields("V_date").Value & " " & tb.Fields("Ent_no").Value & " " & t_cred
                        'End If
                        T_Deb = T_Deb + TB.Fields("Total").Value
                        t_cred = t_cred - TB.Fields("Total").Value
                    End If
                   
                    If Round(t_cred) < 0 Then
                        Exit Do
                    ElseIf Round(t_cred) = 0 Then
                        TB.MoveNext
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            TB.MoveNext
            Loop
' MsgBox "Test"
            If Not TB.EOF Then
            Do While Not TB.EOF
                    Ssql = "select Days,Val(Sum(Qty)) as Quantity,Val(Sum(Bales)) as Bales1,Val(Sum(Amount)) as Total from Sales where V_no = " & TB.Fields("V_no").Value
                    Ssql = Ssql & " Group by Days"
                    Set tb_s = Db_M.OpenRecordset(Ssql)
                    'MsgBox Ssql
                If Not tb_s.EOF Then
                    DoEvents
                    Cntl2.SimpleText = Cntl2.SimpleText & " Total Paid = " & t_cred
                    'If tb.Fields("Total").Value <> 0 Then
                    
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = TB.Fields("v_date").Value
                    tb_t.Fields("Sale_no").Value = TB.Fields("V_no").Value
                    tb_t.Fields("duedate").Value = TB.Fields("v_date").Value + tb_s.Fields("days").Value
                    tb_t.Fields("Due_Days").Value = tb_s.Fields("days").Value
                    tb_t.Fields("accode").Value = AcCode
                    tb_t.Fields("name").Value = Blm.party1(TB_C.Fields("Code").Value)
                    tb_t.Fields("Quantity").Value = tb_s.Fields("Quantity").Value
                    tb_t.Fields("Quantity").Value = tb_s.Fields("Bales1").Value
                    tb_t.Fields("DaysGone").Value = Abs(DateDiff("d", TB.Fields("V_date"), V_Date))
                    tb_t.Fields("OverDays").Value = Abs(DateDiff("d", TB.Fields("V_date") + tb_s.Fields("Days").Value, V_Date))
                    
                    If t_cred < 0 Then
                        t_cred = Abs(t_cred) 'tb.Fields("Total").Value + t_cred
                    Else
                        If Not IsNull(TB.Fields("Total").Value) Then
                        t_cred = TB.Fields("Total").Value
                        End If
                    End If
'                    MsgBox t_cred
                    tb_t.Fields("Amount").Value = t_cred
                    tb_t.Fields("End_Date").Value = V_Date
                    tb_t.Update
                    
                Else
                    'MsgBox "No Bill Details Found"
                End If
                    'End If
                    tb_s.Close
                
                TB.MoveNext
'                If Not tb.EOF Then
'                    If Not IsNull(tb.Fields("Total").Value) Then
'                        t_cred = tb.Fields("Total").Value
'                    End If
'                End If
            Loop
            End If
            't_cred = 0
     
        TB.Close
    End If
      TB_C.MoveNext
      CNTL.Value = CNTL.Value + 1
      DoEvents
    Loop
End If
TB_C.Close
tb_t.Close
db_t.Close
Db_M.Close
End Sub

Public Sub CostSheet(S_Date As Date, E_Date As Date, CNTL As Object)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim Ssql As String
Dim B As Double
Dim ProductionKGS As Double, IssueKGS As Double, IssueAMT As Double
Dim ExpencesAMT As Double

CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)

'Sales Value
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2) in ('21','22')"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 1
        tb_t.Fields("Code").Value = 1
        tb_t.Fields("name").Value = "Sales"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close

'Purchase Value
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2)='45'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 1
        tb_t.Fields("Code").Value = 2
        tb_t.Fields("name").Value = "Purchases & Expences"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close

'Stocks Value
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2)='20'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 2
        tb_t.Fields("Code").Value = 1
        tb_t.Fields("name").Value = "Stocks of Raw Materials"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close


'Stock Web
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2)='20'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 2
        tb_t.Fields("Code").Value = 2
        tb_t.Fields("name").Value = "Stock Web"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close

'Receiveables
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,5)='31001'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 2
        tb_t.Fields("Code").Value = 3
        tb_t.Fields("name").Value = "Receiveable"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close

'Cash With Banks
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,5)='36002'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 2
        tb_t.Fields("Code").Value = 4
        tb_t.Fields("name").Value = "Cash With Banks"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close


'Cash in Hand
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,5)='36001'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 2
        tb_t.Fields("Code").Value = 5
        tb_t.Fields("name").Value = "Cash in Hand"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close

'Material Cost
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2)='22'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    IssueAMT = Abs(TB.Fields("Bal").Value)
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 3
        tb_t.Fields("Code").Value = 1
        tb_t.Fields("name").Value = "Material Cost"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close

'Material Issued KGS
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2)='22'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    IssueKGS = Abs(TB.Fields("Bal").Value)
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 3
        tb_t.Fields("Code").Value = 2
        tb_t.Fields("name").Value = "Material Issued (KGS)"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close


'Production
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2)='22'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    ProductionKGS = Abs(TB.Fields("Bal").Value)
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 3
        tb_t.Fields("Code").Value = 3
        tb_t.Fields("name").Value = "Production"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close


'Expences
Ssql = "Select Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
Ssql = Ssql & " and Mid(Party,1,2)='58'"
Set TB = Db_M.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    ExpencesAMT = Abs(TB.Fields("Bal").Value)
    tb_t.AddNew
        tb_t.Fields("S_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("h_code").Value = 3
        tb_t.Fields("Code").Value = 4
        tb_t.Fields("name").Value = "Expences"
        If TB.Fields("Bal").Value > 0 Then B = TB.Fields("Bal").Value
        If TB.Fields("Bal").Value < 0 Then B = TB.Fields("Bal").Value * -1
        tb_t.Fields("Bal").Value = B
    tb_t.Update
    B = 0
End If
TB.Close
tb_t.Close

Ssql = "Update Trial Set Op_Deb=" & IssueAMT
db_t.Execute Ssql
Ssql = "Update Trial Set Op_Cred=" & IssueKGS
db_t.Execute Ssql
Ssql = "Update Trial Set Debit=" & ProductionKGS
db_t.Execute Ssql
Ssql = "Update Trial Set Credit=" & ExpencesAMT
db_t.Execute Ssql


db_t.Close
Db_M.Close

End Sub
Public Sub ClosingStocks(S_Date As Date, E_Date As Date, CNTL As Object, Optional HCode As String, Optional SHCode As String)
Dim DBT As Database
Dim DBM As Database

Dim PQty As Double, PBales As Double
Dim TBS As Recordset
Dim TBSR As Recordset
Dim TBD As Recordset
Dim TBI As Recordset
Dim TBDR As Recordset
Dim TBDSH As Recordset
Dim TBIR As Recordset
Dim TBP As Recordset
Dim TBPR As Recordset
Dim TBT As Recordset

Dim TBCS As Recordset
Dim TBCSR As Recordset
Dim TBCD As Recordset
Dim TBCI As Recordset
Dim TBCDSH As Recordset
Dim TBCDR As Recordset
Dim TBCIR As Recordset

Dim TBCP As Recordset
Dim TBCPR As Recordset
Dim TBCT As Recordset


Dim Itemname As String
Dim MillName As String
Dim Ssql As String
Dim TBItm As Recordset
Dim Tbmill As Recordset
Dim StockQty As Double, StockBales As Double, ARate As Double
Dim StockQty2 As Double, StockBales2 As Double, ARate2 As Double
Dim StockAmount As Double
Dim IssueQty As Double, IssueBales As Double
Dim SaleQty As Double, SaleBales As Double
Dim ProductionQty As Double, ProductionBales As Double
Dim OpBales As Double, OpQty As Double, OpAmount As Double
Dim PurQty As Double, PurBales As Double
Dim INQty As Double, InBales As Double, InAmount As Double
Dim OutQty As Double, OutBales As Double, OutAmount As Double
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")
CNTL.Value = 0
Ssql = "Delete from Stock"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("Stock", dbOpenTable)

Ssql = "Select a.*,b.Name as GroupName,c.name as SubGroupName from Items a,Groups b, SubGroups c where a.GroupCode=b.Code and a.SubGroupCode=c.Code "
If Len(HCode) > 0 Then
    Ssql = Ssql & " and a.GroupCode in (" & HCode & ")"
End If
If Len(SHCode) > 0 Then
    Ssql = Ssql & " and a.SubGroupCode in (" & SHCode & ")"
End If
Ssql = Ssql & " Order by a.Code"
Set TBItm = DBM.OpenRecordset(Ssql)
If TBItm.EOF = False Then
    TBItm.MoveLast
    CNTL.Max = TBItm.RecordCount
    TBItm.MoveFirst
End If

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales)as B,Sum(Amount) as Amt from Purchase Where V_date < #" & S_Date & "# Group By Item"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    TBP.MoveFirst
End If

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales)as B,Sum(Amount) as Amt from PurchaseReturn Where V_date < #" & S_Date & "# Group By Item"
Set TBPR = DBM.OpenRecordset(Ssql)
If Not TBPR.EOF Then
    TBPR.MoveLast
    TBPR.MoveFirst
End If


Ssql = "Select Item,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from Sales Where V_date < #" & S_Date & "# Group By Item"
Set TBS = DBM.OpenRecordset(Ssql)
If Not TBS.EOF Then
    TBS.MoveLast
    TBS.MoveFirst
End If

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from SalesReturn Where V_date < #" & S_Date & "# Group By Item"
Set TBSR = DBM.OpenRecordset(Ssql)
If Not TBSR.EOF Then
    TBSR.MoveLast
    TBSR.MoveFirst
End If

Ssql = "Select ItemCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from Issue Where V_date < #" & S_Date & "# Group By ItemCode"
Set TBI = DBM.OpenRecordset(Ssql)
If Not TBI.EOF Then
    TBI.MoveLast
    TBI.MoveFirst
End If

Ssql = "Select ItemCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Qty*Rate) as Amt from Production Where V_date < #" & S_Date & "# Group By ItemCode"
Set TBD = DBM.OpenRecordset(Ssql)
If Not TBD.EOF Then
    TBD.MoveLast
    TBD.MoveFirst
End If

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales) as B,Sum(Qty*Rate) as Amt from SHProduction Where V_date < #" & S_Date & "# Group By Item"
Set TBDSH = DBM.OpenRecordset(Ssql)
If Not TBDSH.EOF Then
    TBDSH.MoveLast
    TBDSH.MoveFirst
End If

Ssql = "Select DrCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from Issue Where V_date < #" & S_Date & "# Group By DrCode"
Set TBIR = DBM.OpenRecordset(Ssql)
If Not TBIR.EOF Then
    TBIR.MoveLast
    TBIR.MoveFirst
End If

Ssql = "Select CrCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Qty*Rate) as Amt from Production Where V_date < #" & S_Date & "# Group By CrCode"
Set TBDR = DBM.OpenRecordset(Ssql)
If Not TBDR.EOF Then
    TBDR.MoveLast
    TBDR.MoveFirst
End If


''''Current

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales)as B,Sum(Amount) as Amt from Purchase Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
Set TBCP = DBM.OpenRecordset(Ssql)
If Not TBCP.EOF Then
    TBCP.MoveLast
    TBCP.MoveFirst
End If

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales)as B,Sum(Amount) as Amt from PurchaseReturn Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
Set TBCPR = DBM.OpenRecordset(Ssql)
If Not TBCPR.EOF Then
    TBCPR.MoveLast
    TBCPR.MoveFirst
End If


Ssql = "Select Item,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from Sales Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
Set TBCS = DBM.OpenRecordset(Ssql)
If Not TBCS.EOF Then
    TBCS.MoveLast
    TBCS.MoveFirst
End If

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from SalesReturn Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
Set TBCSR = DBM.OpenRecordset(Ssql)
If Not TBCSR.EOF Then
    TBCSR.MoveLast
    TBCSR.MoveFirst
End If

Ssql = "Select ItemCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from Issue Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By ItemCode"
Set TBCI = DBM.OpenRecordset(Ssql)
If Not TBCI.EOF Then
    TBCI.MoveLast
    TBCI.MoveFirst
End If

Ssql = "Select ItemCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Qty*rate) as Amt from Production Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By ItemCode"
Set TBCD = DBM.OpenRecordset(Ssql)
If Not TBCD.EOF Then
    TBCD.MoveLast
    TBCD.MoveFirst
End If

Ssql = "Select Item,Sum(Qty) as Q,Sum(bales) as B,Sum(Qty*rate) as Amt from SHProduction Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
Set TBCDSH = DBM.OpenRecordset(Ssql)
If Not TBCDSH.EOF Then
    TBCDSH.MoveLast
    TBCDSH.MoveFirst
End If


Ssql = "Select DrCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Amount) as Amt from Issue Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By DrCode"
Set TBCIR = DBM.OpenRecordset(Ssql)
If Not TBCIR.EOF Then
    TBCIR.MoveLast
    TBCIR.MoveFirst
End If

Ssql = "Select CrCode,Sum(Qty) as Q,Sum(bales) as B,Sum(Qty*rate) as Amt from Production Where V_date Between #" & S_Date & "# and #" & E_Date & "# Group By CrCode"
Set TBCDR = DBM.OpenRecordset(Ssql)
If Not TBCDR.EOF Then
    TBCDR.MoveLast
    TBCDR.MoveFirst
End If




Do While Not TBItm.EOF
            
            Itemname = ""
            OpBales = 0
            OpQty = 0
            OpAmount = 0
                        
            Itemname = TBItm.Fields("Name").Value & ""
            OpBales = Val(TBItm.Fields("OPBales").Value & "")
            OpQty = Val(TBItm.Fields("OPWT").Value & "")
            OpAmount = Val(TBItm.Fields("Rate").Value & "") ' * Val(TBItm.Fields("OpWT").Value & "")
            
            If TBI.RecordCount > 0 Then
                TBI.MoveFirst
                TBI.FindFirst "ItemCode=" & TBItm.Fields("Code").Value
                If TBI.NoMatch = False Then
                    OpQty = OpQty - (TBI.Fields("Q").Value)
                    OpBales = OpBales - (TBI.Fields("B").Value)
                    OpAmount = OpAmount - (TBI.Fields("Amt").Value)
                End If
            End If
            
           If TBIR.RecordCount > 0 Then
               TBIR.MoveFirst
                TBIR.FindFirst "DrCode=" & TBItm.Fields("Code").Value
                If TBIR.NoMatch = False Then
                    OpQty = OpQty + (TBIR.Fields("Q").Value)
                    OpBales = OpBales + (TBIR.Fields("B").Value)
                    OpAmount = OpAmount + (TBIR.Fields("Amt").Value)
                End If
            End If

    
            If TBP.RecordCount > 0 Then
                TBP.MoveFirst
                TBP.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBP.NoMatch = False Then
                    OpQty = OpQty + (TBP.Fields("Q").Value)
                    OpBales = OpBales + (TBP.Fields("B").Value)
                    OpAmount = OpAmount + (TBP.Fields("Amt").Value)
                End If
            End If
            
            If TBPR.RecordCount > 0 Then
                TBPR.MoveFirst
                TBPR.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBPR.NoMatch = False Then
                    OpQty = OpQty - (TBPR.Fields("Q").Value)
                    OpBales = OpBales - (TBPR.Fields("B").Value)
                    OpAmount = OpAmount - (TBPR.Fields("Amt").Value)
                End If
            End If
    
            If TBS.RecordCount > 0 Then
                TBS.MoveFirst
                TBS.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBS.NoMatch = False Then
                    OpQty = OpQty - (TBS.Fields("Q").Value)
                    OpBales = OpBales - (TBS.Fields("B").Value)
                    OpAmount = OpAmount - (TBS.Fields("Amt").Value)
                End If
            End If
            
            If TBSR.RecordCount > 0 Then
                TBSR.MoveFirst
                TBSR.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBSR.NoMatch = False Then
                    OpQty = OpQty + (TBSR.Fields("Q").Value)
                    OpBales = OpBales + (TBSR.Fields("B").Value)
                    OpAmount = OpAmount + (TBSR.Fields("Amt").Value)
                End If
            End If
    
            If TBD.RecordCount > 0 Then
                TBD.MoveFirst
                TBD.FindFirst "ItemCode=" & TBItm.Fields("Code").Value
                If TBD.NoMatch = False Then
                    OpQty = OpQty + (TBD.Fields("Q").Value)
                    OpBales = OpBales + (TBD.Fields("B").Value)
                    OpAmount = OpAmount + (TBD.Fields("Amt").Value)
                End If
            End If
            
            If TBDSH.RecordCount > 0 Then
                TBDSH.MoveFirst
                TBDSH.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBDSH.NoMatch = False Then
                    OpQty = OpQty + (TBDSH.Fields("Q").Value)
                    OpBales = OpBales + (TBDSH.Fields("B").Value)
                    OpAmount = OpAmount + (TBDSH.Fields("Amt").Value)
                End If
            End If
            
            If TBDR.RecordCount > 0 Then
                TBDR.MoveFirst
                TBDR.FindFirst "CrCode=" & TBItm.Fields("Code").Value
                If TBDR.NoMatch = False Then
                    OpQty = OpQty - (TBDR.Fields("Q").Value)
                    OpBales = OpBales - (TBDR.Fields("B").Value)
                    OpAmount = OpAmount - (TBDR.Fields("Amt").Value)
                End If
            End If
    '-----------------Current
            If TBCI.RecordCount > 0 Then
                TBCI.MoveFirst
                TBCI.FindFirst "ItemCode=" & TBItm.Fields("Code").Value
                If TBCI.NoMatch = False Then
                    OutQty = OutQty + (TBCI.Fields("Q").Value)
                    OutBales = OutBales + (TBCI.Fields("B").Value)
                    OutAmount = OutAmount + (TBCI.Fields("Amt").Value)
                End If
            End If
            
            If TBCIR.RecordCount > 0 Then
                TBCIR.MoveFirst
                TBCIR.FindFirst "DrCode=" & TBItm.Fields("Code").Value
                If TBCIR.NoMatch = False Then
                    INQty = INQty + (TBCIR.Fields("Q").Value)
                    InBales = InBales + (TBCIR.Fields("B").Value)
                    InAmount = InAmount + (TBCIR.Fields("Amt").Value)
                End If
            End If
    
            If TBCP.RecordCount > 0 Then
                TBCP.MoveFirst
                TBCP.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBCP.NoMatch = False Then
                    INQty = INQty + (TBCP.Fields("Q").Value)
                    InBales = InBales + (TBCP.Fields("B").Value)
                    InAmount = InAmount + (TBCP.Fields("Amt").Value)
                End If
            End If
            
            If TBCPR.RecordCount > 0 Then
                TBCPR.MoveFirst
                TBCPR.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBCPR.NoMatch = False Then
                    OutQty = OutQty + (TBCPR.Fields("Q").Value)
                    OutBales = OutBales + (TBCPR.Fields("B").Value)
                    OutAmount = OutAmount + (TBCPR.Fields("Amt").Value)
                End If
            End If
    
            If TBCS.RecordCount > 0 Then
                TBCS.MoveFirst
                TBCS.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBCS.NoMatch = False Then
                    OutQty = OutQty + (TBCS.Fields("Q").Value)
                    OutBales = OutBales + (TBCS.Fields("B").Value)
                    OutAmount = OutAmount + (TBCS.Fields("Amt").Value)
                End If
            End If
            
            If TBCSR.RecordCount > 0 Then
                TBCSR.MoveFirst
                TBCSR.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBCSR.NoMatch = False Then
                    INQty = INQty + (TBCSR.Fields("Q").Value)
                    InBales = InBales + (TBCSR.Fields("B").Value)
                    InAmount = InAmount + (TBCSR.Fields("Amt").Value)
                End If
            End If
    
            If TBCD.RecordCount > 0 Then
                TBCD.MoveFirst
                TBCD.FindFirst "ItemCode=" & TBItm.Fields("Code").Value
                If TBCD.NoMatch = False Then
                    INQty = INQty + (TBCD.Fields("Q").Value)
                    InBales = InBales + (TBCD.Fields("B").Value)
                    InAmount = InAmount + (TBCD.Fields("Amt").Value)
                End If
            End If
            
            If TBCDSH.RecordCount > 0 Then
                TBCDSH.MoveFirst
                TBCDSH.FindFirst "Item=" & TBItm.Fields("Code").Value
                If TBCDSH.NoMatch = False Then
                    INQty = INQty + (TBCDSH.Fields("Q").Value)
                    InBales = InBales + (TBCDSH.Fields("B").Value)
                    InAmount = InAmount + (TBCDSH.Fields("Amt").Value)
                End If
            End If
            
            If TBCDR.RecordCount > 0 Then
                TBCDR.MoveFirst
                TBCDR.FindFirst "CrCode=" & TBItm.Fields("Code").Value
                If TBCDR.NoMatch = False Then
                    OutQty = OutQty + (TBCDR.Fields("Q").Value)
                    OutBales = OutBales + (TBCDR.Fields("B").Value)
                    OutAmount = OutAmount + (TBCDR.Fields("Amt").Value)
                End If
            End If
'    If TBItm.Fields("Code").Value = 290010001 Then MsgBox "Test"
    StockQty = (OpQty + INQty) - (OutQty)
    StockBales = (OpBales + InBales) - (OutBales)
    StockAmount = (OpAmount + InAmount) - (OutAmount)
'    If TBItm.Fields("GroupCode").Value = 29 Then MsgBox "Test"
'If StockQty > 0 Or StockBales > 0 Or StockAmount > 0 Then
        TBT.AddNew
        TBT.Fields("SDate").Value = S_Date
        TBT.Fields("EDate").Value = E_Date
        TBT.Fields("Code").Value = TBItm.Fields("Code").Value
        TBT.Fields("name").Value = Itemname
        TBT.Fields("HCode").Value = TBItm.Fields("GroupCode").Value
        TBT.Fields("Hname").Value = TBItm.Fields("GroupName").Value
        TBT.Fields("SHCode").Value = TBItm.Fields("SubGroupCode").Value
        TBT.Fields("SHname").Value = TBItm.Fields("SubGroupName").Value
        
        TBT.Fields("OpenBales").Value = OpBales
        TBT.Fields("OpenQty").Value = OpQty
        TBT.Fields("OpenAMount").Value = OpAmount
        
        TBT.Fields("CurrentBales").Value = InBales
        TBT.Fields("CurrentQty").Value = INQty
        TBT.Fields("CurrentAmount").Value = InAmount
        
        TBT.Fields("CurrentBales2").Value = (OutBales)
        TBT.Fields("CurrentQty2").Value = (OutQty)
        TBT.Fields("CurrentAmount2").Value = (OutAmount)
        
        TBT.Fields("ClosingBales").Value = StockBales
        TBT.Fields("ClosingQty").Value = StockQty
        TBT.Fields("ClosingAmount").Value = StockAmount
        
    TBT.Update
'End If
        OpBales = 0
        OpQty = 0
        OpAmount = 0
        
        InBales = 0
        INQty = 0
        InAmount = 0
        
        OutBales = 0
        OutQty = 0
        OutAmount = 0
        
        StockBales = 0
        StockQty = 0
        StockAmount = 0
        
    CNTL.Value = CNTL.Value + 1
    TBItm.MoveNext
    DoEvents
Loop
TBItm.Close
TBT.Close
DBT.Close

TBD.Close
TBI.Close
TBS.Close
TBP.Close
TBPR.Close
TBSR.Close

TBCD.Close
TBCI.Close
TBCS.Close
TBCP.Close
TBCPR.Close
TBCSR.Close

DBM.Close
End Sub

Public Sub PurchaseJobs(E_Date As Date, CNTL As Object)
Dim DBT As Database
Dim DBM As Database

Dim PQty As Double, PBales As Double
Dim TBJ As Recordset
Dim TBP As Recordset
Dim TBT As Recordset
Dim TBC As Recordset
Dim Itemname As String
Dim Ssql As String
Dim TBItm As Recordset
Dim JobType As String, JobInfo As String, JobStatus As Byte, JobResult As String
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Ssql = "Delete from Jobs"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("Jobs", dbOpenTable)

Ssql = "Select * from PurJob Order by V_No"
Set TBJ = DBM.OpenRecordset(Ssql)
CNTL.Value = 0
If Not TBJ.EOF Then
    TBJ.MoveLast
    CNTL.Max = TBJ.RecordCount
    TBJ.MoveFirst
End If

Ssql = "Select JobNo,Count(*) as C from Purchase where V_Date <= #" & E_Date & "# Group by jobNo"
Set TBC = DBM.OpenRecordset(Ssql)
If Not TBC.EOF Then
    TBC.MoveLast
    TBC.MoveFirst
End If
Ssql = "Select JobNo,Item,Sum(Qty) as Q,Sum(Bales) as B from Purchase where V_Date <= #" & E_Date & "# Group by JobNo,Item"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    TBP.MoveFirst
End If

Do While Not TBJ.EOF
    If TBP.RecordCount > 0 Then
        TBP.FindFirst "Jobno=" & TBJ.Fields("V_No").Value & " and Item = " & TBJ.Fields("Item").Value
        If TBP.NoMatch = False Then
            PQty = TBP.Fields("Q").Value
            PBales = TBP.Fields("B").Value
        Else
            PQty = 0
            PBales = 0
        End If
    Else
            PQty = 0
            PBales = 0
    End If
    
    Select Case TBJ.Fields("CTerms").Value
    
        Case 1 'WT Based
            JobType = "Weight Based"
            JobInfo = "Contracted Weight : " & TBJ.Fields("Qty").Value
            If TBJ.Fields("Qty").Value <= PQty Then JobStatus = 1
            JobResult = "Recvd. WT. : " & PQty
        Case 2
            JobType = "Consignments"
            JobInfo = "Contracted Consignments : " & TBJ.Fields("Consignments").Value
            If TBC.RecordCount > 0 Then
                TBC.FindFirst "JobNo=" & TBJ.Fields("V_No")
                If TBC.NoMatch = False Then
                    If TBC.Fields("C").Value >= TBJ.Fields("Consignments").Value Then
                        JobStatus = 1
                        JobResult = "Rcvd. Consignments : " & TBC.Fields("C").Value
                    End If
                End If
            End If
        Case 3
            JobType = "Periodic"
            JobInfo = "From : " & Format(TBJ.Fields("SDate").Value, "dd-MMM-yyyy") & " To : " & Format(TBJ.Fields("EDate").Value, "dd-MMM-yyyy")
            If E_Date > TBJ.Fields("EDate").Value Then JobStatus = 1
            JobResult = "Recvd. WT. : " & PQty
        Case 4
            JobType = "Others"
            JobInfo = TBJ.Fields("OtherTerms").Value & ""
            JobResult = "Recvd. WT. : " & PQty
        Case 5
            JobType = "Rate Based"
            JobInfo = "Contracted Rate : " & TBJ.Fields("Rate").Value & ""
            JobResult = "Recvd. WT. : " & PQty
    End Select
    
    
    TBT.AddNew
    
        TBT.Fields("EDate").Value = E_Date
        TBT.Fields("JobNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("Party").Value = Blm.party1(TBJ.Fields("Buyer").Value)
        TBT.Fields("Item").Value = Blm.item1(TBJ.Fields("Item").Value)
        TBT.Fields("CBales").Value = TBJ.Fields("STRate").Value
        TBT.Fields("CQty").Value = TBJ.Fields("Qty").Value
        TBT.Fields("CRate").Value = TBJ.Fields("Rate").Value
        TBT.Fields("Bales").Value = PBales
        TBT.Fields("Qty").Value = PQty
        TBT.Fields("JobType").Value = JobType
        TBT.Fields("JobInfo").Value = JobInfo
        TBT.Fields("JobResult").Value = JobResult
        TBT.Fields("Status").Value = JobStatus
    TBT.Update
    CNTL.Value = CNTL.Value + 1
    JobType = ""
    JobInfo = ""
    JobResult = ""
    JobStatus = 0
    TBJ.MoveNext
    DoEvents
Loop

TBC.Close
TBT.Close
DBT.Close
TBJ.Close
TBP.Close
DBM.Close
End Sub
Public Sub PurchaseJobLedger(CNTL As Object, Optional ContractNo As Long, Optional PartyCode As Long, Optional ItemCode As Long)
Dim DBT As Database
Dim DBM As Database

Dim PQty As Double, PBales As Double
Dim TBJ As Recordset
Dim TBP As Recordset
Dim TBP2 As Recordset
Dim TBT As Recordset
Dim TBC As Recordset
Dim Itemname As String
Dim Ssql As String
Dim TBItm As Recordset
Dim JobType As String, JobInfo As String, JobStatus As Byte, JobResult As String
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Ssql = "Delete from JobLedger"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("JobLedger", dbOpenTable)

Ssql = "Select * from PurJob where V_No>0 "
If ContractNo > 0 Then
    Ssql = Ssql & " and V_No=" & ContractNo
End If
If PartyCode > 0 Then
    Ssql = Ssql & " and Buyer=" & PartyCode
End If
If ItemCode > 0 Then
    Ssql = Ssql & " and Item=" & ItemCode
End If


Set TBJ = DBM.OpenRecordset(Ssql)
CNTL.Value = 0
If Not TBJ.EOF Then
    TBJ.MoveLast
    CNTL.Max = TBJ.RecordCount
    TBJ.MoveFirst
End If

Ssql = "Select *  from Purchase where jobNo > 0 Order By V_Date"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    
    TBP.MoveFirst
End If

CNTL.Value = 0
If Not TBJ.EOF Then
    Do While Not TBJ.EOF
    Select Case TBJ.Fields("CTerms").Value
    
        Case 1 'WT Based
            JobType = "Weight"
            JobInfo = "Quantity" 'TBJ.Fields("Qty").Value
            
        Case 2
            JobType = "Consignments"
            JobInfo = "Quantity" 'TBJ.Fields("Consignments").Value
            
        Case 3
            JobType = Format(TBJ.Fields("SDate").Value, "dd-MMM-yyyy") & " to " & Format(TBJ.Fields("EDate").Value, "dd-MMM-yyyy")
            JobInfo = "Quantity"
            
        Case 4
            JobType = "Others"
            JobInfo = TBJ.Fields("OtherTerms").Value & ""
            
        Case 5
            JobType = "Rate Based"
            JobInfo = TBJ.Fields("Rate").Value & ""
            
    End Select
    
    TBT.AddNew
        'TBT.Fields("SDate").Value = S_Date
        'TBT.Fields("EDate").Value = E_Date
        TBT.Fields("ContractNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("VDate").Value = TBJ.Fields("V_Date").Value
        TBT.Fields("VNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("PartyCode").Value = TBJ.Fields("Buyer").Value
        TBT.Fields("ItemCode").Value = TBJ.Fields("Item").Value
        TBT.Fields("PartyName").Value = Blm.party1(TBJ.Fields("Buyer").Value)
        TBT.Fields("ItemName").Value = Blm.item1(TBJ.Fields("Item").Value)
        TBT.Fields("CRate").Value = TBJ.Fields("Rate").Value
        
        If TBJ.Fields("CTerms").Value = 1 Then
            TBT.Fields("BQuantity").Value = TBJ.Fields("Qty").Value
            TBT.Fields("Quantity").Value = 0
            TBT.Fields("CQuantity").Value = TBJ.Fields("Qty").Value
        End If
        If TBJ.Fields("CTerms").Value = 2 Then
            TBT.Fields("BQuantity").Value = TBJ.Fields("Consignments").Value
            TBT.Fields("Quantity").Value = 0
            TBT.Fields("CQuantity").Value = TBJ.Fields("Consignments").Value
        End If
        If TBJ.Fields("CTerms").Value = 3 Then
            TBT.Fields("Quantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBJ.Fields("SDate"))) '+ 1
            TBT.Fields("CQuantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBJ.Fields("SDate"))) + 1
        End If
        TBT.Fields("Weight").Value = 0
        TBT.Fields("JobInfo").Value = JobType
        TBT.Fields("JobInfo2").Value = JobInfo
        TBT.Fields("CTerm").Value = TBJ.Fields("Cterms").Value
        TBT.Fields("EType").Value = 1
    TBT.Update

TBP.Filter = "JobNo=" & TBJ.Fields("V_No")
Set TBP2 = TBP.OpenRecordset()
If TBP2.EOF = False Then
Do While Not TBP2.EOF
    
    TBT.AddNew
        'TBT.Fields("SDate").Value = S_Date
        'TBT.Fields("EDate").Value = E_Date
        TBT.Fields("ContractNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("VDate").Value = TBP2.Fields("V_Date").Value
        TBT.Fields("VNo").Value = TBP2.Fields("V_No").Value
        TBT.Fields("PartyCode").Value = TBP2.Fields("Seller").Value
        TBT.Fields("ItemCode").Value = TBP2.Fields("Item").Value
        TBT.Fields("PartyName").Value = Blm.party1(TBP2.Fields("Seller").Value)
        TBT.Fields("ItemName").Value = Blm.item1(TBP2.Fields("Item").Value)
        TBT.Fields("CTerm").Value = TBJ.Fields("Cterms").Value
        If TBJ.Fields("CTerms").Value = 1 Then
            TBT.Fields("Quantity").Value = TBP2.Fields("Qty").Value
            TBT.Fields("SQuantity").Value = Format(TBP2.Fields("Qty").Value, "#")
        End If
        If TBJ.Fields("CTerms").Value = 2 Then
            TBT.Fields("Quantity").Value = 1
            TBT.Fields("SQuantity").Value = 1
        End If
        If TBJ.Fields("CTerms").Value = 3 Then
            TBT.Fields("Quantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBP2.Fields("V_Date"))) '+ 1
            TBT.Fields("SQuantity").Value = Format(TBP2.Fields("V_Date").Value, "dd/MM/yyyy")
        End If
        TBT.Fields("Weight").Value = TBP2.Fields("Qty").Value
        TBT.Fields("BQuantity").Value = 0
        TBT.Fields("Rate").Value = TBP2.Fields("Rate").Value
        TBT.Fields("VehicalNo").Value = TBP2.Fields("LorryNo").Value & ""
        
        If TBJ.Fields("CTerms").Value = 1 Then
            TBT.Fields("CQuantity").Value = TBJ.Fields("Qty").Value
        End If
        If TBJ.Fields("CTerms").Value = 2 Then
            TBT.Fields("CQuantity").Value = TBJ.Fields("Consignments").Value
        End If
        If TBJ.Fields("CTerms").Value = 3 Then
            TBT.Fields("CQuantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBJ.Fields("SDate"))) + 1
        End If

        
        TBT.Fields("JobInfo").Value = JobType
        TBT.Fields("JobInfo2").Value = JobInfo
        TBT.Fields("CRate").Value = TBJ.Fields("Rate").Value
        TBT.Fields("EType").Value = 2
     '   MsgBox TBJ.Fields("V_No").Value
    TBT.Update
    
    TBP2.MoveNext
    
Loop
End If
TBP2.Close
CNTL.Value = CNTL.Value + 1
DoEvents
TBJ.MoveNext
Loop
End If

TBT.Close
DBT.Close
TBJ.Close
TBP.Close
DBM.Close

End Sub
Public Sub SaleJobLedger(CNTL As Object, Optional ContractNo As Long, Optional PartyCode As Long, Optional ItemCode As Long)
Dim DBT As Database
Dim DBM As Database

Dim PQty As Double, PBales As Double
Dim TBJ As Recordset
Dim TBP As Recordset
Dim TBP2 As Recordset
Dim TBT As Recordset
Dim TBC As Recordset
Dim Itemname As String
Dim Ssql As String
Dim TBItm As Recordset
Dim JobType As String, JobInfo As String, JobStatus As Byte, JobResult As String
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Ssql = "Delete from JobLedger"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("JobLedger", dbOpenTable)

Ssql = "Select * from SaleJob where V_No>0 "
If ContractNo > 0 Then
    Ssql = Ssql & " and V_No=" & ContractNo
End If
If PartyCode > 0 Then
    Ssql = Ssql & " and Buyer=" & PartyCode
End If
If ItemCode > 0 Then
    Ssql = Ssql & " and Item=" & ItemCode
End If


Set TBJ = DBM.OpenRecordset(Ssql)
CNTL.Value = 0
If Not TBJ.EOF Then
    TBJ.MoveLast
    CNTL.Max = TBJ.RecordCount
    TBJ.MoveFirst
End If

Ssql = "Select *  from Sales where jobNo > 0 Order By V_Date"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    
    TBP.MoveFirst
End If

CNTL.Value = 0
If Not TBJ.EOF Then
    Do While Not TBJ.EOF
    Select Case TBJ.Fields("CTerms").Value
    
        Case 1 'WT Based
            JobType = "Weight"
            JobInfo = "Quantity" 'TBJ.Fields("Qty").Value
            
        Case 2
            JobType = "Consignments"
            JobInfo = "Quantity" 'TBJ.Fields("Consignments").Value
            
        Case 3
            JobType = Format(TBJ.Fields("SDate").Value, "dd-MMM-yyyy") & " to " & Format(TBJ.Fields("EDate").Value, "dd-MMM-yyyy")
            JobInfo = "Quantity"
            
        Case 4
            JobType = "Others"
            JobInfo = TBJ.Fields("OtherTerms").Value & ""
            
        Case 5
            JobType = "Rate Based"
            JobInfo = TBJ.Fields("Rate").Value & ""
            
    End Select
    
    TBT.AddNew
        'TBT.Fields("SDate").Value = S_Date
        'TBT.Fields("EDate").Value = E_Date
        TBT.Fields("ContractNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("VDate").Value = TBJ.Fields("V_Date").Value
        TBT.Fields("VNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("PartyCode").Value = TBJ.Fields("Buyer").Value
        TBT.Fields("ItemCode").Value = TBJ.Fields("Item").Value
        TBT.Fields("PartyName").Value = Blm.party1(TBJ.Fields("Buyer").Value)
        TBT.Fields("ItemName").Value = Blm.item1(TBJ.Fields("Item").Value)
        TBT.Fields("CRate").Value = TBJ.Fields("Rate").Value
        
        If TBJ.Fields("CTerms").Value = 1 Then
            TBT.Fields("BQuantity").Value = TBJ.Fields("Qty").Value
            TBT.Fields("Quantity").Value = 0
            TBT.Fields("CQuantity").Value = TBJ.Fields("Qty").Value
        End If
        If TBJ.Fields("CTerms").Value = 2 Then
            TBT.Fields("BQuantity").Value = TBJ.Fields("Consignments").Value
            TBT.Fields("Quantity").Value = 0
            TBT.Fields("CQuantity").Value = TBJ.Fields("Consignments").Value
        End If
        If TBJ.Fields("CTerms").Value = 3 Then
            TBT.Fields("Quantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBJ.Fields("SDate"))) '+ 1
            TBT.Fields("CQuantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBJ.Fields("SDate"))) + 1
        End If
        TBT.Fields("Weight").Value = 0
        TBT.Fields("JobInfo").Value = JobType
        TBT.Fields("JobInfo2").Value = JobInfo
        TBT.Fields("CTerm").Value = TBJ.Fields("Cterms").Value
        TBT.Fields("EType").Value = 1
    TBT.Update

TBP.Filter = "JobNo=" & TBJ.Fields("V_No")
Set TBP2 = TBP.OpenRecordset()
If TBP2.EOF = False Then
Do While Not TBP2.EOF
    
    TBT.AddNew
        'TBT.Fields("SDate").Value = S_Date
        'TBT.Fields("EDate").Value = E_Date
        TBT.Fields("ContractNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("VDate").Value = TBP2.Fields("V_Date").Value
        TBT.Fields("VNo").Value = TBP2.Fields("V_No").Value
        TBT.Fields("PartyCode").Value = TBP2.Fields("Seller").Value
        TBT.Fields("ItemCode").Value = TBP2.Fields("Item").Value
        TBT.Fields("PartyName").Value = Blm.party1(TBP2.Fields("Seller").Value)
        TBT.Fields("ItemName").Value = Blm.item1(TBP2.Fields("Item").Value)
        TBT.Fields("CTerm").Value = TBJ.Fields("Cterms").Value
        If TBJ.Fields("CTerms").Value = 1 Then
            TBT.Fields("Quantity").Value = TBP2.Fields("Qty").Value
            TBT.Fields("SQuantity").Value = Format(TBP2.Fields("Qty").Value, "#")
        End If
        If TBJ.Fields("CTerms").Value = 2 Then
            TBT.Fields("Quantity").Value = 1
            TBT.Fields("SQuantity").Value = 1
        End If
        If TBJ.Fields("CTerms").Value = 3 Then
            TBT.Fields("Quantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBP2.Fields("V_Date"))) '+ 1
            TBT.Fields("SQuantity").Value = Format(TBP2.Fields("V_Date").Value, "dd/MM/yyyy")
        End If
        TBT.Fields("Weight").Value = TBP2.Fields("Qty").Value
        TBT.Fields("BQuantity").Value = 0
        TBT.Fields("Rate").Value = TBP2.Fields("Rate").Value
        TBT.Fields("VehicalNo").Value = TBP2.Fields("LorryNo").Value & ""
        
        If TBJ.Fields("CTerms").Value = 1 Then
            TBT.Fields("CQuantity").Value = TBJ.Fields("Qty").Value
        End If
        If TBJ.Fields("CTerms").Value = 2 Then
            TBT.Fields("CQuantity").Value = TBJ.Fields("Consignments").Value
        End If
        If TBJ.Fields("CTerms").Value = 3 Then
            TBT.Fields("CQuantity").Value = Abs(DateDiff("d", TBJ.Fields("EDate").Value, TBJ.Fields("SDate"))) + 1
        End If

        
        TBT.Fields("JobInfo").Value = JobType
        TBT.Fields("JobInfo2").Value = JobInfo
        TBT.Fields("CRate").Value = TBJ.Fields("Rate").Value
        TBT.Fields("EType").Value = 2
     '   MsgBox TBJ.Fields("V_No").Value
    TBT.Update
    
    TBP2.MoveNext
    
Loop
End If
TBP2.Close
CNTL.Value = CNTL.Value + 1
DoEvents
TBJ.MoveNext
Loop
End If

TBT.Close
DBT.Close
TBJ.Close
TBP.Close
DBM.Close
End Sub

Public Sub SaleJobs(E_Date As Date, CNTL As Object)
Dim DBT As Database
Dim DBM As Database

Dim PQty As Double, PBales As Double
Dim TBJ As Recordset
Dim TBP As Recordset
Dim TBT As Recordset
Dim TBC As Recordset
Dim Itemname As String
Dim Ssql As String
Dim TBItm As Recordset
Dim JobType As String, JobInfo As String, JobStatus As Byte, JobResult As String
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Ssql = "Delete from Jobs"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("Jobs", dbOpenTable)

Ssql = "Select * from SaleJob Order by V_No"
Set TBJ = DBM.OpenRecordset(Ssql)
CNTL.Value = 0
If Not TBJ.EOF Then
    TBJ.MoveLast
    CNTL.Max = TBJ.RecordCount
    TBJ.MoveFirst
End If

Ssql = "Select JobNo,Count(*) as C from Sales where V_Date <= #" & E_Date & "# Group by jobNo"
Set TBC = DBM.OpenRecordset(Ssql)
If Not TBC.EOF Then
    TBC.MoveLast
    TBC.MoveFirst
End If
Ssql = "Select JobNo,Item,Sum(Qty) as Q,Sum(Bales) as B from Sales where V_Date <= #" & E_Date & "# Group by JobNo,Item"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    TBP.MoveFirst
End If

Do While Not TBJ.EOF
    If TBP.RecordCount > 0 Then
        TBP.FindFirst "Jobno=" & TBJ.Fields("V_No").Value & " and Item = " & TBJ.Fields("Item").Value
        If TBP.NoMatch = False Then
            PQty = TBP.Fields("Q").Value
            PBales = TBP.Fields("B").Value
        Else
            PQty = 0
            PBales = 0
        End If
    Else
            PQty = 0
            PBales = 0
    End If
    
    Select Case TBJ.Fields("CTerms").Value
    
        Case 1 'WT Based
            JobType = "Weight Based"
            JobInfo = "Contracted Weight : " & TBJ.Fields("Qty").Value
            If TBJ.Fields("Qty").Value <= PQty Then JobStatus = 1
            JobResult = "Sold. WT. : " & PQty
        Case 2
            JobType = "Consignments"
            JobInfo = "Contracted Consignments : " & TBJ.Fields("Consignments").Value
            If TBC.RecordCount > 0 Then
                TBC.FindFirst "JobNo=" & TBJ.Fields("V_No")
                If TBC.NoMatch = False Then
                    If TBC.Fields("C").Value >= TBJ.Fields("Consignments").Value Then
                        JobStatus = 1
                        JobResult = "Sold. Consignments : " & TBC.Fields("C").Value
                    End If
                End If
            End If
        Case 3
            JobType = "Periodic"
            JobInfo = "From : " & Format(TBJ.Fields("SDate").Value, "dd-MMM-yyyy") & " To : " & Format(TBJ.Fields("EDate").Value, "dd-MMM-yyyy")
            If E_Date > TBJ.Fields("EDate").Value Then JobStatus = 1
            JobResult = "Sold. WT. : " & PQty
        Case 4
            JobType = "Others"
            JobInfo = TBJ.Fields("OtherTerms").Value & ""
            JobResult = "Sold. WT. : " & PQty
        Case 5
            JobType = "Rate Based"
            JobInfo = "Contracted Rate : " & TBJ.Fields("Rate").Value & ""
            JobResult = "Sold. WT. : " & PQty
    End Select
    
    
    TBT.AddNew
    
        TBT.Fields("EDate").Value = E_Date
        TBT.Fields("JobNo").Value = TBJ.Fields("V_No").Value
        TBT.Fields("Party").Value = Blm.party1(TBJ.Fields("Buyer").Value)
        TBT.Fields("Item").Value = Blm.item1(TBJ.Fields("Item").Value)
        TBT.Fields("CBales").Value = TBJ.Fields("STRate").Value
        TBT.Fields("CQty").Value = TBJ.Fields("Qty").Value
        TBT.Fields("CRate").Value = TBJ.Fields("Rate").Value
        TBT.Fields("Bales").Value = PBales
        TBT.Fields("Qty").Value = PQty
        TBT.Fields("JobType").Value = JobType
        TBT.Fields("JobInfo").Value = JobInfo
        TBT.Fields("JobResult").Value = JobResult
        TBT.Fields("Status").Value = JobStatus
    TBT.Update
    CNTL.Value = CNTL.Value + 1
    JobType = ""
    JobInfo = ""
    JobResult = ""
    JobStatus = 0
    TBJ.MoveNext
    DoEvents
Loop

TBC.Close
TBT.Close
DBT.Close
TBJ.Close
TBP.Close
DBM.Close
End Sub
Public Sub MonthlyProduction(S_Date As Date, E_Date As Date, CNTL As Object, Optional RefNo As Long, Optional ItemCode As Long)
Dim DBT As Database
Dim DBM As Database
Dim OPStockQty As Double
Dim opStockbales As Double
Dim OpAmt As Double
Dim PQty As Double, PBales As Double
Dim PAmt As Double
Dim TBP As Recordset
Dim TBI As Recordset
Dim TBT As Recordset
Dim Itemname As String
Dim PurQty As Double, PurBales As Double, PurAmount As Double
Dim IssueQty As Double, IssueBales As Double, IssueAmount As Double, Issuerate As Double, AvgBaleWT As Double, BaleStock As Double, QtyStock As Double, StockAmount As Double, StockRate As Double
Dim Ssql As String
Dim TBItm As Recordset
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Ssql = "Delete from ItemProduction"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("ItemProduction", dbOpenTable)
Ssql = "Select * from Items Where Code in (Select Distinct ItemCode from Production)"
If ItemCode > 0 Then
Ssql = Ssql & " and Code = " & ItemCode
End If
Ssql = Ssql & "Order By Code"
Set TBItm = DBM.OpenRecordset(Ssql)
Ssql = "Select ItemCode,V_Date,Sum(Qty) as Q,Sum(Bales) as B,Sum(Rate * Qty) as A from Production where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
If RefNo > 0 Then
    Ssql = Ssql & " and RefNo=" & RefNo
End If
If ItemCode > 0 Then
    Ssql = Ssql & " and ItemCode=" & ItemCode
End If
Ssql = Ssql & " Group by ItemCode,V_Date"
MsgBox Ssql
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    TBP.MoveFirst
End If

Ssql = "Select Item,V_Date,Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from Sales where V_Date Between #" & S_Date & "# and #" & E_Date & "#"
If ItemCode > 0 Then
    Ssql = Ssql & " and Item=" & ItemCode
End If
Ssql = Ssql & " Group by Item,V_Date"
Set TBI = DBM.OpenRecordset(Ssql)
If Not TBI.EOF Then
    TBI.MoveLast
    TBI.MoveFirst
End If

Dim ThisDate As Date
Dim d As Integer
Dim Diff As Integer

Diff = Abs(DateDiff("d", S_Date, E_Date)) + 1

If Not TBItm.EOF Then
Do While Not TBItm.EOF
QtyStock = 0
BalesStock = 0
CNTL.Value = 0
CNTL.Max = Abs(DateDiff("d", S_Date, E_Date)) + 1
OPStockQty = Blm.ClosingStockEndDate(TBItm.Fields("Code").Value, S_Date - 1)
opStockbales = Blm.ClosingStockBalesEndDate(TBItm.Fields("Code").Value, S_Date - 1)
OpAmt = Blm.ClosingStockAmountEndDate(TBItm.Fields("Code").Value, S_Date - 1)
PQty = OPStockQty
PBales = opStockbales
PAmt = OpAmt

For d = 1 To Diff
    ThisDate = DateAdd("d", d - 1, S_Date)
   
    Ssql = "V_Date = #" & ThisDate & "# and Item=" & TBItm.Fields("Code").Value
    If Not TBI.EOF Then
        TBI.MoveFirst
        TBI.FindFirst Ssql
        If TBI.NoMatch Then
            IssueQty = 0
            IssueBales = 0
            IssueAmount = 0
        Else
            IssueQty = TBI.Fields("Q").Value
            IssueBales = TBI.Fields("B").Value
            IssueAmount = TBI.Fields("A").Value
        End If
    End If
    Ssql = "V_Date = #" & ThisDate & "# and ItemCode=" & TBItm.Fields("Code").Value
    If Not TBP.EOF Then
        TBP.MoveFirst
        TBP.FindFirst Ssql
'        MsgBox ssql
        If TBP.NoMatch Then
            PurQty = 0
            PurBales = 0
            PurAmount = 0
        Else
            PurQty = TBP.Fields("Q").Value
            PurBales = TBP.Fields("B").Value
            PurAmount = TBP.Fields("A").Value
'         MsgBox ThisDate
        End If
'         MsgBox ThisDate
   End If
    
    QtyStock = (PQty + PurQty) - IssueQty
    BalesStock = (PBales + PurBales) - IssueBales
    StockAmount = (PAmount + ProdAmount) - IssueAmount
    If IssueAmount <> 0 Or ProdAmount <> 0 Then
    
    TBT.AddNew
    
        TBT.Fields("S_Date").Value = S_Date
        TBT.Fields("E_Date").Value = E_Date
        TBT.Fields("V_Date").Value = ThisDate
        TBT.Fields("ProdBales").Value = PurBales
        TBT.Fields("ProdQty").Value = PurQty
        TBT.Fields("ProdAmount").Value = ProdAmount
        TBT.Fields("SaleBales").Value = IssueBales
        TBT.Fields("SaleQty").Value = IssueQty
        TBT.Fields("SaleAMount").Value = IssueAmount
        TBT.Fields("QtyStock").Value = QtyStock
        TBT.Fields("BaleStock").Value = BalesStock
        TBT.Fields("AmountStock").Value = StockAmount
        TBT.Fields("opQty").Value = PQty
        TBT.Fields("opBales").Value = PBales
        TBT.Fields("OpAMount").Value = PAmount
        TBT.Fields("ItemName").Value = TBItm.Fields("Name").Value & ""
        If RefNo > 0 Then
            TBT.Fields("RefNo").Value = RefNo
        End If
    TBT.Update
    End If
    CNTL.Value = CNTL.Value + 1
    DoEvents
    PQty = QtyStock
    PBales = BalesStock
    PAmount = StockAmount
    PurQty = 0
    PurBales = 0
    PurAmount = 0
    ProdAmount = 0
    IssueQty = 0
    IssueBales = 0
    IssueAmount = 0
Next d
TBItm.MoveNext
DoEvents
Loop
End If
TBItm.Close
TBT.Close
DBT.Close
TBI.Close
TBP.Close
DBM.Close
End Sub
Public Sub MonthlyPurchases(S_Date As Date, E_Date As Date, CNTL As Object)
Dim DBT As Database
Dim DBM As Database
Dim OPStockQty As Double
Dim opStockbales As Double
Dim PQty As Double, PBales As Double
Dim TBP As Recordset
Dim TBI As Recordset
Dim TBT As Recordset
Dim Itemname As String
Dim PurQty As Double, PurBales As Double, PurAmount As Double
Dim IssueQty As Double, IssueBales As Double, IssueAmount As Double, Issuerate As Double, AvgBaleWT As Double, BaleStock As Double, QtyStock As Double, StockAmount As Double, StockRate As Double
Dim Ssql As String
Dim TBItm As Recordset, Tbmill As Recordset
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Ssql = "Delete from ItemProduction"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("ItemProduction", dbOpenTable)
Ssql = "Select * from Items where IType=0 Order by Code"
Set TBItm = DBM.OpenRecordset(Ssql)
If Not TBItm.EOF Then
    TBItm.MoveLast
    TBItm.MoveFirst
End If
Ssql = "Select Item,V_Date,Sum(Qty) as Q,Sum(Bales) as B from Purchase where V_Date Between #" & S_Date & "# and #" & E_Date & "# Group by Item,V_Date"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    TBP.MoveFirst
End If

Ssql = "Select Item,V_Date,Sum(Qty) as Q,Sum(Bales) as B from Sales where V_Date Between #" & S_Date & "# and #" & E_Date & "# Group by Item,V_Date"
Set TBI = DBM.OpenRecordset(Ssql)
If Not TBI.EOF Then
    TBI.MoveLast
    TBI.MoveFirst
End If

Dim ThisDate As Date
Dim d As Integer
Dim Diff As Integer

Diff = Abs(DateDiff("d", S_Date, E_Date)) + 1
If Not TBItm.EOF Then
Do While Not TBItm.EOF
QtyStock = 0
BalesStock = 0
CNTL.Value = 0
CNTL.Max = Abs(DateDiff("d", S_Date, E_Date)) + 1
OPStockQty = Blm.ClosingStockEndDate(TBItm.Fields("Code").Value, S_Date - 1)
opStockbales = Blm.ClosingStockBalesEndDate(TBItm.Fields("Code").Value, S_Date - 1)
PQty = OPStockQty
PBales = opStockbales

For d = 1 To Diff
    ThisDate = DateAdd("d", d - 1, S_Date)
'    MsgBox ThisDate
    Ssql = "V_Date = #" & ThisDate & "# and Item=" & TBItm.Fields("Code").Value
    If TBI.RecordCount > 0 Then
        TBI.MoveFirst
        TBI.FindFirst Ssql
        If TBI.NoMatch Then
            IssueQty = 0
            IssueBales = 0
            
        Else
            IssueQty = TBI.Fields("Q").Value
            IssueBales = TBI.Fields("B").Value

        End If
    End If
    Ssql = "V_Date = #" & ThisDate & "# and Item=" & TBItm.Fields("Code").Value
    If TBP.RecordCount > 0 Then
        TBP.MoveFirst
        TBP.FindFirst Ssql
        If TBP.NoMatch Then
            PurQty = 0
            PurBales = 0
            
        Else
            PurQty = TBP.Fields("Q").Value
            PurBales = TBP.Fields("B").Value

        End If
    End If
    
    QtyStock = (PQty + PurQty) - IssueQty
    BalesStock = (PBales + PurBales) - IssueBales
    
    TBT.AddNew
    
        TBT.Fields("S_Date").Value = S_Date
        TBT.Fields("E_Date").Value = E_Date
        TBT.Fields("V_Date").Value = ThisDate
        TBT.Fields("ProdBales").Value = PurBales
        TBT.Fields("ProdQty").Value = PurQty
        TBT.Fields("SaleBales").Value = IssueBales
        TBT.Fields("SaleQty").Value = IssueQty
        TBT.Fields("QtyStock").Value = QtyStock
        TBT.Fields("BaleStock").Value = BalesStock
        TBT.Fields("opQty").Value = PQty
        TBT.Fields("opBales").Value = PBales
        TBT.Fields("ItemName").Value = TBItm.Fields("Name").Value & ""
    TBT.Update
    CNTL.Value = CNTL.Value + 1
    DoEvents
    PQty = QtyStock
    PBales = BalesStock
    PurQty = 0
    PurBales = 0
    IssueQty = 0
    IssueBales = 0
Next d
TBItm.MoveNext
DoEvents
Loop
End If
TBItm.Close
TBT.Close
DBT.Close
TBI.Close
TBP.Close
DBM.Close
End Sub

Public Sub BankBalances(E_Date As Date)
Dim DBT As Database
Dim DBM As Database
Dim TBP As Recordset
Dim TBT As Recordset
Dim Ssql As String
Dim TBItm As Recordset
Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Ssql = "Delete from BankBalance"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("BankBalance", dbOpenTable)
Ssql = "Select Party,Sum(Debit) - Sum(Credit) as Bal from VouDTL where V_Date <= #" & E_Date & "# and Mid(party,1,5)='" & BanksHead & "' Group By Party"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
Do While Not TBP.EOF
    TBT.AddNew
        TBT.Fields("BankTitle").Value = Blm.party1(TBP.Fields("Party").Value)
        TBT.Fields("Balance").Value = TBP.Fields("Bal").Value
    TBT.Update
    
TBP.MoveNext
DoEvents
Loop
End If
TBT.Close
DBT.Close
TBP.Close
DBM.Close
End Sub

Public Sub ItemWiseIssue(ItemCode As Double, S_Date As Date, E_Date As Date, CNTL As Object, Optional RefNo As Integer)
Dim DBT As Database
Dim DBM As Database

Dim TBP As Recordset
Dim TBI As Recordset
Dim TBT As Recordset
Dim Itemname As String
Dim PurQty As Double, PurBales As Double, PurAmount As Double
Dim IssueQty As Double, IssueBales As Double, IssueAmount As Double, Issuerate As Double, AvgBaleWT As Double, BaleStock As Double, QtyStock As Double, StockAmount As Double, StockRate As Double
Dim IssueRatio As String
Dim Ssql As String

Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
Set DBT = OpenDatabase(App.path & "\Book.mdb")

Itemname = Blm.item1(ItemCode)

Ssql = "Delete from Itemissue"
DBT.Execute Ssql

Set TBT = DBT.OpenRecordset("ItemIssue", dbOpenTable)

Ssql = "Select * from Purchase where Item=" & ItemCode & " and V_Date Between #" & S_Date & "# and #" & E_Date & "# Order by V_Date"
Set TBP = DBM.OpenRecordset(Ssql)
If Not TBP.EOF Then
    TBP.MoveLast
    TBP.MoveFirst
End If

Ssql = "Select * from Issue where ItemCode=" & ItemCode & " and V_Date Between #" & S_Date & "# and #" & E_Date & "# and RefNo= " & RefNo & " Order by V_Date"
Set TBI = DBM.OpenRecordset(Ssql)
If Not TBI.EOF Then
    TBI.MoveLast
    TBI.MoveFirst
End If

Dim ThisDate As Date
Dim d As Integer
Dim Diff As Integer
CNTL.Value = 0
Diff = Abs(DateDiff("d", S_Date, E_Date)) + 1
CNTL.Max = Abs(DateDiff("d", S_Date, E_Date)) + 1

For d = 1 To Diff
    ThisDate = DateAdd("d", d - 1, S_Date)
'    MsgBox ThisDate
    Ssql = "V_Date = #" & ThisDate & "#"
    If TBP.RecordCount > 0 Then
        TBP.MoveFirst
        TBP.FindFirst Ssql
        If TBP.NoMatch Then
            PurQty = 0
            PurBales = 0
            PurAmount = 0
        Else
            PurQty = TBP.Fields("Qty").Value
            PurBales = TBP.Fields("Bales").Value
            PurAmount = TBP.Fields("Amount").Value + Val(TBP.Fields("Freight").Value & "")
        End If
    End If
    Ssql = "V_Date = #" & ThisDate & "#"
'    MsgBox ThisDate
    If TBI.RecordCount > 0 Then
        TBI.MoveFirst
        TBI.FindFirst Ssql
        If TBI.NoMatch Then
            Issuerate = 0
            IssueRatio = ""
            IssueQty = 0
            IssueBales = 0
            IssueAmount = 0
            Issuerate = 0
            
        Else
            IssueRatio = TBI.Fields("Percent").Value & ""
            IssueQty = TBI.Fields("Qty").Value
            IssueBales = TBI.Fields("Bales").Value
            IssueAmount = TBI.Fields("Amount").Value
            Issuerate = TBI.Fields("Rate").Value
            BaleStock = Val(TBI.Fields("BaleStock").Value & "")
            QtyStock = Val(TBI.Fields("QtyStock").Value & "")
            AvgBaleWT = Val(TBI.Fields("AvgBaleWT").Value & "")
            StockAmount = Issuerate * QtyStock
            StockRate = Issuerate
'            MsgBox "Check"
        End If
    End If
    If IssueQty = 0 And IssueBales = 0 Then
        If PurBales > 0 And PurQty > 0 Then
            StockAmount = StockAmount + PurAmount
            BaleStock = BaleStock + PurBales
            QtyStock = QtyStock + PurQty
            AvgBaleWT = Round(QtyStock / BaleStock, 3)
            StockRate = Round(StockAmount / QtyStock, 3)
        End If
    
    End If
    TBT.AddNew
    
        TBT.Fields("SDate").Value = S_Date
        TBT.Fields("EDate").Value = E_Date
        TBT.Fields("V_Date").Value = ThisDate
        TBT.Fields("SDate").Value = S_Date
        TBT.Fields("PurBales").Value = PurBales
        TBT.Fields("PurQty").Value = PurQty
        TBT.Fields("PurAmount").Value = PurAmount
        TBT.Fields("IssueBales").Value = IssueBales
        TBT.Fields("IssueQty").Value = IssueQty
        TBT.Fields("IssueAmount").Value = IssueAmount
        TBT.Fields("IssueRate").Value = Issuerate
       ' TBT.Fields("ItemRatio").Value = Val(IssueRatio)
        TBT.Fields("Rate").Value = StockRate
        TBT.Fields("AvgBaleWT").Value = AvgBaleWT
        TBT.Fields("Stock").Value = QtyStock
        TBT.Fields("BalesStock").Value = BaleStock
        TBT.Fields("Amount").Value = StockAmount
        TBT.Fields("ItemName").Value = Itemname
        TBT.Fields("ItemCode").Value = ItemCode
    TBT.Update
    CNTL.Value = CNTL.Value + 1
    DoEvents
Next d
TBT.Close
DBT.Close
TBI.Close
TBP.Close
DBM.Close
End Sub
Public Function LastDayofmonth(M As Integer, Y As Integer) As Integer
Select Case M

    Case 1
        LastDayofmonth = 31
    Case 2
        If (Y Mod 4) <> 0 Then
            LastDayofmonth = 28
        Else
            LastDayofmonth = 29
        End If
    Case 3
        LastDayofmonth = 31
    Case 4
        LastDayofmonth = 30
    Case 5
        LastDayofmonth = 31
    Case 6
        LastDayofmonth = 30
    Case 7
        LastDayofmonth = 31
    Case 8
        LastDayofmonth = 31
    Case 9
        LastDayofmonth = 30
    Case 10
        LastDayofmonth = 31
    Case 11
        LastDayofmonth = 30
    Case 12
        LastDayofmonth = 31
End Select
End Function
Public Function CheckHoliday(d As Date) As String
Select Case Weekday(d, vbMonday)
    Case 1
        CheckHoliday = "*"
    Case Else
        CheckHoliday = "N"
End Select
End Function

Public Sub AtdRegister(S_Date As Date, E_Date As Date, CNTL As Object)
Dim DBM As Database
Dim DBT As Database
Dim PCounts As Integer
Dim HOlidays As String
Dim HRst As String
Dim LDay As Integer
Dim RST As Recordset
Dim RS As Recordset
Dim RSSal As Recordset
Dim RSM As Recordset
Dim R As Long, d As Integer
Dim BSal As Double
Dim Ssql As String

Dim CH As String, TDate As Date, G As String

CNTL.Value = 0
Set DBT = OpenDatabase(App.path & "\Book.mdb")
DBT.Execute "Delete from ATDReg"
Set RST = DBT.OpenRecordset("AtdReg", dbOpenTable)

Set DBM = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")

Ssql = "Select * from Acchart where Mid(Code,1,2)='" & EmpHead & "' Order by Code"
Set RSM = DBM.OpenRecordset(Ssql)

If Not RSM.EOF Then
RSM.MoveLast
CNTL.Max = RSM.RecordCount
RSM.MoveFirst
Do While Not RSM.EOF
    
    RST.AddNew
        RST.Fields("SHCode").Value = Val(Mid(RSM.Fields("Code"), 1, 5))
        RST.Fields("SHName").Value = Blm.SubHeadName(Val(Mid(RSM.Fields("Code"), 1, 5)))
        RST.Fields("EmpCode").Value = RSM.Fields("Code")
        RST.Fields("EmpName").Value = RSM.Fields("Name") & ""
        RST.Fields("Hours").Value = RSM.Fields("Hours") & ""
        RST.Fields("SalRate").Value = RSM.Fields("SalRate")
        Ssql = "Select A_Date,Ac_Code,Status,Day(A_Date) as ADay from EmpATD where Ac_Code = " & RSM.Fields("Code") & " and A_date Between #" & S_Date & "# and #" & E_Date & "# Order by A_date"
        Set RS = DBM.OpenRecordset(Ssql)
        If Not RS.EOF Then
        RS.MoveLast
        RS.MoveFirst
        For d = 1 To 31
            LDay = LastDayofmonth(Month(S_Date), Year(S_Date))
            If d <= LDay Then
            TDate = CDate(Month(S_Date) & "/" & d & "/" & Year(S_Date))
            CH = CheckHoliday(TDate)
'            If CH = "*" Then
'                RST.Fields("D" & d).Value = "*"
'                PCounts = PCounts + 1
'            Else
'            RS.MoveFirst
            RS.FindFirst "ADay=" & d
            
            If RS.NoMatch Then
                RST.Fields("D" & d).Value = "A"
            ElseIf RS.Fields("Status").Value = "P" Then
                PCounts = PCounts + 1
                RST.Fields("D" & d).Value = RS.Fields("Status").Value & ""
            ElseIf RS.Fields("Status").Value = "L" Then
                PCounts = PCounts + 1
                RST.Fields("D" & d).Value = RS.Fields("Status").Value & ""
            ElseIf RS.Fields("Status").Value = "*" Then
                PCounts = PCounts + 1
                RST.Fields("D" & d).Value = RS.Fields("Status").Value & ""
                'HRst = GetHolidayName(RS.Fields("A_Date"))
                'HOlidays = Holiday & vbCrLf & HRst & " on " & Format(RS.Fields("A_Date"), "dd-MMM-yyyy")
            ElseIf RS.Fields("Status").Value = "C" Then
                PCounts = PCounts + 1
                RST.Fields("D" & d).Value = RS.Fields("Status").Value & ""
            ElseIf RS.Fields("Status").Value = "Y" Then
                PCounts = PCounts + 1
                RST.Fields("D" & d).Value = RS.Fields("Status").Value & ""
            ElseIf IsNull(RS.Fields("Status").Value) Then
                
                RST.Fields("D" & d).Value = "A"
            Else
                RST.Fields("D" & d).Value = "A"
            End If
            'End If
            End If
            RS.MoveFirst
        Next d
    End If
    RS.Close
    
    RST.Fields("DaysWorked").Value = PCounts
    RST.Fields("Holidays") = HOlidays
    HOlidays = ""
    RST.Update
    PCounts = 0
    CNTL.Value = CNTL.Value + 1
    RSM.MoveNext
Loop
End If
RSM.Close

RST.Close
DBT.Close
End Sub

Public Sub day_due_TEMP(V_Date As Date, CNTL As Control, Cntl2 As Control)
Dim Db_M As Database
Dim db_t As Database
Dim test As Integer
Dim TB As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim t_cred As Currency
Dim Ssql As String
Dim tbtrial As Recordset
Dim TB_C As Recordset
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from vou_temp"
db_t.Execute Ssql
CNTL.Value = 0
Ssql = "delete from trial"
db_t.Execute (Ssql)
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
Ssql = "SELECT * FROM ACCHART ORDER BY CODE"
Set TB_C = Db_M.OpenRecordset(Ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
CNTL.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    Cntl2.SimpleText = TB_C.Fields("name").Value
        Ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUdtl WHERE V_DATE <= #" & V_Date & "#"
        Ssql = Ssql & " AND PARTY = " & TB_C.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) < 0 "
        Set TB = Db_M.OpenRecordset(Ssql)
        If Not TB.EOF Then
        If Not IsNull(TB.Fields("DEB").Value) Then
            t_cred = TB.Fields("DEB").Value
        End If
        End If
'        MsgBox TB_C.Fields("DESCRIPTION").Value & " pAID = " & t_cred
        TB.Close
    Cntl2.SimpleText = Cntl2.SimpleText & " TOTAL PAID = " & t_cred
     'TOTAL PURCHASED ITEMS ADJUSTMENTS
     If t_cred < 0 Then
        t_cred = t_cred * -1
        'MsgBox t_cred
        Ssql = "SELECT * FROM Voudtl WHERE PARTY = " & TB_C.Fields("CODE").Value
        Ssql = Ssql & " AND V_DATE <= #" & V_Date & "# ORDER BY V_DATE"
        Set TB = Db_M.OpenRecordset(Ssql)
        If Not TB.EOF Then
           TB.MoveLast
            Do While Not TB.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
                    t_cred = t_cred - (TB.Fields("Credit").Value)
                'End If
 '               MsgBox t_cred
                If t_cred <= 0 Then Exit Do
            Else
                Exit Do
            End If
            TB.MovePrevious
            Loop
            
            If TB.BOF Then TB.MoveFirst
  '          MsgBox tb.Fields("PARTY").Value
            If Not TB.BOF Then
            Do While Not TB.EOF
                    DoEvents
                    Cntl2.SimpleText = Cntl2.SimpleText & " tOTAL pAID = " & t_cred
                    If TB.Fields("Pakki_no").Value > 0 Then
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = TB.Fields("v_date").Value
                    tb_t.Fields("due_date").Value = TB.Fields("v_date").Value + TB.Fields("days").Value
                    tb_t.Fields("code").Value = TB.Fields("party").Value
                    tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
                    tb_t.Fields("remarks").Value = TB.Fields("Remarks").Value & " Amt = " & TB.Fields("Credit").Value
                    If t_cred < 0 Then
                        t_cred = TB.Fields("Credit").Value + t_cred
                    Else
                        t_cred = TB.Fields("Credit").Value
                    End If
   '                 MsgBox t_cred
                    tb_t.Fields("CREDIT").Value = t_cred
                    tb_t.Update
                    End If
                't_cred = -1
                TB.MoveNext
                If Not TB.EOF Then
                    t_cred = TB.Fields("Credit").Value
                End If
            Loop
            End If
            't_cred = 0
        End If
        TB.Close
    End If
      TB_C.MoveNext
      CNTL.Value = CNTL.Value + 1
      DoEvents
    Loop
End If
TB_C.Close
tb_t.Close
db_t.Close
Db_M.Close
End Sub
Public Sub day_due2_TEMP(V_Date As Date, CNTL As Control, Cntl2 As Control)
Dim Db_M As Database
Dim db_t As Database
Dim test As Integer
Dim TB As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim t_cred As Currency
Dim Ssql As String
Dim tbtrial As Recordset
Dim TB_C As Recordset
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from vou_temp"
db_t.Execute Ssql
CNTL.Value = 0
Ssql = "delete from trial"
db_t.Execute (Ssql)
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
Ssql = "SELECT * FROM ACCHART ORDER BY CODE"
Set TB_C = Db_M.OpenRecordset(Ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
CNTL.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    Cntl2.SimpleText = TB_C.Fields("Name").Value
        Ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUdtl WHERE V_DATE <= #" & V_Date & "#"
        Ssql = Ssql & " AND PARTY = " & TB_C.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) > 0 "
        Set TB = Db_M.OpenRecordset(Ssql)
        If Not TB.EOF Then
        If Not IsNull(TB.Fields("DEB").Value) Then
            t_cred = TB.Fields("DEB").Value
        End If
        End If
'        MsgBox TB_C.Fields("DESCRIPTION").Value & " pAID = " & t_cred
        TB.Close
    Cntl2.SimpleText = Cntl2.SimpleText & " TOTAL PAID = " & t_cred
     'TOTAL Sales ITEMS ADJUSTMENTS
     If t_cred > 0 Then
        
 '       MsgBox t_cred
        Ssql = "SELECT * FROM voudtl WHERE PARTY = " & TB_C.Fields("CODE").Value
        Ssql = Ssql & " AND V_DATE <= #" & V_Date & "# ORDER BY V_DATE"
        Set TB = Db_M.OpenRecordset(Ssql)
        If Not TB.EOF Then
           TB.MoveLast
            Do While Not TB.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
                    t_cred = t_cred - (TB.Fields("Debit").Value)
                'End If
 '               MsgBox t_cred
                If t_cred <= 0 Then Exit Do
            Else
                Exit Do
            End If
            TB.MovePrevious
            Loop
            
            If TB.BOF Then TB.MoveFirst
'            MsgBox tb.Fields("PARTY").Value & " " & t_cred
            If Not TB.BOF Then
            Do While Not TB.EOF
                    DoEvents
                    Cntl2.SimpleText = Cntl2.SimpleText & " tOTAL pAID = " & t_cred
                    
                    If TB.Fields("Pakki_no").Value > 0 Then
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = TB.Fields("v_date").Value
                    tb_t.Fields("due_date").Value = TB.Fields("v_date").Value + TB.Fields("days").Value
                    tb_t.Fields("code").Value = TB.Fields("party").Value
                    tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
                    tb_t.Fields("remarks").Value = TB.Fields("Remarks").Value & " Amt = " & TB.Fields("Debit").Value
                    If t_cred < 0 Then
                        t_cred = (TB.Fields("Debit").Value) + t_cred
                    Else
                        t_cred = TB.Fields("Debit").Value
                    End If
'                    MsgBox t_cred
                    tb_t.Fields("CREDIT").Value = t_cred
                    tb_t.Update
                    End If
                't_cred = -1
                TB.MoveNext
                If Not TB.EOF Then
                    t_cred = TB.Fields("Debit").Value
                End If
            Loop
            End If
            't_cred = 0
        End If
        TB.Close
    End If
      TB_C.MoveNext
      CNTL.Value = CNTL.Value + 1
      DoEvents
    Loop
End If
TB_C.Close
tb_t.Close
db_t.Close
Db_M.Close
End Sub


Public Property Get report_path() As String
    report_path = App.path & "\"
'report_path = CurDir & "\"
End Property
Public Property Get path() As String
    path = t_path
End Property
Public Sub PL(E_Date As Date, CNTL As Control)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
    
'Vouchers

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_date <= #" & E_Date & "# and "
    Ssql = Ssql & "mid(Party,1,2) in ('21','22','36','58') group by party"
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
            tb_t.Fields("e_date").Value = E_Date
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
            tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
            tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
            tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
            tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
            tb_t.Fields("Bal").Value = TB.Fields("bal").Value
            If TB.Fields("bal").Value > 0 Then
                tb_t.Fields("debit").Value = TB.Fields("bal").Value
            Else
                tb_t.Fields("debit").Value = 0
            End If
            If TB.Fields("bal").Value < 0 Then
                tb_t.Fields("credit").Value = TB.Fields("bal").Value * -1
            Else
                tb_t.Fields("credit").Value = 0
            End If
        tb_t.Update
        TB.MoveNext
        CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close

tb_t.Close
db_t.Close
Db_M.Close
End Sub
Public Sub trial(E_Date As Date, CNTL As Control)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
    
'Vouchers

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_date <= #" & E_Date & "# group by party"
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("code").Value = TB.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("debit").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("credit").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("credit").Value = 0
        End If
    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close

tb_t.Close
db_t.Close
Db_M.Close
End Sub
Public Sub trial2(S_Date As Date, E_Date As Date, CNTL As Control, Optional HeadCode As String, Optional SubHeadCode As String)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
    
'Vouchers

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_date < #" & S_Date & "#"
    If Len(HeadCode) > 0 Then
        Ssql = Ssql & " and Mid(Party,1,2) in (" & HeadCode & ")"
    End If
    If Len(SubHeadCode) > 0 Then
        Ssql = Ssql & " and Mid(Party,1,5) in (" & SubHeadCode & ")"
    End If
    Ssql = Ssql & " group by party"
    'MsgBox Ssql
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("s_date").Value = S_Date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = TB.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = Left(Blm.party1(TB.Fields("party").Value), 50)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("op_deb").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("op_deb").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("op_cred").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("op_cred").Value = 0
        End If
    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close

tb_t.Close
CNTL.Value = 0
Dim tb_3 As Recordset
Dim ssql2 As String
    Ssql = "select party,sum(debit) as D,Sum(credit) as c from voudtl where  v_date between #" & S_Date & "# and #" & E_Date & "#"
    If Len(HeadCode) > 0 Then
        Ssql = Ssql & " and Mid(Party,1,2) in (" & HeadCode & ")"
    End If
    If Len(SubHeadCode) > 0 Then
        Ssql = Ssql & " and Mid(Party,1,5) in (" & SubHeadCode & ")"
    End If
    Ssql = Ssql & " group by party"
    'MsgBox Ssql
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        
        Do While Not TB.EOF
        
        DoEvents
        Ssql = "select * from trial where code = " & TB.Fields("party").Value
        Set tb_3 = db_t.OpenRecordset(Ssql)
        'If TB.Fields("Party").Value = 180020001 Then MsgBox "Test"
        If Not tb_3.EOF Then
            If Not IsNull(TB.Fields("d").Value) Then
                ssql2 = "update trial set debit = " & TB.Fields("D").Value
            Else
                ssql2 = "update trial set debit = 0"
            End If
            ssql2 = ssql2 & " where code = " & TB.Fields("party").Value
            db_t.Execute ssql2
            'If TB.Fields("Party").Value = 200030001 Then MsgBox ssql2
            'If TB.Fields("Party").Value = 200030002 Then MsgBox ssql2
           If Not IsNull(TB.Fields("C").Value) Then
                ssql2 = "update trial set credit = " & TB.Fields("C").Value
            Else
                ssql2 = "update trial set Credit = 0"
           End If
           
           ssql2 = ssql2 & " where code = " & TB.Fields("party").Value
           db_t.Execute ssql2
           'If TB.Fields("Party").Value = 200030001 Then MsgBox ssql2
           'If TB.Fields("Party").Value = 200030002 Then MsgBox ssql2
        Else
            Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
            tb_t.AddNew
            tb_t.Fields("e_date").Value = E_Date
            tb_t.Fields("s_date").Value = S_Date
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("op_deb").Value = 0
            tb_t.Fields("op_cred").Value = 0
            tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
            tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
            tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
            tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
            tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
                If Not IsNull(TB.Fields("D").Value) Then
                    tb_t.Fields("debit").Value = TB.Fields("d").Value
                Else
                    tb_t.Fields("debit").Value = 0
                End If
                If Not IsNull(TB.Fields("D").Value) Then
                    tb_t.Fields("Credit").Value = TB.Fields("c").Value
                Else
                    tb_t.Fields("Credit").Value = 0
                End If
            
            
            tb_t.Update
            tb_t.Close
        End If
        
        
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close


db_t.Close
Db_M.Close
End Sub

Public Sub open_trial(CNTL As Control)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
    
'Vouchers

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_no = 1 and v_type = 10 group by party"
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = FStartDate
        tb_t.Fields("code").Value = TB.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = Left(Blm.party1(TB.Fields("party").Value), 50)
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("debit").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("credit").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("credit").Value = 0
        End If
    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close

tb_t.Close
db_t.Close
Db_M.Close
End Sub

Public Sub ProfitLoss(E_Date As Date, CNTL As Control)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)

'sales
    Ssql = "select sum(debit) as sales from voucher where e_type = 1 or e_type = 4 and v_date <= #" & E_Date & "#"
    Set TB = Db_M.OpenRecordset(Ssql)
    If Not IsNull(TB.Fields("sales").Value) Then
        tb_t.AddNew
            tb_t.Fields("e_date").Value = E_Date
            tb_t.Fields("code").Value = 5000
            tb_t.Fields("Name").Value = "Sales For The Day"
            tb_t.Fields("debit").Value = 0#
            tb_t.Fields("credit").Value = TB.Fields("sales").Value
            
        tb_t.Update
   End If
TB.Close
'purchases
    Ssql = "select sum(credit) as purchases from voucher where e_type between 2 and 3 and v_date <= #" & E_Date & "#"
    Set TB = Db_M.OpenRecordset(Ssql)
    If Not IsNull(TB.Fields("purchases").Value) Then
        tb_t.AddNew
            tb_t.Fields("e_date").Value = E_Date
            tb_t.Fields("code").Value = 6000
            tb_t.Fields("Name").Value = "Purchases For The Day"
            tb_t.Fields("debit").Value = TB.Fields("purchases").Value
            tb_t.Fields("credit").Value = 0
            
        tb_t.Update
    End If
   TB.Close
'Vouchers

    Ssql = "select subcode,sum(debit-credit) as bal from vou_view where party <> 10000 and h_type = 2 and v_date <= #" & E_Date & "# group by subcode"
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = E_Date
        'tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("subcode").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("subcode").Value), 1, 2))))
        tb_t.Fields("s_code").Value = TB.Fields("subcode").Value
        tb_t.Fields("s_desc").Value = Blm.heads(TB.Fields("subcode").Value)
        'tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("debit").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("credit").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("credit").Value = 0
        End If
    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close

tb_t.Close
db_t.Close
Db_M.Close
End Sub

Public Sub day_book(V_Date As Date, CNTL As Control)
Dim DB As Database
    Dim TB As Recordset
    Dim Ssql As String
    Dim t_sal As Currency
    Dim t_pur As Currency
    Dim dif_sp As Currency
    Dim t_t_r As Currency
    Dim t_t_s As Currency
    Set Blm = Nothing
    Set Blm = New bloom1
    Dim DBT As Database
    Dim TBT As Recordset
    Set DB = OpenDatabase(Blm.patHmain)
    Set TB = DB.OpenRecordset("Select * from voucher where v_date = #" & V_Date & "#")
    Set DBT = OpenDatabase(report_path & "book.mdb")
    DBT.Execute ("delete from vou_temp")
    Set TBT = DBT.OpenRecordset("vou_temp", dbOpenTable)
    CNTL.Value = 0
    
    If TB.EOF = False Then
    TB.MoveLast
    CNTL.Max = TB.RecordCount
    TB.MoveFirst
        Do While Not TB.EOF
            TBT.AddNew
                TBT.Fields("V_dATE").Value = V_Date
                TBT.Fields("E_TYPE").Value = TB.Fields("E_TYPE").Value
                TBT.Fields("PRE_CASH").Value = TB.Fields("PRE_CASH").Value
                TBT.Fields("CODE").Value = TB.Fields("PARTY").Value
                TBT.Fields("NAME").Value = Blm.party1(TB.Fields("PARTY").Value)
                TBT.Fields("DEBIT").Value = TB.Fields("DEBIT").Value
                TBT.Fields("CREDIT").Value = TB.Fields("CREDIT").Value
                TBT.Fields("REMARKS").Value = TB.Fields("REMARKS").Value
            TBT.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
        Loop
    Else
        MsgBox "NO RECORD FOR THIS DATE"
    End If
    TBT.Close
    TB.Close
    DBT.Close
    DB.Close
    End Sub
Public Sub ledger(code As Long, CNTL As Control)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
'If Not IsMissing(s_date) Then
 '   ssql = "select sum(debit-credit) as bal from voucher where party = " & code & " and v_date < #" & s_date & "#"
 '   Set tb = db_m.OpenRecordset(ssql)
 '   If Not IsNull(tb.Fields("bal").Value) Then
 '       op_bal = tb.Fields("bal").Value
 '   End If
 '   tb.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
 Ssql = "select * from voudtl where party = " & code
'    MsgBox ssql
'Else
 '   ssql = "select * from voucher where party = " & code & " and v_Date >= #" & s_date & "#" & " and v_date <= #" & e_date & "#"
'End If
'MsgBox ssql
Set TB = Db_M.OpenRecordset(Ssql)
If Not TB.EOF Then
    CNTL.Value = 0
    TB.MoveLast
    CNTL.Max = TB.RecordCount
    TB.MoveFirst
    Do While Not TB.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = TB.Fields("v_date").Value
            tb_t.Fields("v_no").Value = TB.Fields("v_no").Value
            tb_t.Fields("v_type").Value = Blm.VoucherTypesRet(TB.Fields("v_type").Value)
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("description").Value = Blm.party1(TB.Fields("party").Value)
            
            If TB.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
        
            Else
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            End If
            If TB.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = TB.Fields("remarks").Value
            tb_t.Fields("Kachi_no").Value = TB.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = TB.Fields("Days").Value
            tb_t.Fields("K_Date").Value = TB.Fields("P_Date").Value
'                MsgBox tb_t.Fields("Credit").Value
            'If Not IsMissing(s_date) Then
             '   tb_t.Fields("s_date").Value = s_date
             '   tb_t.Fields("e_date").Value = e_date
           ' End If
        tb_t.Update
    TB.MoveNext
    CNTL.Value = CNTL.Value + 1
    Loop
End If
    TB.Close
    tb_t.Close
    db_t.Close
    Db_M.Close
End Sub
Public Sub ICTLedger(ICT As String, CNTL As Control)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
'If Not IsMissing(s_date) Then
 '   ssql = "select sum(debit-credit) as bal from voucher where party = " & code & " and v_date < #" & s_date & "#"
 '   Set tb = db_m.OpenRecordset(ssql)
 '   If Not IsNull(tb.Fields("bal").Value) Then
 '       op_bal = tb.Fields("bal").Value
 '   End If
 '   tb.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
 Ssql = "select * from voudtl where ICT = '" & ICT & "'"
'    MsgBox ssql
'Else
 '   ssql = "select * from voucher where party = " & code & " and v_Date >= #" & s_date & "#" & " and v_date <= #" & e_date & "#"
'End If
'MsgBox ssql
Set TB = Db_M.OpenRecordset(Ssql)
If Not TB.EOF Then
    CNTL.Value = 0
    TB.MoveLast
    CNTL.Max = TB.RecordCount
    TB.MoveFirst
    Do While Not TB.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = TB.Fields("v_date").Value
            tb_t.Fields("v_no").Value = TB.Fields("v_no").Value
            tb_t.Fields("v_type").Value = Blm.VoucherTypesRet(TB.Fields("v_type").Value)
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("description").Value = Blm.party1(TB.Fields("party").Value)
            
            If TB.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
        
            Else
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            End If
            If TB.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = TB.Fields("remarks").Value
            'tb_t.Fields("Kachi_no").Value = Tb.Fields("ICT").Value & ""
            tb_t.Fields("ICt").Value = TB.Fields("ICT").Value & ""
            
'                MsgBox tb_t.Fields("Credit").Value
            'If Not IsMissing(s_date) Then
             '   tb_t.Fields("s_date").Value = s_date
             '   tb_t.Fields("e_date").Value = e_date
           ' End If
        tb_t.Update
    TB.MoveNext
    CNTL.Value = CNTL.Value + 1
    Loop
End If
    TB.Close
    tb_t.Close
    db_t.Close
    Db_M.Close
End Sub
Private Function BreakRemarks(S As String, Bales As Integer, Weight As Double, Rate As Double, Itemname As String, STrate As Double, VehicalNo As String) As Boolean
    Dim p As Integer
    Dim PP As Integer
    
    p = InStr(S, "B:")
    If p <= 0 Then
        BreakRemarks = False
        Exit Function
    End If
    
    p = InStr(1, S, "B:", vbBinaryCompare)
    If p > 0 Then
        Itemname = LTrim(Mid(S, 2, p - 2))
        PP = p
    Else
        PP = -1
    End If
    
    If PP > -1 Then
        p = InStr(PP, S, " ")
        If p > 0 Then
            Bales = Val(Mid(S, PP + 2, p - (PP + 2)))
            PP = p
        Else
            PP = -1
        End If
    End If

    If PP > -1 Then
        p = InStr(PP, S, "@", vbBinaryCompare)
'        MsgBox "Test"
        If p > 0 Then
            Weight = Val(Mid(S, PP + 3, p - (PP + 3)))
            PP = p
        Else
            PP = -1
        End If
    End If

    If PP > -1 Then
        p = InStr(PP, S, " ", vbBinaryCompare)
        If p > 0 Then
            Rate = Val(Mid(S, PP + 1, p - (PP + 1)))
            PP = p
        Else
            PP = -1
        End If
    End If
    
    
    If PP > -1 Then
        p = InStr(PP + 3, S, " ", vbBinaryCompare)
        
        If p > 0 Then
            STrate = Val(Mid(S, PP + 3, p - (PP + 3)))
            PP = p
        Else
            PP = -1
        End If
        'MsgBox "TestST RATE"
    End If
    
    
    If PP > -1 Then
        p = InStr(PP, S, "V#", vbBinaryCompare)
        If p > 0 Then
           VehicalNo = Mid(S, PP + 3, Len(S) - (p + 1))
           'MsgBox "Test"
            PP = p
        Else
            PP = -1
        End If
    End If
    BreakRemarks = True
    
    
'    MsgBox S & " --> " & Itemname & " " & Bales & " " & Weight & " " & Rate & " " & STrate & " " & VehicalNo
End Function

Public Sub ledger2(code As Long, CNTL As Control, S_Date As Date, E_Date As Date, Optional AgeGroup As Integer)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim A As Integer
Dim AA As Integer
Dim Ssql As String
Dim B As Boolean
Dim Itemname As String
Dim Bales As Integer, Weight As Double, Rate As Double, STrate As Double, VehicalNo As String
Dim PDate As Date, Diff As Double
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
PDate = S_Date
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
'If Not IsMissing(s_date) Then
    Ssql = "select sum(debit-credit) as bal from vouDTL where party = " & code & " and v_date < #" & S_Date & "#"
    Set TB = Db_M.OpenRecordset(Ssql)
    If Not IsNull(TB.Fields("bal").Value) Then
    tb_t.AddNew
        op_bal = TB.Fields("bal").Value
        If op_bal < 0 Then
            tb_t.Fields("credit").Value = op_bal * -1
            tb_t.Fields("DEBIT").Value = 0
        End If
        If op_bal > 0 Then
            tb_t.Fields("debit").Value = op_bal
            tb_t.Fields("CREDIT").Value = 0
        End If
            'tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("Kachi_No").Value = 1
            tb_t.Fields("v_date").Value = S_Date
            tb_t.Fields("code").Value = code 'tb.Fields("party").Value
            tb_t.Fields("description").Value = Blm.party1(code)
            tb_t.Fields("remarks").Value = "Opening Balance" ' On .." & Format(S_Date, "dd/MM/yyyy")
'            MsgBox "Test"
     tb_t.Update
    End If
    TB.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    Ssql = "select * from vouDTL where party = " & code & " and v_Date >= #" & S_Date & "#" & " and v_date <= #" & E_Date & "# Order by V_date"
'End If
'MsgBox ssql
AA = 1
Set TB = Db_M.OpenRecordset(Ssql)
If Not TB.EOF Then
    CNTL.Value = 0
    TB.MoveLast
    CNTL.Max = TB.RecordCount
    TB.MoveFirst
    Do While Not TB.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = TB.Fields("v_date").Value
            tb_t.Fields("v_no").Value = TB.Fields("v_no").Value
            tb_t.Fields("v_type").Value = Blm.VoucherTypesRet(TB.Fields("v_type").Value)
        
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("description").Value = Blm.party1(TB.Fields("party").Value)
            If TB.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            Else
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            End If
            If TB.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            End If
            If Not IsNull(TB.Fields("REMARKS").Value) Then
            B = BreakRemarks(TB.Fields("remarks").Value, Bales, Weight, Rate, Itemname, STrate, VehicalNo)
            If B = False Then
                tb_t.Fields("remarks").Value = TB.Fields("remarks").Value
                tb_t.Fields("Weight").Value = 0
                tb_t.Fields("BaleS").Value = 0
                tb_t.Fields("Rate").Value = 0
                tb_t.Fields("STrate").Value = 0
                tb_t.Fields("VehicalNo").Value = ""
            Else
                tb_t.Fields("remarks").Value = Itemname
                tb_t.Fields("Weight").Value = Weight
                tb_t.Fields("BaleS").Value = Bales
                tb_t.Fields("Rate").Value = Rate
                tb_t.Fields("STrate").Value = STrate
                tb_t.Fields("VehicalNo").Value = VehicalNo
            End If
            End If
            tb_t.Fields("Kachi_no").Value = TB.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = TB.Fields("Days").Value
            tb_t.Fields("K_Date").Value = TB.Fields("P_Date").Value
            tb_t.Fields("Kachi_No").Value = AA
            'If Not IsMissing(s_date) Then
                tb_t.Fields("s_date").Value = S_Date
                tb_t.Fields("e_date").Value = E_Date
           ' End If
        tb_t.Update
    TB.MoveNext
    If Not TB.EOF Then
    If TB.Fields("V_Date").Value <> PDate Then
 '   MsgBox "Test"
        Diff = DateDiff("d", TB.Fields("V_date").Value, PDate)
        If Abs(Diff) < 1000 Then
        If Abs(Diff) >= AgeGroup Then
        
            AA = AA + 1
            PDate = TB.Fields("V_Date").Value
        End If
        Else
            MsgBox Abs(Diff)
        End If
'        MsgBox Diff & " " & AA
    End If
     
    End If
   
'    If AgeGroup > 0 And AA >= AgeGroup Then
'       A = A + 1
'       AA = 0
'    End If
'    MsgBox A
    CNTL.Value = CNTL.Value + 1
    Loop
End If
    TB.Close
    tb_t.Close
    db_t.Close
    Db_M.Close
End Sub
Public Sub CashBook(CNTL As Control, S_Date As Date, E_Date As Date, PartyCode As Long)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
'If Not IsMissing(s_date) Then
    Ssql = "select sum(debit-credit) as bal from vouDTL where party = " & PartyCode & " and v_date < #" & S_Date & "#"
    Set TB = Db_M.OpenRecordset(Ssql)
    If Not IsNull(TB.Fields("bal").Value) Then
    
    tb_t.AddNew
        op_bal = TB.Fields("bal").Value
        If op_bal > 0 Then
            tb_t.Fields("Debit").Value = op_bal
            tb_t.Fields("Credit").Value = 0
        End If
        If op_bal < 0 Then
            tb_t.Fields("Credit").Value = op_bal * -1
            tb_t.Fields("Debit").Value = 0
        End If
            'tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = S_Date
            tb_t.Fields("v_No").Value = 0
            tb_t.Fields("v_Type").Value = 0
            tb_t.Fields("code").Value = PartyCode 'tb.Fields("party").Value
            tb_t.Fields("description").Value = Blm.party1(PartyCode)
            tb_t.Fields("remarks").Value = "Cash B/F on  .." & Format(S_Date, "dd/MM/yyyy")
            tb_t.Fields("s_date").Value = S_Date
            tb_t.Fields("e_date").Value = E_Date
     tb_t.Update
'     MsgBox "Test"
    End If
    TB.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    Ssql = "select * from voudtl where party=" & PartyCode & " and v_type=3 and v_Date >= #" & S_Date & "#" & " and v_date <= #" & E_Date & "#"
'End If
'MsgBox ssql

Set TB = Db_M.OpenRecordset(Ssql)
If Not TB.EOF Then
    CNTL.Value = 0
    TB.MoveLast
    CNTL.Max = TB.RecordCount
    TB.MoveFirst
    Do While Not TB.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = TB.Fields("v_date").Value
            tb_t.Fields("v_no").Value = TB.Fields("v_no").Value
            tb_t.Fields("v_type").Value = Blm.VoucherTypesRet(TB.Fields("v_type").Value)
        
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("description").Value = Blm.party1(TB.Fields("party").Value)
            If TB.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            Else
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            End If
            If TB.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = TB.Fields("remarks").Value
            tb_t.Fields("Kachi_no").Value = TB.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = TB.Fields("Days").Value
            tb_t.Fields("K_Date").Value = TB.Fields("P_Date").Value
            
            'If Not IsMissing(s_date) Then
                tb_t.Fields("s_date").Value = S_Date
                tb_t.Fields("e_date").Value = E_Date
           ' End If
        tb_t.Update
    TB.MoveNext
    CNTL.Value = CNTL.Value + 1
    Loop
End If
    TB.Close
    tb_t.Close
    db_t.Close
    Db_M.Close
End Sub

Public Sub ledger3(CNTL As Control, S_Date As Date, E_Date As Date)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim tb_t2 As Recordset
Dim op_bal As Currency
Dim Ssql As String
Dim B As Boolean
Dim Itemname As String
Dim Bales As Integer, Weight As Double, Rate As Double, STrate As Double, VehicalNo As String

CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
Ssql = "delete from opbal"
db_t.Execute Ssql

Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
Set tb_t2 = db_t.OpenRecordset("opbal", dbOpenTable)
'If Not IsMissing(s_date) Then
    Ssql = "select party,sum(debit-credit) as bal from voudtl where v_date < #" & S_Date & "# group by party"
    Set TB = Db_M.OpenRecordset(Ssql)
     If Not TB.EOF Then
        Do While Not TB.EOF
          tb_t.AddNew
            tb_t.Fields("v_date").Value = S_Date
            tb_t.Fields("s_date").Value = S_Date
            tb_t.Fields("e_date").Value = E_Date
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
            tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
            tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
            tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
            tb_t.Fields("description").Value = Blm.party1(TB.Fields("party").Value)
            tb_t.Fields("remarks").Value = "Opening Balance" ' On .." & Format(S_Date, "dd/MM/yyyy")
            If TB.Fields("bal").Value > 0 Then
                tb_t.Fields("debit").Value = TB.Fields("bal").Value
                tb_t.Fields("credit").Value = 0
            End If
            If TB.Fields("bal").Value < 0 Then
                tb_t.Fields("credit").Value = TB.Fields("bal").Value * -1
                tb_t.Fields("debit").Value = 0
            End If
            tb_t.Update
            TB.MoveNext
        Loop
     End If
    TB.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    Ssql = "select * from voudtl where v_Date between #" & S_Date & "#" & " and #" & E_Date & "#"
'End If
'MsgBox ssql

Set TB = Db_M.OpenRecordset(Ssql)
If Not TB.EOF Then
    CNTL.Value = 0
    TB.MoveLast
    CNTL.Max = TB.RecordCount
    TB.MoveFirst
    Do While Not TB.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = TB.Fields("v_date").Value
            tb_t.Fields("v_no").Value = TB.Fields("v_no").Value
            tb_t.Fields("v_type").Value = Blm.VoucherTypesRet(TB.Fields("v_type").Value)
            
            tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
            tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
            tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
            tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
            tb_t.Fields("code").Value = TB.Fields("party").Value
            tb_t.Fields("description").Value = Blm.party1(TB.Fields("party").Value)
            If TB.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            Else
            tb_t.Fields("debit").Value = TB.Fields("debit").Value
            End If
            If TB.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = TB.Fields("credit").Value
            End If
            If Len(TB.Fields("Remarks").Value) > 0 Then
            B = BreakRemarks(TB.Fields("remarks").Value, Bales, Weight, Rate, Itemname, STrate, VehicalNo)
            If B = False Then
                tb_t.Fields("remarks").Value = TB.Fields("remarks").Value
                tb_t.Fields("Weight").Value = 0
                tb_t.Fields("BaleS").Value = 0
                tb_t.Fields("Rate").Value = 0
                tb_t.Fields("STrate").Value = 0
                tb_t.Fields("VehicalNo").Value = ""
            Else
                tb_t.Fields("remarks").Value = Itemname
                tb_t.Fields("Weight").Value = Weight
                tb_t.Fields("BaleS").Value = Bales
                tb_t.Fields("Rate").Value = Rate
                tb_t.Fields("STrate").Value = STrate
                tb_t.Fields("VehicalNo").Value = VehicalNo
            End If
            End If
            tb_t.Fields("Kachi_no").Value = TB.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = TB.Fields("Days").Value
            tb_t.Fields("K_Date").Value = TB.Fields("P_Date").Value
            
            'If Not IsMissing(s_date) Then
                tb_t.Fields("s_date").Value = S_Date
                tb_t.Fields("e_date").Value = E_Date
           ' End If
        tb_t.Update
    TB.MoveNext
    CNTL.Value = CNTL.Value + 1
    Loop
End If
    TB.Close
    tb_t.Close
    db_t.Close
    Db_M.Close
End Sub
Public Property Get FinancePath() As String
FinancePath = App.path & "\Finance.mdb"
End Property
Public Sub ProfitLossState(E_Date As Date, CNTL As Object)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset
Dim NetPlYear As Currency, NetPLMonth As Currency
Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(FinancePath)

Ssql = "delete from PLYear"
db_t.Execute Ssql
Ssql = "delete from PLMONTH"
db_t.Execute Ssql
Ssql = "delete from PLAMT"
db_t.Execute Ssql


Set tb_t = db_t.OpenRecordset("PLYear", dbOpenTable)
    
'Vouchers For the Year

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Val(Mid(Party,1,2)) > 36 and v_date < #" & E_Date & "# group by party Having sum(debit-credit) < 0 "
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        NetPlYear = 0
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("Group1").Value = 1
        tb_t.Fields("e_date").Value = E_Date
        'tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = TB.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        NetPlYear = NetPlYear + TB.Fields("Bal").Value
    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close


    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Val(Mid(Party,1,2)) > 36 and v_date < #" & E_Date & "# group by party Having sum(debit-credit) > 0 "
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        NetPlYear = 0
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("Group1").Value = 2
        tb_t.Fields("e_date").Value = E_Date
        'tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = TB.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        NetPlYear = NetPlYear + TB.Fields("Bal").Value
    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close

tb_t.Close
'Cntl.Value = 0
'Set tb_t = db_t.OpenRecordset("PLMonth", dbOpenTable)
'Dim S_Date As Date
'
'S_Date = CDate(Month(E_Date) & "/01/" & Year(E_Date))
'
''Vouchers For the Month
'
'    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Val(Mid(Party,1,2))> 36 and v_date between #" & S_Date & "# and #" & E_Date & "# group by party"
'    Set TB = Db_M.OpenRecordset(Ssql)
'        If TB.EOF = False Then
'        TB.MoveLast
'        Cntl.Max = TB.RecordCount
'        TB.MoveFirst
'        NetPLMonth = 0
'        Do While Not TB.EOF
'        DoEvents
'        tb_t.AddNew
'        tb_t.Fields("e_date").Value = E_Date
'    '    tb_t.Fields("s_date").Value = S_date
''        MsgBox tb.Fields("party").Value
'        tb_t.Fields("code").Value = TB.Fields("party").Value
'        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
'        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
'        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
'        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
'        tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
'        tb_t.Fields("debit").Value = 0
'        tb_t.Fields("credit").Value = 0
'        If TB.Fields("bal").Value > 0 Then
'            tb_t.Fields("Debit").Value = TB.Fields("bal").Value
'        Else
'            tb_t.Fields("Debit").Value = 0
'        End If
'        If TB.Fields("bal").Value < 0 Then
'            tb_t.Fields("Credit").Value = TB.Fields("bal").Value * -1
'        Else
'            tb_t.Fields("Credit").Value = 0
'        End If
'        NetPLMonth = NetPLMonth + TB.Fields("Bal").Value
'    tb_t.Update
'            TB.MoveNext
'            Cntl.Value = Cntl.Value + 1
'    Loop
'    End If
'
'    TB.Close
'
'tb_t.Close


Set tb_t = db_t.OpenRecordset("PLAMT", dbOpenTable)
tb_t.AddNew
tb_t.Fields("PLYear").Value = NetPlYear
tb_t.Fields("PLMonth").Value = NetPLMonth
tb_t.Update
tb_t.Close


db_t.Close
Db_M.Close

End Sub
Public Sub BalanceSheetState(E_Date As Date, CNTL As Object)
Dim Db_M As Database
Dim db_t As Database
Dim TB As Recordset
Dim tb_t As Recordset

Dim Ssql As String
CNTL.Value = 0
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(FinancePath)

Ssql = "delete from BalYear"
db_t.Execute Ssql
Ssql = "delete from BalMONTH"
db_t.Execute Ssql


Set tb_t = db_t.OpenRecordset("BalYear", dbOpenTable)
    
'Vouchers For the Year

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Val(Mid(Party,1,2)) <=36 and v_date < #" & E_Date & "# group by party"
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst

        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = E_Date
        'tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = TB.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If

    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close
Set TB = db_t.OpenRecordset("PLAMT", dbOpenDynaset)
tb_t.AddNew
    tb_t.Fields("Code").Value = 0
    tb_t.Fields("Name").Value = "Net Profit Loss For the Year"
    If TB.Fields("PLYear").Value < 0 Then
        tb_t.Fields("E_date").Value = E_Date
        tb_t.Fields("Credit").Value = TB.Fields("PlYear").Value * -1
        tb_t.Fields("Debit").Value = 0
    Else
        tb_t.Fields("E_date").Value = E_Date
        tb_t.Fields("Debit").Value = TB.Fields("PlYear").Value
        tb_t.Fields("Credit").Value = 0
    End If
tb_t.Update
TB.Close

tb_t.Close

CNTL.Value = 0

Set tb_t = db_t.OpenRecordset("BalMonth", dbOpenTable)
Dim S_Date As Date

S_Date = CDate(Month(E_Date) & "/01/" & Year(E_Date))

'Vouchers For the Month

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Val(Mid(Party,1,2)) <=36 and v_date between #" & S_Date & "# and #" & E_Date & "# group by party"
    Set TB = Db_M.OpenRecordset(Ssql)
        If TB.EOF = False Then
        TB.MoveLast
        CNTL.Max = TB.RecordCount
        TB.MoveFirst
        
        Do While Not TB.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = E_Date
    '    tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = TB.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(TB.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = Blm.heads(Val((Mid(CStr(TB.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = Blm.party1(TB.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If TB.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = TB.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If TB.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = TB.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        
    tb_t.Update
            TB.MoveNext
            CNTL.Value = CNTL.Value + 1
    Loop
    End If
    
    TB.Close

Set TB = db_t.OpenRecordset("PLAMT", dbOpenDynaset)
tb_t.AddNew
    tb_t.Fields("Code").Value = 0
    tb_t.Fields("Name").Value = "Net Profit Loss For the Month"
    If TB.Fields("PLMonth").Value < 0 Then
        tb_t.Fields("E_date").Value = E_Date
        tb_t.Fields("Credit").Value = TB.Fields("PlMonth").Value * -1
        tb_t.Fields("Debit").Value = 0
    Else
        tb_t.Fields("E_date").Value = E_Date
        tb_t.Fields("Debit").Value = TB.Fields("PlMonth").Value
        tb_t.Fields("Credit").Value = 0
    End If
tb_t.Update
TB.Close
tb_t.Close




db_t.Close
Db_M.Close

End Sub
Public Sub OverAllStockNew(E_Date As Date)
'On Error Resume Next
Dim Db_M As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim Ssql As String
Dim F As String
Dim tb_m As Recordset
Dim uname As String
Dim tb_p As Recordset
Dim tb_pr As Recordset
Dim tb_pk As Recordset
Dim tbop As Recordset
Dim tb_s As Recordset
Dim tb_i As Recordset
Dim tb_g As Recordset
Dim oprate As Currency
Dim Sstk As Currency
Dim Pstk As Currency
Dim PRstk As Currency
Dim qpk As Currency
Dim qop As Currency
Dim GNAME As String
Dim itname As String
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)




db_t.Execute "Drop Table TempItem"
Db_M.Execute "Select '" & Format(E_Date, "dd-MMM-yyyy") & "' As EDate,Code,Name,Unit,Stock into TempItem in '" & t_path & "' from Items Order by Code"

'Inward
Ssql = "Drop Table PurchaseStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as PQTY,Avg(Rate)as AVGPPRice into PurchaseStk in '" & t_path & "' from Purchase where V_Date <= #" & E_Date & "# Group BY item"
Db_M.Execute Ssql
'Sales
Ssql = "Drop Table SALEStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as SALEQTY,Avg(Rate)as SRATE into SALESTK in '" & t_path & "' from Sales where v_Date <= #" & E_Date & "# Group BY item"
Db_M.Execute Ssql
'Consumption
Ssql = "Drop Table ConsumeStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as CQTY,Avg(Rate)as SRATE into ConsumeSTK in '" & t_path & "' from Consume where v_Date <= #" & E_Date & "# Group BY item"
Db_M.Execute Ssql


Db_M.Close
db_t.Close





End Sub
Public Sub PeriodicStockNew(S_Date As Date, E_Date As Date)
'On Error Resume Next
Dim Db_M As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim Ssql As String
Dim F As String
Dim tb_m As Recordset
Dim uname As String
Dim tb_p As Recordset
Dim tb_pr As Recordset
Dim tb_pk As Recordset
Dim tbop As Recordset
Dim tb_s As Recordset
Dim tb_i As Recordset
Dim tb_g As Recordset
Dim oprate As Currency
Dim Sstk As Currency
Dim Pstk As Currency
Dim PRstk As Currency
Dim qpk As Currency
Dim qop As Currency
Dim GNAME As String
Dim opstk As Double
Dim itname As String
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)




db_t.Execute "Drop Table TempItem"
Db_M.Execute "Select '" & Format(S_Date, "dd-MMM-yyyy") & "' As SDate, '" & Format(E_Date, "dd-MMM-yyyy") & "' As EDate,Code,Name,Unit,Stock into TempItem in '" & t_path & "' from Items Order by Code"

'Inward
Ssql = "Drop Table PurchaseStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as PQTY,Avg(Rate)as AVGPPRice into PurchaseStk in '" & t_path & "' from Purchase where V_Date Between #" & S_Date & "# and #" & E_Date & "# Group BY item"
Db_M.Execute Ssql
'Sales
Ssql = "Drop Table SALEStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as SALEQTY,Avg(Rate)as SRATE into SALESTK in '" & t_path & "' from Sales where v_Date Between #" & S_Date & "# and #" & E_Date & "# Group BY item"
Db_M.Execute Ssql
'Consumption
Ssql = "Drop Table ConsumeStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as CQTY,Avg(Rate)as SRATE into ConsumeSTK in '" & t_path & "' from Consume where v_Date Between #" & S_Date & "# and #" & E_Date & "# Group BY item"
Db_M.Execute Ssql

Set tb_i = db_t.OpenRecordset("TempItem", dbOpenDynaset)
If Not tb_i.EOF Then
    Do While Not tb_i.EOF
        tb_i.Edit
            tb_i.Fields("Stock").Value = Blm.ITEMstocks(tb_i.Fields("Code").Value, S_Date)
        tb_i.Update
        tb_i.MoveNext
    Loop

End If
tb_i.Close
Db_M.Close
db_t.Close





End Sub

Public Sub ItemLedger(S_Date As Date, E_Date As Date, Itm As Double, p As Object)
Dim Db_M As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim tb_i As Recordset
Dim tb_a As Recordset
Dim Ssql As String
Dim F As String
Dim tb_m As Recordset
Dim Tbmill As Recordset
Dim itname As String
Dim uname As String
Dim MillName As String
Dim pname As String
Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "Delete from itmLedger"
db_t.Execute Ssql
p.Value = 0

Set tb_t = db_t.OpenRecordset("itmLedger", dbOpenTable)

Set tb_i = Db_M.OpenRecordset("Select * from Items Order by Code")
Set tb_a = Db_M.OpenRecordset("Select * from Acchart Order by Code")


'Purchases
Ssql = "Select * from Purchase where item = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by V_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        F = "Code = " & tb_m.Fields("Item").Value
        tb_i.FindFirst F
        If tb_i.NoMatch Then
            itname = ""
        Else
            itname = tb_i.Fields("Name").Value
            uname = tb_i.Fields("Unit").Value
        End If
        
        F = "Code = " & tb_m.Fields("Seller").Value
        tb_a.FindFirst F
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = tb_m.Fields("Item").Value
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Seller").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("Qty").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = Blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        tb_t.Fields("Inv_No").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Fields("Bales").Value = tb_m.Fields("Bales").Value
        
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'sales

'Purchases
Ssql = "Select * from Sales where item = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by V_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        F = "Code = " & tb_m.Fields("Item").Value
        tb_i.FindFirst F
        If tb_i.NoMatch Then
            itname = ""
        Else
            itname = tb_i.Fields("Name").Value
            uname = tb_i.Fields("Unit").Value
        End If
        
        F = "Code = " & tb_m.Fields("Seller").Value
        tb_a.FindFirst F
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = tb_m.Fields("Item").Value
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Seller").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("Qty").Value = tb_m.Fields("Qty").Value * -1
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = Blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        tb_t.Fields("Inv_No").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Fields("Bales").Value = tb_m.Fields("Bales").Value * -1
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'Issue

Ssql = "Select * from Issue where itemCode = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by v_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        F = "Code = " & tb_m.Fields("ItemCode").Value
        tb_i.FindFirst F
        If tb_i.NoMatch Then
            itname = ""
        Else
            itname = tb_i.Fields("Name").Value
            uname = tb_i.Fields("Unit").Value
        End If
                
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("Vno").Value
        tb_t.Fields("Item").Value = tb_m.Fields("ItemCode").Value
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("Qty").Value = tb_m.Fields("Qty").Value * -1
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = "IS"
        tb_t.Fields("Bales").Value = tb_m.Fields("Bales").Value * -1
        
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close

'Production

Ssql = "Select * from Production where itemCode = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by v_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        F = "Code = " & tb_m.Fields("ItemCode").Value
        tb_i.FindFirst F
        If tb_i.NoMatch Then
            itname = ""
        Else
            itname = tb_i.Fields("Name").Value
            uname = tb_i.Fields("Unit").Value
        End If
        
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("Vno").Value
        tb_t.Fields("Item").Value = tb_m.Fields("ItemCode").Value
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Qty").Value = tb_m.Fields("Qty").Value * -1
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = "PD"
        tb_t.Fields("Bales").Value = tb_m.Fields("Bales").Value
        
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close





tb_i.Close
tb_a.Close

tb_t.Close
Db_M.Close
db_t.Close
End Sub
Public Sub ItemLedgerNew(S_Date As Date, E_Date As Date, Itm As Double, p As Object)
Dim Db_M As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim tb_i As Recordset
Dim tb_a As Recordset
Dim Ssql As String
Dim F As String
Dim tb_m As Recordset
Dim Tbmill As Recordset
Dim itname As String
Dim uname As String
Dim MillName As String
Dim pname As String
Dim OpBales As Double
Dim OpWT As Double
Dim OpAmt As Double

Set Db_M = OpenDatabase(Blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "Delete from itmLedger2"
db_t.Execute Ssql
p.Value = 0

Set tb_t = db_t.OpenRecordset("itmLedger2", dbOpenTable)

Set tb_i = Db_M.OpenRecordset("Select * from Items where Code=" & Itm)
If Not tb_i.EOF Then
    itname = tb_i.Fields("Name").Value & ""
    OpWT = Val(tb_i.Fields("OPWT").Value & "")
    OpBales = Val(tb_i.Fields("OpBales").Value & "")
    OpAmt = Val(tb_i.Fields("Rate").Value & "")
End If
tb_i.Close
Set tb_a = Db_M.OpenRecordset("Select * from Acchart Order by Code")

'---------Opening-----------------
'Purchases
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from Purchase where item = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT + Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales + Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt + Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close

'Sale Return
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from SalesReturn where item = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT + Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales + Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt + Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close

'IssueProcessing
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from Issue where DrCode = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT + Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales + Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt + Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close

'Production
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Rate * Qty) as A from Production where ItemCode = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT + Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales + Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt + Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close

'SH Production
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from SHProduction where Item = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT + Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales + Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt + Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close


'Sales
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from Sales where item = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT - Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales - Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt - Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close

'Purchase Returns
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from PurchaseReturn where item = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT - Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales - Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt - Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close


'Issue
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Amount) as A from Issue where ItemCode = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT - Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales - Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt - Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close

'Production
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B,Sum(Rate * Qty) as A from Production where CrCode = " & Itm
Ssql = Ssql & " and v_date < #" & S_Date & "#"
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    OpWT = OpWT - Val(tb_m.Fields("Q").Value & "")
    OpBales = OpBales - Val(tb_m.Fields("B").Value & "")
    OpAmt = OpAmt - Val(tb_m.Fields("A").Value & "")
End If
tb_m.Close


        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = S_Date
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("inWeight").Value = OpWT
        tb_t.Fields("InBales").Value = OpBales
        'tb_t.Fields("v_TYPE").Value = Blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        
        tb_t.Fields("InAmount").Value = OpAmt
        tb_t.Fields("OutWeight").Value = 0
        tb_t.Fields("OutBales").Value = 0
        tb_t.Fields("OutAmount").Value = 0
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Remarks").Value = "Opening Balance" ' as On " & Format(S_Date, "dd-MMM-yyyy")
        tb_t.Update

'MsgBox "Test"

'--------End of Opening------------
'Purchases
Ssql = "Select * from Purchase where item = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by V_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        
        F = "Code = " & tb_m.Fields("Seller").Value
        tb_a.FindFirst F
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Seller").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("InWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = Blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        tb_t.Fields("Inv_No").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Fields("InBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("InAmount").Value = tb_m.Fields("Amount").Value
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("OutWeight").Value = 0
        tb_t.Fields("OutBales").Value = 0
        tb_t.Fields("OutAmount").Value = 0
        tb_t.Fields("Remarks").Value = tb_m.Fields("LorryNo") & " Job No." & tb_m.Fields("JobNo") 'tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close

'Sales Ret
Ssql = "Select * from SalesReturn where item = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by V_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        
        F = "Code = " & tb_m.Fields("Seller").Value
        tb_a.FindFirst F
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Seller").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("InWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = Blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        tb_t.Fields("Inv_No").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Fields("InBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("InAmount").Value = tb_m.Fields("Amount").Value
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("OutWeight").Value = 0
        tb_t.Fields("OutBales").Value = 0
        tb_t.Fields("OutAmount").Value = 0
        tb_t.Fields("Remarks").Value = tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close



'sales

Ssql = "Select * from Sales where item = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by V_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        
        F = "Code = " & tb_m.Fields("Seller").Value
        tb_a.FindFirst F
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Seller").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("OutWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("OutBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("OutAmount").Value = tb_m.Fields("Amount").Value
        tb_t.Fields("InWeight").Value = 0
        tb_t.Fields("inBales").Value = 0
        tb_t.Fields("inAmount").Value = 0
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = Blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        tb_t.Fields("Inv_No").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        
        tb_t.Fields("Remarks").Value = tb_m.Fields("LorryNo") & " Job No." & tb_m.Fields("JobNo") 'tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'Purchase Return

Ssql = "Select * from PurchaseReturn where item = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by V_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        
        F = "Code = " & tb_m.Fields("Seller").Value
        tb_a.FindFirst F
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Seller").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("OutWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = Blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        tb_t.Fields("Inv_No").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Fields("OutBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("OutAmount").Value = tb_m.Fields("Amount").Value
        tb_t.Fields("InWeight").Value = 0
        tb_t.Fields("inBales").Value = 0
        tb_t.Fields("inAmount").Value = 0
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Remarks").Value = tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'Issue

Ssql = "Select * from Issue where itemCode = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by v_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
                
        F = "Code = " & tb_m.Fields("DrCode").Value
        tb_a.FindFirst F
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("Vno").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("OutWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = "IS"
        tb_t.Fields("OutBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("OutAmount").Value = tb_m.Fields("Amount").Value
        tb_t.Fields("InWeight").Value = 0
        tb_t.Fields("inBales").Value = 0
        tb_t.Fields("inAmount").Value = 0
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Remarks").Value = "Ref. No." & tb_m.Fields("RefNo").Value ' tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'Issue Processing

Ssql = "Select * from Issue where DrCode = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by v_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
                
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("Vno").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("OutWeight").Value = 0
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = "IP"
        tb_t.Fields("OutBales").Value = 0
        tb_t.Fields("OutAmount").Value = 0
        tb_t.Fields("InWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("inBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("inAmount").Value = tb_m.Fields("Amount").Value
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Remarks").Value = "Ref. No." & tb_m.Fields("RefNo").Value ' tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'Production

Ssql = "Select * from Production where itemCode = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by v_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("Vno").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("InWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("OutWeight").Value = 0
        tb_t.Fields("OutBales").Value = 0
        tb_t.Fields("OutAmount").Value = 0
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = "PD"
        tb_t.Fields("InBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("InAmount").Value = tb_m.Fields("Rate").Value * tb_m.Fields("Qty").Value
        tb_t.Fields("Remarks").Value = "Ref. No." & tb_m.Fields("RefNo").Value ' tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close

'Production Processing

Ssql = "Select * from Production where CrCode = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by v_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("Vno").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("InWeight").Value = 0
        tb_t.Fields("OutWeight").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("OutBales").Value = tb_m.Fields("Bales").Value
        tb_t.Fields("OutAmount").Value = tb_m.Fields("Rate").Value * tb_m.Fields("Qty").Value
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = "PP"
        tb_t.Fields("InBales").Value = 0
        tb_t.Fields("InAmount").Value = 0
        tb_t.Fields("Remarks").Value = "Ref. No." & tb_m.Fields("RefNo").Value ' tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close

'SH Production (Transfer)

Ssql = "Select * from SHProduction where item = " & Itm
Ssql = Ssql & " and v_date Between #" & S_Date & "# and #" & E_Date & "# Order by v_date"
p.Value = 0
Set tb_m = Db_M.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        
        tb_t.AddNew
        tb_t.Fields("s_date").Value = S_Date
        tb_t.Fields("e_date").Value = E_Date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = itname
        If tb_m.Fields("Qty").Value > 0 Then
            tb_t.Fields("InWeight").Value = tb_m.Fields("Qty").Value
            tb_t.Fields("InBales").Value = tb_m.Fields("Bales").Value
            tb_t.Fields("InAmount").Value = tb_m.Fields("Rate").Value * tb_m.Fields("Qty").Value
            tb_t.Fields("OutWeight").Value = 0
            tb_t.Fields("OutBales").Value = 0
            tb_t.Fields("OutAmount").Value = 0
        End If
        If tb_m.Fields("Qty").Value < 0 Then
            tb_t.Fields("OutWeight").Value = Abs(tb_m.Fields("Qty").Value)
            tb_t.Fields("outBales").Value = Abs(tb_m.Fields("Bales").Value)
            tb_t.Fields("OutAmount").Value = Abs(tb_m.Fields("Rate").Value * tb_m.Fields("Qty").Value)
            tb_t.Fields("InWeight").Value = 0
            tb_t.Fields("InBales").Value = 0
            tb_t.Fields("InAmount").Value = 0
        End If
        
        tb_t.Fields("Qty").Value = 0
        tb_t.Fields("Bales").Value = 0
        tb_t.Fields("Amount").Value = 0
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = "TR"
        
        tb_t.Fields("Remarks").Value = tb_m.Fields("Remarks").Value & ""
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close





tb_a.Close

tb_t.Close
Db_M.Close
db_t.Close
End Sub

Private Sub Class_Initialize()
Set Blm = New bloom1
m_path = Blm.patHmain
't_path = "c:\soft\Accounts\book.mdb"
't_path = CurDir & "\" & "book.mdb"
t_path = App.path & "\book.mdb"
End Sub

Private Sub Class_Terminate()
Set Blm = Nothing
End Sub
