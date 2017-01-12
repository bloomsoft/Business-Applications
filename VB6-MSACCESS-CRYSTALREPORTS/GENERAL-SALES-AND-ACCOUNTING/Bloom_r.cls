VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bloom_r"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Blm1 As New bloom1
Public Sub shortfallPurchase(v_date As Date, cntl As Control)
Dim DB_M As Database
Dim tb_m As Recordset
Dim tb_m2 As Recordset
Dim tb_m3 As Recordset
Dim tb_m4 As Recordset
Dim tb_t As Recordset
Dim tb_pak As Recordset
Dim tb_k As Recordset
Dim tb_p As Recordset
Dim tb_itm As Recordset
Dim tb_kf As Recordset
Dim tb_pakf As Recordset
Dim tb_emp As Recordset
Dim db_t As Database
Dim Ssql As String
Dim t_gaz As Currency
Dim t_r As Currency
Dim t_d As Currency
Dim cgaz As Currency
Dim t_ded As Currency
cntl.Value = 0
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
 Ssql = "delete from contract"
 db_t.Execute Ssql
 Set tb_t = db_t.OpenRecordset("contract", dbOpenTable)
 
    Ssql = "Select * from item order by COde"
    Set tb_itm = DB_M.OpenRecordset(Ssql)
    
    Ssql = "Select * from Parties order by COde"
    Set tb_p = DB_M.OpenRecordset(Ssql)
    
    Ssql = "select a.Party,B.JobNo,sum(b.Quantity-b.LKamiValue)as rec from In_MST a,In_DTL b where a.P_No=b.P_No and "
    Ssql = Ssql & " a.v_Date <= #" & v_date & "# Group by a.Party,b.Jobno"
    
    Set tb_k = DB_M.OpenRecordset(Ssql)
    
    
    Ssql = "select * from Pcontract where Cont_date <= #" & v_date & "# ORDER BY SellerCode,cont_no"
    Set tb_m = DB_M.OpenRecordset(Ssql)
    If Not tb_m.EOF Then
    tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
    t_gaz = 0
    t_d = 0
    t_r = 0
    t_ded = 0
    If Not tb_k.EOF Then
        tb_k.Filter = "Jobno = " & tb_m.Fields("Cont_no").Value & " And Party = " & tb_m.Fields("SellerCode").Value
        If tb_k.EOF Then
            t_gaz = 0
        Else
            t_gaz = Val(tb_k.Fields("Rec").Value & "")
            tb_k.MoveFirst
        End If
        tb_k.Filter = ""
    End If
        
   
   'If tb_m.Fields("Complete").Value <> 1 Then
    'If cgaz < tb_m.Fields("quantity").Value Then
   
    tb_t.AddNew
    tb_t.Fields("v_date").Value = v_date
    tb_t.Fields("cont_no").Value = tb_m.Fields("cont_no").Value
    tb_t.Fields("cont_date").Value = tb_m.Fields("Cont_date").Value
    p$ = ""
    If Not tb_p.EOF Then
        tb_p.FindFirst "Code = " & tb_m.Fields("SellerCOde").Value
        If tb_p.NoMatch Then
            p$ = ""
        Else
            p$ = tb_p.Fields("Name").Value
            
        End If
        tb_p.MoveFirst
    End If
    tb_t.Fields("party").Value = p$ 'blm.party1(tb_m.Fields("party").Value)
    tb_t.Fields("rate").Value = tb_m.Fields("Clothrate").Value
    tb_t.Fields("del_date").Value = tb_m.Fields("deldate").Value
    p$ = ""
    If Not tb_itm.EOF Then
        tb_itm.FindFirst "Code = " & tb_m.Fields("ClothCode").Value
        If tb_itm.NoMatch Then
            p$ = ""
        Else
            p$ = tb_itm.Fields("Name").Value '& "   " & tb_itm.Fields("Width").Value
            
        End If
        tb_itm.MoveFirst
    End If
    tb_t.Fields("item").Value = p$ 'blm.item1(tb_m.Fields("item").Value)
    tb_t.Fields("it_code").Value = tb_m.Fields("ClothCode").Value
    tb_t.Fields("quantity").Value = tb_m.Fields("quantity").Value
    tb_t.Fields("rec_gazana").Value = t_gaz ' - t_r - t_d - t_ded
    tb_t.Update
    
    'End If
 'End If
 tb_m.MoveNext
 Loop
End If
tb_p.Close
tb_k.Close

tb_itm.Close

db_t.Close
tb_m.Close


End Sub
Public Sub shortfallSale(v_date As Date, cntl As Control)
Dim DB_M As Database
Dim tb_m As Recordset
Dim tb_m2 As Recordset
Dim tb_m3 As Recordset
Dim tb_m4 As Recordset
Dim tb_t As Recordset
Dim tb_pak As Recordset
Dim tb_k As Recordset
Dim tb_p As Recordset
Dim tb_itm As Recordset
Dim tb_kf As Recordset
Dim tb_pakf As Recordset
Dim tb_emp As Recordset
Dim db_t As Database
Dim Ssql As String
Dim t_gaz As Currency
Dim t_r As Currency
Dim t_d As Currency
Dim cgaz As Currency
Dim t_ded As Currency
cntl.Value = 0
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
 Ssql = "delete from contract"
 db_t.Execute Ssql
 Set tb_t = db_t.OpenRecordset("contract", dbOpenTable)
 
    Ssql = "Select * from item order by COde"
    Set tb_itm = DB_M.OpenRecordset(Ssql)
    
    Ssql = "Select * from Parties order by COde"
    Set tb_p = DB_M.OpenRecordset(Ssql)
    
    Ssql = "select a.Party,B.JobNo,sum(b.Qty-b.LKamiValue)as rec from Sale_1 a,Sale_2 b where a.MNO=b.MNO and a.Inv_No=b.Inv_No and "
    Ssql = Ssql & " a.inv_Date <= #" & v_date & "# Group by a.Party,b.Jobno"
    
    Set tb_k = DB_M.OpenRecordset(Ssql)
    
    
    Ssql = "select * from Scontract where Cont_date <= #" & v_date & "# ORDER BY SellerCode,cont_no"
    Set tb_m = DB_M.OpenRecordset(Ssql)
    If Not tb_m.EOF Then
    tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
    t_gaz = 0
    t_d = 0
    t_r = 0
    t_ded = 0
    If Not tb_k.EOF Then
        tb_k.Filter = "Jobno = " & tb_m.Fields("Cont_no").Value & " And Party = " & tb_m.Fields("SellerCode").Value
        If tb_k.EOF Then
            t_gaz = 0
        Else
            t_gaz = Val(tb_k.Fields("Rec").Value & "")
            tb_k.MoveFirst
        End If
        tb_k.Filter = ""
    End If
        
   
   'If tb_m.Fields("Complete").Value <> 1 Then
    'If cgaz < tb_m.Fields("quantity").Value Then
   
    tb_t.AddNew
    tb_t.Fields("v_date").Value = v_date
    tb_t.Fields("cont_no").Value = tb_m.Fields("cont_no").Value
    tb_t.Fields("cont_date").Value = tb_m.Fields("Cont_date").Value
    p$ = ""
    If Not tb_p.EOF Then
        tb_p.FindFirst "Code = " & tb_m.Fields("SellerCOde").Value
        If tb_p.NoMatch Then
            p$ = ""
        Else
            p$ = tb_p.Fields("Name").Value
            
        End If
        tb_p.MoveFirst
    End If
    tb_t.Fields("party").Value = p$ 'blm.party1(tb_m.Fields("party").Value)
    tb_t.Fields("rate").Value = tb_m.Fields("Clothrate").Value
    tb_t.Fields("del_date").Value = tb_m.Fields("deldate").Value
    p$ = ""
    If Not tb_itm.EOF Then
        tb_itm.FindFirst "Code = " & tb_m.Fields("ClothCode").Value
        If tb_itm.NoMatch Then
            p$ = ""
        Else
            p$ = tb_itm.Fields("Name").Value '& "   " & tb_itm.Fields("Width").Value
            
        End If
        tb_itm.MoveFirst
    End If
    tb_t.Fields("item").Value = p$ 'blm.item1(tb_m.Fields("item").Value)
    tb_t.Fields("it_code").Value = tb_m.Fields("ClothCode").Value
    tb_t.Fields("quantity").Value = tb_m.Fields("quantity").Value
    tb_t.Fields("rec_gazana").Value = t_gaz ' - t_r - t_d - t_ded
    tb_t.Update
    
    'End If
 'End If
 tb_m.MoveNext
 Loop
End If
tb_p.Close
tb_k.Close

tb_itm.Close

db_t.Close
tb_m.Close


End Sub

Public Sub day_due2_TEMP(v_date As Date, cntl As Control, CNTL2 As Control)
Dim DB_M As Database
Dim ddate As Date
Dim db_t As Database
Dim test As Integer
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim t_cred As Currency
Dim Ssql As String
Dim tbtrial As Recordset
Dim tb_c As Recordset
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Ssql = "delete from vou_temp"
db_t.Execute Ssql
cntl.Value = 0
Ssql = "delete from trial"
db_t.Execute (Ssql)
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
Ssql = "SELECT * FROM PARTIES ORDER BY CODE"
Set tb_c = DB_M.OpenRecordset(Ssql)
If Not tb_c.EOF Then
tb_c.MoveLast
cntl.Max = tb_c.RecordCount
tb_c.MoveFirst
    Do While Not tb_c.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    CNTL2.SimpleText = tb_c.Fields("NAME").Value
        Ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUCHER WHERE V_DATE <= #" & v_date & "#"
        Ssql = Ssql & " AND PARTY = " & tb_c.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) > 0 "
        Set tb = DB_M.OpenRecordset(Ssql)
        If Not tb.EOF Then
        If Not IsNull(tb.Fields("DEB").Value) Then
            t_cred = tb.Fields("DEB").Value
        End If
        End If
'        MsgBox TB_C.Fields("DESCRIPTION").Value & " pAID = " & t_cred
        tb.Close
    CNTL2.SimpleText = CNTL2.SimpleText & " tOTAL pAID = " & t_cred
     'TOTAL Sales ITEMS ADJUSTMENTS
     If t_cred > 0 Then
        
 '       MsgBox t_cred
        Ssql = "SELECT a.Inv_No,a.MNO,Sum((b.Qty-b.LKamiValue)*b.Rate) as Amt FROM SALE_1 a,SALE_2 b WHERE a.Inv_No=b.Inv_No and a.MNO=b.MNO and a.PARTY = " & tb_c.Fields("CODE").Value
        Ssql = Ssql & " and a.INV_DATE <= #" & v_date & "# GROUP BY a.INV_NO,a.MNO"
        'Clipboard.SetText ssql
        'MsgBox ssql
        Set tb = DB_M.OpenRecordset(Ssql)
        If Not tb.EOF Then
           tb.MoveLast
            Do While Not tb.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
                ddate = DateAdd("d", 1, tb.Fields("INVV_date").Value)
                'If ddate <= v_date Then
                    t_cred = t_cred - (tb.Fields("AMT").Value)
                'End If
                'End If
 '               MsgBox t_cred
                If t_cred <= 0 Then Exit Do
            Else
                Exit Do
            End If
            tb.MovePrevious
            Loop
            
            If tb.BOF Then tb.MoveFirst
'            MsgBox tb.Fields("PARTY").Value & " " & t_cred
            If Not tb.BOF Then
            Do While Not tb.EOF
                    DoEvents
                    CNTL2.SimpleText = CNTL2.SimpleText & " tOTAL pAID = " & t_cred
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = tb.Fields("invv_date").Value
                    tb_t.Fields("due_date").Value = tb.Fields("invv_date").Value + 1 'tb.Fields("days").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
'                    ssql = blm.Item1(tb.Fields("item").Value) & " "
'                    ssql = ssql & "Rate = " & tb.Fields("rate").Value & " "
'                    ssql = ssql & "Quantity = " & tb.Fields("quantity").Value & " "
'                    ssql = ssql & "Days = " & tb.Fields("days").Value
'                    If tb.Fields("item").Value < 3000 Then
'                        ssql = ssql & " Amount = " & tb.Fields("rate").Value * tb.Fields("quantity").Value
'                    Else
'                        ssql = ssql & " Amount = " & tb.Fields("rate").Value * tb.Fields("quantity").Value / 100
'                    End If
'                    tb_t.Fields("remarks").Value = ssql
                    tb_t.Fields("Remarks").Value = "Bill No." & tb.Fields("Inv_No").Value & " of " & MonthName(tb.Fields("MNO"))
                    If t_cred < 0 Then
                        t_cred = (tb.Fields("AMT").Value) + t_cred
                    Else
                        t_cred = tb.Fields("AMT").Value
                    End If
'                    MsgBox t_cred
                    tb_t.Fields("CREDIT").Value = t_cred
                    tb_t.Update
                't_cred = -1
                tb.MoveNext
                If Not tb.EOF Then
                    t_cred = tb.Fields("AMT").Value
                End If
            Loop
            End If
            Else
            
'            MsgBox t_cred
             tb_t.AddNew
             tb_t.Fields("code").Value = tb_c.Fields("Code").Value
             tb_t.Fields("name").Value = Blm1.party1(tb_c.Fields("Code").Value)
             'tb_T.Fields("Desc").Value = s
'            MsgBox t_cred
             tb_t.Fields("CREDIT").Value = Format(t_cred, "#.00")
             tb_t.Update
            't_cred = 0
        End If
        tb.Close
    End If
      tb_c.MoveNext
      cntl.Value = cntl.Value + 1
      DoEvents
    Loop
End If
tb_c.Close
tb_t.Close
db_t.Close
DB_M.Close
End Sub

Public Sub DayBook(ByVal VDate As Date)
Dim DB_M As Database
Dim db_t As Database
Dim Ssql As String
Dim TBT As Recordset
Dim tb As Recordset
Dim PReCash As Double
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")
Ssql = "Delete from Voucher"
db_t.Execute Ssql

Set TBT = db_t.OpenRecordset("Voucher", dbOpenTable)
Ssql = "Select * from Pre_Cash where V_Date=#" & VDate & "#"
Set tb = DB_M.OpenRecordset(Ssql)
If Not tb.EOF Then
    PReCash = tb.Fields("Opening").Value
    TBT.AddNew
        TBT.Fields("V_Date").Value = tb.Fields("V_Date").Value
        TBT.Fields("Party").Value = "Cash Brought Forward"
        TBT.Fields("Remarks").Value = "Opening Cash in Hand"
        TBT.Fields("Debit").Value = 0
        TBT.Fields("Credit").Value = PReCash
    TBT.Update
End If
tb.Close

Ssql = "Select * from Voucher where E_type=9 and V_Date=#" & VDate & "#"
Set tb = DB_M.OpenRecordset(Ssql)
If Not tb.EOF Then
    Do While Not tb.EOF
    TBT.AddNew
        TBT.Fields("V_Date").Value = tb.Fields("V_Date").Value
        TBT.Fields("Party").Value = Blm1.party1(tb.Fields("Party").Value)
        TBT.Fields("Remarks").Value = tb.Fields("Remarks").Value & ""
        TBT.Fields("Debit").Value = tb.Fields("Debit").Value
        TBT.Fields("Credit").Value = tb.Fields("Credit").Value
    TBT.Update
    tb.MoveNext
    Loop
End If
tb.Close
TBT.Close
db_t.Close
DB_M.Close
BankBalances VDate
End Sub
Public Sub BankBalances(ByVal VDate As Date)
Dim DB_M As Database
Dim db_t As Database
Dim Ssql As String
Dim TBT As Recordset
Dim tb As Recordset
Dim PReCash As Double
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")
Ssql = "Delete from Banks"
db_t.Execute Ssql

Set TBT = db_t.OpenRecordset("Banks", dbOpenTable)

Ssql = "Select b.Code,Sum(a.Debit) - Sum(a.Credit) as Bal from Voucher a,Parties b where a.Party=b.Code and b.Atype=5 and a.V_Date<=#" & VDate & "# Group by b.Code"
Set tb = DB_M.OpenRecordset(Ssql)
If Not tb.EOF Then
    Do While Not tb.EOF
    TBT.AddNew
        TBT.Fields("Party").Value = Blm1.party1(tb.Fields("Code").Value)
        TBT.Fields("Debit").Value = tb.Fields("Bal").Value
            
        
    TBT.Update
    tb.MoveNext
    Loop
End If
tb.Close
TBT.Close
db_t.Close
DB_M.Close
End Sub

Public Sub ItemLedger(ByVal S_Date As Date, ByVal E_Date As Date, ItemCOde As Double)
Dim DB_M As Database
Dim db_t As Database
Dim Ssql As String
Dim OpStock As Currency
Dim OpStockB As Double
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")
Dim TBT As Recordset
Ssql = "Drop Table SaleLedger"
db_t.Execute Ssql

Ssql = "Drop Table PurchaseLedger"
db_t.Execute Ssql


OpStock = Blm1.ITEMstocks(ItemCOde, S_Date - 1)
OpStockB = Blm1.ITEMBalesstocks(ItemCOde, S_Date - 1)
Ssql = "Select a.Inv_no,a.Inv_Date,a.Inv_Type,a.Party,a.MNO,b.Item,(b.Qty * -1) as Qty,b.Rate,(b.DiscountItm * -1) as DiscountItm,c.name as ItmName,d.name as PartyName,e.Name as wareHousename into SaleLedger in '" & App.Path & "\Book.mdb' from Sale_1 a,Sale_2 b,Item c, Parties d,warehouse e where a.Inv_no=b.inv_No and a.MNO=b.MNO and a.Inv_Type=b.Inv_Type and a.Party=D.Code and b.WareHouse=e.Code and b.Item=c.Code and a.Inv_Date Between #" & S_Date & "# and #" & E_Date & "# and b.Item=" & ItemCOde
'Clipboard.SetText ssql
'MsgBox ssql
DB_M.Execute Ssql

Ssql = "Select a.P_no,a.v_Date,a.Inv_Type,a.Party,b.Item,b.Quantity,b.Rate,b.Scheme as DiscountItm,c.name as ItmName,d.name as PartyName,e.name as WareHouseName into PurchaseLedger in '" & App.Path & "\Book.mdb' from In_MST a,In_DTL b,Item c, Parties d,warehouse e where a.P_no=b.P_No and a.Inv_Type=b.Inv_Type and a.Party=D.Code and b.Item=c.Code and b.WareHouse=e.Code and a.v_Date Between #" & S_Date & "# and #" & E_Date & "# and b.Item=" & ItemCOde
DB_M.Execute Ssql

'ssql = "Select 0 as Inv_no," & S_Date & " as inv_Date,0 as Inv_Type,0 as Party,0 as MNO," & ItemCOde & " as ItemCode," & OpStock & ",0 as Rate,0 as DiscountItm," & Blm1.Item1(ItemCode) & " as ItmName,'' as PartyName into OpStockLedger in '" & App.Path & "\Book.mdb' from Sale_1 a,Sale_2 b,Item c, Parties d where a.Inv_no=b.inv_No and a.MNO=b.MNO and a.Inv_Type=b.Inv_Type and a.Party=D.Code and b.Item=c.Code and a.Inv_Date Between #" & S_Date & "# and #" & E_Date & "# and b.Item=" & ItemCOde
'db_m.Execute ssql
Ssql = "delete from OpStockLedger"
db_t.Execute Ssql
Set TBT = db_t.OpenRecordset("OpStockLedger", dbOpenTable)
TBT.AddNew
    TBT.Fields("Item").Value = ItemCOde
    TBT.Fields("ItmName").Value = Blm1.Item1(CLng(ItemCOde))
    TBT.Fields("Quantity").Value = OpStock
    TBT.Fields("DiscountITm").Value = OpStockB
    TBT.Fields("V_date").Value = S_Date - 1
TBT.Update
TBT.Close
db_t.Close


DB_M.Close

End Sub
Public Sub CreateStock(ByVal E_Date As Date)
Dim DB_M As Database
Dim db_t As Database
Dim Ssql As String

Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")

Ssql = "Drop Table Items"
db_t.Execute Ssql

Ssql = "Drop Table Groups"
db_t.Execute Ssql

Ssql = "Drop table PSTK"
db_t.Execute Ssql

Ssql = "Drop table SSTK"
db_t.Execute Ssql

Ssql = "Drop table RawSTK"
db_t.Execute Ssql

db_t.Close

Ssql = "Select * into Items in '" & App.Path & "\Book.mdb' from Item"
DB_M.Execute Ssql

Ssql = "Select * into Groups in '" & App.Path & "\Book.mdb' from Groups"
DB_M.Execute Ssql

Ssql = "select Item,Sum(Quantity) as QSTOCK,Sum(Scheme) as BStock,Sum(Quantity*rate) as Amt into PSTK in '" & App.Path & "\Book.mdb" & "' From In_DTl Where v_date <=#" & E_Date & "# Group By Item"
'MsgBox ssql
DB_M.Execute Ssql

Ssql = "select Item,Sum(Qty) as QSTOCK,Sum(Discountitm) as BStock,Sum(Qty*rate) as Amt into SSTK in '" & App.Path & "\Book.mdb" & "' From Sale_2 Where Inv_date <=#" & E_Date & "# Group By Item"
DB_M.Execute Ssql

Ssql = "select RawCode,Sum(QtyUsed)as QSTOCK into RawSTK in '" & App.Path & "\Book.mdb" & "' From Rawsales Where Inv_date <=#" & E_Date & "# Group By RawCode"
DB_M.Execute Ssql


DB_M.Close
End Sub

Public Sub CreateStockBetweenDates(ByVal S_Date As Date, ByVal E_Date As Date)
Dim DB_M As Database
Dim db_t As Database
Dim Ssql As String

Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")

Ssql = "Drop table PSTK"
db_t.Execute Ssql

Ssql = "Drop table SSTK"
db_t.Execute Ssql

Ssql = "Drop table RawSTK"
db_t.Execute Ssql

Ssql = "Drop table CPSTK"
db_t.Execute Ssql

Ssql = "Drop table CSSTK"
db_t.Execute Ssql

Ssql = "Drop table CRawSTK"
db_t.Execute Ssql

db_t.Close

'Opening Stocks
Ssql = "select Item,Sum(Scheme) as BSTOCK,Sum(Quantity) as QSTOCK,Sum(Quantity*rate) as Amt into PSTK in '" & App.Path & "\Book.mdb" & "' From In_DTl Where v_date <#" & S_Date & "# Group By Item"
'MsgBox ssql
DB_M.Execute Ssql

Ssql = "select Item,Sum(DiscountItm) as BSTOCK,Sum(Qty) as QSTOCK,Sum(Qty*rate) as Amt into SSTK in '" & App.Path & "\Book.mdb" & "' From Sale_2 Where Inv_date <#" & S_Date & "# Group By Item"
DB_M.Execute Ssql

Ssql = "select RawCode,Sum(QtyUsed)as QSTOCK into RawSTK in '" & App.Path & "\Book.mdb" & "' From Rawsales Where Inv_date <#" & S_Date & "# Group By RawCode"
DB_M.Execute Ssql

'Stocks Between Dates

Ssql = "select Item,Sum(Scheme) as BSTOCK,Sum(Quantity) as QSTOCK,Sum(Quantity*rate) as Amt into CPSTK in '" & App.Path & "\Book.mdb" & "' From In_DTl Where v_date  Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
'MsgBox ssql
DB_M.Execute Ssql

Ssql = "select Item,Sum(DiscountItm) as BSTOCK,Sum(Qty) as QSTOCK,Sum(Qty*rate) as Amt into CSSTK in '" & App.Path & "\Book.mdb" & "' From Sale_2 Where Inv_date Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
DB_M.Execute Ssql

Ssql = "select RawCode,Sum(QtyUsed)as QSTOCK into CRawSTK in '" & App.Path & "\Book.mdb" & "' From Rawsales Where Inv_date between #" & S_Date & "# and #" & E_Date & "# Group By RawCode"
'MsgBox ssql
DB_M.Execute Ssql


DB_M.Close
End Sub

Public Sub LedgerBetweenDates(ByVal S_Date As Date, ByVal E_Date As Date, ByVal Accode As Long)
Dim DB_M As Database
Dim db_t As Database
Dim Ssql As String
Dim TBT As Recordset
Dim Tbm As Recordset
Dim Acname As String
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")

Ssql = "Drop table Ledger"
db_t.Execute Ssql
Acname = Blm1.party1(Accode)

db_t.Close

'Ledger Between Dates
Ssql = "Select *,'" & Acname & "' as AcName into Ledger in '" & App.Path & "\Book.mdb" & "' from Voucher where v_date Between #" & S_Date & "# and #" & E_Date & "# and Party = " & Accode & " Order By V_date"
DB_M.Execute Ssql


Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Ssql = "Select Sum(Debit - Credit) as Bal from Voucher where Party = " & Accode & " and v_date < #" & S_Date & "#"
Set Tbm = DB_M.OpenRecordset(Ssql)
Set TBT = db_t.OpenRecordset("Ledger", dbOpenTable)
If Not IsNull(Tbm.Fields("Bal").Value) Then
    TBT.AddNew
        TBT.Fields("v_date").Value = S_Date - 1
        TBT.Fields("Remarks").Value = "Opening Balance"
        TBT.Fields("Account").Value = Accode
        TBT.Fields("Acname").Value = Acname
        If Tbm.Fields("Bal").Value < 0 Then
            TBT.Fields("Credit").Value = Tbm.Fields("Bal").Value * -1
            TBT.Fields("Debit").Value = 0
        Else
            TBT.Fields("Debit").Value = Tbm.Fields("Bal").Value
            TBT.Fields("Credit").Value = 0
        End If
    TBT.Update
End If
TBT.Close
db_t.Close
Tbm.Close
DB_M.Close
End Sub
Public Sub TrialBalance(ByVal E_Date As Date)
Dim DB_M As Database
Dim db_t As Database
Dim Ssql As String
Dim OPCash As Currency, CashBF As Currency
Set db_t = OpenDatabase(App.Path & "\Book.mdb")
Set DB_M = OpenDatabase(App.Path & "\Bloom.mdb")

Ssql = "Drop table Trial"
db_t.Execute Ssql

Ssql = "select a.Party,b.name as AcName,Sum(Debit - Credit) as Bal into Trial in '" & App.Path & "\Book.Mdb' from Voucher a,Parties b where a.Party = b.code and a.v_date <= #" & E_Date & "# Group by a.Party,b.name"
DB_M.Execute Ssql
DB_M.Close
Blm1.Cash E_Date, OPCash, CashBF
Ssql = "Insert into Trial (Party,ACName,Bal)Values(10000,'Cash In Hand'," & CashBF & ")"
db_t.Execute Ssql

db_t.Close
End Sub

