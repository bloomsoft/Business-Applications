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

Public Sub ItemLedger(ByVal S_Date As Date, ByVal E_Date As Date, ItemCOde As Double)
Dim db_m As Database
Dim Db_T As Database
Dim ssql As String
Dim OpStock As Currency

Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")
Dim TBT As Recordset
ssql = "Drop Table SaleLedger"
Db_T.Execute ssql

ssql = "Drop Table PurchaseLedger"
Db_T.Execute ssql


OpStock = Blm1.ITEMstocks(ItemCOde, S_Date - 1)
ssql = "Select a.Inv_no,a.Inv_Date,a.Inv_Type,a.Party,a.MNO,b.Item,(b.Qty * -1) as Qty,b.Rate,b.DiscountItm,c.name as ItmName,d.name as PartyName into SaleLedger in '" & App.Path & "\Book.mdb' from Sale_1 a,Sale_2 b,Item c, Parties d where a.Inv_no=b.inv_No and a.MNO=b.MNO and a.Inv_Type=b.Inv_Type and a.Party=D.Code and b.Item=c.Code and a.Inv_Date Between #" & S_Date & "# and #" & E_Date & "# and b.Item=" & ItemCOde
'Clipboard.SetText ssql
'MsgBox ssql
db_m.Execute ssql

ssql = "Select a.P_no,a.v_Date,a.Inv_Type,a.Party,b.Item,b.Quantity,b.Rate,0 as DiscountItm,c.name as ItmName,d.name as PartyName into PurchaseLedger in '" & App.Path & "\Book.mdb' from In_MST a,In_DTL b,Item c, Parties d where a.P_no=b.P_No and a.Inv_Type=b.Inv_Type and a.Party=D.Code and b.Item=c.Code and a.v_Date Between #" & S_Date & "# and #" & E_Date & "# and b.Item=" & ItemCOde
db_m.Execute ssql

'ssql = "Select 0 as Inv_no," & S_Date & " as inv_Date,0 as Inv_Type,0 as Party,0 as MNO," & ItemCOde & " as ItemCode," & OpStock & ",0 as Rate,0 as DiscountItm," & Blm1.Item1(ItemCode) & " as ItmName,'' as PartyName into OpStockLedger in '" & App.Path & "\Book.mdb' from Sale_1 a,Sale_2 b,Item c, Parties d where a.Inv_no=b.inv_No and a.MNO=b.MNO and a.Inv_Type=b.Inv_Type and a.Party=D.Code and b.Item=c.Code and a.Inv_Date Between #" & S_Date & "# and #" & E_Date & "# and b.Item=" & ItemCOde
'db_m.Execute ssql
ssql = "delete from OpStockLedger"
Db_T.Execute ssql
Set TBT = Db_T.OpenRecordset("OpStockLedger", dbOpenTable)
TBT.AddNew
    TBT.Fields("Item").Value = ItemCOde
    TBT.Fields("ItmName").Value = Blm1.Item1(CLng(ItemCOde))
    TBT.Fields("Quantity").Value = OpStock
    TBT.Fields("V_date").Value = S_Date - 1
TBT.Update
TBT.Close
Db_T.Close


db_m.Close

End Sub
Public Sub CreateStock(ByVal E_Date As Date)
Dim db_m As Database
Dim Db_T As Database
Dim ssql As String

Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")

ssql = "Drop Table Items"
Db_T.Execute ssql

ssql = "Drop Table Groups"
Db_T.Execute ssql

ssql = "Drop table PSTK"
Db_T.Execute ssql

ssql = "Drop table SSTK"
Db_T.Execute ssql

ssql = "Drop table RawSTK"
Db_T.Execute ssql

Db_T.Close

ssql = "Select * into Items in '" & App.Path & "\Book.mdb' from Item"
db_m.Execute ssql

ssql = "Select * into Groups in '" & App.Path & "\Book.mdb' from Groups"
db_m.Execute ssql

ssql = "select Item,Sum(Quantity) as QSTOCK,Sum(Quantity*rate) as Amt into PSTK in '" & App.Path & "\Book.mdb" & "' From In_DTl Where v_date <=#" & E_Date & "# Group By Item"
'MsgBox ssql
db_m.Execute ssql

ssql = "select Item,Sum(Qty) as QSTOCK,Sum(Qty*rate) as Amt into SSTK in '" & App.Path & "\Book.mdb" & "' From Sale_2 Where Inv_date <=#" & E_Date & "# Group By Item"
db_m.Execute ssql

ssql = "select RawCode,Sum(QtyUsed)as QSTOCK into RawSTK in '" & App.Path & "\Book.mdb" & "' From Rawsales Where Inv_date <=#" & E_Date & "# Group By RawCode"
db_m.Execute ssql


db_m.Close
End Sub

Public Sub CreateStockBetweenDates(ByVal S_Date As Date, ByVal E_Date As Date)
Dim db_m As Database
Dim Db_T As Database
Dim ssql As String

Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")

ssql = "Drop table PSTK"
Db_T.Execute ssql

ssql = "Drop table SSTK"
Db_T.Execute ssql

ssql = "Drop table RawSTK"
Db_T.Execute ssql

ssql = "Drop table CPSTK"
Db_T.Execute ssql

ssql = "Drop table CSSTK"
Db_T.Execute ssql

ssql = "Drop table CRawSTK"
Db_T.Execute ssql

Db_T.Close

'Opening Stocks
ssql = "select Item,Sum(Quantity) as QSTOCK,Sum(Quantity*rate) as Amt into PSTK in '" & App.Path & "\Book.mdb" & "' From In_DTl Where v_date <#" & S_Date & "# Group By Item"
'MsgBox ssql
db_m.Execute ssql

ssql = "select Item,Sum(Qty) as QSTOCK,Sum(Qty*rate) as Amt into SSTK in '" & App.Path & "\Book.mdb" & "' From Sale_2 Where Inv_date <#" & S_Date & "# Group By Item"
db_m.Execute ssql

ssql = "select RawCode,Sum(QtyUsed)as QSTOCK into RawSTK in '" & App.Path & "\Book.mdb" & "' From Rawsales Where Inv_date <#" & S_Date & "# Group By RawCode"
db_m.Execute ssql

'Stocks Between Dates

ssql = "select Item,Sum(Quantity) as QSTOCK,Sum(Quantity*rate) as Amt into CPSTK in '" & App.Path & "\Book.mdb" & "' From In_DTl Where v_date  Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
'MsgBox ssql
db_m.Execute ssql

ssql = "select Item,Sum(Qty) as QSTOCK,Sum(Qty*rate) as Amt into CSSTK in '" & App.Path & "\Book.mdb" & "' From Sale_2 Where Inv_date Between #" & S_Date & "# and #" & E_Date & "# Group By Item"
db_m.Execute ssql

ssql = "select RawCode,Sum(QtyUsed)as QSTOCK into CRawSTK in '" & App.Path & "\Book.mdb" & "' From Rawsales Where Inv_date between #" & S_Date & "# and #" & E_Date & "# Group By RawCode"
'MsgBox ssql
db_m.Execute ssql


db_m.Close
End Sub

Public Sub LedgerBetweenDates(ByVal S_Date As Date, ByVal E_Date As Date, ByVal Accode As Long)
Dim db_m As Database
Dim Db_T As Database
Dim ssql As String
Dim TBT As Recordset
Dim Tbm As Recordset
Dim Acname As String
Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")

ssql = "Drop table Ledger"
Db_T.Execute ssql
Acname = Blm1.party1(Accode)

Db_T.Close

'Ledger Between Dates
ssql = "Select *,'" & Acname & "' as AcName into Ledger in '" & App.Path & "\Book.mdb" & "' from Voucher where v_date Between #" & S_Date & "# and #" & E_Date & "# and Party = " & Accode & " Order By V_date"
db_m.Execute ssql


Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
ssql = "Select Sum(Debit - Credit) as Bal from Voucher where Party = " & Accode & " and v_date < #" & S_Date & "#"
Set Tbm = db_m.OpenRecordset(ssql)
Set TBT = Db_T.OpenRecordset("Ledger", dbOpenTable)
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
Db_T.Close
Tbm.Close
db_m.Close
End Sub
Public Sub DayBook(ByVal V_Date As Date)
Dim db_m As Database
Dim Db_T As Database
Dim ssql As String
Dim TBT As Recordset
Dim Tbm As Recordset
Dim Acname As String
Dim PCashBF As Currency, CashBF As Currency
Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")

ssql = "Drop table Voucher"
Db_T.Execute ssql
Db_T.Close
Blm1.Cash V_Date, PCashBF, CashBF
'MsgBox CashBF
'Ledger Between Dates
ssql = "Select a.*,b.Name as PartyName," & CashBF & " as CashBF into Voucher in '" & App.Path & "\Book.mdb" & "' from Voucher a,Parties b where a.Party=b.Code and a.v_date = #" & V_Date & "# Order By V_date"
db_m.Execute ssql




db_m.Close
End Sub

Public Sub CaseLedger(SerialNo As Double)
Dim db_m As Database
Dim Db_T As Database
Dim ssql As String
Dim TBT As Recordset
Dim Tbm As Recordset
Dim Acname As String
Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")

ssql = "Drop table CaseLedger"
Db_T.Execute ssql
Db_T.Close
Dim BillNo As Double, BillDate As String, PartyName As String, Goods As String, Pacakges As String
ssql = "Select a.*,b.BillNo,b.BillDate from Docs a Left Join Invoices b on a.SrNo=b.SerialNo where a.SrNo=" & SerialNo
Set TBT = db_m.OpenRecordset(ssql)
If Not TBT.EOF Then
    BillNo = TBT.Fields("BillNo").Value & ""
    If Not IsNull(TBT.Fields("BillDate").Value) Then BillDate = Format(TBT.Fields("BillDate").Value, "dd-MMM-yyyy")
    PartyName = Blm1.party1(TBT.Fields("PartyCode").Value)
    Goods = TBT.Fields("Goods").Value & ""
    Packages = TBT.Fields("Packages").Value & ""
End If
TBT.Close
'Case Ledger
ssql = "Select *," & BillNo & " as BiLLNo,'" & BillDate & "' as BiLLDate,'" & PartyName & "' as PartyName,'" & Goods & "' as Goods,'" & Packages & "' as Packages into CaseLedger in '" & App.Path & "\Book.mdb" & "' from Voucher where SerialNo=" & SerialNo & " Order By V_date"
db_m.Execute ssql


db_m.Close
End Sub

Public Sub TrialBalance(ByVal E_Date As Date)
Dim db_m As Database
Dim Db_T As Database
Dim ssql As String
Dim OPCash As Currency, CashBF As Currency
Set Db_T = OpenDatabase(App.Path & "\Book.mdb")
Set db_m = OpenDatabase(App.Path & "\Bloom.mdb")

ssql = "Drop table Trial"
Db_T.Execute ssql

ssql = "select a.Party,b.name as AcName,Sum(Debit - Credit) as Bal into Trial in '" & App.Path & "\Book.Mdb' from Voucher a,Parties b where a.Party = b.code and a.v_date <= #" & E_Date & "# Group by a.Party,b.name"
db_m.Execute ssql
db_m.Close
Blm1.Cash E_Date, OPCash, CashBF
ssql = "Insert into Trial (Party,ACName,Bal)Values(10000,'Cash In Hand'," & CashBF & ")"
Db_T.Execute ssql

Db_T.Close
End Sub

