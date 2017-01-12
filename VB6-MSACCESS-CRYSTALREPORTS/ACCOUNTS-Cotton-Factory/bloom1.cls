VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bloom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_pathmain As String
Public Function GetEmpSalRate(EmpCode As Double) As Double
Dim DB As Database
Dim RST As Recordset
Dim Ssql As String

Set DB = OpenDatabase(patHmain)
Ssql = "Select * from Acchart where Code = " & EmpCode
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    GetEmpSalRate = RST.Fields("SalRate").Value & ""
End If
RST.Close
DB.Close
End Function
Public Function GetEmpAdvance(EmpCode As Double, SDate As Date, EDate As Date) As Double
Dim DB As Database
Dim RST As Recordset
Dim Ssql As String
Dim V As Double
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Debit)-Sum(Credit) as Bal from VouDTL where Party = " & EmpCode
Ssql = Ssql & " and V_Date <= #" & EDate - 1 & "#"
'Ssql = Ssql & " and V_type=11"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Bal").Value) Then
    V = RST.Fields("Bal").Value
Else
    V = 0
End If
RST.Close
DB.Close
GetEmpAdvance = V
End Function

Public Function GetEmpOverTime(EmpCode As Double, SDate As Date, EDate As Date) As Double
Dim DB As Database
Dim RST As Recordset
Dim Ssql As String
Dim V As Double
Set DB = OpenDatabase(patHmain)
Ssql = "Select * from OverTime where Accode = " & EmpCode
Ssql = Ssql & " and ADate Between #" & SDate & "# and #" & EDate & "#"

Set RST = DB.OpenRecordset(Ssql)
'MsgBox "Test"
If Not RST.EOF Then
    Do While Not RST.EOF
        V = V + RST.Fields("OHours").Value
        RST.MoveNext
    Loop
End If
RST.Close
DB.Close
GetEmpOverTime = V
End Function
Public Function GetEmpAdvDeduction(EmpCode As Double, SDate As Date, EDate As Date) As Double

Dim DB As Database
Dim RST As Recordset
Dim Ssql As String
Dim V As Double
Set DB = OpenDatabase(patHmain)
Ssql = "Select * from OverTime where Accode = " & EmpCode
Ssql = Ssql & " and ADate Between #" & SDate & "# and #" & EDate & "#"

Set RST = DB.OpenRecordset(Ssql)
'MsgBox "Test"
If Not RST.EOF Then
    Do While Not RST.EOF
        V = V + Val(RST.Fields("AdvDeduction").Value & "")
        RST.MoveNext
    Loop
End If
RST.Close
DB.Close
GetEmpAdvDeduction = V
End Function

Public Function GetEmpShorTime(EmpCode As Double, SDate As Date, EDate As Date) As Double
Dim DB As Database
Dim RST As Recordset
Dim Ssql As String
Dim V As Double
Set DB = OpenDatabase(patHmain)
Ssql = "Select * from OverTime where Accode = " & EmpCode
Ssql = Ssql & " and ADate Between #" & SDate & "# and #" & EDate & "#"

Set RST = DB.OpenRecordset(Ssql)
'MsgBox "Test"
If Not RST.EOF Then
    Do While Not RST.EOF
        V = V + RST.Fields("SHours").Value
        RST.MoveNext
    Loop
End If
RST.Close
DB.Close
GetEmpShorTime = V
End Function

Public Function ITEMstocks(ByVal C As Double, E_Date As Date) As Currency
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim ssale As Currency
Dim spur As Currency
Dim ssaleret As Currency
Dim slessstock As Currency, purRet As Currency
Dim spack As Currency
Dim sopen As Currency

Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Qty) as Q from Purchase where Item = " & C
Ssql = Ssql & " and v_date < #" & E_Date & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    spur = TB.Fields("Q").Value
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Sales where Item = " & C
Ssql = Ssql & " and v_date < #" & E_Date & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    ssale = TB.Fields("Q").Value
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Consume where Item = " & C
Ssql = Ssql & " and v_date < #" & E_Date & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    slessstock = TB.Fields("Q").Value
End If
TB.Close
DB.Close
'MsgBox "Test"
ITEMstocks = (Opstocks(C) + (spur + ssaleret + spack)) - (ssale + slessstock + purRet + sopen)


End Function
Public Function CurrentBalance(AcCode As Long) As String
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim AMT As String
Set DB = OpenDatabase(patHmain)

Ssql = "select Sum(Debit - Credit) as Bal from Voudtl where Party = " & AcCode
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Bal").Value) Then
    If TB.Fields("Bal").Value > 0 Then
        AMT = Format(TB.Fields("Bal").Value, "#.00") & " DR"
    ElseIf TB.Fields("Bal").Value < 0 Then
        AMT = Format(TB.Fields("Bal").Value, "#.00") & " CR"
    End If
Else
    AMT = "..."
End If
TB.Close
DB.Close
CurrentBalance = AMT
End Function

Public Function ClosingStock(ByVal C As Double) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim PQ As Double
Dim SQ As Double
Dim OQ As Double
Dim IQ As Double
Dim DQ As Double
OQ = Opstocks(C)
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Qty) as Q from Purchase where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PQ = TB.Fields("Q").Value
Else
    PQ = 0
End If
TB.Close
Ssql = "Select Sum(Qty) as Q from PurchaseReturn where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PRQ = TB.Fields("Q").Value
Else
    PRQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Sales where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SQ = TB.Fields("Q").Value
Else
    SQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from SalesReturn where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SRQ = TB.Fields("Q").Value
Else
    SRQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Issue where DrCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IPQ = TB.Fields("Q").Value
Else
    IPQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Issue where ItemCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IQ = TB.Fields("Q").Value
Else
    IQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Production where ItemCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    DQ = TB.Fields("Q").Value
Else
    DQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Production where CrCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PrCQ = TB.Fields("Q").Value
Else
    PrCQ = 0
End If
TB.Close


Ssql = "Select Sum(Qty) as Q from SHProduction where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SHPQ = TB.Fields("Q").Value
Else
    SHPQ = 0
End If
TB.Close

DB.Close

ClosingStock = (OQ + PQ + DQ + SRQ + IPQ + SHPQ) - (SQ + IQ + PRQ + PrCQ)
End Function
Public Function ClosingStockSH(ByVal C As Double) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim PQ As Double
Dim SQ As Double, SRQ As Double
Dim OQ As Double
Dim IQ As Double
Dim DQ As Double
OQ = OpstocksSH(C)
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Qty) as Q from Purchase where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PQ = TB.Fields("Q").Value
Else
    PQ = 0
End If
TB.Close
Ssql = "Select Sum(Qty) as Q from PurchaseReturn where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PRQ = TB.Fields("Q").Value
Else
    PRQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Sales where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SQ = TB.Fields("Q").Value
Else
    SQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from SalesReturn where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SRQ = TB.Fields("Q").Value
Else
    SRQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Issue where Mid(DrCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IPQ = TB.Fields("Q").Value
Else
    IPQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Issue where Mid(ItemCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IQ = TB.Fields("Q").Value
Else
    IQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Production where Mid(ItemCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    DQ = TB.Fields("Q").Value
Else
    DQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Production where Mid(CrCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PrCQ = TB.Fields("Q").Value
Else
    PrCQ = 0
End If
TB.Close


Ssql = "Select Sum(Qty) as Q from SHProduction where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SHPQ = TB.Fields("Q").Value
Else
    SHPQ = 0
End If
TB.Close

DB.Close
ClosingStockSH = (OQ + PQ + DQ + SRQ + IPQ + SHPQ) - (SQ + IQ + PRQ + PrCQ)
End Function

Public Function ClosingStockEndDate(ByVal C As Double, EndDate As Date) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim PQ As Double
Dim SQ As Double
Dim OQ As Double
Dim IQ As Double
Dim DQ As Double
OQ = Opstocks(C)
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Qty) as Q from Purchase where Item=" & C & " and V_date <= #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PQ = TB.Fields("Q").Value
Else
    PQ = 0
End If
TB.Close
Ssql = "Select Sum(Qty) as Q from Sales where Item=" & C & " and V_Date = #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SQ = TB.Fields("Q").Value
Else
    SQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Issue where ItemCode=" & C & " and V_Date = #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IQ = TB.Fields("Q").Value
Else
    IQ = 0
End If
TB.Close

Ssql = "Select Sum(Qty) as Q from Production where ItemCode=" & C & " and V_Date = #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    DQ = TB.Fields("Q").Value
Else
    DQ = 0
End If
TB.Close
DB.Close

ClosingStockEndDate = (OQ + PQ + DQ) - (SQ + IQ)
End Function
Public Function ClosingStockAmountEndDate(ByVal C As Double, EndDate As Date) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim PQ As Double
Dim SQ As Double
Dim OQ As Double
Dim IQ As Double
Dim DQ As Double
OQ = Opstocks(C)
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Amount) as Q from Purchase where Item=" & C & " and V_date <= #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PQ = TB.Fields("Q").Value
Else
    PQ = 0
End If
TB.Close
Ssql = "Select Sum(Qty) as Q from Sales where Item=" & C & " and V_Date = #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SQ = TB.Fields("Amount").Value
Else
    SQ = 0
End If
TB.Close

Ssql = "Select Sum(Amount) as Q from Issue where ItemCode=" & C & " and V_Date = #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IQ = TB.Fields("Q").Value
Else
    IQ = 0
End If
TB.Close

Ssql = "Select Sum(Rate * Qty) as Q from Production where ItemCode=" & C & " and V_Date = #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    DQ = TB.Fields("Q").Value
Else
    DQ = 0
End If
TB.Close
DB.Close

ClosingStockAmountEndDate = (OQ + PQ + DQ) - (SQ + IQ)
End Function

Public Function ClosingStockBales(ByVal C As Double) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim PQ As Double
Dim PRQ As Double
Dim SQ As Double
Dim OQ As Double
Dim IQ As Double
Dim DQ As Double
Dim SRQ As Double
Dim IPQ As Double
Dim SHPQ As Double
Dim PrCQ As Double
OQ = opStockbales(C)
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Bales) as Q from Purchase where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PQ = TB.Fields("Q").Value
Else
    PQ = 0
End If
TB.Close
Ssql = "Select Sum(Bales) as Q from PurchaseReturn where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PRQ = TB.Fields("Q").Value
Else
    PRQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Sales where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SQ = TB.Fields("Q").Value
Else
    SQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from SalesReturn where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SRQ = TB.Fields("Q").Value
Else
    SRQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Issue where DrCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IPQ = TB.Fields("Q").Value
Else
    IPQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Issue where ItemCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IQ = TB.Fields("Q").Value
Else
    IQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Production where ItemCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    DQ = TB.Fields("Q").Value
Else
    DQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Production where CrCode=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PrCQ = TB.Fields("Q").Value
Else
    PrCQ = 0
End If
TB.Close


Ssql = "Select Sum(Bales) as Q from SHProduction where Item=" & C
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SHPQ = TB.Fields("Q").Value
Else
    SHPQ = 0
End If
TB.Close

DB.Close

ClosingStockBales = (OQ + PQ + DQ + SRQ + IPQ + SHPQ) - (SQ + IQ + PRQ + PrCQ)
End Function
Public Function ClosingStockBalesSH(ByVal C As Double) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim PQ As Double
Dim SQ As Double, SRQ As Double
Dim OQ As Double
Dim IQ As Double
Dim DQ As Double
OQ = OpstocksBalesSH(C)
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Bales) as Q from Purchase where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
'Clipboard.SetText Ssql
'MsgBox Ssql
If Not IsNull(TB.Fields("Q").Value) Then
    PQ = TB.Fields("Q").Value
Else
    PQ = 0
End If
TB.Close
Ssql = "Select Sum(Bales) as Q from PurchaseReturn where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PRQ = TB.Fields("Q").Value
Else
    PRQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Sales where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SQ = TB.Fields("Q").Value
Else
    SQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from SalesReturn where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SRQ = TB.Fields("Q").Value
Else
    SRQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Issue where Mid(DrCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IPQ = TB.Fields("Q").Value
Else
    IPQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Issue where Mid(ItemCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IQ = TB.Fields("Q").Value
Else
    IQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Production where Mid(ItemCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    DQ = TB.Fields("Q").Value
Else
    DQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Production where Mid(CrCode,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PrCQ = TB.Fields("Q").Value
Else
    PrCQ = 0
End If
TB.Close


Ssql = "Select Sum(Bales) as Q from SHProduction where Mid(Item,1,5)='" & C & "'"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SHPQ = TB.Fields("Q").Value
Else
    SHPQ = 0
End If
TB.Close

DB.Close
'MsgBox OQ & " " & PQ + DQ + SRQ + IPQ + SHPQ & " " & SQ + IQ + PRQ + PrCQ
ClosingStockBalesSH = (OQ + PQ + DQ + SRQ + IPQ + SHPQ) - (SQ + IQ + PRQ + PrCQ)
End Function

Public Function ClosingStockBalesEndDate(ByVal C As Double, EndDate As Date) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String
Dim PQ As Double
Dim SQ As Double
Dim OQ As Double
Dim IQ As Double
Dim DQ As Double
OQ = Opstocks(C)
Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(Bales) as Q from Purchase where Item=" & C & " and V_date <= #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    PQ = TB.Fields("Q").Value
Else
    PQ = 0
End If
TB.Close
Ssql = "Select Sum(Bales) as Q from Sales where Item=" & C & " and V_date <= #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    SQ = TB.Fields("Q").Value
Else
    SQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Issue where ItemCode=" & C & " and V_date <= #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    IQ = TB.Fields("Q").Value
Else
    IQ = 0
End If
TB.Close

Ssql = "Select Sum(Bales) as Q from Production where ItemCode=" & C & " and V_date <= #" & EndDate & "#"
Set TB = DB.OpenRecordset(Ssql)
If Not IsNull(TB.Fields("Q").Value) Then
    DQ = TB.Fields("Q").Value
Else
    DQ = 0
End If
TB.Close
DB.Close

ClosingStockBalesEndDate = (OQ + PQ + DQ) - (SQ + IQ)
End Function

Public Function Opstocks(ByVal C As Double) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(patHmain)
Ssql = "Select * from items where Code = " & C
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
'    MsgBox "Test"
    Opstocks = TB.Fields("OpWT").Value
Else
    Opstocks = 0
End If
TB.Close
DB.Close

End Function

Public Function opStockbales(ByVal C As Double) As Double
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(patHmain)
Ssql = "Select * from items where Code = " & C
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    opStockbales = TB.Fields("OpBales").Value
Else
    opStockbales = 0
End If
TB.Close
DB.Close

End Function

Public Function OpstocksSH(ByVal C As Double) As Currency
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(opWT) as W from items where Mid(Code,1,5) = '" & C & "'"
Set TB = DB.OpenRecordset(Ssql)

If Not IsNull(TB.Fields("W")) Then
    OpstocksSH = TB.Fields("W").Value
Else
    OpstocksSH = 0
End If
TB.Close
DB.Close

End Function
Public Function OpstocksBalesSH(ByVal C As Double) As Currency
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(patHmain)
Ssql = "Select Sum(OpBales) as B from items where Mid(Code,1,5) = '" & C & "'"
Set TB = DB.OpenRecordset(Ssql)

If Not IsNull(TB.Fields("B").Value) Then
    OpstocksBalesSH = TB.Fields("B").Value
Else
    OpstocksBalesSH = 0
End If
TB.Close
DB.Close

End Function

Public Sub vouchercomb(CNTL As Control)
CNTL.clear
CNTL.AddItem "Journal Voucher"
CNTL.ItemData(CNTL.NewIndex) = 1
CNTL.AddItem "Bank Voucher"
CNTL.ItemData(CNTL.NewIndex) = 2
CNTL.AddItem "Cash Voucher"
CNTL.ItemData(CNTL.NewIndex) = 3
CNTL.ListIndex = 0
End Sub
Public Function VoucherTypesRet(C As Integer) As String
Select Case C
    Case 1
        VoucherTypesRet = "JV"
    Case 2
        VoucherTypesRet = "BV"
    Case 3
        VoucherTypesRet = "CV"
    Case 4
        VoucherTypesRet = "PJV"
    Case 5
        VoucherTypesRet = "SJV"
    Case 6
        VoucherTypesRet = "CJV"
    Case 10
        VoucherTypesRet = "OB"
    Case 11
        VoucherTypesRet = "ADV"
    Case 15
        VoucherTypesRet = "SL"
    Case 16
        VoucherTypesRet = "PRET"
    Case 17
        VoucherTypesRet = "SRET"
    Case 18
        VoucherTypesRet = "STKT"
    Case 21
        VoucherTypesRet = "PRD"
    Case 20
        VoucherTypesRet = "ISS"
        
End Select
    
End Function

Public Sub fill_comb(Ssql As String, CNTL As Control, itemcol As String, datacol As String)
    Dim DB As Database
    Dim TB As Recordset
    CNTL.clear
    Set DB = OpenDatabase(patHmain)
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then
        Do While Not TB.EOF
            CNTL.AddItem TB.Fields(itemcol).Value
            CNTL.ItemData(CNTL.NewIndex) = TB.Fields(datacol).Value
            TB.MoveNext
        Loop
        CNTL.ListIndex = 0
    End If
    TB.Close
    DB.Close
        
End Sub
Public Sub fill_comb2(Ssql As String, CNTL As Control, itemcol As String, datacol As String)
    Dim DB As Database
    Dim TB As Recordset
    CNTL.clear
    Set DB = OpenDatabase(patHmain)
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then
        Do While Not TB.EOF
            CNTL.AddItem TB.Fields(datacol).Value & " - " & TB.Fields(itemcol).Value
            CNTL.ItemData(CNTL.NewIndex) = TB.Fields(datacol).Value
            TB.MoveNext
        Loop
        CNTL.ListIndex = 0
    End If
    TB.Close
    DB.Close
        
End Sub

Public Function cash1(V_Date As Date) As Currency
    Dim DB As Database
    Dim TB As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim Ssql As String
    Dim m_date As Date
    Dim op_cash As Currency
    
    Set DB = OpenDatabase(patHmain)
    Ssql = "select sum(debit - credit) as bal from voucher where party = 10000 "
    Set TB = DB.OpenRecordset(Ssql)
    If IsNull(TB.Fields("bal").Value) = False Then
        op_cash = TB.Fields("bal").Value
    End If
    TB.Close
    MsgBox op_cash
    Ssql = "Select max(v_Date) as dat1 from voucher where v_date < #" & V_Date & "#"
    
    Set TB = DB.OpenRecordset(Ssql)
    If Not IsNull(TB.Fields("dat1").Value) Then
    
        m_date = TB.Fields("dat1").Value
      MsgBox m_date
    Else
        cash1 = op_cash
        Exit Function
        'm_date = v_date
    End If
    TB.Close
    
    Ssql = "select pre_cash,sum(debit) as deb,sum(credit) as cred from voucher where e_type = 5 and v_date = #" & m_date & "# group by pre_cash"
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then

        pre_1 = TB.Fields("pre_cash").Value
        deb1 = TB.Fields("deb").Value
        cred1 = TB.Fields("cred").Value
           MsgBox "pre" & " " & pre_1 & "cred" & " " & cred1 & "deb" & " " & deb1
    End If
    TB.Close
    DB.Close
    
    cash1 = ((cred1 + pre_1) - deb1) + op_cash
    MsgBox cash1
End Function

Public Function cash2(V_Date As Date) As Currency
    Dim DB As Database
    Dim TB As Recordset
    Dim Tb2 As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim Ssql As String
    Dim m_date As Date
    Dim op_cash As Currency
    Dim op As Integer
    op = 0
    
    Set DB = OpenDatabase(patHmain)
    Ssql = "select sum(debit - credit) as bal from voucher where party = 10000 "
    Set TB = DB.OpenRecordset(Ssql)
    If IsNull(TB.Fields("bal").Value) = False Then
    
        op_cash = TB.Fields("bal").Value
    End If
    TB.Close
    
    
    MsgBox op_cash
    
    Ssql = "select * from voucher where e_type=6"
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then
    MsgBox TB.Fields("V_date").Value
        If V_Date = TB.Fields("v_date").Value Then
            op = 1
        End If
    End If
    TB.Close
    
    Ssql = "select count(*) as cnt from voucher where v_Date = #" & V_Date & "#"
    Set Tb2 = DB.OpenRecordset(Ssql)
    If Tb2.Fields("cnt").Value <> 0 Then
        m_date = V_Date
    Else
    
    Ssql = "Select max(v_Date) as dat1 from voucher where v_date < #" & V_Date & "#"

    Set TB = DB.OpenRecordset(Ssql)
    If IsNull(TB.Fields("dat1").Value) = False Then
    
        m_date = TB.Fields("dat1").Value
 '   MsgBox m_date
    Else
        m_date = V_Date
    End If
    TB.Close
    End If
'    MsgBox m_date
    Tb2.Close
    
    
    Ssql = "select pre_cash,sum(debit) as deb,sum(credit) as cred from voucher where e_type = 5 and v_date = #" & m_date & "# group by pre_cash"
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then

        pre_1 = TB.Fields("pre_cash").Value
        deb1 = TB.Fields("deb").Value
        cred1 = TB.Fields("cred").Value
'            MsgBox "pre" & " " & pre_1 & "cred" & " " & cred1 & "deb" & " " & deb1
    End If
    TB.Close
    DB.Close
   MsgBox (cred1 + pre_1) - deb1
   If op = 1 Then
    cash2 = (((cred1 + pre_1) - deb1)) + op_cash
   Else
    cash2 = (((cred1 + pre_1) - deb1))
   End If
End Function
Public Property Get patHmain() As String
'patHmain = "C:\sOFT\Accounts\BLOOM.MDB"
patHmain = App.path & "\Years\" & YearN & "\Bloom.mdb"

End Property
Public Property Get SettingsPath() As String
'patHmain = "C:\sOFT\Accounts\BLOOM.MDB"
SettingsPath = App.path & "\Settings.mdb"

End Property
Public Function orgname() As String
Dim DB As Database
Dim TB As Recordset

Set DB = OpenDatabase(App.path & "\Settings.mdb")
Set TB = DB.OpenRecordset("Options", dbOpenTable)
If Not TB.EOF Then
    orgname = TB.Fields("OrgInfo").Value & ""
End If
TB.Close
DB.Close
End Function
Public Function CashAcCode() As Long
Dim DB As Database
Dim TB As Recordset

Set DB = OpenDatabase(App.path & "\Settings.mdb")
Set TB = DB.OpenRecordset("Options", dbOpenTable)
If Not TB.EOF Then
    CashAcCode = TB.Fields("CashAc").Value & ""
End If
TB.Close
DB.Close
End Function

Public Function orgAddress() As String
Dim DB As Database
Dim TB As Recordset

Set DB = OpenDatabase(App.path & "\Settings.mdb")
Set TB = DB.OpenRecordset("Options", dbOpenTable)
If Not TB.EOF Then
    orgAddress = TB.Fields("Address").Value & ""
End If
TB.Close
DB.Close
End Function
Public Function orgGSTNo() As String
Dim DB As Database
Dim TB As Recordset

Set DB = OpenDatabase(App.path & "\Settings.mdb")
Set TB = DB.OpenRecordset("Options", dbOpenTable)
If Not TB.EOF Then
    orgGSTNo = TB.Fields("GSTNo").Value & ""
End If
TB.Close
DB.Close
End Function
Public Function orgPhone() As String
Dim DB As Database
Dim TB As Recordset

Set DB = OpenDatabase(App.path & "\Settings.mdb")
Set TB = DB.OpenRecordset("Options", dbOpenTable)
If Not TB.EOF Then
    orgPhone = TB.Fields("Phone").Value & ""
End If
TB.Close
DB.Close
End Function

Public Function item1(code As Double) As String
  'On Error Resume Next
    Dim DB As Database
    Dim TB As Recordset
    Set DB = OpenDatabase(patHmain)
    Dim Ssql As String
    Ssql = "select * from items WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
             If TB.EOF = False Then
                     item1 = TB.Fields("name").Value
                     
            Else
                    item1 = "NOT"
            End If

      TB.Close
    DB.Close
End Function
Public Function SubGroupName(code As Double) As String
  'On Error Resume Next
    Dim DB As Database
    Dim TB As Recordset
    Set DB = OpenDatabase(patHmain)
    Dim Ssql As String
    Ssql = "select * from SubGroups WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
             If TB.EOF = False Then
                     SubGroupName = TB.Fields("name").Value
                     
            Else
                     SubGroupName = "NOT"
            End If

      TB.Close
    DB.Close
End Function
Public Function SubHeadName(code As Double) As String
  'On Error Resume Next
    Dim DB As Database
    Dim TB As Recordset
    Set DB = OpenDatabase(patHmain)
    Dim Ssql As String
    Ssql = "select * from Heads WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
             If TB.EOF = False Then
                     SubHeadName = TB.Fields("name").Value
                     
            Else
                     SubHeadName = "NOT"
            End If

      TB.Close
    DB.Close
End Function

Public Function GroupName(code As Integer) As String
  'On Error Resume Next
    Dim DB As Database
    Dim TB As Recordset
    Set DB = OpenDatabase(patHmain)
    Dim Ssql As String
    Ssql = "select * from Groups WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
             If TB.EOF = False Then
                     GroupName = TB.Fields("name").Value
                     
            Else
                     GroupName = "NOT"
            End If

      TB.Close
    DB.Close
End Function

Public Function Mill1(code As Integer) As String
  'On Error Resume Next
    Dim DB As Database
    Dim TB As Recordset
    Set DB = OpenDatabase(patHmain)
    Dim Ssql As String
    Ssql = "select * from Mills WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
             If TB.EOF = False Then
                    Mill1 = TB.Fields("name").Value
                     
            Else
                    Mill1 = "NOT"
            End If

      TB.Close
    DB.Close
End Function

Public Function party1(code As Long) As String
  'On Error Resume Next
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
        Set DB = OpenDatabase(patHmain)
        Ssql = "select * from ACCHART WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
            If TB.EOF = False Then
                    party1 = TB.Fields("name").Value
            Else
                    party1 = "NOT"
            End If
      TB.Close
    DB.Close
End Function
Public Function GetEmpStatus(code As Long) As Integer
  'On Error Resume Next
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
        Set DB = OpenDatabase(patHmain)
        Ssql = "select * from ACCHART WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
            If TB.EOF = False Then
                    GetEmpStatus = Val(TB.Fields("Status").Value & "")
            Else
                    GetEmpStatus = 0
            End If
      TB.Close
    DB.Close
End Function

Public Function SalesTaxNo(code As Long) As String
  'On Error Resume Next
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
        Set DB = OpenDatabase(patHmain)
        Ssql = "select * from ACCHART WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
            If TB.EOF = False Then
                    SalesTaxNo = TB.Fields("STNo").Value & ""
            Else
                    SalesTaxNo = "===="
            End If
      TB.Close
    DB.Close
End Function
Public Function heads(code As Long) As String
  'On Error Resume Next
    Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
        Set DB = OpenDatabase(patHmain)
        Ssql = "select * from heads WHERE CODE = " & code
        Set TB = DB.OpenRecordset(Ssql)
            If TB.EOF = False Then
                    heads = TB.Fields("name").Value
            Else
                    heads = "NOT"
            End If
      TB.Close
    DB.Close
End Function

