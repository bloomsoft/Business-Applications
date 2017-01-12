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
Public Function ITEMstocks(ByVal C As Integer, e_date As Date) As Currency
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim ssale As Currency
Dim spur As Currency
Dim ssaleret As Currency
Dim slessstock As Currency, purRet As Currency
Dim spack As Currency
Dim sopen As Currency

Set db = OpenDatabase(patHmain)
ssql = "Select Sum(Qty) as Q from Purchase where Item = " & C
ssql = ssql & " and v_date < #" & e_date & "#"
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Q").Value) Then
    spur = tb.Fields("Q").Value
End If
tb.Close

ssql = "Select Sum(Qty) as Q from Sales where Item = " & C
ssql = ssql & " and v_date < #" & e_date & "#"
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Q").Value) Then
    ssale = tb.Fields("Q").Value
End If
tb.Close

ssql = "Select Sum(Qty) as Q from Consume where Item = " & C
ssql = ssql & " and v_date < #" & e_date & "#"
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Q").Value) Then
    slessstock = tb.Fields("Q").Value
End If
tb.Close
db.Close
'MsgBox "Test"
ITEMstocks = (Opstocks(C) + (spur + ssaleret + spack)) - (ssale + slessstock + purRet + sopen)


End Function
Public Function CurrentBalance(AcCode As Long) As String
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim Amt As String
Set db = OpenDatabase(patHmain)

ssql = "select Sum(Debit - Credit) as Bal from Voudtl where Party = " & AcCode
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Bal").Value) Then
    If tb.Fields("Bal").Value > 0 Then
        Amt = Format(tb.Fields("Bal").Value, "#.00") & " DR"
    ElseIf tb.Fields("Bal").Value < 0 Then
        Amt = Format(tb.Fields("Bal").Value, "#.00") & " CR"
    End If
Else
    Amt = "..."
End If
tb.Close
db.Close
CurrentBalance = Amt
End Function

Public Function ClosingStock(ByVal C As Integer) As Double
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim PQ As Double
Dim SQ As Double
Dim OQ As Double
OQ = Opstocks(C)
Set db = OpenDatabase(patHmain)
ssql = "Select Sum(Qty) as Q from Purchase where Item=" & C
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Q").Value) Then
    PQ = tb.Fields("Q").Value
Else
    PQ = 0
End If
tb.Close
ssql = "Select Sum(Qty) as Q from Sales where Item=" & C
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Q").Value) Then
    SQ = tb.Fields("Q").Value
Else
    SQ = 0
End If
tb.Close
db.Close
ClosingStock = (OQ + PQ) - SQ
End Function
Public Function Opstocks(ByVal C As Integer) As Currency
Dim db As Database
Dim tb As Recordset
Dim ssql As String

Set db = OpenDatabase(patHmain)
ssql = "Select * from items where Code = " & C
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    Opstocks = tb.Fields("Stock").Value
Else
    Opstocks = 0
End If
tb.Close
db.Close

End Function

Public Sub vouchercomb(cntl As Control)
cntl.clear
cntl.AddItem "Journal Voucher"
cntl.ItemData(cntl.NewIndex) = 1
cntl.AddItem "Bank Voucher"
cntl.ItemData(cntl.NewIndex) = 2
cntl.AddItem "Cash Voucher"
cntl.ItemData(cntl.NewIndex) = 3
cntl.ListIndex = 0
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
        
End Select
    
End Function

Public Sub fill_comb(ssql As String, cntl As Control, itemcol As String, datacol As String)
    Dim db As Database
    Dim tb As Recordset
    cntl.clear
    Set db = OpenDatabase(patHmain)
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            cntl.AddItem tb.Fields(itemcol).Value
            cntl.ItemData(cntl.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        cntl.ListIndex = 0
    End If
    tb.Close
    db.Close
        
End Sub
Public Function cash1(v_date As Date) As Currency
    Dim db As Database
    Dim tb As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim ssql As String
    Dim m_date As Date
    Dim op_cash As Currency
    
    Set db = OpenDatabase(patHmain)
    ssql = "select sum(debit - credit) as bal from voucher where party = 10000 "
    Set tb = db.OpenRecordset(ssql)
    If IsNull(tb.Fields("bal").Value) = False Then
        op_cash = tb.Fields("bal").Value
    End If
    tb.Close
    MsgBox op_cash
    ssql = "Select max(v_Date) as dat1 from voucher where v_date < #" & v_date & "#"
    
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("dat1").Value) Then
    
        m_date = tb.Fields("dat1").Value
      MsgBox m_date
    Else
        cash1 = op_cash
        Exit Function
        'm_date = v_date
    End If
    tb.Close
    
    ssql = "select pre_cash,sum(debit) as deb,sum(credit) as cred from voucher where e_type = 5 and v_date = #" & m_date & "# group by pre_cash"
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then

        pre_1 = tb.Fields("pre_cash").Value
        deb1 = tb.Fields("deb").Value
        cred1 = tb.Fields("cred").Value
           MsgBox "pre" & " " & pre_1 & "cred" & " " & cred1 & "deb" & " " & deb1
    End If
    tb.Close
    db.Close
    
    cash1 = ((cred1 + pre_1) - deb1) + op_cash
    MsgBox cash1
End Function

Public Function cash2(v_date As Date) As Currency
    Dim db As Database
    Dim tb As Recordset
    Dim tb2 As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim ssql As String
    Dim m_date As Date
    Dim op_cash As Currency
    Dim op As Integer
    op = 0
    
    Set db = OpenDatabase(patHmain)
    ssql = "select sum(debit - credit) as bal from voucher where party = 10000 "
    Set tb = db.OpenRecordset(ssql)
    If IsNull(tb.Fields("bal").Value) = False Then
    
        op_cash = tb.Fields("bal").Value
    End If
    tb.Close
    
    
    MsgBox op_cash
    
    ssql = "select * from voucher where e_type=6"
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
    MsgBox tb.Fields("V_date").Value
        If v_date = tb.Fields("v_date").Value Then
            op = 1
        End If
    End If
    tb.Close
    
    ssql = "select count(*) as cnt from voucher where v_Date = #" & v_date & "#"
    Set tb2 = db.OpenRecordset(ssql)
    If tb2.Fields("cnt").Value <> 0 Then
        m_date = v_date
    Else
    
    ssql = "Select max(v_Date) as dat1 from voucher where v_date < #" & v_date & "#"

    Set tb = db.OpenRecordset(ssql)
    If IsNull(tb.Fields("dat1").Value) = False Then
    
        m_date = tb.Fields("dat1").Value
 '   MsgBox m_date
    Else
        m_date = v_date
    End If
    tb.Close
    End If
'    MsgBox m_date
    tb2.Close
    
    
    ssql = "select pre_cash,sum(debit) as deb,sum(credit) as cred from voucher where e_type = 5 and v_date = #" & m_date & "# group by pre_cash"
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then

        pre_1 = tb.Fields("pre_cash").Value
        deb1 = tb.Fields("deb").Value
        cred1 = tb.Fields("cred").Value
'            MsgBox "pre" & " " & pre_1 & "cred" & " " & cred1 & "deb" & " " & deb1
    End If
    tb.Close
    db.Close
   MsgBox (cred1 + pre_1) - deb1
   If op = 1 Then
    cash2 = (((cred1 + pre_1) - deb1)) + op_cash
   Else
    cash2 = (((cred1 + pre_1) - deb1))
   End If
End Function
Public Property Get patHmain() As String
'patHmain = "C:\sOFT\Accounts\BLOOM.MDB"
patHmain = App.path & "\Bloom.mdb"

End Property
Public Property Get SettingsPath() As String
'patHmain = "C:\sOFT\Accounts\BLOOM.MDB"
SettingsPath = App.path & "\Settings.mdb"

End Property
Public Function orgname() As String
Dim db As Database
Dim tb As Recordset

Set db = OpenDatabase(App.path & "\Settings.mdb")
Set tb = db.OpenRecordset("Options", dbOpenTable)
If Not tb.EOF Then
    orgname = tb.Fields("OrgInfo").Value & ""
End If
tb.Close
db.Close
End Function
Public Function CashAcCode() As Long
Dim db As Database
Dim tb As Recordset

Set db = OpenDatabase(App.path & "\Settings.mdb")
Set tb = db.OpenRecordset("Options", dbOpenTable)
If Not tb.EOF Then
    CashAcCode = tb.Fields("CashAc").Value & ""
End If
tb.Close
db.Close
End Function

Public Function orgAddress() As String
Dim db As Database
Dim tb As Recordset

Set db = OpenDatabase(App.path & "\Settings.mdb")
Set tb = db.OpenRecordset("Options", dbOpenTable)
If Not tb.EOF Then
    orgAddress = tb.Fields("Address").Value & ""
End If
tb.Close
db.Close
End Function
Public Function orgGSTNo() As String
Dim db As Database
Dim tb As Recordset

Set db = OpenDatabase(App.path & "\Settings.mdb")
Set tb = db.OpenRecordset("Options", dbOpenTable)
If Not tb.EOF Then
    orgGSTNo = tb.Fields("GSTNo").Value & ""
End If
tb.Close
db.Close
End Function
Public Function orgPhone() As String
Dim db As Database
Dim tb As Recordset

Set db = OpenDatabase(App.path & "\Settings.mdb")
Set tb = db.OpenRecordset("Options", dbOpenTable)
If Not tb.EOF Then
    orgPhone = tb.Fields("Phone").Value & ""
End If
tb.Close
db.Close
End Function

Public Function item1(code As Integer) As String
  'On Error Resume Next
    Dim db As Database
    Dim tb As Recordset
    Set db = OpenDatabase(patHmain)
    Dim ssql As String
    ssql = "select * from items WHERE CODE = " & code
        Set tb = db.OpenRecordset(ssql)
             If tb.EOF = False Then
                     item1 = tb.Fields("name").Value
                     
            Else
                    item1 = "NOT"
            End If

      tb.Close
    db.Close
End Function
Public Function party1(code As Long) As String
  'On Error Resume Next
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
        Set db = OpenDatabase(patHmain)
        ssql = "select * from ACCHART WHERE CODE = " & code
        Set tb = db.OpenRecordset(ssql)
            If tb.EOF = False Then
                    party1 = tb.Fields("name").Value
            Else
                    party1 = "NOT"
            End If
      tb.Close
    db.Close
End Function
Public Function SalesTaxNo(code As Long) As String
  'On Error Resume Next
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
        Set db = OpenDatabase(patHmain)
        ssql = "select * from ACCHART WHERE CODE = " & code
        Set tb = db.OpenRecordset(ssql)
            If tb.EOF = False Then
                    SalesTaxNo = tb.Fields("STNo").Value & ""
            Else
                    SalesTaxNo = "===="
            End If
      tb.Close
    db.Close
End Function
Public Function heads(code As Long) As String
  'On Error Resume Next
    Dim ssql As String
    Dim db As Database
    Dim tb As Recordset
        Set db = OpenDatabase(patHmain)
        ssql = "select * from heads WHERE CODE = " & code
        Set tb = db.OpenRecordset(ssql)
            If tb.EOF = False Then
                    heads = tb.Fields("name").Value
            Else
                    heads = "NOT"
            End If
      tb.Close
    db.Close
End Function

