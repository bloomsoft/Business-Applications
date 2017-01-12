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
Private m_path As String
Dim ii() As Integer

Public Sub Cash(EntryDate As Date, Opening As Currency, Closing As Currency)
Dim db As Database
Dim tb As Recordset
Dim ssql As String

ssql = "select Opening,Closing from Pre_Cash where v_date = (Select Max(V_date) from Pre_cash where v_date < #" & EntryDate & "#)"
Set db = OpenDatabase(pathMain)
Set tb = db.OpenRecordset(ssql)
If tb.EOF Then
    Opening = 0
    Closing = 0
Else
    Opening = tb.Fields("Opening").Value
    Closing = tb.Fields("Closing").Value
End If
tb.Close
db.Close
End Sub

Public Sub LessStock(itm As Long, LQTY As Currency, E_Date As Date)
Dim db As Database
Dim ssql As String
Dim tb As Recordset
Dim cstk As Currency
'Set db = OpenDatabase(m_path)
'ssql = "select CStock from Item WHere Code = " & itm
'Set tb = db.OpenRecordset(ssql)
'If Not tb.EOF Then
'If Not IsNull(tb.Fields("CSTOCK").Value) Then
'    cstk = tb.Fields("CSTOCK").Value
    'MsgBox cstk
'End If
'End If
'tb.Close
'If cstk > 0 Then
'ssql = "update item set l_date = #" & e_date & "#, cstock = " & (cstk - LQTY) & " where Code = " & itm
'db.Execute ssql
'End If
'db.Close

End Sub
Public Sub UpdateStock(itm As Long, QTY As Currency, E_Date As Date)
Dim db As Database
Dim ssql As String
Dim tb As Recordset
Dim cstk As Currency
'Set db = OpenDatabase(m_path)
'ssql = "update item set l_date = #" & e_date & "#, cstock = " & QTY & " where Code = " & itm
'db.Execute ssql
'db.Close
End Sub
Public Function Opstocks(C As Long) As Currency
Dim db As Database
Dim tb As Recordset
Dim ssql As String

Set db = OpenDatabase(m_path)
ssql = "Select * from item where Code = " & C
Set tb = db.OpenRecordset(ssql)
If Not tb.EOF Then
    Opstocks = tb.Fields("Stock").Value
Else
    Opstocks = 0
End If
tb.Close
db.Close

End Function
Public Function ITEMstocks(ByVal C As Long, E_Date As Date) As Currency
Dim db As Database
Dim tb As Recordset
Dim ssql As String
Dim ssale As Currency
Dim spur As Currency
Dim ssaleret As Currency
Dim slessstock As Currency, purRet As Currency
Dim spack As Currency
Dim sopen As Currency

Set db = OpenDatabase(m_path)
ssql = "Select Sum(Quantity) as Q from In_DTL where Item = " & C
ssql = ssql & " and v_date <= #" & E_Date & "#"
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Q").Value) Then
    spur = tb.Fields("Q").Value
End If
tb.Close


ssql = "Select Sum(Qty) as Q from Sale_2 where Item = " & C
ssql = ssql & " and inv_date <= #" & E_Date & "#"
Set tb = db.OpenRecordset(ssql)
If Not IsNull(tb.Fields("Q").Value) Then
    ssale = tb.Fields("Q").Value
End If
tb.Close









db.Close

ITEMstocks = (Opstocks(C) + (spur + ssaleret + spack)) - (ssale + slessstock + purRet + sopen)


End Function


Public Function Reasons(C As Byte) As String
Select Case C
    Case 1
        Reasons = "Entertainment"
    Case 2
        Reasons = "Breakage"
    Case 3
        Reasons = "Expired"
    Case 4
        Reasons = "Shortage"
    Case 5
        Reasons = "Theft"
End Select
End Function

Public Function CTNCheck(CTN As Long) As Boolean
Dim db_m As Database
Dim tb As Recordset
Dim ssql As String
CTNCheck = False
Set db_m = OpenDatabase(m_path)
ssql = "select * from in_dtl where CTN_no = " & CTN
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
        CTNCheck = True
End If
tb.Close
ssql = "select * from rec_2 where CTN_no = " & CTN
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
        CTNCheck = True
End If
tb.Close
db_m.Close
End Function
Public Function OutStand_DyeStock(C As Long, itm As Long) As Currency
Dim db_m As Database
Dim tb As Recordset
Dim ss As Currency
Dim qs As Currency
Dim qr As Currency
Dim ssql As String
Set db_m = OpenDatabase(m_path)
ssql = "select sum(quantity)as q from sent_view_1 where party = " & C & " and item = " & itm
Set tb = db_m.OpenRecordset(ssql)
'MsgBox ssql
If Not IsNull(tb.Fields("q").Value) Then
'MsgBox qs
        qs = tb.Fields("q").Value
        ssql = "select sum(quantity) as q from rec_view_final where raw_item = " & itm
        ssql = ssql & " and party = " & C
        Set tb2 = db_m.OpenRecordset(ssql)
            If Not IsNull(tb2.Fields("q").Value) Then
                qr = tb2.Fields("q").Value
            End If
        tb2.Close
End If
'   MsgBox qr
        ss = qs - qr
        OutStand_DyeStock = ss
tb.Close
db_m.Close
End Function
Public Function Stock(C As Long, V_Date As Date) As Currency
    Dim db As Database
    Dim tb As Recordset
    Dim ssql As String
    Dim In1 As Currency
    Dim sale1 As Currency
    Dim Rec1 As Currency
    Dim Sent1 As Currency
    Dim fstock As Currency
    
    Set db = OpenDatabase(m_path)
    'Inward
    ssql = "select sum(quantity)as q from in_dtl where item = " & C
    ssql = ssql & " and v_Date <= #" & V_Date & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("q").Value) Then
        In1 = tb.Fields("q").Value
    End If
    tb.Close
    'Sale
    ssql = "select sum(quantity)as q from sale_2 where item = " & C
    ssql = ssql & " and v_Date <= #" & V_Date & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("q").Value) Then
        sale1 = tb.Fields("q").Value
    End If
    tb.Close
    'Sent
    ssql = "select sum(quantity)as q from sent_2 where item = " & C
    ssql = ssql & " and v_Date <= #" & V_Date & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("q").Value) Then
        Sent1 = tb.Fields("q").Value
    End If
    tb.Close
    'Rec
    ssql = "select sum(quantity)as q from rec_2 where item = " & C
    ssql = ssql & " and v_Date <= #" & V_Date & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("q").Value) Then
        Rec1 = tb.Fields("q").Value
    End If
    tb.Close
    'MsgBox In1
    'MsgBox Rec1
    'MsgBox Sent1
    'MsgBox sale1
    
    fstock = (In1 + Rec1) - (Sent1 + sale1)
    Stock = fstock
    db.Close
End Function
Public Function RPAD(s As String, length As Byte) As String
    Dim l As Integer
    Dim I As Integer
    Dim e As Integer
    
    l = Len(s)
    e = length - l
    s = Trim(s)
    For I = 1 To e
        s = s & "_"
    Next I
    RPAD = Mid(s, 1, length)
End Function
Public Property Get report_path() As String
report_path = "C:\soft\Hosiery\"
End Property
Public Function cash1(V_Date As Date) As Currency
    Dim db As Database
    Dim tb As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim ssql As String
    Dim m_date As Date
    ssql = "Select max(v_Date) as dat1 from Pre_Cash where v_date <= #" & V_Date & "#"
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("dat1").Value) Then
    
        m_date = tb.Fields("dat1").Value
 
    Else
        cash1 = 0
        Exit Function
 
    End If
    tb.Close
    ssql = "select amount from pre_cash where v_date = #" & m_date & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
        pre_1 = tb.Fields("amount").Value
    End If
    tb.Close
    
    ssql = "select sum(debit) as deb,sum(credit) as cred from voucher where e_type=5 and v_date = #" & m_date & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then

        
        If Not IsNull(tb.Fields("deb").Value) Then
            deb1 = tb.Fields("deb").Value
        End If
        If Not IsNull(tb.Fields("cred").Value) Then
            cred1 = tb.Fields("cred").Value
        End If
'           MsgBox "pre" & " " & pre_1 & "cred" & " " & cred1 & "deb" & " " & deb1
    End If
    tb.Close
    db.Close
 '   MsgBox (cred1 + pre_1) - deb1
    cash1 = (cred1 + pre_1) - deb1
 '   MsgBox cash1
End Function

Public Function cash2(V_Date As Date) As Currency
    Dim db As Database
    Dim tb As Recordset
    Dim tb2 As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim ssql As String
    Dim m_date As Date
    Set db = OpenDatabase(m_path)
    
    ssql = "Select max(v_Date) as dat1 from pre_cash where v_date <= #" & V_Date & "#"

    Set tb = db.OpenRecordset(ssql)
    If IsNull(tb.Fields("dat1").Value) = False Then
    
        m_date = tb.Fields("dat1").Value
 '   MsgBox m_date
    Else
        m_date = V_Date
    End If
    tb.Close
    
'    MsgBox m_date
    
    pre_1 = cash1(m_date)
'    MsgBox pre_1
    ssql = "select sum(debit) as deb,sum(credit) as cred from voucher where e_type=5 and v_date = #" & m_date & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then

        'pre_1 = tb.Fields("pre_cash").Value
If Not IsNull(tb.Fields("deb").Value) Then
        deb1 = tb.Fields("deb").Value
End If
If Not IsNull(tb.Fields("cred").Value) Then
        cred1 = tb.Fields("cred").Value
End If
     '      MsgBox "pre" & " " & pre_1 & "cred" & " " & cred1 & "deb" & " " & deb1
    End If
    tb.Close
    db.Close
 '   MsgBox (cred1 + pre_1) - deb1
    'cash2 = (cred1 + pre_1) - deb1
    cash2 = pre_1
    
End Function


Public Property Get pathMain() As String
    pathMain = m_path
End Property

Private Sub Class_Initialize()
    m_path = App.Path & "\Bloom.mdb"
   
End Sub

Public Function comb_contracts(CNTL As Control, code As Long) As Integer()
    Dim ssql As String
    
    Dim j As Long
    Dim db As Database
    Dim tb As Recordset
    CNTL.clear
    Set db = OpenDatabase(m_path)
    ssql = "select * from contract where party = " & code
    Set tb = db.OpenRecordset(ssql)
    j = 0
        If Not tb.EOF Then
        tb.MoveLast
        j = tb.RecordCount
        ReDim ii(j)
        j = 0
            Do While Not tb.EOF
                CNTL.AddItem Item1(tb.Fields("cloth").Value)
                CNTL.ItemData(CNTL.NewIndex) = tb.Fields("cont_no")
                'ReDim Preserve ii(UBound(ii) + 1)
                ii(j) = tb.Fields("type").Value
                tb.MoveNext
                j = j + 1
            Loop
            CNTL.ListIndex = 0
        End If
        tb.Close
        db.Close
        comb_contracts = ii()
End Function
Public Function party1(C As Long) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from Parties where code = " & C)
    If Not tb.EOF Then
        party1 = tb.Fields("Name").Value
    Else
        party1 = "NOT FOUND"
    End If
tb.Close
db.Close
End Function
Public Function Address1(C As Long) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from Parties where code = " & C)
    If Not tb.EOF Then
        Address1 = tb.Fields("Address").Value & ""
    Else
        Address1 = "NOT"
    End If
tb.Close
db.Close
End Function
Public Function GST1(C As Long) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from Parties where code = " & C)
    If Not tb.EOF Then
        GST1 = tb.Fields("GSTNo").Value & ""
    Else
        GST1 = "NOT"
    End If
tb.Close
db.Close
End Function

Public Function PartyofCase(C As Double) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from Docs where SrNo = " & C)
    If Not tb.EOF Then
        PartyofCase = party1(tb.Fields("PartyCode").Value)
    Else
        PartyofCase = "NOT FOUND"
    End If
tb.Close
db.Close
End Function

Public Function Item1(C As Long) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from item where code = " & C)
    If Not tb.EOF Then
        Item1 = tb.Fields("name").Value & "  [" & tb.Fields("Unit").Value & "]"
    Else
        Item1 = "NOT FOUND"
    End If
tb.Close
db.Close
End Function
Public Function WareHouse(ByVal C As Long) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from WareHouse where code = " & C)
    If Not tb.EOF Then
        WareHouse = tb.Fields("name").Value
    Else
        WareHouse = "NOT FOUND"
    End If
tb.Close
db.Close
End Function

Public Sub fill_comb(ssql As String, CNTL As Control, itemcol As String, datacol As String)
    Dim db As Database
    Dim tb As Recordset
    CNTL.clear
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            CNTL.AddItem tb.Fields(itemcol).Value
            CNTL.ItemData(CNTL.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        CNTL.ListIndex = 0
    End If
    tb.Close
    db.Close
        
End Sub
Public Sub fill_comb_Item(ssql As String, CNTL As Control, itemcol As String, datacol As String)
    Dim db As Database
    Dim tb As Recordset
    CNTL.clear
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            CNTL.AddItem tb.Fields(itemcol).Value & " {" & tb.Fields("Unit").Value & "}"
            CNTL.ItemData(CNTL.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        CNTL.ListIndex = 0
    End If
    tb.Close
    db.Close
        
End Sub

Public Sub fill_comb2(ssql As String, CNTL As Control, itemcol As String, itemcol2 As String, itemcol3 As String, datacol As String)
    Dim db As Database
    Dim tb As Recordset
    CNTL.clear
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            CNTL.AddItem tb.Fields(itemcol).Value & "-----" & tb.Fields(itemcol2).Value & "-----" & Item1(tb.Fields(itemcol3).Value)
            CNTL.ItemData(CNTL.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        CNTL.ListIndex = 0
    End If
    tb.Close
    db.Close
        
End Sub

Public Sub cont_type_comb(CNTL As Control)
    CNTL.clear
        CNTL.AddItem "Recieved"
        CNTL.ItemData(CNTL.NewIndex) = 1
        CNTL.AddItem "Sent"
        CNTL.ItemData(CNTL.NewIndex) = 2
    CNTL.ListIndex = 0
End Sub
Public Function cont_type_ret(type1 As Byte) As String
    
    Select Case type1
        Case 1
            cont_type_ret = "Recieved"
        Case 2
            cont_type_ret = "Sent"
    End Select
End Function

Public Sub Bro_type_comb(CNTL As Control)
    CNTL.clear
        CNTL.AddItem "/Unit"
        CNTL.ItemData(CNTL.NewIndex) = 1
        CNTL.AddItem "Total Amount"
        CNTL.ItemData(CNTL.NewIndex) = 2
    CNTL.ListIndex = 0
End Sub
Public Function Balance(C As Long) As Currency
Dim db As Database
Dim tb As Recordset
Dim ssql As String

    Set db = OpenDatabase(m_path)
    ssql = "select sum(debit - credit) as bal from voucher where party = " & C
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("bal").Value) Then
        Balance = tb.Fields("Bal").Value
    Else
        Balance = 0
    End If
tb.Close
db.Close

End Function

