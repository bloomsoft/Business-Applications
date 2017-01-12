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
Public Function OutStand_DyeStock(c As Long, ITM As Long) As Currency
Dim db_m As Database
Dim tb As Recordset
Dim ss As Currency
Dim qs As Currency
Dim qr As Currency
Dim ssql As String
Set db_m = OpenDatabase(m_path)
ssql = "select sum(quantity)as q from sent_view_1 where party = " & c & " and item = " & ITM
Set tb = db_m.OpenRecordset(ssql)
'MsgBox ssql
If Not IsNull(tb.Fields("q").Value) Then
'MsgBox qs
        qs = tb.Fields("q").Value
        ssql = "select sum(quantity) as q from rec_view_final where raw_item = " & ITM
        ssql = ssql & " and party = " & c
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
Public Function Stock(c As Long, V_DATE As Date) As Currency
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
    ssql = "select sum(quantity)as q from in_dtl where item = " & c
    ssql = ssql & " and v_Date <= #" & V_DATE & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("q").Value) Then
        In1 = tb.Fields("q").Value
    End If
    tb.Close
    'Sale
    ssql = "select sum(quantity)as q from sale_2 where item = " & c
    ssql = ssql & " and v_Date <= #" & V_DATE & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("q").Value) Then
        sale1 = tb.Fields("q").Value
    End If
    tb.Close
    'Sent
    ssql = "select sum(quantity)as q from sent_2 where item = " & c
    ssql = ssql & " and v_Date <= #" & V_DATE & "#"
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("q").Value) Then
        Sent1 = tb.Fields("q").Value
    End If
    tb.Close
    'Rec
    ssql = "select sum(quantity)as q from rec_2 where item = " & c
    ssql = ssql & " and v_Date <= #" & V_DATE & "#"
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
Public Function RPAD(S As String, length As Byte) As String
    Dim l As Integer
    Dim i As Integer
    Dim e As Integer
    
    l = Len(S)
    e = length - l
    S = Trim(S)
    For i = 1 To e
        S = S & "_"
    Next i
    RPAD = Mid(S, 1, length)
End Function
Public Property Get report_path() As String
    report_path = "c:\soft\Makkah\"
'report_path = CurDir & "\"
End Property
Public Function cash1(V_DATE As Date) As Currency
    Dim db As Database
    Dim tb As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim ssql As String
    Dim m_date As Date
    ssql = "Select max(v_Date) as dat1 from Pre_Cash where v_date <= #" & V_DATE & "#"
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

Public Function cash2(V_DATE As Date) As Currency
    Dim db As Database
    Dim tb As Recordset
    Dim tb2 As Recordset
    Dim pre_1 As Currency
    Dim deb1 As Currency
    Dim cred1 As Currency
    Dim ssql As String
    Dim m_date As Date
    Set db = OpenDatabase(m_path)
    
    ssql = "Select max(v_Date) as dat1 from pre_cash where v_date <= #" & V_DATE & "#"

    Set tb = db.OpenRecordset(ssql)
    If IsNull(tb.Fields("dat1").Value) = False Then
    
        m_date = tb.Fields("dat1").Value
 '   MsgBox m_date
    Else
        m_date = V_DATE
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
   m_path = App.path & "\BLOOMNEW.mdb"
    'm_path = CurDir & "\" & "main.mdb"
    'm_path = "c:\soft\Makkah\Years\" & UserN & "\BLOOMNEW.mdb"
End Sub

Public Function comb_contracts(cntl As Control, code As Long) As Integer()
    Dim ssql As String
    
    Dim j As Long
    Dim db As Database
    Dim tb As Recordset
    cntl.clear
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
                cntl.AddItem Item1(tb.Fields("cloth").Value)
                cntl.ItemData(cntl.NewIndex) = tb.Fields("cont_no")
                'ReDim Preserve ii(UBound(ii) + 1)
                ii(j) = tb.Fields("type").Value
                tb.MoveNext
                j = j + 1
            Loop
            cntl.ListIndex = 0
        End If
        tb.Close
        db.Close
        comb_contracts = ii()
End Function
Public Function party1(c As Long) As String
    Dim tb As ADODB.Recordset
    Set tb = CN.Execute("select * from acchart where code = " & c)
    If Not tb.EOF Then
        party1 = tb.Fields("Name").Value
    Else
        party1 = "NOT FOUND"
    End If
tb.Close
End Function
Public Sub PartDet(c As Long, PType As Control, PName As Control, Pno As Control)
    Dim tb As ADODB.Recordset
    Set tb = CN.Execute("select * from parts where partcode = " & c)
    If Not tb.EOF Then
    If tb.Fields("parttype").Value = 0 Then
    PType = "Needle"
    ElseIf tb.Fields("parttype").Value = 1 Then
    PType = "Sinker"
    End If
    
    PName = tb.Fields("partname").Value
    Pno = tb.Fields("partno").Value & ""
    
    End If
tb.Close
End Sub

Public Function GST1(c As Long) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from acchart where code = " & c)
    If Not tb.EOF Then
        GST1 = tb.Fields("GSTNo").Value
    Else
        GST1 = "NOT FOUND"
    End If
tb.Close
db.Close
End Function

Public Function Item1(c As Long) As String
    Dim tb As ADODB.Recordset
    
    Set tb = CN.Execute("select * from Yarn where code = " & c)
    If Not tb.EOF Then
        Item1 = tb.Fields("name").Value
    Else
        Item1 = "NOT FOUND"
    End If
tb.Close
End Function
Public Function Cloth(c As Long) As String
    Dim tb As ADODB.Recordset
    Set tb = CN.Execute("select * from Cloth where code = " & c)
    If Not tb.EOF Then
        Cloth = tb.Fields("name").Value & " " & tb.Fields("WIdth").Value & ""
    Else
        Cloth = "NOT FOUND"
    End If
tb.Close
End Function

Public Function ClothWidth(c As Long) As String
    Dim tb As ADODB.Recordset
    Set tb = CN.Execute("select * from Cloth where code = " & c)
    If Not tb.EOF Then
        ClothWidth = tb.Fields("Width").Value & ""
    End If
tb.Close
End Function

Public Function Yarn(c As Long) As String
    Dim tb As ADODB.Recordset
    Set tb = CN.Execute("select * from Yarn where Y_Type=1 and code = " & c)
    If Not tb.EOF Then
        Yarn = tb.Fields("name").Value & " " & tb.Fields("Count").Value & ""
    Else
        Yarn = "NOT FOUND"
    End If
tb.Close

End Function
Public Function Lycra(c As Long) As String
    Dim tb As ADODB.Recordset
    
    Set tb = CN.Execute("select * from Yarn where Y_Type=2 and code = " & c)
    If Not tb.EOF Then
        Lycra = tb.Fields("name").Value & " " & tb.Fields("Count").Value & " "
    Else
        Lycra = "NOT FOUND"
    End If
tb.Close
End Function


Public Function machine(c As Long) As String
    Dim tb As ADODB.Recordset
    
    Set tb = CN.Execute("select * from Machine where code = " & c)
    If Not tb.EOF Then
        machine = tb.Fields("name").Value
    Else
        machine = "NOT FOUND"
    End If
tb.Close

End Function

Public Function YarnCount(c As Long) As String
    Dim tb As ADODB.Recordset
    Set tb = CN.Execute("select * from Yarn where code = " & c)
    If Not tb.EOF Then
        YarnCount = tb.Fields("Count").Value & ""
    End If
tb.Close

End Function


Public Function WareHouse(ByVal c As Long) As String
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset("select * from WareHouse where code = " & c)
    If Not tb.EOF Then
        WareHouse = tb.Fields("name").Value
    Else
        WareHouse = "NOT FOUND"
    End If
tb.Close
db.Close
End Function

Public Sub fill_comb(ssql As String, cntl As Control, itemcol As String, datacol As String)
    Dim tb As ADODB.Recordset
    cntl.clear
    Set tb = CN.Execute(ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            cntl.AddItem tb.Fields(itemcol).Value
            cntl.ItemData(cntl.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        cntl.ListIndex = 0
    End If
    tb.Close
        
End Sub
Public Sub fill_comb2(ssql As String, cntl As Control, itemcol As String, itemcol2 As String, itemcol3 As String, datacol As String)
    Dim db As Database
    Dim tb As Recordset
    cntl.clear
    Set db = OpenDatabase(m_path)
    Set tb = db.OpenRecordset(ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            cntl.AddItem tb.Fields(itemcol).Value & "-----" & tb.Fields(itemcol2).Value & "-----" & Item1(tb.Fields(itemcol3).Value)
            cntl.ItemData(cntl.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        cntl.ListIndex = 0
    End If
    tb.Close
    db.Close
        
End Sub

Public Sub cont_type_comb(cntl As Control)
    cntl.clear
        cntl.AddItem "Recieved"
        cntl.ItemData(cntl.NewIndex) = 1
        cntl.AddItem "Sent"
        cntl.ItemData(cntl.NewIndex) = 2
    cntl.ListIndex = 0
End Sub
Public Function cont_type_ret(type1 As Byte) As String
    
    Select Case type1
        Case 1
            cont_type_ret = "Recieved"
        Case 2
            cont_type_ret = "Sent"
    End Select
End Function

Public Sub Bro_type_comb(cntl As Control)
    cntl.clear
        cntl.AddItem "/Unit"
        cntl.ItemData(cntl.NewIndex) = 1
        cntl.AddItem "Total Amount"
        cntl.ItemData(cntl.NewIndex) = 2
    cntl.ListIndex = 0
End Sub
Public Function Balance(c As Long) As Currency
Dim db As Database
Dim tb As Recordset
Dim ssql As String

    Set db = OpenDatabase(m_path)
    ssql = "select sum(debit - credit) as bal from voucher where party = " & c
    Set tb = db.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("bal").Value) Then
        Balance = tb.Fields("Bal").Value
    Else
        Balance = 0
    End If
tb.Close
db.Close

End Function

