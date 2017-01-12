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
Private blm As bloom1
Public Sub Cont_Sum(cont_no As Long, c_type As Byte, cntl As Control)
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim tb_m As ADODB.Recordset
Dim tb As ADODB.Recordset
cntl.Value = 0

Set db_t = OpenDatabase(t_path)
ssql = "Delete from Cont"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("cont", dbOpenTable)
ssql = "select * from cont_1 where cont_no = " & cont_no & " and E_Type = " & c_type
Set tb_m = CN.Execute(ssql)

If Not tb_m.EOF Then
    'inward
    ssql = "select * from Inward where Cont_no  = " & cont_no & " and E_Type = " & c_type
    Set tb = CN.Execute(ssql)
    If Not tb.EOF Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
            tb_t.AddNew
                tb_t.Fields("Cont_no").Value = cont_no
                tb_t.Fields("E_Type").Value = c_type
                tb_t.Fields("C_Date").Value = tb_m.Fields("V_Date").Value
                tb_t.Fields("Accode").Value = tb_m.Fields("Party").Value
                tb_t.Fields("AcName").Value = blm.party1(tb_m.Fields("Party").Value)
                tb_t.Fields("Del_Date").Value = tb_m.Fields("Del_Date").Value
                tb_t.Fields("CQuantity").Value = tb_m.Fields("CQuantity").Value
                tb_t.Fields("YQuantity").Value = tb_m.Fields("YQuantity").Value
                tb_t.Fields("ClothName").Value = blm.Cloth(tb_m.Fields("Item").Value)
                tb_t.Fields("in_no").Value = tb.Fields("in_no").Value
                tb_t.Fields("in_Date").Value = tb.Fields("in_Date").Value
                tb_t.Fields("Cloth").Value = tb.Fields("Quantity").Value
                tb_t.Fields("CRolls").Value = tb.Fields("Bags").Value
                'tb_t.Fields("Cloth").Value = tb.Fields("Quantity").Value
                'tb_t.Fields("CRolls").Value = tb.Fields("Bags").Value
                tb_t.Fields("Lycra").Value = tb.Fields("Lycra").Value
                tb_t.Fields("LBags").Value = tb.Fields("L_bags").Value
                tb_t.Fields("In_Type").Value = 100
                tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
                
            tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
        Loop
    End If
    tb.Close

    'outward
    cntl.Value = 0
    ssql = "select * from outward where Cont_no  = " & cont_no & " and E_Type = " & c_type
    Set tb = CN.Execute(ssql)
    If Not tb.EOF Then
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
        Do While Not tb.EOF
            tb_t.AddNew
                tb_t.Fields("Cont_no").Value = cont_no
                tb_t.Fields("E_Type").Value = c_type
                tb_t.Fields("C_Date").Value = tb_m.Fields("V_Date").Value
                tb_t.Fields("Accode").Value = tb_m.Fields("Party").Value
                tb_t.Fields("AcName").Value = blm.party1(tb_m.Fields("Party").Value)
                tb_t.Fields("Del_Date").Value = tb_m.Fields("Del_Date").Value
                tb_t.Fields("in_no").Value = tb.Fields("out_no").Value
                tb_t.Fields("in_Date").Value = tb.Fields("v_Date").Value
                tb_t.Fields("Cloth").Value = tb.Fields("Quantity").Value
                tb_t.Fields("LBags").Value = 0
                tb_t.Fields("CRolls").Value = tb.Fields("rolls").Value
                tb_t.Fields("In_Type").Value = 200
                tb_t.Fields("Rate").Value = tb.Fields("Rate").Value
            tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
        Loop
    End If
    tb.Close

End If
tb_m.Close
tb_t.Close
db_t.Close
End Sub
Public Sub Cont_Sum_Party(Party As Long, c_type As Byte, cntl As Control)
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim tb_m As ADODB.Recordset
Dim tb As ADODB.Recordset
cntl.Value = 0

Set db_t = OpenDatabase(t_path)
ssql = "Delete from Cont"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("cont", dbOpenTable)
ssql = "select * from cont_1 where Party = " & Party & " and E_Type = " & c_type
Set tb_m = CN.Execute(ssql)

If Not tb_m.EOF Then
'tb_m.MoveLast
'cntl.Max =
'MsgBox tb_m.RecordCount
'tb_m.MoveFirst
Do While Not tb_m.EOF
cntl.Value = 0
 '   MsgBox tb_m.Fields("Cont_no").Value
    'inward
    ssql = "select * from Inward where Cont_no  = " & tb_m.Fields("cont_no").Value & " and E_Type = " & c_type
    Set tb = CN.Execute(ssql)
    If Not tb.EOF Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
            tb_t.AddNew
                tb_t.Fields("Cont_no").Value = tb_m.Fields("cont_no").Value
                tb_t.Fields("E_Type").Value = c_type
                tb_t.Fields("C_Date").Value = tb_m.Fields("V_Date").Value
                tb_t.Fields("Accode").Value = tb_m.Fields("Party").Value
                tb_t.Fields("AcName").Value = blm.party1(tb_m.Fields("Party").Value)
                tb_t.Fields("Del_Date").Value = tb_m.Fields("Del_Date").Value
                tb_t.Fields("CQuantity").Value = tb_m.Fields("CQuantity").Value
                tb_t.Fields("YQuantity").Value = tb_m.Fields("YQuantity").Value
                tb_t.Fields("ClothName").Value = blm.Cloth(tb_m.Fields("Item").Value)
                tb_t.Fields("in_no").Value = tb.Fields("in_no").Value
                tb_t.Fields("in_Date").Value = tb.Fields("in_Date").Value
                tb_t.Fields("Cloth").Value = tb.Fields("Quantity").Value
                tb_t.Fields("CRolls").Value = tb.Fields("Bags").Value
                'tb_t.Fields("Cloth").Value = tb.Fields("Quantity").Value
                'tb_t.Fields("CRolls").Value = tb.Fields("Bags").Value
                tb_t.Fields("Lycra").Value = tb.Fields("Lycra").Value
                tb_t.Fields("LBags").Value = tb.Fields("L_bags").Value
                tb_t.Fields("In_Type").Value = 100
                tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
                
            tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
        Loop
    End If
    tb.Close

    'outward
    cntl.Value = 0
    ssql = "select * from outward where Cont_no  = " & tb_m.Fields("cont_no").Value & " and E_Type = " & c_type
    Set tb = CN.Execute(ssql)
    If Not tb.EOF Then
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
        Do While Not tb.EOF
            tb_t.AddNew
                tb_t.Fields("Cont_no").Value = tb_m.Fields("cont_no").Value
                tb_t.Fields("E_Type").Value = c_type
                tb_t.Fields("C_Date").Value = tb_m.Fields("V_Date").Value
                tb_t.Fields("Accode").Value = tb_m.Fields("Party").Value
                tb_t.Fields("AcName").Value = blm.party1(tb_m.Fields("Party").Value)
                tb_t.Fields("Del_Date").Value = tb_m.Fields("Del_Date").Value
                tb_t.Fields("in_no").Value = tb.Fields("out_no").Value
                tb_t.Fields("in_Date").Value = tb.Fields("v_Date").Value
                tb_t.Fields("Cloth").Value = tb.Fields("Quantity").Value
                tb_t.Fields("CRolls").Value = tb.Fields("rolls").Value
                tb_t.Fields("In_Type").Value = 200
                  tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
            tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
        Loop
    End If
    tb.Close
tb_m.MoveNext
Loop
End If
tb_m.Close
tb_t.Close
db_t.Close
End Sub

Public Sub iteminout(s_date As Date, e_date As Date, cntl As Control, Itcode As Long)
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim tb_m As ADODB.Recordset
Dim tb_p As ADODB.Recordset
Dim tb_s As ADODB.Recordset
Dim tb_i As ADODB.Recordset
Dim tb_ac As ADODB.Recordset
Dim opq As Currency, opctn As Currency
Set db_t = OpenDatabase(t_path)
cntl.Value = 0
ssql = "delete from ctns"
db_t.Execute ssql

ssql = "select * from item order by code"
Set tb_i = CN.Execute(ssql)

ssql = "select * from acchart order by code"
Set tb_ac = CN.Execute(ssql)

Set tb_t = db_t.OpenRecordset("ctns", dbOpenTable)
ssql = "select item,sum(quantity) as q,count(ctn_no)As c from Purc where Item = " & Itcode
ssql = ssql & " and v_date < #" & s_date & "# Group by Item"
Set tb_p = CN.Execute(ssql)

ssql = "select item,sum(quantity) as q,count(ctn_no)As c from Sale where Item = " & Itcode
ssql = ssql & " and v_date < #" & s_date & "# Group by Item"
Set tb_s = CN.Execute(ssql)
If Not tb_p.EOF Then
If Not IsNull(tb_p.Fields("q").Value) Then
    opq = tb_p.Fields("q").Value
    opctn = tb_p.Fields("c").Value
End If
If Not tb_s.EOF Then
If Not IsNull(tb_s.Fields("q").Value) Then
    opq = opq - tb_s.Fields("q").Value
    opctn = opctn - tb_s.Fields("c").Value
End If
End If
End If
tb_t.AddNew
            tb_t.Fields("lot_no").Value = opctn 'open cartons
            tb_t.Fields("rate").Value = opq     'open quantity
            tb_t.Fields("s_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("v_date").Value = s_date
            tb_t.Fields("itcode").Value = Itcode
                If Not tb_i.EOF Then
                    tb_i.Find "Code = " & Itcode
                    If Not tb_i.EOF Then
                        tb_t.Fields("Itname").Value = tb_i.Fields("Name").Value
                    End If
                End If
        tb_t.Fields("Quantity").Value = opq
        tb_t.Fields("ctn_no").Value = opctn
        tb_t.Fields("Name").Value = "Openings...."
tb_t.Update
cntl.Value = 0
'Purchases
ssql = "select a.v_date,party,sum(quantity)as q,count(ctn_no)as c from in_view_1 where a.v_date between #" & s_date & "# and #" & e_date & "# and Item = " & Itcode
ssql = ssql & " group by a.v_date,party"
Set tb_m = CN.Execute(ssql)

If Not tb_m.EOF Then
tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF
        tb_t.AddNew
            tb_t.Fields("lot_no").Value = opctn 'open cartons
            tb_t.Fields("rate").Value = opq     'open quantity
            tb_t.Fields("s_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
            tb_t.Fields("itcode").Value = Itcode
                If Not tb_i.EOF Then
                    tb_i.Find "Code = " & Itcode
                    If Not tb_i.EOF Then
                        tb_t.Fields("Itname").Value = tb_i.Fields("Name").Value
                    End If
                End If
            tb_t.Fields("E_Type").Value = 2
            tb_t.Fields("Quantity").Value = tb_m.Fields("Q").Value
            tb_t.Fields("ctn_no").Value = tb_m.Fields("c").Value
            tb_t.Fields("Code").Value = tb_m.Fields("Party").Value
                If Not tb_ac.EOF Then
                    tb_ac.Find "Code = " & tb_m.Fields("Party").Value
                    If Not tb_ac.EOF Then
                        tb_t.Fields("name").Value = tb_ac.Fields("Name").Value
                    End If
                End If
            tb_t.Update
            tb_m.MoveNext
            cntl.Value = cntl.Value + 1
Loop
End If
tb_m.Close
cntl.Value = 0
'Sales
ssql = "select a.v_date,party,sum(quantity)as q,count(ctn_no)as c from sale_view_1 where a.v_date between #" & s_date & "# and #" & e_date & "# and Item = " & Itcode
ssql = ssql & " group by a.v_date,party"
Set tb_m = CN.Execute(ssql)

If Not tb_m.EOF Then
tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF
        tb_t.AddNew
            tb_t.Fields("lot_no").Value = opctn 'open cartons
            tb_t.Fields("rate").Value = opq     'open quantity
            tb_t.Fields("s_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
            tb_t.Fields("itcode").Value = Itcode
                If Not tb_i.EOF Then
                    tb_i.Find "Code = " & Itcode
                    If Not tb_i.EOF Then
                        tb_t.Fields("Itname").Value = tb_i.Fields("Name").Value
                    End If
                End If
            tb_t.Fields("E_Type").Value = 1
            tb_t.Fields("Quantity").Value = tb_m.Fields("Q").Value * -1
            tb_t.Fields("ctn_no").Value = tb_m.Fields("c").Value * -1
            tb_t.Fields("Code").Value = tb_m.Fields("Party").Value
                If Not tb_ac.EOF Then
                    tb_ac.Find "Code = " & tb_m.Fields("Party").Value
                    If Not tb_ac.EOF Then
                        tb_t.Fields("name").Value = tb_ac.Fields("Name").Value
                    End If
                End If
            tb_t.Update
            tb_m.MoveNext
            cntl.Value = cntl.Value + 1
Loop
End If
tb_m.Close
cntl.Value = 0
'sent
ssql = "select a.v_date,party,sum(quantity)as q,count(ctn_no)as c from sent_view_1 where a.v_date between #" & s_date & "# and #" & e_date & "# and Item = " & Itcode
ssql = ssql & " group by a.v_date,party"
Set tb_m = CN.Execute(ssql)

If Not tb_m.EOF Then
tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF
        tb_t.AddNew
            tb_t.Fields("lot_no").Value = opctn 'open cartons
            tb_t.Fields("rate").Value = opq     'open quantity
            tb_t.Fields("s_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
            tb_t.Fields("itcode").Value = Itcode
                If Not tb_i.EOF Then
                    tb_i.Find "Code = " & Itcode
                    If Not tb_i.EOF Then
                        tb_t.Fields("Itname").Value = tb_i.Fields("Name").Value
                    End If
                End If
            tb_t.Fields("E_Type").Value = 3
            tb_t.Fields("Quantity").Value = tb_m.Fields("Q").Value * -1
            tb_t.Fields("ctn_no").Value = tb_m.Fields("c").Value * -1
            tb_t.Fields("Code").Value = tb_m.Fields("Party").Value
                If Not tb_ac.EOF Then
                    tb_ac.Find "Code = " & tb_m.Fields("Party").Value
                    If Not tb_ac.EOF Then
                        tb_t.Fields("name").Value = tb_ac.Fields("Name").Value
                    End If
                End If
            tb_t.Update
            tb_m.MoveNext
            cntl.Value = cntl.Value + 1
Loop
End If
tb_m.Close
cntl.Value = 0
'Rec
ssql = "select a.v_date,party,sum(quantity)as q,count(ctn_no)as c from rec_view_1 where a.v_date between #" & s_date & "# and #" & e_date & "# and Item = " & Itcode
ssql = ssql & " group by a.v_date,party"
Set tb_m = CN.Execute(ssql)

If Not tb_m.EOF Then
tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF
        tb_t.AddNew
            tb_t.Fields("lot_no").Value = opctn 'open cartons
            tb_t.Fields("rate").Value = opq     'open quantity
            tb_t.Fields("s_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
            tb_t.Fields("itcode").Value = Itcode
                If Not tb_i.EOF Then
                    tb_i.Find "Code = " & Itcode
                    If Not tb_i.EOF Then
                        tb_t.Fields("Itname").Value = tb_i.Fields("Name").Value
                    End If
                End If
            tb_t.Fields("E_Type").Value = 4
            tb_t.Fields("Quantity").Value = tb_m.Fields("Q").Value
            tb_t.Fields("ctn_no").Value = tb_m.Fields("c").Value
            tb_t.Fields("Code").Value = tb_m.Fields("Party").Value
                If Not tb_ac.EOF Then
                    tb_ac.Find "Code = " & tb_m.Fields("Party").Value
                    If Not tb_ac.EOF Then
                        tb_t.Fields("name").Value = tb_ac.Fields("Name").Value
                    End If
                End If
            tb_t.Update
            tb_m.MoveNext
            cntl.Value = cntl.Value + 1
Loop
End If
tb_m.Close

tb_t.Close
tb_ac.Close
tb_i.Close
tb_p.Close
tb_s.Close
db_t.Close

End Sub
Public Sub PL(s_date As Date, e_date As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim prate As Currency
Dim pdate As Date
Dim tb_m As Recordset
Dim tb_p As Recordset
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)

ssql = "delete from PL"
db_t.Execute ssql

cntl.Value = 0
Set tb_t = db_t.OpenRecordset("PL", dbOpenTable)

ssql = "select * from Sale where v_date between #" & s_date & "# and #" & e_date & "# and e_type = 1 order by v_date"
Set tb_m = db_m.OpenRecordset(ssql)

ssql = "select * from Purc where v_date between #" & s_date & "# and #" & e_date & "# order by Ctn_no"
Set tb_p = db_m.OpenRecordset(ssql)

If Not tb_m.EOF Then
    tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF
    If Not tb_p.EOF Then
    If Not IsNull(tb_m.Fields("Ctn_no").Value) Then
        tb_p.FindFirst "Ctn_no = " & tb_m.Fields("Ctn_no").Value
        If tb_p.NoMatch Then
            prate = 0
        Else
            
                prate = tb_p.Fields("Rate").Value
                pdate = tb_p.Fields("V_date").Value
        End If
    End If
    End If
    tb_t.AddNew
        tb_t.Fields("ST_Date").Value = s_date
        tb_t.Fields("E_Date").Value = e_date
        tb_t.Fields("P_Date").Value = pdate
        tb_t.Fields("S_Date").Value = tb_m.Fields("V_date").Value
        tb_t.Fields("Ctn_no").Value = tb_m.Fields("Ctn_no").Value
        tb_t.Fields("Lot_no").Value = tb_m.Fields("Lot_no").Value
        tb_t.Fields("Prate").Value = prate 'b_m.Fields("Ctn_no").Value
        tb_t.Fields("Srate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("Quantity").Value = tb_m.Fields("Quantity").Value
    tb_t.Update
    tb_m.MoveNext
    cntl.Value = cntl.Value + 1
Loop
End If
tb_t.Close
tb_p.Close
tb_m.Close
db_t.Close
db_m.Close
End Sub
Public Sub recsent(s_date As Date, e_date As Date, p As Long, cntl As Control, ITM As Long)
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim tb_m As Recordset
Dim tb_rbal As Recordset
Dim tb_sbal As Recordset
Dim bs As Currency
Dim br As Currency


Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)

ssql = "delete from sentrec"
db_t.Execute ssql

cntl.Value = 0
Set tb_t = db_t.OpenRecordset("sentrec", dbOpenTable)

ssql = "select sum(quantity) as q from sent_view_1 where a.v_date < #" & s_date & "# and party = " & p & " and item = " & ITM
Set tb_sbal = db_m.OpenRecordset(ssql)

If Not IsNull(tb_sbal.Fields("q").Value) Then
       bs = tb_sbal.Fields("q").Value
End If

ssql = "select sum(quantity) as q from rec_view_1 where a.v_date < #" & s_date & "#  and party = " & p & " and raw_item = " & ITM
Set tb_rbal = db_m.OpenRecordset(ssql)

If Not IsNull(tb_rbal.Fields("q").Value) Then
       br = tb_rbal.Fields("q").Value
End If

tb_t.AddNew
        tb_t.Fields("v_date").Value = s_date
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("Code").Value = p
        tb_t.Fields("Name").Value = blm.party1(p)
        tb_t.Fields("Quantity").Value = bs - br
        tb_t.Fields("E_type").Value = 10   'opbal
        tb_t.Fields("ITEM").Value = ITM
        tb_t.Fields("ITnAME").Value = blm.Item1(ITM)
        
    
   'MsgBox bs - br & " " & tb_m.Fields("iTEM")
    
    tb_t.Update
    
ssql = "select * from sent_rec_view where e_type=3 and v_date between #" & s_date & "# and #" & e_date & "#  and party = " & p & " and item = " & ITM & " order by item"
Set tb_m = db_m.OpenRecordset(ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF

    tb_t.AddNew
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("Trans_no").Value = tb_m.Fields("Issue_no").Value
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("Item").Value = tb_m.Fields("Item")
        tb_t.Fields("ItName").Value = tb_m.Fields("ItName")
        tb_t.Fields("Code").Value = tb_m.Fields("Party")
        tb_t.Fields("Name").Value = tb_m.Fields("PName")
        tb_t.Fields("Raw_Name").Value = tb_m.Fields("RawName")
        tb_t.Fields("Raw_Item").Value = tb_m.Fields("Raw_Item").Value
        tb_t.Fields("Unit").Value = tb_m.Fields("Unit")
     
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate")
        If E_type = 3 Then
            tb_t.Fields("Quantity").Value = tb_m.Fields("Quantity").Value * -1
        Else
            tb_t.Fields("Quantity").Value = tb_m.Fields("Quantity").Value
        End If
        tb_t.Fields("Dye_Rate").Value = tb_m.Fields("Dye_Rate")
        tb_t.Fields("E_type").Value = tb_m.Fields("E_Type")
    
   'MsgBox bs - br & " " & tb_m.Fields("iTEM")
    
    tb_t.Update
    tb_m.MoveNext
    cntl.Value = cntl.Value + 1
Loop
End If
cntl.Value = 0
ssql = "select * from sent_rec_view where e_type=4 and v_date between #" & s_date & "# and #" & e_date & "#  and party = " & p & " and raw_item = " & ITM & " order by item"
Set tb_m = db_m.OpenRecordset(ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF

    tb_t.AddNew
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("Trans_no").Value = tb_m.Fields("Issue_no").Value
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("Item").Value = tb_m.Fields("Item")
        tb_t.Fields("ItName").Value = tb_m.Fields("ItName")
        tb_t.Fields("Code").Value = tb_m.Fields("Party")
        tb_t.Fields("Name").Value = tb_m.Fields("PName")
        tb_t.Fields("Raw_Name").Value = tb_m.Fields("RawName")
        tb_t.Fields("Raw_Item").Value = tb_m.Fields("Raw_Item").Value
        tb_t.Fields("Unit").Value = tb_m.Fields("Unit")
     
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate")
        If E_type = 3 Then
            tb_t.Fields("Quantity").Value = tb_m.Fields("Quantity").Value
        Else
            tb_t.Fields("Quantity").Value = tb_m.Fields("Quantity").Value * -1
        End If
        tb_t.Fields("Dye_Rate").Value = tb_m.Fields("Dye_Rate")
        tb_t.Fields("E_type").Value = tb_m.Fields("E_Type")
    
   'MsgBox bs - br & " " & tb_m.Fields("iTEM")
    
    tb_t.Update
    tb_m.MoveNext
    cntl.Value = cntl.Value + 1
Loop
End If

tb_sbal.Close
tb_rbal.Close
tb_m.Close
tb_t.Close
db_m.Close
db_t.Close

End Sub
Public Sub recsentALL(s_date As Date, e_date As Date, p As Long, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim tb_m As Recordset
Dim tb_rbal As Recordset
Dim tb_sbal As Recordset
Dim bs As Currency
Dim br As Currency


Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)

ssql = "delete from sentrec"
db_t.Execute ssql

cntl.Value = 0
Set tb_t = db_t.OpenRecordset("sentrec", dbOpenTable)

ssql = "select ITEM,sum(quantity) as q from sent_view_1 where a.v_date < #" & s_date & "# and party = " & p & " GROUP BY ITEM"
Set tb_sbal = db_m.OpenRecordset(ssql)


ssql = "select RAW_ITEM,sum(quantity) as q from rec_view_1 where a.v_date < #" & s_date & "#  and party = " & p & " GROUP BY RAW_ITEM"
Set tb_rbal = db_m.OpenRecordset(ssql)


If Not tb_sbal.EOF Then

Do While Not tb_sbal.EOF
bs = tb_sbal.Fields("q").Value
If tb_sbal.Fields("ITEM").Value = 67 Then
'MsgBox tb_sbal.Fields("ITEM").Value
End If
If Not tb_rbal.EOF Then
    tb_rbal.FindFirst "RAW_ITEM = " & tb_sbal.Fields("ITEM").Value
    If tb_rbal.NoMatch = False Then
        br = tb_rbal.Fields("Q").Value
        If tb_rbal.Fields("raw_iteM").Value = 67 Then
 '           MsgBox br
        End If
    Else
        br = 0
    End If
End If
tb_t.AddNew
        tb_t.Fields("v_date").Value = s_date
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("Code").Value = p
        tb_t.Fields("Name").Value = blm.party1(p)
        tb_t.Fields("Raw_Item").Value = tb_sbal.Fields("iTEM").Value
        tb_t.Fields("Raw_Name").Value = blm.Item1(tb_sbal.Fields("iTEM").Value)
        tb_t.Fields("ITEM").Value = tb_sbal.Fields("iTEM").Value
        tb_t.Fields("ITnAME").Value = blm.Item1(tb_sbal.Fields("iTEM").Value)
        tb_t.Fields("Quantity").Value = bs - br
        tb_t.Fields("E_type").Value = 10   'opbal
'If tb_sbal.Fields("ITEM").Value = 67 Then
'   MsgBox bs - br & " " & tb_sbal.Fields("iTEM")
'End If
tb_t.Update
tb_sbal.MoveNext
Loop
End If

'tb_t.AddNew
'        tb_t.Fields("v_date").Value = s_date
'        tb_t.Fields("s_date").Value = s_date
'        tb_t.Fields("e_date").Value = e_date
'        tb_t.Fields("Quantity").Value = bs - br
'        tb_t.Fields("E_type").Value = 10   'opbal
    
   'MsgBox bs - br & " " & tb_m.Fields("iTEM")
    
'tb_t.Update
    
ssql = "select * from sent_rec_view where v_date between #" & s_date & "# and #" & e_date & "#  and party = " & p & " order by item"
Set tb_m = db_m.OpenRecordset(ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
    cntl.Max = tb_m.RecordCount
    tb_m.MoveFirst
Do While Not tb_m.EOF

    tb_t.AddNew
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("Trans_no").Value = tb_m.Fields("Issue_no").Value
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("Item").Value = tb_m.Fields("Item")
        tb_t.Fields("ItName").Value = tb_m.Fields("ItName")
        tb_t.Fields("Code").Value = tb_m.Fields("Party")
        tb_t.Fields("Name").Value = tb_m.Fields("PName")
        tb_t.Fields("Raw_Name").Value = tb_m.Fields("RawName")
        tb_t.Fields("Raw_Item").Value = tb_m.Fields("Raw_Item").Value
        tb_t.Fields("Unit").Value = tb_m.Fields("Unit")
     
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate")
        If tb_m.Fields("e_type").Value = 3 Then
            tb_t.Fields("Quantity").Value = tb_m.Fields("Quantity").Value
            tb_t.Fields("Raw_Item").Value = tb_m.Fields("Item")
            tb_t.Fields("Raw_Name").Value = tb_m.Fields("ItName").Value
        
        End If
        If tb_m.Fields("e_type").Value = 4 Then
            
            tb_t.Fields("Quantity").Value = tb_m.Fields("Quantity").Value * -1
        End If
        tb_t.Fields("Dye_Rate").Value = tb_m.Fields("Dye_Rate")
        tb_t.Fields("E_type").Value = tb_m.Fields("E_Type")
    
   'MsgBox bs - br & " " & tb_m.Fields("iTEM")
    
    tb_t.Update
    tb_m.MoveNext
    cntl.Value = cntl.Value + 1
Loop
End If
tb_sbal.Close
tb_rbal.Close
tb_m.Close
db_m.Close
db_t.Close

End Sub

Public Sub Itm_Inout(s_date As Date, e_date As Date, ITM As Long, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim tb_itm As Recordset
Dim tb_sale As Recordset
Dim tb_pur As Recordset
Dim tb_w As Recordset
Dim tb_p As Recordset
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)

ssql = "delete from Ctns"
db_t.Execute ssql
cntl.Value = 0
Set tb_t = db_t.OpenRecordset("CTns", dbOpenTable)
ssql = "select * from item order by code"
Set tb_itm = db_m.OpenRecordset(ssql)

ssql = "select * from acchart order by code"
Set tb_p = db_m.OpenRecordset(ssql)

ssql = "select * from WareHouse order by code"
Set tb_w = db_m.OpenRecordset(ssql)

ssql = "Select * from purc where item = " & ITM
ssql = ssql & " and v_date between #" & s_date & "# and #" & e_date & "# order by ctn_no"
Set tb_pur = db_m.OpenRecordset(ssql)

ssql = "Select * from sale where item = " & ITM
ssql = ssql & " and v_date between #" & s_date & "# and #" & e_date & "# order by ctn_no"
Set tb_sale = db_m.OpenRecordset(ssql)
If Not tb_pur.EOF Then
tb_pur.MoveLast
cntl.Max = tb_pur.RecordCount
tb_pur.MoveFirst
Do While Not tb_pur.EOF
tb_itm.FindFirst "Code = " & tb_pur.Fields("Item").Value
tb_w.FindFirst "Code = " & tb_pur.Fields("w_h_code").Value
tb_p.FindFirst "Code = " & tb_pur.Fields("Party").Value
tb_t.AddNew
    
    tb_t.Fields("ItCode").Value = tb_pur.Fields("Item").Value
    tb_t.Fields("ItName").Value = tb_itm.Fields("Name").Value
    tb_t.Fields("V_date").Value = tb_pur.Fields("v_date").Value
    tb_t.Fields("Ctn_no").Value = tb_pur.Fields("Ctn_no").Value
    tb_t.Fields("Lot_no").Value = tb_pur.Fields("Lot_no").Value
    tb_t.Fields("Bobins").Value = tb_pur.Fields("Bobins").Value
    tb_t.Fields("Rate").Value = tb_pur.Fields("Rate").Value
    tb_t.Fields("Quantity").Value = tb_pur.Fields("Quantity").Value
    tb_t.Fields("Wh_Name").Value = tb_w.Fields("Name").Value
    tb_t.Fields("Code").Value = tb_pur.Fields("Party").Value
    tb_t.Fields("Name").Value = tb_p.Fields("Name").Value
    tb_t.Fields("e_type").Value = tb_pur.Fields("e_type").Value
    tb_t.Fields("S_date").Value = s_date
    tb_t.Fields("e_date").Value = e_date
tb_t.Update

ssql = "Ctn_no = " & tb_pur.Fields("Ctn_no").Value
If Not tb_sale.EOF Then
    tb_sale.FindFirst ssql
    If tb_sale.NoMatch Then
    
    Else
tb_itm.FindFirst "Code = " & tb_sale.Fields("Item").Value
tb_w.FindFirst "Code = " & tb_sale.Fields("w_h_code").Value
tb_p.FindFirst "Code = " & tb_sale.Fields("Party").Value
        
        tb_t.AddNew
    
            tb_t.Fields("ItCode").Value = tb_sale.Fields("Item").Value
            tb_t.Fields("ItName").Value = tb_itm.Fields("Name").Value
            tb_t.Fields("V_date").Value = tb_sale.Fields("v_date").Value
            tb_t.Fields("Ctn_no").Value = tb_sale.Fields("Ctn_no").Value
            tb_t.Fields("Lot_no").Value = tb_sale.Fields("Lot_no").Value
            tb_t.Fields("Bobins").Value = tb_sale.Fields("Bobins").Value
            tb_t.Fields("Rate").Value = tb_sale.Fields("Rate").Value
            tb_t.Fields("Quantity").Value = tb_sale.Fields("Quantity").Value
            tb_t.Fields("Wh_Name").Value = tb_w.Fields("Name").Value
            tb_t.Fields("Code").Value = tb_sale.Fields("Party").Value
            tb_t.Fields("Name").Value = tb_p.Fields("Name").Value
            tb_t.Fields("e_type").Value = tb_sale.Fields("e_type").Value
            tb_t.Fields("S_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
        tb_t.Update

            
    End If
End If
tb_pur.MoveNext
cntl.Value = cntl.Value + 1
Loop
End If
tb_pur.Close
tb_sale.Close
tb_itm.Close
tb_p.Close
tb_w.Close
tb_t.Close
db_m.Close
db_t.Close
End Sub
Public Sub Itm_Inout_ctn(CTN As Long, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim ssql As String
Dim tb_itm As Recordset
Dim tb_sale As Recordset
Dim tb_pur As Recordset
Dim tb_w As Recordset
Dim tb_p As Recordset
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)

ssql = "delete from Ctns"
db_t.Execute ssql
cntl.Value = 0
Set tb_t = db_t.OpenRecordset("CTns", dbOpenTable)
ssql = "select * from item order by code"
Set tb_itm = db_m.OpenRecordset(ssql)

ssql = "select * from acchart order by code"
Set tb_p = db_m.OpenRecordset(ssql)

ssql = "select * from WareHouse order by code"
Set tb_w = db_m.OpenRecordset(ssql)

ssql = "Select * from purc where ctn_no = " & CTN
'ssql = ssql & " and v_date between #" & s_date & "# and #" & e_date & "# order by ctn_no"
Set tb_pur = db_m.OpenRecordset(ssql)

ssql = "Select * from sale where ctn_no = " & CTN
'ssql = ssql & " and v_date between #" & s_date & "# and #" & e_date & "# order by ctn_no"
Set tb_sale = db_m.OpenRecordset(ssql)
If Not tb_pur.EOF Then
tb_pur.MoveLast
cntl.Max = tb_pur.RecordCount
tb_pur.MoveFirst
Do While Not tb_pur.EOF
tb_itm.FindFirst "Code = " & tb_pur.Fields("Item").Value
tb_w.FindFirst "Code = " & tb_pur.Fields("w_h_code").Value
tb_p.FindFirst "Code = " & tb_pur.Fields("Party").Value
tb_t.AddNew
    
    tb_t.Fields("ItCode").Value = tb_pur.Fields("Item").Value
    tb_t.Fields("ItName").Value = tb_itm.Fields("Name").Value
    tb_t.Fields("V_date").Value = tb_pur.Fields("v_date").Value
    tb_t.Fields("Ctn_no").Value = tb_pur.Fields("Ctn_no").Value
    tb_t.Fields("Lot_no").Value = tb_pur.Fields("Lot_no").Value
    tb_t.Fields("Bobins").Value = tb_pur.Fields("Bobins").Value
    tb_t.Fields("Rate").Value = tb_pur.Fields("Rate").Value
    tb_t.Fields("Quantity").Value = tb_pur.Fields("Quantity").Value
    tb_t.Fields("Wh_Name").Value = tb_w.Fields("Name").Value
    tb_t.Fields("Code").Value = tb_pur.Fields("Party").Value
    tb_t.Fields("Name").Value = tb_p.Fields("Name").Value
    tb_t.Fields("e_type").Value = tb_pur.Fields("e_type").Value
    tb_t.Fields("trans_no").Value = tb_pur.Fields("p_no").Value
    
    tb_t.Fields("S_date").Value = s_date
    tb_t.Fields("e_date").Value = e_date
tb_t.Update

ssql = "Ctn_no = " & tb_pur.Fields("Ctn_no").Value
If Not tb_sale.EOF Then
    tb_sale.FindFirst ssql
    If tb_sale.NoMatch Then
    
    Else
tb_itm.FindFirst "Code = " & tb_sale.Fields("Item").Value
tb_w.FindFirst "Code = " & tb_sale.Fields("w_h_code").Value
tb_p.FindFirst "Code = " & tb_sale.Fields("Party").Value
        
        tb_t.AddNew
    
            tb_t.Fields("ItCode").Value = tb_sale.Fields("Item").Value
            tb_t.Fields("ItName").Value = tb_itm.Fields("Name").Value
            tb_t.Fields("V_date").Value = tb_sale.Fields("v_date").Value
            tb_t.Fields("Ctn_no").Value = tb_sale.Fields("Ctn_no").Value
            tb_t.Fields("Lot_no").Value = tb_sale.Fields("Lot_no").Value
            tb_t.Fields("Bobins").Value = tb_sale.Fields("Bobins").Value
            tb_t.Fields("Rate").Value = tb_sale.Fields("Rate").Value
            tb_t.Fields("Quantity").Value = tb_sale.Fields("Quantity").Value
            tb_t.Fields("Wh_Name").Value = tb_w.Fields("Name").Value
            tb_t.Fields("Code").Value = tb_sale.Fields("Party").Value
            tb_t.Fields("Name").Value = tb_p.Fields("Name").Value
            tb_t.Fields("e_type").Value = tb_sale.Fields("e_type").Value
            tb_t.Fields("trans_no").Value = tb_sale.Fields("sale_no").Value
            tb_t.Fields("S_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
        tb_t.Update

            
    End If
End If
tb_pur.MoveNext
cntl.Value = cntl.Value + 1
Loop
End If
tb_pur.Close
tb_sale.Close
tb_itm.Close
tb_p.Close
tb_w.Close
tb_t.Close
db_m.Close
db_t.Close
End Sub

Public Sub Dye_STock(V_DATE As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb2 As Recordset
Set blm = New bloom1
Dim tb_t As Recordset
Dim ss As Currency
Dim sc As Long
Dim rc As Long
Dim qs As Currency
Dim qr As Currency
Dim ssql As String
Dim tb_itm As Recordset
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from Stock"
db_t.Execute ssql
cntl.Value = 0
Set tb_t = db_t.OpenRecordset("Stock", dbOpenTable)
ssql = "select * from item order by code"
Set tb_itm = db_m.OpenRecordset(ssql)
ssql = "select party,item,count(ctn_no)as c,sum(quantity)as q from sent_view_1 group by party,item"
Set tb = db_m.OpenRecordset(ssql)
ssql = "select party,raw_item,count(ctn_no)as c,sum(quantity) as q from rec_view_final group by party,Raw_item"

Set tb2 = db_m.OpenRecordset(ssql)
        
If Not tb.EOF Then
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
    sc = 0
    rc = 0
            If Not tb2.EOF Then
                tb2.FindFirst "Party = " & tb.Fields("Party").Value & " and Raw_Item = " & tb.Fields("Item").Value
                If tb2.NoMatch Then
                
                Else
                    qr = tb2.Fields("q").Value
                    rc = tb2.Fields("C").Value
                End If
            End If

        qs = tb.Fields("q").Value
        sc = tb.Fields("C").Value
        ss = qs - qr
        If Not tb_itm.EOF Then
            tb_itm.FindFirst "COde = " & tb.Fields("Item").Value
            If tb_itm.NoMatch Then
                it$ = ""
            Else
                it$ = tb_itm.Fields("Name").Value
            End If
        End If
        tb_t.AddNew
            tb_t.Fields("e_date").Value = V_DATE
            tb_t.Fields("code").Value = tb.Fields("item").Value
            tb_t.Fields("Name").Value = it$ 'blm.Item1(tb.Fields("Item").Value)
            tb_t.Fields("quantity").Value = ss
            tb_t.Fields("Accode").Value = tb.Fields("party").Value
            tb_t.Fields("acname").Value = blm.party1(tb.Fields("Party").Value)
            tb_t.Fields("Ctns").Value = sc - rc
        tb_t.Update
        tb.MoveNext
        cntl.Value = cntl.Value + 1
    Loop
End If
tb.Close
        tb2.Close
tb_t.Close
db_t.Close
db_m.Close
End Sub
Public Sub Count_STock(V_DATE As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_in As Recordset
Dim tb_ot As Recordset
Set blm = New bloom1
Dim tb_t As Recordset
Dim ss As Currency
Dim pc As Long
Dim rc As Long
Dim ssql As String
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from Item"
db_t.Execute ssql
cntl.Value = 0
Set tb_t = db_t.OpenRecordset("Item", dbOpenTable)
ssql = "select * from item_vw order by code"

Set tb = db_m.OpenRecordset(ssql)
ssql = "select item,count(ctn_no) as c,sum(quantity) as q,avg(rate) as R from ctns where v_date <= #" & V_DATE & "# group by item"
Set tb_in = db_m.OpenRecordset(ssql)
ssql = "select item,count(ctn_no) as c,sum(quantity) as q from ctns2 where v_date <= #" & V_DATE & "# group by item"
Set tb_out = db_m.OpenRecordset(ssql)

If Not tb.EOF Then
tb.MoveLast
cntl.Max = tb.RecordCount
tb.MoveFirst
    Do While Not tb.EOF
        'ss = blm.Stock(tb.Fields("Code").Value, v_date)
        ss = 0
        If Not tb_in.EOF Then
            tb_in.FindFirst "Item = " & tb.Fields("Code").Value
            If tb_in.NoMatch Then
            
            Else
                ss = tb_in.Fields("Q").Value
                pc = tb_in.Fields("C").Value
            End If
        End If
        If Not tb_out.EOF Then
            tb_out.FindFirst "Item = " & tb.Fields("Code").Value
            If tb_out.NoMatch Then
            
            Else
                ss = ss - tb_out.Fields("Q").Value
                rc = tb_out.Fields("C").Value
            End If
        End If
        tb_t.AddNew
            tb_t.Fields("Gr_COde").Value = tb.Fields("Group").Value
            tb_t.Fields("Gr_Name").Value = tb.Fields("b.name").Value
            tb_t.Fields("Item").Value = tb.Fields("a.Name").Value
            tb_t.Fields("Stock").Value = ss
            tb_t.Fields("Code").Value = tb.Fields("COde").Value
            tb_t.Fields("Unit").Value = tb.Fields("Unit").Value
            tb_t.Fields("e_date").Value = V_DATE
            tb_t.Fields("ctns").Value = pc - rc
        tb_t.Update
        tb.MoveNext
        cntl.Value = cntl.Value + 1
        pc = 0
        rc = 0
    Loop
End If
tb.Close
tb_in.Close
tb_out.Close
tb_t.Close
db_t.Close
db_m.Close

End Sub
Public Sub ITMSTOCK(V_DATE As Date, cntl As Control, ITM As Long)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_in As Recordset
Dim tb_ot As Recordset
Set blm = New bloom1
Dim tb_t As Recordset
Dim ss As Currency
Dim pc As Long
Dim rc As Long
Dim ssql As String
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from Item"
db_t.Execute ssql
cntl.Value = 0
Set tb_t = db_t.OpenRecordset("Item", dbOpenTable)
ssql = "select * from item_vw where code = " & ITM

Set tb = db_m.OpenRecordset(ssql)
ssql = "select item,ctn_no,Quantity,w_h_code from ctns where item = " & ITM & " and cTN_NO NOT IN (select ctn_no from ctns2 WHere item = " & ITM & " and v_date <= #" & V_DATE & "#)"

Set tb_in = db_m.OpenRecordset(ssql)
'MsgBox ssql
If Not tb_in.EOF Then
tb_in.MoveLast
cntl.Max = tb_in.RecordCount
tb_in.MoveFirst
    Do While Not tb_in.EOF
                tb_t.AddNew
                tb_t.Fields("Gr_COde").Value = tb.Fields("Group").Value
                tb_t.Fields("Gr_Name").Value = tb.Fields("b.name").Value
                tb_t.Fields("Item").Value = tb.Fields("a.Name").Value
                tb_t.Fields("Stock").Value = tb_in.Fields("Quantity").Value
                tb_t.Fields("Code").Value = tb.Fields("Code").Value
                tb_t.Fields("Unit").Value = tb.Fields("Unit").Value
                tb_t.Fields("e_date").Value = V_DATE
                tb_t.Fields("ctns").Value = tb_in.Fields("ctn_no").Value
                tb_t.Fields("W_H_Code").Value = tb_in.Fields("W_H_Code").Value
                tb_t.Fields("WHName").Value = blm.WareHouse(tb_in.Fields("W_H_Code").Value)
                tb_t.Update
            
        tb_in.MoveNext
        cntl.Value = cntl.Value + 1
        DoEvents
    Loop
End If
tb.Close
tb_in.Close

tb_t.Close
db_t.Close
db_m.Close

End Sub

Public Sub Count_STock1(V_DATE As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_in As Recordset
Dim tb_inf As Recordset
Dim tb_outf As Recordset
Dim tb_out As Recordset
Dim tb_wh As Recordset
Set blm = New bloom1
Dim tb_t As Recordset
Dim grc As Long
Dim grn As String
Dim ITM As String
Dim pc As Long
Dim rc As Long
Dim ss As Currency
Dim w As Integer
Dim wh As String
Dim ssql As String
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from Item"
db_t.Execute ssql
cntl.Value = 0
Set tb_t = db_t.OpenRecordset("Item", dbOpenTable)
ssql = "select * from item_vw order by code"

Set tb = db_m.OpenRecordset(ssql)

ssql = "select * from warehouse order by code"
Set tb_wh = db_m.OpenRecordset(ssql)
ssql = "select item,W_H_Code,count(ctn_no) as c,sum(quantity) as q,avg(rate) as R from ctns where v_date <= #" & V_DATE & "# group by item,W_H_Code"
Set tb_in = db_m.OpenRecordset(ssql)
ssql = "select item,W_H_Code,count(ctn_no) as c,sum(quantity) as q from ctns2 where v_date <= #" & V_DATE & "# group by item,W_H_Code"
Set tb_out = db_m.OpenRecordset(ssql)

If Not tb_in.EOF Then
tb_in.MoveLast
cntl.Max = tb_in.RecordCount
tb_in.MoveFirst
    Do While Not tb_in.EOF
    w = 0
        'ss = blm.Stock(tb.Fields("Code").Value, v_date)
        ss = 0
        
    
        
            w = 0
                ss = tb_in.Fields("Q").Value
                pc = tb_in.Fields("C").Value
                w = tb_in.Fields("W_H_COde").Value
    
        
        If Not tb_wh.EOF Then
            tb_wh.FindFirst "Code = " & w
            If tb_wh.NoMatch Then
            
            Else
                wh = tb_wh.Fields("Name").Value
            End If
        End If
        
        If Not tb.EOF Then
            tb.FindFirst "Code = " & tb_in.Fields("Item").Value
            If tb.NoMatch Then
            
            Else
                grc = tb.Fields("group").Value
                grn = tb.Fields("b.name").Value
                ITM = tb.Fields("a.name").Value
                u$ = tb.Fields("Unit").Value
            End If
        End If
        
        If Not tb_out.EOF Then
            tb_out.FindFirst "Item = " & tb_in.Fields("item").Value & " and W_H_Code = " & w
            
            If tb_out.NoMatch Then
            
            Else
                ss = ss - tb_out.Fields("Q").Value
                pc = pc - tb_out.Fields("C").Value
            End If
        End If
        tb_t.AddNew
            tb_t.Fields("Gr_COde").Value = grc
            tb_t.Fields("Gr_Name").Value = grn
            If tb.Fields("code").Value >= 40 And tb.Fields("COde").Value <= 60 Then
'            MsgBox tb.Fields("code").Value
            End If
            tb_t.Fields("Item").Value = ITM
            tb_t.Fields("Stock").Value = ss
            tb_t.Fields("Code").Value = tb_in.Fields("Item").Value
            tb_t.Fields("Unit").Value = u$
            tb_t.Fields("e_date").Value = V_DATE
            tb_t.Fields("W_H_Code").Value = tb_in.Fields("W_H_Code").Value
            tb_t.Fields("WHName").Value = wh
            tb_t.Fields("Ctns").Value = pc
        tb_t.Update
        tb_in.MoveNext
        cntl.Value = cntl.Value + 1
        pc = 0
        rc = 0
    Loop
End If
tb.Close
tb_in.Close
tb_out.Close
tb_t.Close
db_t.Close
db_m.Close

End Sub

Public Sub Prices(V_DATE As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim p1 As Single
Dim p2 As Single
Dim tb_itm As Recordset
Dim p_f As Single
Dim tb_t As Recordset
Dim tb2 As Recordset
Dim ssql As String
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from trial"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
ssql = "select * from item order by code"
Set tb_itm = db_m.OpenRecordset(ssql)

ssql = "select item,avg(rate)As p1 from in_dtl where v_date <= #" & V_DATE & "# group by Item"
Set tb = db_m.OpenRecordset(ssql)
ssql = "select item,avg(rate) As p1,avg(dye_rate)as d1 from rec_2 where v_date <= #" & V_DATE & "# and item = " & tb.Fields("item").Value & " Group by Item"
Set tb2 = db_m.OpenRecordset(ssql)
        
cntl.Value = 0
If Not tb.EOF Then
tb.MoveLast
cntl.Max = tb.RecordCount
tb.MoveFirst
    Do While Not tb.EOF
        p1 = tb.Fields("p1").Value
        p_f = p1
        If Not tb2.EOF Then
            tb2.FindFirst "Item = " & tb.Fields("Item").Value
            If tb2.NoMatch Then
                p_f = p1
            Else
                p_f = (p1 + tb2.Fields("p1").Value) / 2
            End If
        End If
        If Not tb_itm.EOF Then
            tb_itm.FindFirst "Code = " & tb.Fields("Item").Value
            If tb_itm.NoMatch Then
                it$ = ""
            Else
                it$ = tb_itm.Fields("Name").Value
            End If
        End If
        tb_t.AddNew
            tb_t.Fields("e_date").Value = V_DATE
            tb_t.Fields("name").Value = it$ 'blm.Item1(tb.Fields("item").Value)
            tb_t.Fields("debit").Value = p_f
        tb_t.Update
        
        tb.MoveNext
        cntl.Value = cntl.Value + 1
    Loop
End If
tb.Close
tb2.Close
tb_itm.Close
tb_t.Close
db_t.Close
db_m.Close
End Sub
Public Sub STock_Price(V_DATE As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb_itm As Recordset
Dim tb_in As Recordset
Dim tb_out As Recordset
Dim ITM As String
Dim p1 As Single
Dim p2 As Single
Dim p_f As Single
Dim stk As Currency
Dim tb_t As Recordset
Dim tb2 As Recordset
Dim ssql As String
'Set blm = New bloom1
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from item"
db_t.Execute ssql
ssql = "select * from item order by Code"
Set tb_itm = db_m.OpenRecordset(ssql)
ssql = "select item,sum(quantity) as q,avg(rate) as R from ctns where v_date <= #" & V_DATE & "# group by item"
Set tb_in = db_m.OpenRecordset(ssql)
ssql = "select item,sum(quantity) as q from ctns2 where v_date <= #" & V_DATE & "# group by item"
Set tb_out = db_m.OpenRecordset(ssql)

Set tb_t = db_t.OpenRecordset("item", dbOpenTable)
cntl.Value = 0
If Not tb_in.EOF Then
tb_in.MoveLast
cntl.Max = tb_in.RecordCount
tb_in.MoveFirst
     Do While Not tb_in.EOF
        ITM = ""
        If Not tb_itm.EOF Then
            tb_itm.FindFirst "Code = " & tb_in.Fields("Item").Value
            If tb_itm.NoMatch Then
            
            Else
                ITM = tb_itm.Fields("Name").Value & "" & tb_itm.Fields("Unit").Value
            End If
        End If
        If Not tb_out.EOF Then
            tb_out.FindFirst "Item = " & tb_in.Fields("Item")
            If tb_out.NoMatch Then
                stk = tb_in.Fields("Q").Value
                p_f = tb_in.Fields("R").Value
            Else
                stk = tb_in.Fields("Q").Value - tb_out.Fields("Q").Value
                p_f = tb_in.Fields("R").Value
            End If
        End If
        tb_t.AddNew
            tb_t.Fields("e_date").Value = V_DATE
            tb_t.Fields("Item").Value = ITM

            tb_t.Fields("Price").Value = p_f
            tb_t.Fields("Stock").Value = stk
        tb_t.Update
        cntl.Value = cntl.Value + 1
        tb_in.MoveNext
    Loop
End If
tb_in.Close
tb_out.Close
tb_itm.Close
tb_t.Close
db_t.Close
db_m.Close

End Sub



Public Property Get report_path() As String
    report_path = App.path & "\"
'report_path = CurDir & "\"
End Property
Public Property Get path() As String
    path = t_path
End Property
Public Property Get Userpath() As String
    Userpath = "c:\soft\Alpha\BLOOMNEW.mdb"
    
End Property


Public Sub trial(e_date As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim ssql As String
Dim wrkjet As Workspace
cntl.Value = 0
Set wrkjet = CreateWorkspace("", "admin", "", dbUseJet)

Set db_m = wrkjet.OpenDatabase(blm.pathMain)
Set db_t = wrkjet.OpenDatabase(t_path)

ssql = "delete from trial"
db_t.Execute ssql, dbSQLPassThrough
db_t.Close
Set db_t = wrkjet.OpenDatabase(t_path, , False)
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
'tb_t.AddNew
'    tb_t.Fields("e_date").Value = e_date
'    tb_t.Fields("code").Value = 10000
'    tb_t.Fields("name").Value = "Cash In Hand"
   ' MsgBox Format(e_date, "dd/MM/yyyy")
'    tb_t.Fields("Debit").Value = blm.cash2(e_date)
'    tb_t.Fields("credit").Value = 0
'tb_t.Update

'sales
    ssql = "select sum(debit) as sales from voucher where e_type = 1 and v_date <= #" & e_date & "#"
    Set tb = db_m.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("sales").Value) Then
        tb_t.AddNew
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("code").Value = 6001
            tb_t.Fields("Name").Value = "Sales Upto " & Format(e_date, "dd/MM/yyyy")
           tb_t.Fields("debit").Value = 0#
           tb_t.Fields("credit").Value = tb.Fields("sales").Value
            
        tb_t.Update
    End If
tb.Close

'purchases
    ssql = "select sum(credit) as purchases from voucher where e_type = 2 and v_date <= #" & e_date & "#"
    Set tb = db_m.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("purchases").Value) Then
        tb_t.AddNew
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("code").Value = 6002
            tb_t.Fields("Name").Value = "Purchases Upto " & Format(e_date, "dd/MM/yyyy")
            tb_t.Fields("debit").Value = tb.Fields("purchases").Value
            tb_t.Fields("credit").Value = 0
            
        tb_t.Update
    End If
 tb.Close
'recieved After Dying
   ssql = "select sum(credit) as Recieved from voucher where e_type = 4 and v_date <= #" & e_date & "#"
    Set tb = db_m.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("Recieved").Value) Then
   
            tb_t.AddNew
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("code").Value = 6003
            tb_t.Fields("name").Value = "Recieved After Dying UpTo " & Format(e_date, "dd/MM/yyyy")
            tb_t.Fields("debit").Value = tb.Fields("Recieved").Value
            tb_t.Fields("credit").Value = 0
            
        tb_t.Update
    
    End If
    tb.Close
'sent for Dying
    ssql = "select sum(debit) as sent from voucher where e_type = 3 and v_date <= #" & e_date & "#"
    Set tb = db_m.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("sent").Value) Then
        tb_t.AddNew
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("code").Value = 6004
           tb_t.Fields("name").Value = "Sent For Dying To " & Format(e_date, "dd/MM/yyyy")
            tb_t.Fields("debit").Value = 0
            tb_t.Fields("credit").Value = tb.Fields("Sent").Value
  
       tb_t.Update
   
    End If
    tb.Close
    
'Vouchers
Dim tb_p As Recordset
    ssql = "Select * from acchart order by code"
    Set tb_p = db_m.OpenRecordset(ssql)
    ssql = "select party,sum(debit-credit) as bal from voucher where party < 5000   and v_date <= #" & e_date & "# group by party Having Sum(debit - credit)<> 0"
    Set tb = db_m.OpenRecordset(ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
        If Not tb_p.EOF Then
            tb_p.FindFirst "Code = " & tb.Fields("Party").Value
            If tb_p.NoMatch Then
                p$ = ""
            Else
                p$ = tb_p.Fields("Name").Value
            End If
        End If
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("name").Value = p$ 'blm.party1(tb.Fields("party").Value)
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("debit").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("debit").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("credit").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("credit").Value = 0
        End If
    tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    tb_p.Close
    tb.Close

tb_t.Close
db_t.Close
db_m.Close
wrkjet.Close

End Sub
Public Sub day_book(V_DATE As Date, cntl As Control)
'On Error Resume Next
Dim db As Database
    Dim tb As Recordset
    Dim ssql As String
    Dim t_sal As Currency
    Dim t_pur As Currency
    Dim dif_sp As Currency
    Dim t_t_r As Currency
    Dim t_t_s As Currency
    Set blm = Nothing
    Set blm = New bloom1
    Dim dbt As Database
    Dim tbt As Recordset
    Set db = OpenDatabase(blm.pathMain)
    Set tb = db.OpenRecordset("Select * from vou_view where v_date = #" & V_DATE & "#")
    Set dbt = OpenDatabase(report_path & "book.mdb")
    dbt.Execute "delete from vou_temp"
    dbt.Close
    Set dbt = OpenDatabase(report_path & "book.mdb")
    Set tbt = dbt.OpenRecordset("vou_temp", dbOpenTable)
    cntl.Value = 0
    
    If tb.EOF = False Then
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
        Do While Not tb.EOF
            tbt.AddNew
                tbt.Fields("V_dATE").Value = V_DATE
                tbt.Fields("E_TYPE").Value = tb.Fields("E_TYPE").Value
                tbt.Fields("PRE_CASH").Value = tb.Fields("PRE_CASH").Value
                tbt.Fields("CODE").Value = tb.Fields("PARTY").Value
                tbt.Fields("NAME").Value = blm.party1(tb.Fields("PARTY").Value)
                tbt.Fields("DEBIT").Value = tb.Fields("DEBIT").Value
                tbt.Fields("CREDIT").Value = tb.Fields("CREDIT").Value
                tbt.Fields("REMARKS").Value = tb.Fields("REMARKS").Value
            tbt.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
        Loop
    Else
        MsgBox "NO RECORD FOR THIS DATE"
    End If
    tbt.Close
    tb.Close
    dbt.Close
    db.Close
    End Sub
Public Sub day_outstand(V_DATE As Date)
Dim db As Database
    Dim tb As Recordset
    Dim ssql As String
    Set blm = Nothing
    Set blm = New bloom1
    Dim dbt As Database
    Dim tbt As Recordset
    Set db = OpenDatabase(blm.pathMain)
    Set tb = db.OpenRecordset("Select * from sale where v_date <= #" & V_DATE - 1 & "#")
    Set dbt = OpenDatabase(report_path & "book.mdb")
    dbt.Execute ("delete from s_P")
    Set tbt = dbt.OpenRecordset("s_p", dbOpenTable)
    
    
    If tb.EOF = False Then
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
        Do While Not tb.EOF
            tbt.AddNew
                tbt.Fields("V_dATE").Value = V_DATE
                tbt.Fields("E_TYPE").Value = tb.Fields("E_TYPE").Value
                tbt.Fields("PRE_CASH").Value = tb.Fields("PRE_CASH").Value
                tbt.Fields("CODE").Value = tb.Fields("PARTY").Value
                tbt.Fields("NAME").Value = blm.party1(tb.Fields("PARTY").Value)
                tbt.Fields("DEBIT").Value = tb.Fields("DEBIT").Value
                tbt.Fields("CREDIT").Value = tb.Fields("CREDIT").Value
                tbt.Fields("REMARKS").Value = tb.Fields("REMARKS").Value
            tbt.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
        Loop
    Else
        MsgBox "NO RECORD FOR THIS DATE"
    End If
    tbt.Close
    tb.Close
    dbt.Close
    db.Close
    End Sub

Public Sub ledger(code As Long, cntl As Control)
'On Error Resume Next
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_in As Recordset
Dim tb_out As Recordset
Dim tb2 As Recordset
Dim op_bal As Currency
Dim tb_p As Recordset
Dim ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from ledger"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
ssql = "select * from acchart order by code"
Set tb_p = db_m.OpenRecordset(ssql)

ssql = "select * from in_mst order by p_no"
Set tb_in = db_m.OpenRecordset(ssql)

ssql = "select * from Sale_1 order by Sale_no"
Set tb_out = db_m.OpenRecordset(ssql)

ssql = "select * from voucher where party = " & code & " order by v_date ASC, ent_no DESC"
Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
        If Not tb_p.EOF Then
            tb_p.FindFirst "Code = " & tb.Fields("Party").Value
            If tb_p.NoMatch Then
                it$ = ""
            Else
                it$ = tb_p.Fields("Name").Value
            End If
        End If
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            tb_t.Fields("remarks").Value = tb.Fields("remarks").Value & " "
            Select Case tb.Fields("e_type").Value
                    Case 1
                    '    ssql = "select * from sale_1 where sale_no = " & tb.Fields("ent_no").Value
                        ssql = "Sale Invoice # " & tb.Fields("ent_no").Value
                        If Not tb_out.EOF Then
                            tb_out.FindFirst "Sale_no = " & tb.Fields("ent_no").Value
                            If tb_out.NoMatch Then
                                
                            Else
                                ssql = " Goods / " & tb_out.Fields("Goods").Value & " "
                            End If
                        End If
                    Case 2
                    '    ssql = "select * from in_mst where p_no = " & tb.Fields("ent_no").Value
                        ssql = "Inward # "
                        If Not tb_in.EOF Then
                            tb_in.FindFirst "P_no = " & tb.Fields("ent_no").Value
                            If tb_in.NoMatch Then
                                
                            Else
                                ssql = " Goods / " & tb_in.Fields("Goods").Value & " "
                     '           MsgBox ssql
                            End If
                        End If
                    Case 3
                    '    ssql = "select * from sent_1 where issue_no = " & tb.Fields("ent_no").Value
                        ssql = "Sent For Dying Note # "
                    Case 4
                    '    ssql = "select * from rec_1 where rec_no = " & tb.Fields("ent_no").Value
                        ssql = "Rec. After Dying Note # "
                    
            End Select
            If tb.Fields("e_type").Value >= 1 And tb.Fields("e_type").Value <= 4 Then
            
                    'MsgBox ssql
                        tb_t.Fields("remarks").Value = tb.Fields("Remarks").Value & " " & ssql & " "
                    'MsgBox tb_t.Fields("remarks").Value
            End If
            
            'If Not IsMissing(s_date) Then
             '   tb_t.Fields("s_date").Value = s_date
             '   tb_t.Fields("e_date").Value = e_date
           ' End If
        tb_t.Update
    tb.MoveNext
    cntl.Value = cntl.Value + 1
    Loop
End If
    tb.Close
    tb_p.Close
    tb_in.Close
    tb_out.Close
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub
Public Sub ledger2(code As Long, cntl As Control, s_date As Date, e_date As Date)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb2 As Recordset
Dim tb_t As Recordset
Dim tb_p As Recordset
Dim op_bal As Currency
Dim ssql As String
Dim ssql2 As String
Dim ss2 As String
Dim ss As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from ledger"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
ssql = "select * from acchart order by code"
Set tb_p = db_m.OpenRecordset(ssql)
'If Not IsMissing(s_date) Then
    ssql = "select sum(debit-credit) as bal from voucher where party = " & code & " and v_date < #" & s_date & "#"
    Set tb = db_m.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("bal").Value) Then
    If Not tb_p.EOF Then
            tb_p.FindFirst "Code = " & code
            If tb_p.NoMatch Then
                it$ = ""
            Else
                it$ = tb_p.Fields("Name").Value
            End If
        End If
    tb_t.AddNew
        op_bal = tb.Fields("bal").Value
        If op_bal < 0 Then
            tb_t.Fields("credit").Value = op_bal * -1
            tb_t.Fields("DEBIT").Value = 0
        End If
        If op_bal > 0 Then
            tb_t.Fields("debit").Value = op_bal
            tb_t.Fields("CREDIT").Value = 0
        End If
            'tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = s_date
            tb_t.Fields("code").Value = code 'tb.Fields("party").Value
            tb_t.Fields("description").Value = it$ 'blm.party1(code)
            tb_t.Fields("remarks").Value = "Opening Balance On .." & Format(s_date, "dd/MM/yyyy")
     tb_t.Update
    End If
    tb.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    ssql = "select * from voucher where party = " & code & " and v_Date >= #" & s_date & "#" & " and v_date <= #" & e_date & "# order by v_date ASC, ent_no DESC"
'End If
'MsgBox ssql

Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
        ss2 = ""
        If Not tb_p.EOF Then
            tb_p.FindFirst "Code = " & tb.Fields("Party").Value
            If tb_p.NoMatch Then
                it$ = ""
            Else
                it$ = tb_p.Fields("Name").Value
            End If
        End If
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            tb_t.Fields("remarks").Value = tb.Fields("remarks").Value & " "
'            MsgBox tb.Fields("e_type").Value
                    
                    If tb.Fields("e_type").Value = 1 Then
                        ssql = "select item,sum(quantity)as q,count(ctn_no)as c,avg(rate)as r from sale_2 where sale_no = " & tb.Fields("ent_no").Value & " group by item"
                        ssql2 = "select * from sale_1 where sale_no = " & tb.Fields("ent_no").Value
                        ss = "Sale # "
                    End If
                    If tb.Fields("e_type").Value = 2 Then
                        ssql = "select item,sum(quantity)as q,count(ctn_no)as c,avg(rate)as r from in_dtl where p_no = " & tb.Fields("ent_no").Value & " group by item"
                        ssql2 = "select * from in_Mst where p_no = " & tb.Fields("ent_no").Value
                        ss = "Inward # "
                    End If
                    If tb.Fields("e_type").Value = 3 Then
                        ssql = "select item,sum(quantity)as q,count(ctn_no)as c,avg(rate)as r from sent_2 where issue_no = " & tb.Fields("ent_no").Value & " group by item"
                        ssql2 = "select * from sent_1 where issue_no = " & tb.Fields("ent_no").Value
                        ss = "Issue # "
                    End If
                    If tb.Fields("e_type").Value = 4 Then
                        ssql = "select item,sum(quantity)as q,count(ctn_no)as c,avg(rate)as r ,avg(dye_rate) as DR from rec_2 where rec_no = " & tb.Fields("ent_no").Value & " group by item"
                        ssql2 = "select * from rec_1 where rec_no = " & tb.Fields("ent_no").Value
                        ss = "Rec. # "
                    End If
                    If tb.Fields("e_type").Value >= 1 And tb.Fields("e_type").Value <= 4 Then
'                        MsgBox ssql
                        Set tb2 = db_m.OpenRecordset(ssql2)
                  
                        If Not tb2.EOF Then
                        
                        If tb.Fields("E_Type").Value = 3 Then
                            tb_t.Fields("Bill_no").Value = tb2.Fields("Bill_no").Value
                            ss2 = ss2 & tb2.Fields("Note").Value
                        Else
                            If tb.Fields("E_Type").Value <= 2 Then
                                If Len(tb2.Fields("Goods").Value) > 0 Then
                                    ss2 = ss2 & "Goods = " & tb2.Fields("Goods").Value & vbCrLf & tb2.Fields("Note").Value
                                End If
                                If tb.Fields("E_Type").Value = 2 Then
                                    
                                    tb_t.Fields("Bill_no").Value = tb2.Fields("Bilty_no").Value
                                End If
                                If tb.Fields("E_Type").Value = 1 Then
                                    tb_t.Fields("Bill_no").Value = tb2.Fields("Bilty_no").Value
                                End If
                            Else
                                If tb.Fields("E_Type").Value = 4 Then
                                    ss2 = ss2 & vbCrLf & tb2.Fields("Note").Value
                                End If
                            End If
                   
                        End If
                        End If
                        tb2.Close
                        Set tb2 = db_m.OpenRecordset(ssql)
 '                           MsgBox tb2.EOF
                            If Not tb2.EOF Then
                            ssql = " "
                                tb_t.Fields("code").Value = tb.Fields("party").Value
                                    tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
                                    tb_t.Fields("Item").Value = blm.Item1(tb2.Fields("item").Value)
                                    tb_t.Fields("Quantity").Value = tb2.Fields("q").Value
                                    tb_t.Fields("Rate").Value = tb2.Fields("R").Value & ""
                                    tb_t.Fields("Ctn_no").Value = tb2.Fields("c").Value
                                   
                                    If tb.Fields("E_Type").Value = 4 Then
                                        tb_t.Fields("Dye_Rate").Value = tb2.Fields("DR").Value
                                    End If

                                tb_t.Fields("remarks").Value = ss & " " & tb.Fields("ent_no").Value & ", " & ss2 & vbCrLf & vbCrLf & ssql
                            Else
                                tb_t.Fields("remarks").Value = tb.Fields("remarks").Value & " "
                                
                            End If
                            tb_t.Fields("s_date").Value = s_date
                            tb_t.Fields("e_date").Value = e_date
                            tb_t.Update
                            If Not tb2.EOF Then
                            tb2.MoveNext
                            Do While Not tb2.EOF
                                    tb_t.AddNew
                                    tb_t.Fields("op_bal").Value = op_bal
                                    tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                                    tb_t.Fields("code").Value = tb.Fields("party").Value
                                    tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
                                    tb_t.Fields("Item").Value = blm.Item1(tb2.Fields("item").Value)
                                    tb_t.Fields("Quantity").Value = tb2.Fields("q").Value
                                    tb_t.Fields("Rate").Value = tb2.Fields("R").Value & ""
                                    tb_t.Fields("Ctn_no").Value = tb2.Fields("c").Value
                                   
                                    If tb.Fields("E_Type").Value = 4 Then
                                        tb_t.Fields("Dye_Rate").Value = tb2.Fields("DR").Value
                                    End If
                        
'                                    MsgBox ssql
                                    tb2.MoveNext
                                    tb_t.Update
                                    
                            Loop
                            End If
                        tb2.Close
                    Else
                        tb_t.Update 'right
                    End If
            

            'If Not IsMissing(s_date) Then
                
           ' End If
        
    tb.MoveNext
    cntl.Value = cntl.Value + 1
    Loop
End If
    tb.Close
    tb_p.Close
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub
Public Sub ledger3(code As Long, cntl As Control, s_date As Date, e_date As Date)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb2 As Recordset
Dim tb_t As Recordset
Dim tb_p As Recordset
Dim op_bal As Currency
Dim ssql As String
Dim ssql2 As String
Dim tb_in As Recordset
Dim tb_in_f As Recordset
Dim tb_out As Recordset
Dim tb_out_f As Recordset
Dim tb_sent As Recordset
Dim tb_sent_f As Recordset
Dim tb_rec As Recordset
Dim tb_rec_f As Recordset
Dim tb_itm As Recordset
Dim ss2 As String
Dim ss As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from ledger"
db_t.Execute ssql


ssql = "select a.p_no,a.v_date,a.gate_no,a.note,a.bilty_no,a.goods,a.party,b.item,b.ctn_no,b.lot_no,b.w_h_code,b.bobins,b.quantity,b.rate from in_mst a,in_dtl b where a.p_no=b.p_no"
Set tb_in = db_m.OpenRecordset(ssql)

ssql = "select a.sale_no,a.v_date,a.gate_no,a.note,a.bilty_no,a.goods,a.party,b.item,b.ctn_no,b.lot_no,b.w_h_code,b.bobins,b.quantity,b.rate from sale_1 a,sale_2 b where a.sale_no=b.sale_no"
Set tb_out = db_m.OpenRecordset(ssql)

ssql = "select a.Issue_no,a.v_date,a.gate_no,a.note,a.bill_no,a.advice_no,a.party,b.item,b.ctn_no,b.lot_no,b.w_h_code,b.bobins,b.quantity,b.rate from sent_1 a,sent_2 b where a.issue_no=b.issue_no"
Set tb_sent = db_m.OpenRecordset(ssql)

ssql = "select a.Rec_no,a.v_date,a.gate_no,a.note,a.party,b.ctn_no,b.lot_no,b.w_h_code,b.item,b.bobins,b.quantity,b.rate,b.dye_rate,b.raw_item from rec_1 a,rec_2 b where a.rec_no=b.rec_no"
Set tb_rec = db_m.OpenRecordset(ssql)


ssql = "select * from acchart order by code"
Set tb_p = db_m.OpenRecordset(ssql)

ssql = "select * from Item order by code"
Set tb_itm = db_m.OpenRecordset(ssql)

Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
'If Not IsMissing(s_date) Then
    ssql = "select sum(debit-credit) as bal from voucher where party = " & code & " and v_date < #" & s_date & "#"
    Set tb = db_m.OpenRecordset(ssql)
    If Not IsNull(tb.Fields("bal").Value) Then
    If Not tb_p.EOF Then
            tb_p.FindFirst "Code = " & code
            If tb_p.NoMatch Then
                it$ = ""
            Else
                it$ = tb_p.Fields("Name").Value
            End If
        End If
    tb_t.AddNew
        op_bal = tb.Fields("bal").Value
        If op_bal < 0 Then
            tb_t.Fields("credit").Value = op_bal * -1
            tb_t.Fields("DEBIT").Value = 0
        End If
        If op_bal > 0 Then
            tb_t.Fields("debit").Value = op_bal
            tb_t.Fields("CREDIT").Value = 0
        End If
            'tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = s_date
            tb_t.Fields("code").Value = code 'tb.Fields("party").Value
            tb_t.Fields("description").Value = it$ 'blm.party1(code)
            tb_t.Fields("remarks").Value = "Opening Balance On .." & Format(s_date, "dd/MM/yyyy")
     tb_t.Update
    End If
    tb.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    ssql = "select * from voucher where party = " & code & " and v_Date >= #" & s_date & "#" & " and v_date <= #" & e_date & "# order by v_date ASC, ent_no DESC"
'End If
'MsgBox ssql

Set tb = db_m.OpenRecordset(ssql)
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
    ss2 = ""
    If Not tb_p.EOF Then
            tb_p.FindFirst "Code = " & tb.Fields("Party").Value
            If tb_p.NoMatch Then
                it$ = ""
            Else
                it$ = tb_p.Fields("Name").Value
            End If
    End If
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            'tb_t.Fields("remarks").Value = tb.Fields("remarks").Value & " "
'            MsgBox tb.Fields("e_type").Value
                    ssql = ""
                    If tb.Fields("e_type").Value = 1 Then
                        If Not tb_out.EOF Then
                            tb_out.Filter = "Sale_no = " & tb.Fields("ent_no").Value
                            Set tb_out_f = tb_out.OpenRecordset()
                        End If
                        ss = "Sale # "
                    End If
                    If tb.Fields("e_type").Value = 2 Then
                        If Not tb_in.EOF Then
                            tb_in.Filter = "p_no = " & tb.Fields("ent_no").Value
                              Set tb_in_f = tb_in.OpenRecordset()
                        End If
                        
                        ss = "Inward # "
                    End If
                    If tb.Fields("e_type").Value = 3 Then
                        If Not tb_sent.EOF Then
                            tb_sent.Filter = "Issue_no = " & tb.Fields("ent_no").Value
                              Set tb_sent_f = tb_sent.OpenRecordset()
                        End If
                        ss = "Issue # "
                    End If
                    If tb.Fields("e_type").Value = 4 Then
                        If Not tb_rec.EOF Then
                            tb_rec.Filter = "Rec_no = " & tb.Fields("ent_no").Value
                              Set tb_rec_f = tb_rec.OpenRecordset()
                        End If
                        ss = "Rec. # "
                    End If
                    If tb.Fields("e_type").Value >= 1 And tb.Fields("e_type").Value <= 4 Then
                    '    Set tb2 = db_m.OpenRecordset(ssql2)
                  
                        
                        
                        If tb.Fields("E_Type").Value = 3 Then
                            If Not tb_sent_f.EOF Then
                            tb_t.Fields("bill_no").Value = tb_sent_f.Fields("Bill_no").Value
                            End If
                        End If
                        If tb.Fields("E_Type").Value = 1 Then
                            If Not tb_out_f.EOF Then
                            tb_t.Fields("bill_no").Value = tb_out_f.Fields("Bilty_no").Value
                            End If
                        End If
                            If tb.Fields("E_Type").Value = 2 Then
                                If Not tb_in_f.EOF Then
                                If Len(tb_in_f.Fields("Goods").Value) > 0 Then
                                    ss2 = ss2 & "Goods = " & tb_in_f.Fields("Goods").Value & vbCrLf & tb_in_f.Fields("Note").Value
                                End If
                                End If
                            End If
                            If tb.Fields("E_Type").Value = 1 Then
                                If Not tb_out_f.EOF Then
                                If Len(tb_out_f.Fields("Goods").Value) > 0 Then
                                    ss2 = ss2 & "Goods = " & tb_out_f.Fields("Goods").Value & vbCrLf & tb_out_f.Fields("Note").Value
                                End If
                                
                                End If
                            End If
                            If tb.Fields("E_Type").Value = 4 Then
                                If Not tb_rec_f.EOF Then
                                      ss2 = ss2 & vbCrLf & tb_rec_f.Fields("Note").Value
                                  End If
                            End If
                          
                        
                    End If
                     tb_t.Fields("remarks").Value = tb.Fields("Remarks").Value & vbCrLf
                    If E_type >= 1 And E_type <= 4 Then
                         tb_t.Fields("remarks").Value = tb_t.Fields("remarks").Value & ss & tb.Fields("ent_no").Value & vbCrLf & ss2 & vbCrLf & ssql
                    End If
                    tb_t.Fields("s_date").Value = s_date
                    tb_t.Fields("e_date").Value = e_date
           ' End If
                    tb_t.Update
'
                    
 '
                    
                                ssql = ""
                              
                                    If tb.Fields("E_Type").Value = 1 Then
                                    If Not tb_out_f.EOF Then
                                    Do While Not tb_out_f.EOF
                                    tb_t.AddNew
                                    If Not tb_itm.EOF Then
                                        tb_itm.FindFirst "Code = " & tb_out_f.Fields("Item").Value
                                        If tb_itm.NoMatch Then
                                            ITM$ = ""
                                        Else
                                            ITM$ = tb_itm.Fields("Name").Value
                                        End If
                                    End If
                                        tb_t.Fields("op_bal").Value = op_bal
                                        tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                                        tb_t.Fields("code").Value = tb.Fields("party").Value
                                        tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
            
                                        tb_t.Fields("Item").Value = ITM$
                                        tb_t.Fields("Quantity").Value = tb_out_f.Fields("quantity").Value
                                        tb_t.Fields("Rate").Value = tb_out_f.Fields("Rate").Value
                                        tb_t.Fields("Ctn_no").Value = tb_out_f.Fields("ctn_no").Value
                                        tb_out_f.MoveNext
                                    tb_t.Update
                                    Loop
                                    End If
                                    End If
                                    
                                    
                                    If tb.Fields("E_Type").Value = 2 Then
                                    If Not tb_in_f.EOF Then
                                     Do While Not tb_in_f.EOF
                                        tb_t.AddNew
                                    If Not tb_itm.EOF Then
                                    
                                        tb_itm.FindFirst "Code = " & tb_in_f.Fields("Item").Value
                                        If tb_itm.NoMatch Then
                                            ITM$ = ""
                                        Else
                                            ITM$ = tb_itm.Fields("Name").Value
                                        End If
                                    End If
                                        tb_t.Fields("op_bal").Value = op_bal
                                        tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                                        tb_t.Fields("code").Value = tb.Fields("party").Value
                                        tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
                
                                        tb_t.Fields("Item").Value = ITM$
                                        tb_t.Fields("Quantity").Value = tb_in_f.Fields("quantity").Value
                                        tb_t.Fields("Rate").Value = tb_in_f.Fields("Rate").Value
                                        tb_t.Fields("Ctn_no").Value = tb_in_f.Fields("ctn_no").Value
                                        tb_in_f.MoveNext
                                        tb_t.Update
                                        Loop
                                    End If
                                    End If
                                   
                                    If tb.Fields("E_Type").Value = 3 Then
                                    If Not tb_sent_f.EOF Then
                                     Do While Not tb_sent_f.EOF
                                        tb_t.AddNew
                                    If Not tb_itm.EOF Then
                                        tb_itm.FindFirst "Code = " & tb_sent_f.Fields("Item").Value
                                        If tb_itm.NoMatch Then
                                            ITM$ = ""
                                        Else
                                            ITM$ = tb_itm.Fields("Name").Value
                                        End If
                                    End If
                                        tb_t.Fields("op_bal").Value = op_bal
                                        tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                                        tb_t.Fields("code").Value = tb.Fields("party").Value
                                        tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
            
                                        tb_t.Fields("Item").Value = ITM$
                                        tb_t.Fields("Quantity").Value = tb_sent_f.Fields("quantity").Value
                                        tb_t.Fields("Rate").Value = tb_sent_f.Fields("Rate").Value
                                        tb_t.Fields("Ctn_no").Value = tb_sent_f.Fields("ctn_no").Value
                                        
                                        tb_sent_f.MoveNext
                                        tb_t.Update
                                        Loop
                                    End If
                                    End If
                                                                      
                                    If tb.Fields("E_Type").Value = 4 Then
                                    If Not tb_rec_f.EOF Then
                                     Do While Not tb_rec_f.EOF
                                     tb_t.AddNew
                                    If Not tb_itm.EOF Then
                                        tb_itm.FindFirst "Code = " & tb_rec_f.Fields("Item").Value
                                        If tb_itm.NoMatch Then
                                            ITM$ = ""
                                        Else
                                            ITM$ = tb_itm.Fields("Name").Value
                                        End If
                                    End If
                                        tb_t.Fields("op_bal").Value = op_bal
                                        tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                                        tb_t.Fields("code").Value = tb.Fields("party").Value
                                        tb_t.Fields("description").Value = it$ 'blm.party1(tb.Fields("party").Value)
            
                                        tb_t.Fields("Item").Value = ITM$
                                        tb_t.Fields("Quantity").Value = tb_rec_f.Fields("quantity").Value
                                        tb_t.Fields("Rate").Value = tb_rec_f.Fields("Rate").Value
                                        tb_t.Fields("Ctn_no").Value = tb_rec_f.Fields("ctn_no").Value
                                        tb_t.Fields("Dye_Rate").Value = tb_rec_f.Fields("Dye_Rate").Value
                                        tb_rec_f.MoveNext
                                        tb_t.Update
                                        Loop
                                    End If
                                    End If
'                                    MsgBox ssql
                                
                           
                                
                            
                            
 If tb.Fields("e_type").Value = 1 Then tb_out_f.Close
If tb.Fields("e_type").Value = 2 Then tb_in_f.Close
If tb.Fields("e_type").Value = 3 Then tb_sent_f.Close
If tb.Fields("e_type").Value = 4 Then tb_rec_f.Close
               
            

            'If Not IsMissing(s_date) Then
                
    tb.MoveNext
    cntl.Value = cntl.Value + 1
    Loop
End If
    tb.Close
    tb_itm.Close
    tb_in.Close
    tb_out.Close
    tb_sent.Close
    tb_rec.Close
    tb_p.Close
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub

Private Sub Class_Initialize()
Set blm = New bloom1
m_path = blm.pathMain
t_path = App.path & "\book.mdb"
't_path = CurDir & "\" & "book.mdb"

End Sub

Private Sub Class_Terminate()
Set blm = Nothing
End Sub
Public Sub day_due_TEMP(V_DATE As Date, cntl As Control, CNTL2 As Control)
Dim db_m As Database
Dim db_t As Database
Dim test As Integer
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim t_cred As Currency
Dim ssql As String
Dim tbtrial As Recordset
Dim TB_C As Recordset
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from vou_temp"
db_t.Execute ssql
cntl.Value = 0
ssql = "delete from trial"
db_t.Execute (ssql)
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
ssql = "SELECT * FROM ACCHART ORDER BY CODE"
Set TB_C = db_m.OpenRecordset(ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
cntl.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    CNTL2.SimpleText = TB_C.Fields("DESCRIPTION").Value
        ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUCHER WHERE V_DATE <= #" & V_DATE & "#"
        ssql = ssql & " AND PARTY = " & TB_C.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) < 0 "
        Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
        If Not IsNull(tb.Fields("DEB").Value) Then
            t_cred = tb.Fields("DEB").Value
        End If
        End If
'        MsgBox TB_C.Fields("DESCRIPTION").Value & " pAID = " & t_cred
        tb.Close
    CNTL2.SimpleText = CNTL2.SimpleText & " tOTAL pAID = " & t_cred
     'TOTAL PURCHASED ITEMS ADJUSTMENTS
     If t_cred < 0 Then
        t_cred = t_cred * -1
        'MsgBox t_cred
        ssql = "SELECT * FROM SALE WHERE PARTY = " & TB_C.Fields("CODE").Value
        ssql = ssql & " AND TYPE=2 AND V_DATE <= #" & V_DATE & "# ORDER BY V_DATE"
        Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
           tb.MoveLast
            Do While Not tb.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
                    t_cred = t_cred - (tb.Fields("QUANTITY").Value * tb.Fields("RATE").Value)
                'End If
 '               MsgBox t_cred
                If t_cred <= 0 Then Exit Do
            Else
                Exit Do
            End If
            tb.MovePrevious
            Loop
            
            If tb.BOF Then tb.MoveFirst
  '          MsgBox tb.Fields("PARTY").Value
            If Not tb.BOF Then
            Do While Not tb.EOF
                    DoEvents
                    CNTL2.SimpleText = CNTL2.SimpleText & " tOTAL pAID = " & t_cred
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("due_date").Value = tb.Fields("v_date").Value + tb.Fields("days").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    ssql = blm.Item1(tb.Fields("item").Value) & " "
                    ssql = ssql & "Rate = " & tb.Fields("rate").Value & " "
                    ssql = ssql & "Quantity = " & tb.Fields("quantity").Value & " "
                    ssql = ssql & "Days = " & tb.Fields("days").Value
                    If tb.Fields("item").Value < 3000 Then
                        ssql = ssql & " Amount = " & tb.Fields("rate").Value * tb.Fields("quantity").Value
                    Else
                        ssql = ssql & " Amount = " & tb.Fields("rate").Value * tb.Fields("quantity").Value / 100
                    End If
                    tb_t.Fields("remarks").Value = ssql
                    If t_cred < 0 Then
                        t_cred = tb.Fields("QUANTITY").Value * tb.Fields("RATE").Value + t_cred
                    Else
                        t_cred = tb.Fields("QUANTITY").Value * tb.Fields("RATE").Value
                    End If
   '                 MsgBox t_cred
                    tb_t.Fields("CREDIT").Value = t_cred
                    tb_t.Update
                't_cred = -1
                tb.MoveNext
                If Not tb.EOF Then
                    t_cred = tb.Fields("QUANTITY").Value * tb.Fields("RATE").Value
                End If
            Loop
            End If
            't_cred = 0
        End If
        tb.Close
    End If
      TB_C.MoveNext
      cntl.Value = cntl.Value + 1
      DoEvents
    Loop
End If
TB_C.Close
tb_t.Close
db_t.Close
db_m.Close
End Sub

Public Sub day_due2_TEMP(V_DATE As Date, cntl As Control, CNTL2 As Control)
Dim db_m As Database
Dim db_t As Database
Dim test As Integer
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim t_cred As Currency
Dim ssql As String
Dim tbtrial As Recordset
Dim TB_C As Recordset
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from vou_temp"
db_t.Execute ssql
cntl.Value = 0
ssql = "delete from trial"
db_t.Execute (ssql)
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
ssql = "SELECT * FROM ACCHART ORDER BY CODE"
Set TB_C = db_m.OpenRecordset(ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
cntl.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    CNTL2.SimpleText = TB_C.Fields("Name").Value
        ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUCHER WHERE V_DATE <= #" & V_DATE & "#"
        ssql = ssql & " AND PARTY = " & TB_C.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) > 0 "
        Set tb = db_m.OpenRecordset(ssql)
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
        ssql = "SELECT * FROM SALE_Sum_View WHERE PARTY = " & TB_C.Fields("CODE").Value
        ssql = ssql & " AND V_DATE <= #" & V_DATE & "# ORDER BY V_DATE"
        Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
           tb.MoveLast
            Do While Not tb.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
        If Not IsNull(tb.Fields("Total").Value) Then
                    t_cred = t_cred - tb.Fields("Total").Value
        End If
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
                    tb_t.Fields("v_Date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("due_date").Value = tb.Fields("v_date").Value + tb.Fields("due_days").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    ssql = "Sale Invoice # " & tb.Fields("Sale_no").Value
                    ssql = ssql & "    DaysPassed = " & DateDiff("d", tb.Fields("V_Date").Value + tb.Fields("Due_days").Value, V_DATE)
                    ssql = ssql & "    Amount = " & tb.Fields("Total").Value
                    ssql = ssql & "    Quatity = " & tb.Fields("Qty").Value
                    tb_t.Fields("remarks").Value = ssql
                    If Not IsNull(tb.Fields("Total").Value) Then
                    If t_cred < 0 Then
                        t_cred = tb.Fields("Total").Value + t_cred
                    Else
                        t_cred = tb.Fields("Total").Value
                    End If
                    End If
'                    MsgBox t_cred
                    tb_t.Fields("CREDIT").Value = t_cred
                    tb_t.Update
                't_cred = -1
                tb.MoveNext
                If Not tb.EOF Then
                    If Not IsNull(tb.Fields("Total").Value) Then
                        t_cred = tb.Fields("Total").Value
                    End If
                End If
            Loop
            End If
            't_cred = 0
        End If
        tb.Close
    End If
      TB_C.MoveNext
      cntl.Value = cntl.Value + 1
      DoEvents
    Loop
End If
TB_C.Close
tb_t.Close
db_t.Close
db_m.Close
End Sub
Public Sub Aging(V_DATE As Date, cntl As Control, CNTL2 As Control, Accode As Long)
Dim db_m As Database
Dim db_t As Database
Dim test As Integer
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim tb_s As Recordset
Dim t_cred As Currency
Dim ssql As String
Dim tbtrial As Recordset
Dim TB_C As Recordset
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from aging"
db_t.Execute ssql
cntl.Value = 0
ssql = "delete from trial"
db_t.Execute (ssql)
Set tb_t = db_t.OpenRecordset("Aging", dbOpenTable)
ssql = "SELECT * FROM ACCHART WHERE CODE = " & Accode & " ORDER BY CODE"
Set TB_C = db_m.OpenRecordset(ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
cntl.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    CNTL2.SimpleText = TB_C.Fields("Name").Value
        ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUCHER WHERE V_DATE <= #" & V_DATE & "#"
        ssql = ssql & " AND PARTY = " & TB_C.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) > 0 "
        Set tb = db_m.OpenRecordset(ssql)
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
        ssql = "SELECT * FROM SALE_Sum_View WHERE PARTY = " & TB_C.Fields("CODE").Value
        ssql = ssql & " AND V_DATE <= #" & V_DATE & "# ORDER BY V_DATE"
                Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
           tb.MoveLast
            Do While Not tb.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
                If Not IsNull(tb.Fields("Total").Value) Then
                    t_cred = t_cred - tb.Fields("Total").Value
                End If
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
                    ssql = "select * from sale_View_1 where a.sale_no = " & tb.Fields("Sale_no").Value
                    Set tb_s = db_m.OpenRecordset(ssql)
                    If Not tb_s.EOF Then
                    DoEvents
                    CNTL2.SimpleText = CNTL2.SimpleText & " Total Paid = " & t_cred
                    Do While Not tb_s.EOF
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = tb_s.Fields("a.v_date").Value
                    tb_t.Fields("Sale_no").Value = tb_s.Fields("a.Sale_no").Value
                    tb_t.Fields("duedate").Value = tb_s.Fields("a.v_date").Value + tb_s.Fields("due_days").Value
                    tb_t.Fields("Due_Days").Value = tb_s.Fields("due_days").Value
                    tb_t.Fields("accode").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    If Not IsNull(tb_s.Fields("item").Value) Then
                    tb_t.Fields("Item").Value = blm.Item1(tb_s.Fields("Item").Value) ' & tb_s.Fields("CTN_no").Value
                    End If
                    tb_t.Fields("Quantity").Value = tb_s.Fields("Quantity").Value
                    tb_t.Fields("Rate").Value = tb_s.Fields("Rate").Value
                    tb_t.Fields("Bobins").Value = tb_s.Fields("Bobins").Value
                    tb_t.Fields("DaysGone").Value = DateDiff("d", tb_s.Fields("a.V_date"), V_DATE)
                    If t_cred < 0 Then
                        t_cred = tb.Fields("Total").Value + t_cred
                    Else
                        If Not IsNull(tb.Fields("Total").Value) Then
                        t_cred = tb.Fields("Total").Value
                        End If
                    End If
'                    MsgBox t_cred
                    tb_t.Fields("Amount").Value = t_cred
                    tb_t.Update
                    tb_s.MoveNext
                    Loop
                    End If
                    tb_s.Close
                't_cred = -1
                tb.MoveNext
                If Not tb.EOF Then
                    If Not IsNull(tb.Fields("Total").Value) Then
                        t_cred = tb.Fields("Total").Value
                    End If
                End If
            Loop
            End If
            't_cred = 0
        End If
        tb.Close
    End If
      TB_C.MoveNext
      cntl.Value = cntl.Value + 1
      DoEvents
    Loop
End If
TB_C.Close
tb_t.Close
db_t.Close
db_m.Close
End Sub

Public Sub sale_pur(s_date As Date, e_date As Date, ITM As Long, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim t_sal As Currency
Dim t_pur As Currency
Dim dif_sp As Currency

Dim pre_1 As Currency
Dim ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from vou_temp"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
ssql2 = "select int_stock,ext_stock from item where code =" & ITM
Set tb_1 = db_m.OpenRecordset(ssql2)
If Not tb_1.EOF Then
    op_ext = tb_1.Fields("ext_stock").Value
    op_int = tb_1.Fields("int_stock").Value
tb_t.AddNew
    tb_t.Fields("v_date").Value = s_date
    tb_t.Fields("remarks").Value = "Opeining STock on " & Format(s_date, "dd/MM/yyyy")
    If op_int > 0 Then
    tb_t.Fields("cred_q").Value = op_int
    tb_t.Fields("deb_q").Value = 0
    Else
    tb_t.Fields("cred_q").Value = 0
    tb_t.Fields("deb_q").Value = op_int
    End If
tb_t.Update
End If
tb_1.Close


'sales

ssql = "select * from sale where type = 1 and v_date between #" & s_date & "# and #" & e_date & "# and item = " & ITM
    Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
        tb.MoveLast
        cntl.Value = 0
        cntl.Max = tb.RecordCount
        tb.MoveFirst
            Do While Not tb.EOF
                tb_t.AddNew
                    tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    tb_t.Fields("debit").Value = tb.Fields("rate").Value * tb.Fields("quantity").Value
                    tb_t.Fields("deb_q").Value = tb.Fields("quantity").Value
                    tb_t.Fields("cred_q").Value = 0
                    tb_t.Fields("credit").Value = 0
                    tb_t.Fields("remarks").Value = "Sale " & blm.Item1(tb.Fields("item").Value) & " " & "Q=" & tb.Fields("quantity").Value & " " & "R=" & tb.Fields("rate").Value & " D.O.(Int) " & tb.Fields("int_d_no").Value & " D.O.(Ext) " & tb.Fields("ext_d_no").Value
                    tb_t.Fields("e_type").Value = tb.Fields("type").Value
                                
                 tb.MoveNext
                 tb_t.Update
                 cntl.Value = cntl.Value + 1
            Loop
        End If
        tb.Close
   tb_t.Close
   cntl.Value = 0
   Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
'sales
ssql = "select * from sale where type = 2 and v_date between #" & s_date & "# and #" & e_date & "# and item = " & ITM
    Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
        tb.MoveLast
        cntl.Value = 0
        cntl.Max = tb.RecordCount
        tb.MoveFirst
            Do While Not tb.EOF
                tb_t.AddNew
                    tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    tb_t.Fields("debit").Value = 0
                    tb_t.Fields("deb_q").Value = 0
                    tb_t.Fields("cred_q").Value = tb.Fields("quantity").Value
                    tb_t.Fields("credit").Value = tb.Fields("rate").Value * tb.Fields("quantity").Value
                    tb_t.Fields("remarks").Value = "Purc. " & blm.Item1(tb.Fields("item").Value) & " " & "Q=" & tb.Fields("quantity").Value & " " & "R=" & tb.Fields("rate").Value & " D.O.(Int) " & tb.Fields("int_d_no").Value & " D.O.(Ext) " & tb.Fields("ext_d_no").Value
                    tb_t.Fields("e_type").Value = tb.Fields("type").Value
                                
                 tb.MoveNext
                 tb_t.Update
                 cntl.Value = cntl.Value + 1
            Loop
        End If
        tb.Close
   tb_t.Close

db_t.Close
db_m.Close

End Sub

Public Sub rep_sale(s_date As Date, e_date As Date, ITM As Long, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim t_sal As Currency
Dim t_pur As Currency
Dim dif_sp As Currency

Dim pre_1 As Currency
Dim ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from vou_temp"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
'sales
ssql = "select * from sale where type = 1 and v_date between #" & s_date & "# and #" & e_date & "# and item = " & ITM
    Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
        tb.MoveLast
        cntl.Value = 0
        cntl.Max = tb.RecordCount
        tb.MoveFirst
            Do While Not tb.EOF
                tb_t.AddNew
                    tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    tb_t.Fields("debit").Value = tb.Fields("rate").Value * tb.Fields("quantity").Value
                    tb_t.Fields("deb_q").Value = tb.Fields("quantity").Value
                    tb_t.Fields("credit").Value = 0
                    tb_t.Fields("remarks").Value = "Sale " & blm.Item1(tb.Fields("item").Value) & " " & "Q=" & tb.Fields("quantity").Value & " " & "R=" & tb.Fields("rate").Value & " D.O.(Int) " & tb.Fields("int_d_no").Value & " D.O.(Ext) " & tb.Fields("ext_d_no").Value
                    tb_t.Fields("e_type").Value = tb.Fields("type").Value
                                
                 tb.MoveNext
                 tb_t.Update
                 cntl.Value = cntl.Value + 1
            Loop
        End If
        tb.Close
   tb_t.Close
db_t.Close
db_m.Close

End Sub

Public Sub rep_purchase(s_date As Date, e_date As Date, ITM As Long, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim t_sal As Currency
Dim t_pur As Currency
Dim dif_sp As Currency
Dim pre_1 As Currency
Dim ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.pathMain)
Set db_t = OpenDatabase(t_path)
ssql = "delete from vou_temp"
db_t.Execute ssql
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
'sales
ssql = "select * from sale where type = 2 and v_date between #" & s_date & "# and #" & e_date & "# and item = " & ITM
    Set tb = db_m.OpenRecordset(ssql)
        If Not tb.EOF Then
        tb.MoveLast
        cntl.Value = 0
        cntl.Max = tb.RecordCount
        tb.MoveFirst
            Do While Not tb.EOF
                tb_t.AddNew
                    tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    tb_t.Fields("debit").Value = 0
                    tb_t.Fields("credit").Value = tb.Fields("rate").Value * tb.Fields("quantity").Value
                    tb_t.Fields("remarks").Value = "Purc. " & blm.Item1(tb.Fields("item").Value) & " " & "Q=" & tb.Fields("quantity").Value & " " & "R=" & tb.Fields("rate").Value & " D.O.(Int) " & tb.Fields("int_d_no").Value & " D.O.(Ext) " & tb.Fields("ext_d_no").Value
                    tb_t.Fields("e_type").Value = tb.Fields("type").Value
                                
                 tb.MoveNext
                 tb_t.Update
                 cntl.Value = cntl.Value + 1
            Loop
        End If
        tb.Close
   tb_t.Close
db_t.Close
db_m.Close
End Sub
Public Sub TrialStock(e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As Recordset
Dim ssql As String
Dim Opsaleyarn As Double
Dim Opsalelycra As Double
Dim Oppuryarn As Double
Dim Oppurlycra As Double
Dim Opinlycra As Double, Opinyarn As Double
Dim Opoutlycra As Double, Opoutyarn As Double
Dim Fabin As Double
Dim Fabout As Double
Dim Usedin As Double
Dim Usedout As Double
Dim tb3 As New ADODB.Recordset
Fabin = 0
Fabout = 0
Usedin = 0
Usedout = 0

    Set db = OpenDatabase(m_path)
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

'''''''''''OPENING BALANCE'''''''''''''''
ssql = "Select party,yarncount,lycracount,sum(opbal) as opbal,sum(oplycra) as oplycra from cont_1 where E_type=2 and v_date<='" & Format(e_date, "dd-MMM-yyyy") & "' group by party,yarncount,lycracount"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
While Not tb1.EOF
tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("yarncode").Value = tb1.Fields("Yarncount").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("lycracode").Value = tb1.Fields("lycracount").Value
    tb.Fields("description2").Value = blm.Lycra(tb1.Fields("lycracount").Value)
    tb.Fields("yarnbal").Value = tb1.Fields("opbal").Value
    tb.Fields("lycrabal").Value = tb1.Fields("oplycra").Value
    tb.Fields("yarnbal2").Value = 0
    tb.Fields("lycrabal2").Value = 0
tb.Update
tb1.MoveNext
Wend
End If
tb1.Close
ssql = "Select party,yarncount,lycracount,sum(opbal) as opbal,sum(oplycra) as oplycra from cont_1 where E_type=1 and v_date<='" & Format(e_date, "dd-MMM-yyyy") & "' group by party,yarncount,lycracount"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
While Not tb1.EOF
tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("yarncode").Value = tb1.Fields("Yarncount").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("lycracode").Value = tb1.Fields("lycracount").Value
    tb.Fields("description2").Value = blm.Lycra(tb1.Fields("lycracount").Value)
    tb.Fields("yarnbal2").Value = tb1.Fields("opbal").Value
    tb.Fields("lycrabal2").Value = tb1.Fields("oplycra").Value
    tb.Fields("yarnbal").Value = 0
    tb.Fields("lycrabal").Value = 0

tb.Update
tb1.MoveNext
Wend

End If
tb1.Close

'Fabric Inward
ssql = "Select accode,sum(quantity)as infab,sum(lyc_used)as lycused from outward where E_type=1 "
ssql = ssql & "and v_date<='" & Format(e_date, "dd-MMM-yyyy") & "'group by accode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)

    tb.Fields("infabric").Value = tb1.Fields("infab").Value
    tb.Fields("inlyc_used").Value = tb1.Fields("lycused").Value
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Yarn & lycra Inward
ssql = "Select accode,yarn,lycracode,sum(bags*44.2)as inbag,sum(addwt)as inlbag,sum(lycra) as lycra,sum(l_bags) as inlycra from Inward where E_type=2 "
ssql = ssql & "and in_date<='" & Format(e_date, "dd-MMM-yyyy") & "'group by accode,yarn,lycracode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
While Not tb1.EOF
tb.AddNew
'Fabin = 0
'Usedin = 0

''Fabric Inward
'ssql = "Select item,sum(quantity)as infab,sum(lyc_used)as lycused from outward where E_type=1 "
'ssql = ssql & "and accode=" & tb1.Fields(Accode).Value & " and v_date<='" & Format(e_date, "dd-MMM-yyyy") & "'group by item"
'Set tb3 = CN.Execute(ssql)
'If Not tb1.EOF Then

'     If Not IsNull(tb3.Fields("infab")) Then
'     Fabin = tb3.Fields("infab").Value
'     Else
'     Fabin = 0
'     End If
'     If Not IsNull(tb3.Fields("lycused")) Then
'     Usedin = tb3.Fields("lycused").Value
'     Else
'     Usedin = 0
'     End If
'End If
'tb3.Close


    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("yarncode").Value = tb1.Fields("Yarn").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("lycracode").Value = tb1.Fields("lycracode").Value
    tb.Fields("description2").Value = blm.Lycra(tb1.Fields("lycracode").Value)
    tb.Fields("inbags").Value = tb1.Fields("inbag").Value + tb1.Fields("inlbag").Value
    tb.Fields("inlycra").Value = tb1.Fields("inlycra").Value
'    tb.Fields("infabric").Value = Fabin
 '   tb.Fields("inlyc_used").Value = Usedin

    tb.Fields("yarnbal").Value = 0
    tb.Fields("yarnbal2").Value = 0
    tb.Fields("lycrabal").Value = 0
    tb.Fields("lycrabal2").Value = 0
    tb.Fields("outbags").Value = 0
    tb.Fields("outlycra").Value = 0
tb.Update
tb1.MoveNext
Wend
End If
tb1.Close


''''''''''''

''Fabric Outward
ssql = "Select accode,sum(quantity)as outfab,sum(lyc_used)as lycused from outward where E_type=2 "
ssql = ssql & "and v_date<='" & Format(e_date, "dd-MMM-yyyy") & "'group by accode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("outfabric").Value = tb1.Fields("outfab").Value
    tb.Fields("outlyc_used").Value = tb1.Fields("lycused").Value
'MsgBox blm.party1(tb1.Fields("accode").Value)

    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close


'Yarn & lycra Outward
ssql = "Select accode,yarn,lycracode,sum(bags*44.2)as outbag,sum(addwt)as outlbag,sum(lycra) as lycra,sum(l_bags) as outlycra from Inward where E_type=1 "
ssql = ssql & "and in_date<='" & Format(e_date, "dd-MMM-yyyy") & "' group by accode,yarn,lycracode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
While Not tb1.EOF
tb.AddNew
'Fabout = 0
'Usedout = 0
'ssql = "Select item,sum(quantity)as outfab,sum(lyc_used)as lycused from outward where E_type=2 "
'ssql = ssql & "and accode=" & tb1.Fields(Accode).Value & " and v_date<='" & Format(e_date, "dd-MMM-yyyy") & "'group by item"
'Set tb3 = CN.Execute(ssql)
'If Not tb1.EOF Then
'     If Not IsNull(tb3.Fields("outfab")) Then
'     Fabout = tb3.Fields("outfab").Value
'     Else
'     Fabout = 0
'     End If
'     If Not IsNull(tb3.Fields("lycused")) Then
'     Usedout = tb3.Fields("lycused").Value
'     Else
'     Usedout = 0
'     End If
'End If
'tb3.Close


    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("yarncode").Value = tb1.Fields("Yarn").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("lycracode").Value = tb1.Fields("lycracode").Value
    tb.Fields("description2").Value = blm.Lycra(tb1.Fields("lycracode").Value)
    tb.Fields("outbags").Value = tb1.Fields("outbag").Value + tb1.Fields("outlbag").Value
    tb.Fields("outlycra").Value = tb1.Fields("outlycra").Value
    tb.Fields("outfabric").Value = Fabout
    tb.Fields("outlyc_used").Value = Usedout

    tb.Fields("yarnbal").Value = 0
    tb.Fields("yarnbal2").Value = 0
    tb.Fields("lycrabal").Value = 0
    tb.Fields("lycrabal2").Value = 0
    tb.Fields("inbags").Value = 0
    tb.Fields("inlycra").Value = 0
tb.Update
tb1.MoveNext
Wend
End If
tb1.Close


tb.Close
db.Close
End Sub

Public Sub inventoryledger(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As Recordset
Dim ssql As String
Dim Opsaleyarn As Double
Dim Opsalelycra As Double
Dim Oppuryarn As Double
Dim Oppurlycra As Double
Dim Opinlycra As Double, Opinyarn As Double
Dim Opoutlycra As Double, Opoutyarn As Double
Dim Infab As Double, Outfab As Double, Fabused As Double
    Set db = OpenDatabase(m_path)
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

'''''''''''OPENING BALANCE'''''''''''''''
ssql = "Select party,sum(opbal) as opbal,sum(oplycra) as oplycra from cont_1 where E_type=2 and party = " & c & "and v_date<'" & Format(s_date, "dd-MMM-yyyy") & "' group by party"

Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
If Not IsNull(tb1.Fields("opbal")) Then
Opsaleyarn = tb1.Fields("opbal").Value
Else
Opsaleyarn = 0
End If
If Not IsNull(tb1.Fields("oplycra")) Then
Opsalelycra = tb1.Fields("oplycra").Value
Else
Opsalelycra = 0
End If
End If

tb1.Close

ssql = "Select party,sum(opbal) as opbal,sum(oplycra) as oplycra from cont_1 where E_type=1 and party = " & c & "and v_date<'" & Format(s_date, "dd-MMM-yyyy") & "' group by party"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
If Not IsNull(tb1.Fields("opbal")) Then
Oppuryarn = tb1.Fields("opbal").Value
Else
Oppuryarn = 0
End If
If Not IsNull(tb1.Fields("oplycra")) Then
Oppurlycra = tb1.Fields("oplycra").Value
Else
Oppurlycra = 0
End If
End If

tb1.Close

'Yarn & lycra & Fabric  Inward
ssql = "Select accode,sum(quantity)as qty from outward where E_type=1 and Accode = " & c
ssql = ssql & "and v_date<'" & Format(s_date, "dd-MMM-yyyy") & "'group by accode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
Infab = tb1.Fields("qty").Value
Else
Infab = 0
End If
tb1.Close

ssql = "Select accode,sum(bags*44.2)as inbag,sum(addwt)as inlbag,sum(lycra) as lycra,sum(l_bags) as inlycra from Inward where E_type=2 and Accode = " & c
ssql = ssql & "and in_date<'" & Format(s_date, "dd-MMM-yyyy") & "'group by accode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
Opinyarn = tb1.Fields("inbag").Value + tb1.Fields("inlbag").Value
Opinlycra = tb1.Fields("inlycra").Value ' + tb1.Fields("lycra").Value
Else
Opinyarn = 0
Opinlycra = 0
End If
tb1.Close
'Yarn & lycra & fabric Outward
ssql = "Select accode,sum(quantity) as qty,sum(lyc_used)as lyc_used from outward where E_type=2 and Accode = " & c
ssql = ssql & "and v_date<'" & Format(s_date, "dd-MMM-yyyy") & "'group by accode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
Outfab = tb1.Fields("qty").Value
Fabused = tb1.Fields("lyc_used").Value
Else
Outfab = 0
Fabused = 0
End If
tb1.Close

ssql = "Select accode,sum(bags*44.2)as outbag,sum(addwt)as outlbag,sum(lycra) as lycra,sum(l_bags) as outlycra from Inward where E_type=1 and Accode = " & c
ssql = ssql & "and in_date<'" & Format(s_date, "dd-MMM-yyyy") & "' group by accode"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
Opoutyarn = tb1.Fields("outbag").Value + tb1.Fields("outlbag").Value
Opoutlycra = tb1.Fields("outlycra").Value '+ tb1.Fields("lycra").Value
Else
Opoutyarn = 0
Opoutlycra = 0
End If
tb1.Close

tb.AddNew
tb.Fields("yarnbal").Value = (Opsaleyarn + Opinyarn + Infab + Opinlycra + Opsalelycra) - (Oppuryarn + Opoutyarn + Outfab + Opoutlycra + Oppurlycra)
tb.Fields("lycrabal").Value = (Opsalelycra + Opinlycra) - (Oppurlycra + Opoutlycra + Fabused)
tb.Update
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''sale conprect opening
ssql = "Select * from cont_1 where E_type=2 and party = " & c & "and v_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    If tb1.Fields("opbal") > 0 Or tb1.Fields("oplycra") > 0 Or tb1.Fields("opfabric") > 0 Then
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("gpin").Value = "OP"
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("inyarn").Value = tb1.Fields("opbal").Value
    tb.Fields("Inlycra").Value = tb1.Fields("oplycra").Value
    'tb.Fields("outfabric").Value = tb1.Fields("opfabric").Value
    tb.Update
    End If
    tb1.MoveNext
    Wend
End If
tb1.Close
''''''''purchase cont opening
ssql = "Select * from cont_1 where E_type=1 and party = " & c & "and v_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    If tb1.Fields("opbal") > 0 Or tb1.Fields("oplycra") > 0 Or tb1.Fields("opfabric") > 0 Then
    
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("gpout").Value = "OP"
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("outyarn").Value = tb1.Fields("opbal").Value
    tb.Fields("outlycra").Value = tb1.Fields("oplycra").Value
'    tb.Fields("infabric").Value = tb1.Fields("opfabric").Value
    tb.Update
    End If
    tb1.MoveNext
    Wend
End If
tb1.Close

'Yarn Inward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.Yarn=b.Code and b.Y_Type=1 and a.E_type=2 and a.Accode = " & c & "and in_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("gpin").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("inbags").Value = tb1.Fields("bags").Value
    tb.Fields("InYarn").Value = tb1.Fields("AddWt").Value
    tb.Fields("InYarnWeight").Value = tb1.Fields("wtPerPackYarn").Value * tb1.Fields("Bags")
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Lycra Inward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.LycraCode=b.Code and b.Y_type=2 and a.E_type=2 and a.Accode = " & c & "and in_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
'Clipboard.SetText ssql
'MsgBox ssql
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("gpin").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("LycraCode").Value)
    tb.Fields("inLycra").Value = tb1.Fields("L_Bags").Value
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Fabric Inward
ssql = "Select * from outward where E_type=1 and Accode = " & c & "and v_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("gpin").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Cloth(tb1.Fields("Item").Value) & " " & tb1.Fields("size1").Value & ""
    tb.Fields("inFabric").Value = tb1.Fields("Quantity").Value
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close


'Yarn Outward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.Yarn=b.Code and b.Y_type=1 and a.E_type=1 and a.Accode = " & c & "and in_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("gpout").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("outbags").Value = tb1.Fields("bags").Value
    tb.Fields("outYarn").Value = tb1.Fields("AddWt").Value
    tb.Fields("outYarnWeight").Value = tb1.Fields("WtPerPackYarn").Value * tb1.Fields("Bags")
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Lycra Outward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.LycraCode=b.Code and b.Y_type=2 and a.E_type=1 and a.Accode = " & c & "and in_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("gpout").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("LycraCode").Value)
    tb.Fields("outLycra").Value = tb1.Fields("L_bags").Value
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Fabric outward
ssql = "Select * from outward where E_type=2 and Accode = " & c & "and v_date between'" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "'"
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("gpout").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Cloth(tb1.Fields("Item").Value) & " " & tb1.Fields("size1").Value & ""
    tb.Fields("outFabric").Value = tb1.Fields("Quantity").Value
    tb.Fields("outRoll").Value = tb1.Fields("Rolls").Value
    tb.Fields("OutAge").Value = tb1.Fields("Lyc_Per").Value
    tb.Fields("OutLyc_Used").Value = tb1.Fields("Lyc_Used").Value
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close


tb.Close
db.Close

End Sub
Public Sub cont_summery(c As Single, T As Single)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As Recordset
Dim ssql As String
Dim Opsaleyarn As Double
Dim Opsalelycra As Double
Dim Oppuryarn As Double
Dim Oppurlycra As Double
Dim Opinlycra As Double, Opinyarn As Double
Dim Opoutlycra As Double, Opoutyarn As Double
    Set db = OpenDatabase(m_path)
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

'''''''''''OPENING BALANCE'''''''''''''''
'ssql = "Select opbal as opbal,oplycra as oplycra from cont_1 where E_type=" & T & "and cont_no = " & c
'Set tb1 = CN.Execute(ssql)
'If Not tb1.EOF Then
'If Not IsNull(tb1.Fields("opbal")) Then
'Opsaleyarn = tb1.Fields("opbal").Value
'Else
'Opsaleyarn = 0
'End If
'If Not IsNull(tb1.Fields("oplycra")) Then
'Opsalelycra = tb1.Fields("oplycra").Value
'Else
'Opsalelycra = 0
'End If
'End If
'tb1.Close
'
'tb.AddNew
'tb.Fields("yarnbal").Value = Opsaleyarn
'tb.Fields("lycrabal").Value = Opsalelycra
'tb.Update
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''sale conprect opening
If T = 2 Then
ssql = "Select * from cont_1 where E_type=" & T & " and cont_no= " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    If tb1.Fields("opbal") > 0 Or tb1.Fields("oplycra") > 0 Or tb1.Fields("opfabric") > 0 Then
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("gpin").Value = "OP"
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("inyarn").Value = tb1.Fields("opbal").Value
    tb.Fields("Inlycra").Value = tb1.Fields("oplycra").Value
    'tb.Fields("outfabric").Value = tb1.Fields("opfabric").Value
    tb.Fields("cont_no").Value = c
    'tb1.Fields("item").Value
    'tb1.Fields("YQUANTITY").value
    'tb1.Fields("cQUANTITY").value
    
    tb.Update
    End If
    tb1.MoveNext
    Wend
End If
tb1.Close
End If
''''''''purchase cont opening
If T = 1 Then
ssql = "Select * from cont_1 where E_type=" & T & " and cont_no= " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    If tb1.Fields("opbal") > 0 Or tb1.Fields("oplycra") > 0 Or tb1.Fields("opfabric") > 0 Then
    
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("gpout").Value = "OP"
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("outyarn").Value = tb1.Fields("opbal").Value
    tb.Fields("outlycra").Value = tb1.Fields("oplycra").Value
'    tb.Fields("infabric").Value = tb1.Fields("opfabric").Value
    tb.Fields("cont_no").Value = c
    
    tb.Update
    End If
    tb1.MoveNext
    Wend
End If
tb1.Close
End If

'Yarn Inward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.Yarn=b.Code and b.Y_Type=1 and a.E_type=2 and a.cont_no = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("inbags").Value = tb1.Fields("bags").Value
    tb.Fields("InYarn").Value = tb1.Fields("AddWt").Value
    tb.Fields("InYarnWeight").Value = tb1.Fields("wtPerPackYarn").Value * tb1.Fields("Bags")
    tb.Fields("cont_no").Value = c
    
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Lycra Inward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.LycraCode=b.Code and b.Y_type=2 and a.E_type=2 and a.cont_no = " & c
'Clipboard.SetText ssql
'MsgBox ssql
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("LycraCode").Value)
    tb.Fields("inLycra").Value = tb1.Fields("L_Bags").Value
    tb.Fields("cont_no").Value = c
    
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Fabric Inward
ssql = "Select * from outward where E_type=1 and cont_no = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Cloth(tb1.Fields("Item").Value)
    tb.Fields("inFabric").Value = tb1.Fields("Quantity").Value
    tb.Fields("cont_no").Value = c
    
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close


'Yarn Outward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.Yarn=b.Code and b.Y_type=1 and a.E_type=1 and a.cont_no = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("outbags").Value = tb1.Fields("bags").Value
    tb.Fields("outYarn").Value = tb1.Fields("AddWt").Value
    tb.Fields("outYarnWeight").Value = tb1.Fields("WtPerPackYarn").Value * tb1.Fields("Bags")
    tb.Fields("cont_no").Value = c
    
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Lycra Outward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.LycraCode=b.Code and b.Y_type=2 and a.E_type=1 and a.cont_no = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("LycraCode").Value)
    tb.Fields("outLycra").Value = tb1.Fields("L_bags").Value
    tb.Fields("cont_no").Value = c
    
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close

'Fabric outward
ssql = "Select * from outward where E_type=2 and cont_no = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Cloth(tb1.Fields("Item").Value)
    tb.Fields("outFabric").Value = tb1.Fields("Quantity").Value
    tb.Fields("outRoll").Value = tb1.Fields("Rolls").Value
    tb.Fields("OutAge").Value = tb1.Fields("Lyc_Per").Value
    tb.Fields("OutLyc_Used").Value = tb1.Fields("Lyc_Used").Value
    tb.Fields("cont_no").Value = c
    
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close


tb.Close
db.Close

End Sub


Public Sub inventoryinward(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As ADODB.Recordset
Dim ssql As String
    
    Set db = OpenDatabase(m_path)
    ssql = "select * from inward where e_type=2 and accode = " & c & " and in_Date Between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by in_date"
    Set tb1 = CN.Execute(ssql)
    ssql = "select * from outward where e_type in(1,3) and accode = " & c & " and v_Date Between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by v_date"
    Set tb2 = CN.Execute(ssql)
    
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    If tb1.Fields("lycracode").Value > 0 Then
    tb.Fields("description").Value = blm.Lycra(tb1.Fields("lycracode").Value)
    End If
    
    If tb1.Fields("yarn").Value > 0 Then
    tb.Fields("description").Value = blm.Yarn(tb1.Fields("yarn").Value)
    End If
    tb.Fields("inbags").Value = tb1.Fields("bags").Value
    tb.Fields("InYarn").Value = tb1.Fields("AddWt").Value
    
    tb.Fields("inlycra").Value = tb1.Fields("l_bags").Value
    tb.Update
    tb1.MoveNext
    Wend
End If
If Not tb2.EOF Then
    While Not tb2.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb2.Fields("accode").Value
    tb.Fields("acname").Value = tb2.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb2.Fields("out_no").Value
    tb.Fields("date").Value = tb2.Fields("v_date").Value
    tb.Fields("description").Value = blm.Cloth(tb2.Fields("item").Value) & " " & tb2.Fields("size1").Value & ""
    tb.Fields("infabric").Value = tb2.Fields("rolls").Value
    tb.Fields("infabricqty").Value = tb2.Fields("quantity").Value
    tb.Fields("inage").Value = tb2.Fields("lyc_per").Value
    tb.Fields("inlyc_used").Value = tb2.Fields("lyc_used").Value
    
    tb.Update
    tb2.MoveNext
    Wend
End If

tb.Close
tb1.Close
tb2.Close
db.Close

End Sub

Public Sub inventoryoutward(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As ADODB.Recordset
Dim ssql As String
    
    Set db = OpenDatabase(m_path)
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)
    
    ssql = "select * from outward where e_type=2 and accode = " & c & " and v_Date Between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by v_date"
    Set tb1 = CN.Execute(ssql)
    ssql = "select * from inward where e_type in(1,3) and accode = " & c & " and in_Date Between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by in_date"
    Set tb2 = CN.Execute(ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Cloth(tb1.Fields("item").Value) & " " & tb1.Fields("size1").Value & ""
    tb.Fields("outroll").Value = tb1.Fields("rolls").Value
    tb.Fields("outfabric").Value = tb1.Fields("quantity").Value
    tb.Fields("outage").Value = tb1.Fields("lyc_per").Value
    tb.Fields("outlyc_used").Value = tb1.Fields("lyc_used").Value
    
    tb.Update
    tb1.MoveNext
    Wend
End If
If Not tb2.EOF Then
    While Not tb2.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb2.Fields("accode").Value
    tb.Fields("acname").Value = tb2.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb2.Fields("in_no").Value
    tb.Fields("date").Value = tb2.Fields("in_date").Value
    If tb2.Fields("lycracode").Value > 0 Then
    tb.Fields("description").Value = blm.Lycra(tb2.Fields("lycracode").Value)
    End If
    
    If tb2.Fields("yarn").Value > 0 Then
    tb.Fields("description").Value = blm.Yarn(tb2.Fields("yarn").Value)
    End If
    If tb2.Fields("e_type") = 3 Then
    tb.Fields("outroll").Value = tb2.Fields("bags").Value
    tb.Fields("outfabric").Value = tb2.Fields("quantity").Value
    End If
    If tb2.Fields("e_type") = 1 Then
    tb.Fields("outbags").Value = tb2.Fields("bags").Value
    tb.Fields("outYarn").Value = tb2.Fields("AddWt").Value
    tb.Fields("outlycra").Value = tb2.Fields("l_bags").Value
    End If
    tb.Update
    tb2.MoveNext
    Wend
End If



tb.Close
tb1.Close
tb2.Close
db.Close

End Sub

Public Sub yinventoryledger(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As Recordset
Dim ssql As String
    
    Set db = OpenDatabase(m_path)
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

''''''''''sale conprect opening
ssql = "Select * from cont_1 where E_type=2 and yarncount = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    If tb1.Fields("opbal") > 0 Or tb1.Fields("oplycra") > 0 Or tb1.Fields("opfabric") > 0 Then
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("gpin").Value = "OP"
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("inyarn").Value = tb1.Fields("opbal").Value
    'tb.Fields("Inlycra").Value = tb1.Fields("oplycra").Value
    'tb.Fields("outfabric").Value = tb1.Fields("opfabric").Value
    tb.Update
    End If
    tb1.MoveNext
    Wend
End If
tb1.Close
''''''''purchase cont opening
ssql = "Select * from cont_1 where E_type=1 and yarncount = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    If tb1.Fields("opbal") > 0 Or tb1.Fields("oplycra") > 0 Or tb1.Fields("opfabric") > 0 Then
    
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("party").Value
    tb.Fields("acname").Value = blm.party1(tb1.Fields("party").Value)
    tb.Fields("gpout").Value = "OP"
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarncount").Value)
    tb.Fields("outyarn").Value = tb1.Fields("opbal").Value
    'tb.Fields("outlycra").Value = tb1.Fields("oplycra").Value
'    tb.Fields("infabric").Value = tb1.Fields("opfabric").Value
    tb.Update
    End If
    tb1.MoveNext
    Wend
End If
tb1.Close



'Yarn Inward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.Yarn=b.Code and b.Y_type=1 and a.E_type=2 and a.Yarn = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("inbags").Value = tb1.Fields("bags").Value
    tb.Fields("InYarn").Value = tb1.Fields("AddWt").Value
    tb.Fields("InYarnWeight").Value = tb1.Fields("wtPerPackYarn").Value * tb1.Fields("Bags")
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close



'Yarn Outward
ssql = "Select a.*,b.y_type from Inward a,Yarn b where a.Yarn=b.Code and b.Y_type=1 and a.E_type=1 and a.Yarn = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Yarn").Value)
    tb.Fields("outbags").Value = tb1.Fields("bags").Value
    tb.Fields("outYarn").Value = tb1.Fields("AddWt").Value
    tb.Fields("outYarnWeight").Value = tb1.Fields("WtPerPackYarn").Value * tb1.Fields("Bags")
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close


tb.Close
db.Close

End Sub
Public Sub finventoryledger(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As Recordset
Dim ssql As String
    
Set db = OpenDatabase(m_path)
ssql = "delete from inventory"
db.Execute (ssql)
Set tb = db.OpenRecordset("inventory", dbOpenTable)


'Fabric Inward
ssql = "Select * from outward where E_type=1 and Item = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Item1(tb1.Fields("Item").Value)
    tb.Fields("inrolls").Value = tb1.Fields("rolls").Value
    tb.Fields("inFabric").Value = tb1.Fields("Quantity").Value
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close



'Fabric outward
ssql = "Select * from outward where E_type=2 and Item = " & c
Set tb1 = CN.Execute(ssql)
If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = blm.Cloth(tb1.Fields("Item").Value)
    tb.Fields("outFabric").Value = tb1.Fields("Quantity").Value
    tb.Fields("outRoll").Value = tb1.Fields("Rolls").Value
    tb.Fields("OutAge").Value = tb1.Fields("Lyc_Per").Value
    tb.Fields("OutLyc_Used").Value = tb1.Fields("Lyc_Used").Value
    tb.Fields("inrolls").Value = 0
    tb.Update
    tb1.MoveNext
    Wend
End If
tb1.Close


tb.Close
db.Close

End Sub

Public Sub yinventoryinward(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim ssql As String
    
    Set db = OpenDatabase(m_path)
    ssql = "select * from inward where e_type=2 and yarn = " & c & " and in_Date between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by in_date"
    Set tb1 = CN.Execute(ssql)
    
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb1.Fields("in_no").Value
    tb.Fields("date").Value = tb1.Fields("in_date").Value
    tb.Fields("description").Value = blm.party1(tb1.Fields("accode").Value)
    tb.Fields("inbags").Value = tb1.Fields("bags").Value
    tb.Fields("InYarn").Value = tb1.Fields("AddWt").Value
    tb.Fields("inlycra").Value = tb1.Fields("l_bags").Value
    tb.Update
    tb1.MoveNext
    Wend
End If

tb.Close
tb1.Close
db.Close

End Sub

Public Sub finventoryinward(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb2 As ADODB.Recordset
Dim ssql As String
    
    Set db = OpenDatabase(m_path)
    ssql = "select * from outward where e_type in(1,3) and item = " & c & " and v_Date between '" & Format(s_date, "dd-MMM-yyyy") & " 'and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by v_date"
    Set tb2 = CN.Execute(ssql)
    
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

If Not tb2.EOF Then
    While Not tb2.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb2.Fields("accode").Value
    tb.Fields("acname").Value = tb2.Fields("acname").Value & ""
    tb.Fields("gpin").Value = tb2.Fields("out_no").Value
    tb.Fields("date").Value = tb2.Fields("v_date").Value
    tb.Fields("description").Value = blm.party1(tb2.Fields("accode").Value)
    tb.Fields("infabric").Value = tb2.Fields("rolls").Value
    tb.Fields("infabricqty").Value = tb2.Fields("quantity").Value
    tb.Fields("inage").Value = tb2.Fields("lyc_per").Value
    tb.Fields("inlyc_used").Value = tb2.Fields("lyc_used").Value
    
    tb.Update
    tb2.MoveNext
    Wend
End If

tb.Close
tb2.Close
db.Close

End Sub

Public Sub yinventoryoutward(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As Recordset
Dim tb2 As ADODB.Recordset
Dim ssql As String
    
    Set db = OpenDatabase(m_path)
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)
    
    ssql = "select * from inward where e_type =1 and yarn = " & c & " or lycracode= " & c & " and in_Date Between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by in_date"
    Set tb2 = CN.Execute(ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

If Not tb2.EOF Then
    While Not tb2.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb2.Fields("accode").Value
    tb.Fields("acname").Value = tb2.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb2.Fields("in_no").Value
    tb.Fields("date").Value = tb2.Fields("in_date").Value
    tb.Fields("description").Value = blm.party1(tb2.Fields("accode").Value)
    If tb2.Fields("e_type") = 3 Then
    tb.Fields("outroll").Value = tb2.Fields("bags").Value
    tb.Fields("outfabric").Value = tb2.Fields("quantity").Value
    End If
    If tb2.Fields("e_type") = 1 Then
    tb.Fields("outbags").Value = tb2.Fields("bags").Value
    tb.Fields("outYarn").Value = tb2.Fields("AddWt").Value
    
    tb.Fields("outlycra").Value = tb2.Fields("l_bags").Value
    End If
    tb.Update
    tb2.MoveNext
    Wend
End If



tb.Close
tb2.Close
db.Close

End Sub

Public Sub finventoryoutward(c As Single, s_date As Date, e_date As Date)
Dim db As Database
Dim tb As Recordset
Dim tb1 As ADODB.Recordset
Dim tb2 As ADODB.Recordset
Dim ssql As String
    
    Set db = OpenDatabase(m_path)
    ssql = "delete from inventory"
    db.Execute (ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)
    
    ssql = "select * from outward where e_type=2 and item = " & c & " and v_Date between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by v_date"
    Set tb1 = CN.Execute(ssql)
    ssql = "select * from inward where e_type =3 and item = " & c & " and in_Date between '" & Format(s_date, "dd-MMM-yyyy") & "' and '" & Format(e_date, "dd-MMM-yyyy") & "' Order by in_date"
    Set tb2 = CN.Execute(ssql)
    Set tb = db.OpenRecordset("inventory", dbOpenTable)

If Not tb1.EOF Then
    While Not tb1.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb1.Fields("accode").Value
    tb.Fields("acname").Value = tb1.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb1.Fields("out_no").Value
    tb.Fields("date").Value = tb1.Fields("v_date").Value
    tb.Fields("description").Value = tb1.Fields("remarks").Value
    tb.Fields("outroll").Value = tb1.Fields("rolls").Value
    tb.Fields("outfabric").Value = tb1.Fields("quantity").Value
    tb.Fields("outage").Value = tb1.Fields("lyc_per").Value
    tb.Fields("outlyc_used").Value = tb1.Fields("lyc_used").Value
    
    tb.Update
    tb1.MoveNext
    Wend
End If
If Not tb2.EOF Then
    While Not tb2.EOF
    tb.AddNew
    tb.Fields("accode").Value = tb2.Fields("accode").Value
    tb.Fields("acname").Value = tb2.Fields("acname").Value & ""
    tb.Fields("gpout").Value = tb2.Fields("in_no").Value
    tb.Fields("date").Value = tb2.Fields("in_date").Value
    tb.Fields("description").Value = tb2.Fields("remarks").Value
    If tb2.Fields("e_type") = 3 Then
    tb.Fields("outroll").Value = tb2.Fields("bags").Value
    tb.Fields("outfabric").Value = tb2.Fields("quantity").Value
    End If
    If tb2.Fields("e_type") = 1 Then
    tb.Fields("outbags").Value = tb2.Fields("bags").Value
    tb.Fields("outlycra").Value = tb2.Fields("l_bags").Value
    End If
    tb.Update
    tb2.MoveNext
    Wend
End If



tb.Close
tb1.Close
tb2.Close
db.Close

End Sub

