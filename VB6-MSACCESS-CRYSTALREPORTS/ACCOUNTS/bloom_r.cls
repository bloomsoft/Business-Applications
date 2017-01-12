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
Public Sub day_due_TEMP(v_date As Date, cntl As Control, CNTL2 As Control)
Dim db_m As Database
Dim db_t As Database
Dim test As Integer
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim t_cred As Currency
Dim Ssql As String
Dim tbtrial As Recordset
Dim TB_C As Recordset
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from vou_temp"
db_t.Execute Ssql
cntl.Value = 0
Ssql = "delete from trial"
db_t.Execute (Ssql)
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
Ssql = "SELECT * FROM ACCHART ORDER BY CODE"
Set TB_C = db_m.OpenRecordset(Ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
cntl.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    CNTL2.SimpleText = TB_C.Fields("name").Value
        Ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUdtl WHERE V_DATE <= #" & v_date & "#"
        Ssql = Ssql & " AND PARTY = " & TB_C.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) < 0 "
        Set tb = db_m.OpenRecordset(Ssql)
        If Not tb.EOF Then
        If Not IsNull(tb.Fields("DEB").Value) Then
            t_cred = tb.Fields("DEB").Value
        End If
        End If
'        MsgBox TB_C.Fields("DESCRIPTION").Value & " pAID = " & t_cred
        tb.Close
    CNTL2.SimpleText = CNTL2.SimpleText & " TOTAL PAID = " & t_cred
     'TOTAL PURCHASED ITEMS ADJUSTMENTS
     If t_cred < 0 Then
        t_cred = t_cred * -1
        'MsgBox t_cred
        Ssql = "SELECT * FROM Voudtl WHERE PARTY = " & TB_C.Fields("CODE").Value
        Ssql = Ssql & " AND V_DATE <= #" & v_date & "# ORDER BY V_DATE"
        Set tb = db_m.OpenRecordset(Ssql)
        If Not tb.EOF Then
           tb.MoveLast
            Do While Not tb.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
                    t_cred = t_cred - (tb.Fields("Credit").Value)
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
                    If tb.Fields("Pakki_no").Value > 0 Then
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("due_date").Value = tb.Fields("v_date").Value + tb.Fields("days").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    tb_t.Fields("remarks").Value = tb.Fields("Remarks").Value & " Amt = " & tb.Fields("Credit").Value
                    If t_cred < 0 Then
                        t_cred = tb.Fields("Credit").Value + t_cred
                    Else
                        t_cred = tb.Fields("Credit").Value
                    End If
   '                 MsgBox t_cred
                    tb_t.Fields("CREDIT").Value = t_cred
                    tb_t.Update
                    End If
                't_cred = -1
                tb.MoveNext
                If Not tb.EOF Then
                    t_cred = tb.Fields("Credit").Value
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
Public Sub day_due2_TEMP(v_date As Date, cntl As Control, CNTL2 As Control)
Dim db_m As Database
Dim db_t As Database
Dim test As Integer
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_tt As Recordset
Dim t_cred As Currency
Dim Ssql As String
Dim tbtrial As Recordset
Dim TB_C As Recordset
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from vou_temp"
db_t.Execute Ssql
cntl.Value = 0
Ssql = "delete from trial"
db_t.Execute (Ssql)
Set tb_t = db_t.OpenRecordset("vou_temp", dbOpenTable)
Ssql = "SELECT * FROM ACCHART ORDER BY CODE"
Set TB_C = db_m.OpenRecordset(Ssql)
If Not TB_C.EOF Then
TB_C.MoveLast
cntl.Max = TB_C.RecordCount
TB_C.MoveFirst
    Do While Not TB_C.EOF
    t_cred = 0
    'SUM OF TOTAL PAID AMONTS
    CNTL2.SimpleText = TB_C.Fields("Name").Value
        Ssql = "SELECT SUM(DEBIT) - SUM(CREDIT) AS DEB FROM VOUdtl WHERE V_DATE <= #" & v_date & "#"
        Ssql = Ssql & " AND PARTY = " & TB_C.Fields("CODE").Value & " HAVING SUM(DEBIT) - SUM(CREDIT) > 0 "
        Set tb = db_m.OpenRecordset(Ssql)
        If Not tb.EOF Then
        If Not IsNull(tb.Fields("DEB").Value) Then
            t_cred = tb.Fields("DEB").Value
        End If
        End If
'        MsgBox TB_C.Fields("DESCRIPTION").Value & " pAID = " & t_cred
        tb.Close
    CNTL2.SimpleText = CNTL2.SimpleText & " TOTAL PAID = " & t_cred
     'TOTAL Sales ITEMS ADJUSTMENTS
     If t_cred > 0 Then
        
 '       MsgBox t_cred
        Ssql = "SELECT * FROM voudtl WHERE PARTY = " & TB_C.Fields("CODE").Value
        Ssql = Ssql & " AND V_DATE <= #" & v_date & "# ORDER BY V_DATE"
        Set tb = db_m.OpenRecordset(Ssql)
        If Not tb.EOF Then
           tb.MoveLast
            Do While Not tb.BOF
            If t_cred > 0 Then
                'If (tb.Fields("V_DATE").Value + tb.Fields("DAYS").Value) >= v_date Then
'                MsgBox "Check"
                    t_cred = t_cred - (tb.Fields("Debit").Value)
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
                    
                    If tb.Fields("Pakki_no").Value > 0 Then
                    tb_t.AddNew
                    tb_t.Fields("v_Date").Value = tb.Fields("v_date").Value
                    tb_t.Fields("due_date").Value = tb.Fields("v_date").Value + tb.Fields("days").Value
                    tb_t.Fields("code").Value = tb.Fields("party").Value
                    tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
                    tb_t.Fields("remarks").Value = tb.Fields("Remarks").Value & " Amt = " & tb.Fields("Debit").Value
                    If t_cred < 0 Then
                        t_cred = (tb.Fields("Debit").Value) + t_cred
                    Else
                        t_cred = tb.Fields("Debit").Value
                    End If
'                    MsgBox t_cred
                    tb_t.Fields("CREDIT").Value = t_cred
                    tb_t.Update
                    End If
                't_cred = -1
                tb.MoveNext
                If Not tb.EOF Then
                    t_cred = tb.Fields("Debit").Value
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


Public Property Get report_path() As String
    report_path = App.path & "\"
'report_path = CurDir & "\"
End Property
Public Property Get path() As String
    path = t_path
End Property

Public Sub trial(e_date As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
    
'Vouchers

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_date <= #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
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
    
    tb.Close

tb_t.Close
db_t.Close
db_m.Close
End Sub
Public Sub trial2(s_date As Date, e_date As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
    
'Vouchers

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_date < #" & s_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("s_date").Value = s_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("op_deb").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("op_deb").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("op_cred").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("op_cred").Value = 0
        End If
    tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close

tb_t.Close
cntl.Value = 0
Dim tb_3 As Recordset
Dim ssql2 As String
    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_date between #" & s_date & "# and #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
        DoEvents
        Ssql = "select * from trial where code = " & tb.Fields("party").Value
        Set tb_3 = db_t.OpenRecordset(Ssql)
        If Not tb_3.EOF Then
            If tb.Fields("bal").Value < 0 Then
                ssql2 = "update trial set debit = 0,credit = " & tb.Fields("bal").Value * -1
            End If
           If tb.Fields("bal").Value > 0 Then
                ssql2 = "update trial set credit = 0, debit = " & tb.Fields("bal").Value
           End If
           If tb.Fields("bal").Value <> 0 Then
           ssql2 = ssql2 & " where code = " & tb.Fields("party").Value
           
           db_t.Execute ssql2
           End If
        Else
        Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("op_deb").Value = 0
        tb_t.Fields("op_cred").Value = 0
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
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
        tb_t.Close
        End If
        
        
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close


db_t.Close
db_m.Close
End Sub

Public Sub open_trial(cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)
    
'Vouchers

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  v_no = 1 and v_type = 10 group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
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
    
    tb.Close

tb_t.Close
db_t.Close
db_m.Close
End Sub

Public Sub ProfitLoss(e_date As Date, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from trial"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("trial", dbOpenTable)

'sales
    Ssql = "select sum(debit) as sales from voucher where e_type = 1 or e_type = 4 and v_date <= #" & e_date & "#"
    Set tb = db_m.OpenRecordset(Ssql)
    If Not IsNull(tb.Fields("sales").Value) Then
        tb_t.AddNew
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("code").Value = 5000
            tb_t.Fields("Name").Value = "Sales For The Day"
            tb_t.Fields("debit").Value = 0#
            tb_t.Fields("credit").Value = tb.Fields("sales").Value
            
        tb_t.Update
   End If
tb.Close
'purchases
    Ssql = "select sum(credit) as purchases from voucher where e_type between 2 and 3 and v_date <= #" & e_date & "#"
    Set tb = db_m.OpenRecordset(Ssql)
    If Not IsNull(tb.Fields("purchases").Value) Then
        tb_t.AddNew
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("code").Value = 6000
            tb_t.Fields("Name").Value = "Purchases For The Day"
            tb_t.Fields("debit").Value = tb.Fields("purchases").Value
            tb_t.Fields("credit").Value = 0
            
        tb_t.Update
    End If
   tb.Close
'Vouchers

    Ssql = "select subcode,sum(debit-credit) as bal from vou_view where party <> 10000 and h_type = 2 and v_date <= #" & e_date & "# group by subcode"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        'tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("subcode").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("subcode").Value), 1, 2))))
        tb_t.Fields("s_code").Value = tb.Fields("subcode").Value
        tb_t.Fields("s_desc").Value = blm.heads(tb.Fields("subcode").Value)
        'tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
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
    
    tb.Close

tb_t.Close
db_t.Close
db_m.Close
End Sub

Public Sub day_book(v_date As Date, cntl As Control)
Dim DB As Database
    Dim tb As Recordset
    Dim Ssql As String
    Dim t_sal As Currency
    Dim t_pur As Currency
    Dim dif_sp As Currency
    Dim t_t_r As Currency
    Dim t_t_s As Currency
    Set blm = Nothing
    Set blm = New bloom1
    Dim dbt As Database
    Dim tbt As Recordset
    Set DB = OpenDatabase(blm.patHmain)
    Set tb = DB.OpenRecordset("Select * from voucher where v_date = #" & v_date & "#")
    Set dbt = OpenDatabase(report_path & "book.mdb")
    dbt.Execute ("delete from vou_temp")
    Set tbt = dbt.OpenRecordset("vou_temp", dbOpenTable)
    cntl.Value = 0
    
    If tb.EOF = False Then
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
        Do While Not tb.EOF
            tbt.AddNew
                tbt.Fields("V_dATE").Value = v_date
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
    DB.Close
    End Sub
Public Sub ledger(code As Long, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
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
Set tb = db_m.OpenRecordset(Ssql)
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("v_no").Value = tb.Fields("v_no").Value
            tb_t.Fields("v_type").Value = blm.VoucherTypesRet(tb.Fields("v_type").Value)
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = blm.party1(tb.Fields("party").Value)
            
            If tb.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
        
            Else
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            End If
            If tb.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = tb.Fields("remarks").Value
            tb_t.Fields("Kachi_no").Value = tb.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = tb.Fields("Days").Value
            tb_t.Fields("K_Date").Value = tb.Fields("P_Date").Value
'                MsgBox tb_t.Fields("Credit").Value
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
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub
Public Sub ICTLedger(ICT As String, cntl As Control)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
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
Set tb = db_m.OpenRecordset(Ssql)
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("v_no").Value = tb.Fields("v_no").Value
            tb_t.Fields("v_type").Value = blm.VoucherTypesRet(tb.Fields("v_type").Value)
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = blm.party1(tb.Fields("party").Value)
            
            If tb.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
        
            Else
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            End If
            If tb.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = tb.Fields("remarks").Value
            'tb_t.Fields("Kachi_no").Value = Tb.Fields("ICT").Value & ""
            tb_t.Fields("ICt").Value = tb.Fields("ICT").Value & ""
            
'                MsgBox tb_t.Fields("Credit").Value
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
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub

Public Sub ledger2(code As Long, cntl As Control, s_date As Date, e_date As Date)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
'If Not IsMissing(s_date) Then
    Ssql = "select sum(debit-credit) as bal from voudtl where party = " & code & " and v_date < #" & s_date & "#"
    Set tb = db_m.OpenRecordset(Ssql)
    If Not IsNull(tb.Fields("bal").Value) Then
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
            tb_t.Fields("description").Value = blm.party1(code)
            tb_t.Fields("remarks").Value = "Opening Balance On .." & Format(s_date, "dd/MM/yyyy")
     tb_t.Update
    End If
    tb.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    Ssql = "select * from voudtl where party = " & code & " and v_Date >= #" & s_date & "#" & " and v_date <= #" & e_date & "#"
'End If
'MsgBox ssql

Set tb = db_m.OpenRecordset(Ssql)
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("v_no").Value = tb.Fields("v_no").Value
            tb_t.Fields("v_type").Value = blm.VoucherTypesRet(tb.Fields("v_type").Value)
        
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = blm.party1(tb.Fields("party").Value)
            If tb.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            Else
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            End If
            If tb.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = tb.Fields("remarks").Value
            tb_t.Fields("Kachi_no").Value = tb.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = tb.Fields("Days").Value
            tb_t.Fields("K_Date").Value = tb.Fields("P_Date").Value
            
            'If Not IsMissing(s_date) Then
                tb_t.Fields("s_date").Value = s_date
                tb_t.Fields("e_date").Value = e_date
           ' End If
        tb_t.Update
    tb.MoveNext
    cntl.Value = cntl.Value + 1
    Loop
End If
    tb.Close
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub
Public Sub CashBook(cntl As Control, s_date As Date, e_date As Date)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim op_bal As Currency
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
'If Not IsMissing(s_date) Then
    Ssql = "select sum(credit-Debit) as bal from voudtl where party = 120040001 and v_date < #" & s_date & "#"
    Set tb = db_m.OpenRecordset(Ssql)
    If Not IsNull(tb.Fields("bal").Value) Then
    
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
            tb_t.Fields("v_No").Value = 0
            tb_t.Fields("v_Type").Value = 0
            tb_t.Fields("code").Value = 160010001 'tb.Fields("party").Value
            tb_t.Fields("description").Value = blm.party1(160010001)
            tb_t.Fields("remarks").Value = "Cash B/F on  .." & Format(s_date, "dd/MM/yyyy")
            tb_t.Fields("s_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
     tb_t.Update
    End If
    tb.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    Ssql = "select * from voudtl where party <>120040001 and v_type=3 and v_Date >= #" & s_date & "#" & " and v_date <= #" & e_date & "#"
'End If
'MsgBox ssql

Set tb = db_m.OpenRecordset(Ssql)
'Clipboard.SetText Ssql
'MsgBox Ssql
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("v_no").Value = tb.Fields("v_no").Value
            tb_t.Fields("v_type").Value = blm.VoucherTypesRet(tb.Fields("v_type").Value)
        
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = blm.party1(tb.Fields("party").Value)
            If tb.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            Else
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            End If
            If tb.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = tb.Fields("remarks").Value
            tb_t.Fields("Kachi_no").Value = tb.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = tb.Fields("Days").Value
            tb_t.Fields("K_Date").Value = tb.Fields("P_Date").Value
            
            'If Not IsMissing(s_date) Then
                tb_t.Fields("s_date").Value = s_date
                tb_t.Fields("e_date").Value = e_date
           ' End If
        tb_t.Update
    tb.MoveNext
    cntl.Value = cntl.Value + 1
    Loop
End If
    tb.Close
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub

Public Sub ledger3(code As Long, cntl As Control, s_date As Date, e_date As Date)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim tb_t2 As Recordset
Dim op_bal As Currency
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "delete from ledger"
db_t.Execute Ssql
Ssql = "delete from opbal"
db_t.Execute Ssql

Set tb_t = db_t.OpenRecordset("ledger", dbOpenTable)
Set tb_t2 = db_t.OpenRecordset("opbal", dbOpenTable)
'If Not IsMissing(s_date) Then
    Ssql = "select party,sum(debit-credit) as bal from voudtl where v_date < #" & s_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
     If Not tb.EOF Then
        Do While Not tb.EOF
          tb_t.AddNew
            tb_t.Fields("v_date").Value = s_date
            tb_t.Fields("s_date").Value = s_date
            tb_t.Fields("e_date").Value = e_date
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
            tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
            tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
            tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
            tb_t.Fields("description").Value = blm.party1(tb.Fields("party").Value)
            tb_t.Fields("remarks").Value = "Opening Balance On .." & Format(s_date, "dd/MM/yyyy")
            If tb.Fields("bal").Value > 0 Then
                tb_t.Fields("debit").Value = tb.Fields("bal").Value
                tb_t.Fields("credit").Value = 0
            End If
            If tb.Fields("bal").Value < 0 Then
                tb_t.Fields("credit").Value = tb.Fields("bal").Value * -1
                tb_t.Fields("debit").Value = 0
            End If
            tb_t.Update
            tb.MoveNext
        Loop
     End If
    tb.Close
'End If
'If IsMissing(Format(s_date, "dd/MM/yyyy")) Then
' ssql = "select * from voucher where party = " & code
 '   MsgBox ssql
'Else
    Ssql = "select * from voudtl where v_Date between #" & s_date & "#" & " and #" & e_date & "#"
'End If
'MsgBox ssql

Set tb = db_m.OpenRecordset(Ssql)
If Not tb.EOF Then
    cntl.Value = 0
    tb.MoveLast
    cntl.Max = tb.RecordCount
    tb.MoveFirst
    Do While Not tb.EOF
        tb_t.AddNew
            tb_t.Fields("op_bal").Value = op_bal
            tb_t.Fields("v_date").Value = tb.Fields("v_date").Value
            tb_t.Fields("v_no").Value = tb.Fields("v_no").Value
            tb_t.Fields("v_type").Value = blm.VoucherTypesRet(tb.Fields("v_type").Value)
            
            tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
            tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
            tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
            tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
            tb_t.Fields("code").Value = tb.Fields("party").Value
            tb_t.Fields("description").Value = blm.party1(tb.Fields("party").Value)
            If tb.Fields("debit").Value > 0 Then
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            Else
            tb_t.Fields("debit").Value = tb.Fields("debit").Value
            End If
            If tb.Fields("credit").Value > 0 Then
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            Else
            tb_t.Fields("credit").Value = tb.Fields("credit").Value
            End If
            tb_t.Fields("remarks").Value = tb.Fields("remarks").Value
            tb_t.Fields("Kachi_no").Value = tb.Fields("Pakki_no").Value
            tb_t.Fields("Days").Value = tb.Fields("Days").Value
            tb_t.Fields("K_Date").Value = tb.Fields("P_Date").Value
            
            'If Not IsMissing(s_date) Then
                tb_t.Fields("s_date").Value = s_date
                tb_t.Fields("e_date").Value = e_date
           ' End If
        tb_t.Update
    tb.MoveNext
    cntl.Value = cntl.Value + 1
    Loop
End If
    tb.Close
    tb_t.Close
    db_t.Close
    db_m.Close
End Sub
Public Property Get FinancePath() As String
FinancePath = App.path & "\Finance.mdb"
End Property
Public Sub ProfitLossState(e_date As Date, cntl As Object)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset
Dim NetPlYear As Currency, NetPLMonth As Currency
Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(FinancePath)

Ssql = "delete from PLYear"
db_t.Execute Ssql
Ssql = "delete from PLMONTH"
db_t.Execute Ssql
Ssql = "delete from PLAMT"
db_t.Execute Ssql


Set tb_t = db_t.OpenRecordset("PLYear", dbOpenTable)
    
'Vouchers For the Year

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Mid(Party,1,2) in (14,15) and v_date < #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        NetPlYear = 0
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        'tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        NetPlYear = NetPlYear + tb.Fields("Bal").Value
    tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close
    
    cntl.Value = 0
    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Mid(Party,1,5) in (12007) and v_date < #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        NetPlYear = 0
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        'tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        NetPlYear = NetPlYear + tb.Fields("Bal").Value
    tb_t.Update
            tb.MoveNext
'            MsgBox cntl.Value
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close

tb_t.Close
cntl.Value = 0
Set tb_t = db_t.OpenRecordset("PLMonth", dbOpenTable)
Dim s_date As Date

s_date = CDate(Month(e_date) & "/01/" & Year(e_date))

'Vouchers For the Month

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Mid(Party,1,2) in (14,15) and v_date between #" & s_date & "# and #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        NetPLMonth = 0
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
    '    tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        NetPLMonth = NetPLMonth + tb.Fields("Bal").Value
    tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close
cntl.Value = 0
    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Mid(Party,1,5) in (12007) and v_date between #" & s_date & "# and #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        NetPLMonth = 0
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
    '    tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        NetPLMonth = NetPLMonth + tb.Fields("Bal").Value
    tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close

tb_t.Close


Set tb_t = db_t.OpenRecordset("PLAMT", dbOpenTable)
tb_t.AddNew
tb_t.Fields("PLYear").Value = NetPlYear
tb_t.Fields("PLMonth").Value = NetPLMonth
tb_t.Update
tb_t.Close


db_t.Close
db_m.Close

End Sub
Public Sub BalanceSheetState(e_date As Date, cntl As Object)
Dim db_m As Database
Dim db_t As Database
Dim tb As Recordset
Dim tb_t As Recordset

Dim Ssql As String
cntl.Value = 0
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(FinancePath)

Ssql = "delete from BalYear"
db_t.Execute Ssql
Ssql = "delete from BalMONTH"
db_t.Execute Ssql


Set tb_t = db_t.OpenRecordset("BalYear", dbOpenTable)
    
'Vouchers For the Year

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Mid(Party,1,2) Not in (14,15) and v_date < #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst

        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
        'tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If

    tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close
Set tb = db_t.OpenRecordset("PLAMT", dbOpenDynaset)
tb_t.AddNew
    tb_t.Fields("Code").Value = 0
    tb_t.Fields("Name").Value = "Net Profit Loss For the Year"
    If tb.Fields("PLYear").Value < 0 Then
        tb_t.Fields("E_date").Value = e_date
        tb_t.Fields("Credit").Value = tb.Fields("PlYear").Value * -1
        tb_t.Fields("Debit").Value = 0
    Else
        tb_t.Fields("E_date").Value = e_date
        tb_t.Fields("Debit").Value = tb.Fields("PlYear").Value
        tb_t.Fields("Credit").Value = 0
    End If
tb_t.Update
tb.Close

tb_t.Close

cntl.Value = 0

Set tb_t = db_t.OpenRecordset("BalMonth", dbOpenTable)
Dim s_date As Date

s_date = CDate(Month(e_date) & "/01/" & Year(e_date))

'Vouchers For the Month

    Ssql = "select party,sum(debit-credit) as bal from voudtl where  Mid(Party,1,2) not in (14,15) and v_date between #" & s_date & "# and #" & e_date & "# group by party"
    Set tb = db_m.OpenRecordset(Ssql)
        If tb.EOF = False Then
        tb.MoveLast
        cntl.Max = tb.RecordCount
        tb.MoveFirst
        
        Do While Not tb.EOF
        DoEvents
        tb_t.AddNew
        tb_t.Fields("e_date").Value = e_date
    '    tb_t.Fields("s_date").Value = S_date
'        MsgBox tb.Fields("party").Value
        tb_t.Fields("code").Value = tb.Fields("party").Value
        tb_t.Fields("h_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 2))
        tb_t.Fields("h_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 2))))
        tb_t.Fields("s_code").Value = Val(Mid(CStr(tb.Fields("party").Value), 1, 5))
        tb_t.Fields("s_desc").Value = blm.heads(Val((Mid(CStr(tb.Fields("party").Value), 1, 5))))
        tb_t.Fields("name").Value = blm.party1(tb.Fields("party").Value)
        tb_t.Fields("debit").Value = 0
        tb_t.Fields("credit").Value = 0
        If tb.Fields("bal").Value > 0 Then
            tb_t.Fields("Debit").Value = tb.Fields("bal").Value
        Else
            tb_t.Fields("Debit").Value = 0
        End If
        If tb.Fields("bal").Value < 0 Then
            tb_t.Fields("Credit").Value = tb.Fields("bal").Value * -1
        Else
            tb_t.Fields("Credit").Value = 0
        End If
        
    tb_t.Update
            tb.MoveNext
            cntl.Value = cntl.Value + 1
    Loop
    End If
    
    tb.Close

Set tb = db_t.OpenRecordset("PLAMT", dbOpenDynaset)
tb_t.AddNew
    tb_t.Fields("Code").Value = 0
    tb_t.Fields("Name").Value = "Net Profit Loss For the Month"
    If tb.Fields("PLMonth").Value < 0 Then
        tb_t.Fields("E_date").Value = e_date
        tb_t.Fields("Credit").Value = tb.Fields("PlMonth").Value * -1
        tb_t.Fields("Debit").Value = 0
    Else
        tb_t.Fields("E_date").Value = e_date
        tb_t.Fields("Debit").Value = tb.Fields("PlMonth").Value
        tb_t.Fields("Credit").Value = 0
    End If
tb_t.Update
tb.Close
tb_t.Close




db_t.Close
db_m.Close

End Sub
Public Sub OverAllStockNew(e_date As Date)
'On Error Resume Next
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim Ssql As String
Dim f As String
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
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)




db_t.Execute "Drop Table TempItem"
db_m.Execute "Select '" & Format(e_date, "dd-MMM-yyyy") & "' As EDate,Code,Name,Unit,Stock into TempItem in '" & t_path & "' from Items Order by Code"

'Inward
Ssql = "Drop Table PurchaseStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as PQTY,Avg(Rate)as AVGPPRice into PurchaseStk in '" & t_path & "' from Purchase where V_Date <= #" & e_date & "# Group BY item"
db_m.Execute Ssql
'Sales
Ssql = "Drop Table SALEStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as SALEQTY,Avg(Rate)as SRATE into SALESTK in '" & t_path & "' from Sales where v_Date <= #" & e_date & "# Group BY item"
db_m.Execute Ssql
'Consumption
Ssql = "Drop Table ConsumeStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as CQTY,Avg(Rate)as SRATE into ConsumeSTK in '" & t_path & "' from Consume where v_Date <= #" & e_date & "# Group BY item"
db_m.Execute Ssql


db_m.Close
db_t.Close





End Sub
Public Sub PeriodicStockNew(s_date As Date, e_date As Date)
'On Error Resume Next
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim Ssql As String
Dim f As String
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
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)




db_t.Execute "Drop Table TempItem"
db_m.Execute "Select '" & Format(s_date, "dd-MMM-yyyy") & "' As SDate, '" & Format(e_date, "dd-MMM-yyyy") & "' As EDate,Code,Name,Unit,Stock into TempItem in '" & t_path & "' from Items Order by Code"

'Inward
Ssql = "Drop Table PurchaseStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as PQTY,Avg(Rate)as AVGPPRice into PurchaseStk in '" & t_path & "' from Purchase where V_Date Between #" & s_date & "# and #" & e_date & "# Group BY item"
db_m.Execute Ssql
'Sales
Ssql = "Drop Table SALEStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as SALEQTY,Avg(Rate)as SRATE into SALESTK in '" & t_path & "' from Sales where v_Date Between #" & s_date & "# and #" & e_date & "# Group BY item"
db_m.Execute Ssql
'Consumption
Ssql = "Drop Table ConsumeStk"
db_t.Execute Ssql
Ssql = "select Item,Sum(Qty)as CQTY,Avg(Rate)as SRATE into ConsumeSTK in '" & t_path & "' from Consume where v_Date Between #" & s_date & "# and #" & e_date & "# Group BY item"
db_m.Execute Ssql

Set tb_i = db_t.OpenRecordset("TempItem", dbOpenDynaset)
If Not tb_i.EOF Then
    Do While Not tb_i.EOF
        tb_i.Edit
            tb_i.Fields("Stock").Value = blm.ITEMstocks(tb_i.Fields("Code").Value, s_date)
        tb_i.Update
        tb_i.MoveNext
    Loop

End If
tb_i.Close
db_m.Close
db_t.Close





End Sub

Public Sub ItemLedger(s_date As Date, e_date As Date, Itm As Integer, p As Object)
Dim db_m As Database
Dim db_t As Database
Dim tb_t As Recordset
Dim tb_i As Recordset
Dim tb_a As Recordset
Dim Ssql As String
Dim f As String
Dim tb_m As Recordset

Dim itname As String
Dim uname As String
Dim pname As String
Set db_m = OpenDatabase(blm.patHmain)
Set db_t = OpenDatabase(t_path)
Ssql = "Delete from itmLedger"
db_t.Execute Ssql
p.Value = 0

Set tb_t = db_t.OpenRecordset("itmLedger", dbOpenTable)

Set tb_i = db_m.OpenRecordset("Select * from Items Order by Code")
Set tb_a = db_m.OpenRecordset("Select * from Acchart Order by Code")
'Purchases
Ssql = "Select * from Purchase where item = " & Itm
Ssql = Ssql & " and v_date Between #" & s_date & "# and #" & e_date & "# Order by V_date"
p.Value = 0
Set tb_m = db_m.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        f = "Code = " & tb_m.Fields("Item").Value
        tb_i.FindFirst f
        If tb_i.NoMatch Then
            itname = ""
        Else
            itname = tb_i.Fields("Name").Value
            uname = tb_i.Fields("Unit").Value
        End If
        
        f = "Code = " & tb_m.Fields("Seller").Value
        tb_a.FindFirst f
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        tb_t.AddNew
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = tb_m.Fields("Item").Value
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Seller").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("Qty").Value = tb_m.Fields("Qty").Value
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = blm.VoucherTypesRet(tb_m.Fields("V_Type").Value)
        tb_t.Fields("Inv_No").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'sales

Ssql = "Select * from Sales where item = " & Itm
Ssql = Ssql & " and v_date Between #" & s_date & "# and #" & e_date & "# Order by v_date"
p.Value = 0
Set tb_m = db_m.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        f = "Code = " & tb_m.Fields("Item").Value
        tb_i.FindFirst f
        If tb_i.NoMatch Then
            itname = ""
        Else
            itname = tb_i.Fields("Name").Value
            uname = tb_i.Fields("Unit").Value
        End If
        
        f = "Code = " & tb_m.Fields("Buyer").Value
        tb_a.FindFirst f
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        tb_t.AddNew
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = tb_m.Fields("Item").Value
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Buyer").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("Qty").Value = tb_m.Fields("Qty").Value * -1
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = blm.VoucherTypesRet(tb_m.Fields("V_type").Value)
        tb_t.Fields("Inv_no").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close


'Consumption

Ssql = "Select * from Consume where item = " & Itm
Ssql = Ssql & " and v_date Between #" & s_date & "# and #" & e_date & "# Order by v_date"
p.Value = 0
Set tb_m = db_m.OpenRecordset(Ssql)
If Not tb_m.EOF Then
    tb_m.MoveLast
        p.Max = tb_m.RecordCount
    tb_m.MoveFirst
    Do While Not tb_m.EOF
        f = "Code = " & tb_m.Fields("Item").Value
        tb_i.FindFirst f
        If tb_i.NoMatch Then
            itname = ""
        Else
            itname = tb_i.Fields("Name").Value
            uname = tb_i.Fields("Unit").Value
        End If
        
        f = "Code = " & tb_m.Fields("Buyer").Value
        tb_a.FindFirst f
        If tb_a.NoMatch Then
            pname = ""
        Else
            pname = tb_a.Fields("Name").Value
            
        End If
        tb_t.AddNew
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("v_date").Value = tb_m.Fields("v_date").Value
        tb_t.Fields("v_no").Value = tb_m.Fields("V_no").Value
        tb_t.Fields("Item").Value = tb_m.Fields("Item").Value
        tb_t.Fields("ItName").Value = itname
        tb_t.Fields("Unit").Value = uname
        tb_t.Fields("Party").Value = tb_m.Fields("Buyer").Value
        tb_t.Fields("PNAMe").Value = pname
        tb_t.Fields("Qty").Value = tb_m.Fields("Qty").Value * -1
        tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("V_TYPE").Value = blm.VoucherTypesRet(tb_m.Fields("V_type").Value)
        tb_t.Fields("Inv_no").Value = tb_m.Fields("Inv_no").Value & ""
        tb_t.Fields("StRate").Value = tb_m.Fields("StRate").Value
        tb_t.Update
        tb_m.MoveNext
        p.Value = p.Value + 1
    Loop
End If
tb_m.Close



'OpenSTocks
Dim qstk As Currency

        tb_t.AddNew
        tb_t.Fields("s_date").Value = s_date
        tb_t.Fields("e_date").Value = e_date
        tb_t.Fields("v_date").Value = s_date - 1
        tb_t.Fields("Item").Value = Itm
        tb_t.Fields("ItName").Value = blm.item1(Itm)
        tb_t.Fields("Unit").Value = uname
        'tb_t.Fields("Party").Value = tb_m.Fields("Party").Value
        'tb_t.Fields("PNAMe").Value = pname
        qstk = blm.ITEMstocks(Itm, s_date)
        tb_t.Fields("Qty").Value = qstk
        'tb_t.Fields("Rate").Value = tb_m.Fields("Rate").Value
        tb_t.Fields("v_TYPE").Value = "APSTK"
        
        tb_t.Update


tb_i.Close
tb_a.Close
tb_t.Close
db_m.Close
db_t.Close
End Sub

Private Sub Class_Initialize()
Set blm = New bloom1
m_path = blm.patHmain
't_path = "c:\soft\Accounts\book.mdb"
't_path = CurDir & "\" & "book.mdb"
t_path = App.path & "\book.mdb"
End Sub

Private Sub Class_Terminate()
Set blm = Nothing
End Sub
