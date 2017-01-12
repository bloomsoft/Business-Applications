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

Public Function YarnName(c As Long) As String
Dim DBT As Database
Dim TBT As Recordset

Set DBT = OpenDatabase(m_path)
Set TBT = DBT.OpenRecordset("Select * from Yarns WHere Code = " & c)
If Not TBT.EOF Then
    YarnName = TBT.Fields("Name").Value
Else
    YarnName = "Wrong"
End If
TBT.Close
DBT.Close

End Function

Public Sub FillCloth(Cntl As Object)
Dim DBT As Database
Dim TBT As Recordset

Set DBT = OpenDatabase(m_path)
Set TBT = DBT.OpenRecordset("Select * from Cloths")
If Not TBT.EOF Then
    Cntl.clear
    Do While Not TBT.EOF
        Cntl.AddItem TBT.Fields("NAME")
        'Cntl.AddItem TBT.Fields("Reed") & " X " & TBT.Fields("Pick") & " / " & TBT.Fields("Warp") & " X " & TBT.Fields("Weft") & " " & TBT.Fields("Width") & "' " & TBT.Fields("CType") & "'" & TBT.Fields("NAME")
        Cntl.ItemData(Cntl.NewIndex) = TBT.Fields("Code")
        TBT.MoveNext
    Loop
    Cntl.ListIndex = 0
End If
TBT.Close
DBT.Close
End Sub

Public Function Factory(c As Long) As String
    Dim DB As Database
    Dim tb As Recordset
    
    Set DB = OpenDatabase(m_path)
    Set tb = DB.OpenRecordset("select * from FactoryChart where code = " & c)
    If Not tb.EOF Then
        Factory = tb.Fields("Name").Value
    Else
        Factory = "NOT FOUND"
    End If
tb.Close
DB.Close
End Function

Public Function Dying(c As Long) As String
    Dim DB As Database
    Dim tb As Recordset
    
    Set DB = OpenDatabase(m_path)
    Set tb = DB.OpenRecordset("select * from DyingChart where code = " & c)
    If Not tb.EOF Then
        Dying = tb.Fields("Name").Value
    Else
        Dying = "NOT FOUND"
    End If
tb.Close
DB.Close
End Function

Public Function FillCloth1(c As Long) As String
    Dim DB As Database
    Dim tb As Recordset
    
    Set DB = OpenDatabase(m_path)
    Set tb = DB.OpenRecordset("select * from Cloths where code = " & c)
    If Not tb.EOF Then
        FillCloth1 = tb.Fields("Name").Value
    Else
        FillCloth1 = "NOT FOUND"
    End If
tb.Close
DB.Close
End Function



Public Function Reasons(c As Byte) As String
Select Case c
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

Public Function RPAD(s As String, length As Byte) As String
    Dim l As Integer
    Dim i As Integer
    Dim e As Integer
    
    l = Len(s)
    e = length - l
    s = Trim(s)
    For i = 1 To e
        s = s & "_"
    Next i
    RPAD = Mid(s, 1, length)
End Function
Public Property Get report_path() As String
report_path = "C:\soft\Hosiery\"
End Property


Public Property Get pathMain() As String
    pathMain = m_path
End Property

Private Sub Class_Initialize()

    m_path = DatabaseDrive & "\Bloom.mdb"
'    MsgBox m_path
End Sub

Public Sub fill_comb(Ssql As String, Cntl As Control, itemcol As String, datacol As String)
    Dim DB As Database
    Dim tb As Recordset
    Cntl.clear
    Set DB = OpenDatabase(m_path)
    Set tb = DB.OpenRecordset(Ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            Cntl.AddItem tb.Fields(itemcol).Value
            Cntl.ItemData(Cntl.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        Cntl.ListIndex = 0
    End If
    tb.Close
    DB.Close
        
End Sub
Public Sub fill_comb_Item(Ssql As String, Cntl As Control, itemcol As String, datacol As String)
    Dim DB As Database
    Dim tb As Recordset
    Cntl.clear
    Set DB = OpenDatabase(m_path)
    Set tb = DB.OpenRecordset(Ssql)
    If Not tb.EOF Then
        Do While Not tb.EOF
            Cntl.AddItem tb.Fields(itemcol).Value & " {" & tb.Fields("Unit").Value & "}"
            Cntl.ItemData(Cntl.NewIndex) = tb.Fields(datacol).Value
            tb.MoveNext
        Loop
        Cntl.ListIndex = 0
    End If
    tb.Close
    DB.Close
        
End Sub

Public Sub fill_comb2(Ssql As String, Cntl As Control, itemcol As String, itemcol2 As String, itemcol3 As String, datacol As String)
'    Dim db As Database
'    Dim tb As Recordset
'    Cntl.clear
'    Set db = OpenDatabase(m_path)
'    Set tb = db.OpenRecordset(Ssql)
'    If Not tb.EOF Then
'        Do While Not tb.EOF
'            Cntl.AddItem tb.Fields(itemcol).Value & "-----" & tb.Fields(itemcol2).Value & "-----" & Item1(tb.Fields(itemcol3).Value)
'            Cntl.ItemData(Cntl.NewIndex) = tb.Fields(datacol).Value
'            tb.MoveNext
'        Loop
'        Cntl.ListIndex = 0
'    End If
'    tb.Close
'    db.Close
        
End Sub

Public Sub cont_type_comb(Cntl As Control)
    Cntl.clear
        Cntl.AddItem "Recieved"
        Cntl.ItemData(Cntl.NewIndex) = 1
        Cntl.AddItem "Sent"
        Cntl.ItemData(Cntl.NewIndex) = 2
    Cntl.ListIndex = 0
End Sub
Public Function cont_type_ret(type1 As Byte) As String
    
    Select Case type1
        Case 1
            cont_type_ret = "Recieved"
        Case 2
            cont_type_ret = "Sent"
    End Select
End Function

Public Sub Bro_type_comb(Cntl As Control)
    Cntl.clear
        Cntl.AddItem "/Unit"
        Cntl.ItemData(Cntl.NewIndex) = 1
        Cntl.AddItem "Total Amount"
        Cntl.ItemData(Cntl.NewIndex) = 2
    Cntl.ListIndex = 0
End Sub
Public Function Balance(c As Long) As Currency
Dim DB As Database
Dim tb As Recordset
Dim Ssql As String

    Set DB = OpenDatabase(m_path)
    Ssql = "select sum(debit - credit) as bal from voucher where party = " & c
    Set tb = DB.OpenRecordset(Ssql)
    If Not IsNull(tb.Fields("bal").Value) Then
        Balance = tb.Fields("Bal").Value
    Else
        Balance = 0
    End If
tb.Close
DB.Close

End Function

