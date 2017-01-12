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
Private blm As New bloom1
Public Function PaymentStatus(Dying As Long, LotNo As Long) As String
Dim DBM As DAO.Database
Dim TBM As DAO.Recordset
Dim Ssql As String
Set DBM = OpenDatabase(blm.pathMain)

Ssql = "Select * from PaymentDying where Dying_Code=" & Dying & " and Lot_No=" & LotNo
Set TBM = DBM.OpenRecordset(Ssql)
If Not TBM.EOF Then
    PaymentStatus = "Paid"
Else
    PaymentStatus = ""
End If
TBM.Close
DBM.Close


End Function
Public Sub DyeReport(SDate As Date, EDate As Date, FactoryCode As Long, Optional ItemCode As Long, Optional PackingLots As Boolean)
Dim TBT As DAO.Recordset
Dim DBT As DAO.Database
Dim DBM As DAO.Database
Dim TBM As DAO.Recordset
Dim FName As String, CName As String, DName As String
DName = blm.Dying(FactoryCode)
CName = blm.FillCloth1(ItemCode)
Set DBT = OpenDatabase(App.Path & "\Book.mdb")
Ssql = "Delete from DyeRep"
DBT.Execute Ssql

Set DBM = OpenDatabase(blm.pathMain)
Set TBT = DBT.OpenRecordset("Dyerep", dbOpenTable)

'Cloth Rec
Ssql = "Select a.*,b.ThanZ as PackThanz,b.Gazana as DGazana,b.P_gazana as PackGazana,b.Dispute from ClothRec a Left Join Packing b on a.Dying_Code=b.Dying_Code and a.Lot_No=b.Lot_No where a.Date_Recieve Between #" & SDate & "# and #" & EDate & "#"
Ssql = Ssql & " and a.Dying_Code=" & FactoryCode
If ItemCode > 0 Then
    Ssql = Ssql & " and a.Cloth_Code=" & ItemCode
End If
Ssql = Ssql & " Order by a.Date_Recieve,a.Rec_No"
'Clipboard.SetText ssql
'MsgBox ssql
Set TBM = DBM.OpenRecordset(Ssql)
If Not TBM.EOF Then
    Do While Not TBM.EOF
        If PackingLots = False Then
        TBT.AddNew
            TBT.Fields("RecNo").Value = TBM.Fields("Rec_No").Value
            TBT.Fields("SDate").Value = SDate
            TBT.Fields("EDate").Value = EDate
            TBT.Fields("GDate").Value = TBM.Fields("Date_Recieve").Value
            TBT.Fields("LotNo").Value = TBM.Fields("Lot_No").Value
            TBT.Fields("DyingCode").Value = TBM.Fields("Dying_Code").Value
            TBT.Fields("DyingName").Value = DName
            TBT.Fields("GFectoryName") = blm.Factory(TBM.Fields("FAC_Code").Value)
            If Val(ItemCode) > 0 Then
                TBT.Fields("ClothCode").Value = TBM.Fields("Cloth_Code").Value
                TBT.Fields("Clothname").Value = CName
            End If
            TBT.Fields("GThans").Value = TBM.Fields("THANS").Value
            TBT.Fields("GGAZANA").Value = TBM.Fields("Gazana").Value
            If Val(ItemCode) = 0 Then
                TBT.Fields("GQuality").Value = blm.FillCloth1(TBM.Fields("Cloth_Code").Value)
            Else
                TBT.Fields("GQuality").Value = CName
            End If
            TBT.Fields("ProgramGiven").Value = TBM.Fields("Program").Value
            
            If Not IsNull(TBM.Fields("DGazana").Value) Then
                TBT.Fields("DGazana").Value = TBM.Fields("DGazana").Value
            End If
            If Not IsNull(TBM.Fields("packGazana").Value) Then
                TBT.Fields("PGazana").Value = TBM.Fields("PackGazana").Value
            End If
            If Not IsNull(TBM.Fields("packThanz").Value) Then
                TBT.Fields("PackThans").Value = TBM.Fields("PackThanz").Value
            End If
            If Not IsNull(TBM.Fields("Lot_No")) Then
                TBT.Fields("Paid").Value = PaymentStatus(TBM.Fields("Dying_Code").Value, TBM.Fields("Lot_No").Value)
            End If
            If Not IsNull(TBM.Fields("Dispute")) Then
                TBT.Fields("Dispute").Value = TBM.Fields("Dispute").Value & ""
            End If
        TBT.Update
        End If
        
        If PackingLots = True Then
        If IsNull(TBM.Fields("packGazana").Value) Then
        TBT.AddNew
            TBT.Fields("SDate").Value = SDate
            TBT.Fields("EDate").Value = EDate
            TBT.Fields("GDate").Value = TBM.Fields("Date_Recieve").Value
            TBT.Fields("LotNo").Value = TBM.Fields("Lot_No").Value
            TBT.Fields("DyingCode").Value = TBM.Fields("Dying_Code").Value
            TBT.Fields("DyingName").Value = FName
            TBT.Fields("GFectoryName") = blm.Factory(TBM.Fields("FAC_Code").Value)
            If Val(ItemCode) > 0 Then
                TBT.Fields("ClothCode").Value = TBM.Fields("Cloth_Code").Value
                TBT.Fields("Clothname").Value = CName
            End If
            TBT.Fields("GThans").Value = TBM.Fields("THANS").Value
            TBT.Fields("GGAZANA").Value = TBM.Fields("Gazana").Value
            If Val(ItemCode) = 0 Then
                TBT.Fields("GQuality").Value = blm.FillCloth1(TBM.Fields("Cloth_Code").Value)
            Else
                TBT.Fields("GQuality").Value = CName
            End If
            TBT.Fields("ProgramGiven").Value = TBM.Fields("Program").Value
            
            If Not IsNull(TBM.Fields("DGazana").Value) Then
                TBT.Fields("DGazana").Value = TBM.Fields("DGazana").Value
            End If
            If Not IsNull(TBM.Fields("packGazana").Value) Then
                TBT.Fields("PGazana").Value = TBM.Fields("PackGazana").Value
            End If
            If Not IsNull(TBM.Fields("packThanz").Value) Then
                TBT.Fields("PackThans").Value = TBM.Fields("PackThanz").Value
            End If
            If Not IsNull(TBM.Fields("Lot_No")) Then
                TBT.Fields("Paid").Value = PaymentStatus(TBM.Fields("Dying_Code").Value, TBM.Fields("Lot_No").Value)
            End If
            If Not IsNull(TBM.Fields("Dispute")) Then
                TBT.Fields("Dispute").Value = TBM.Fields("Dispute").Value & ""
            End If
        TBT.Update
        End If
        End If
        
    TBM.MoveNext
    Loop
End If
TBM.Close
TBT.Close
DBT.Close
DBM.Close


End Sub

Public Sub GreyReport(SDate As Date, EDate As Date, FactoryCode As Long, ItemCode As Long)
Dim TBT As DAO.Recordset
Dim DBT As DAO.Database
Dim DBM As DAO.Database
Dim TBM As DAO.Recordset
Dim FName As String, CName As String
FName = blm.Factory(FactoryCode)
CName = blm.FillCloth1(ItemCode)
Set DBT = OpenDatabase(App.Path & "\Book.mdb")
Ssql = "Delete from GreyRep"
DBT.Execute Ssql

Set DBM = OpenDatabase(blm.pathMain)
Set TBT = DBT.OpenRecordset("Greyrep", dbOpenTable)

'Yarn Issue
Ssql = "Select * from YarnIssue where Issue_Date Between #" & SDate & "# and #" & EDate & "#"
Ssql = Ssql & " and FAC_Code=" & FactoryCode
Ssql = Ssql & " and Cloth_Code=" & ItemCode
Ssql = Ssql & " Order by Issue_Date,Issue_No"

Set TBM = DBM.OpenRecordset(Ssql)
If Not TBM.EOF Then
    Do While Not TBM.EOF

        TBT.AddNew
            TBT.Fields("SDate").Value = SDate
            TBT.Fields("EDate").Value = EDate
            TBT.Fields("VDate").Value = TBM.Fields("Issue_Date").Value
            TBT.Fields("VNo").Value = TBM.Fields("Issue_No").Value
            TBT.Fields("FACTORYCode").Value = TBM.Fields("FAc_Code").Value
            TBT.Fields("FACTORYName").Value = FName
            TBT.Fields("ItemCode").Value = TBM.Fields("Cloth_Code").Value
            TBT.Fields("ItemName").Value = CName
            TBT.Fields("TANA").Value = TBM.Fields("T_Qty").Value
            TBT.Fields("BANA").Value = TBM.Fields("B_Qty").Value
            TBT.Fields("GAZANA").Value = 0
            TBT.Fields("Amount").Value = 0
            
            
        TBT.Update
    TBM.MoveNext
    Loop
End If
TBM.Close

'Cloth Rec
Ssql = "Select * from ClothRec where Date_Recieve Between #" & SDate & "# and #" & EDate & "#"
Ssql = Ssql & " and FAC_Code=" & FactoryCode
Ssql = Ssql & " and Cloth_Code=" & ItemCode
Ssql = Ssql & " Order by Date_Recieve,Rec_No"

Set TBM = DBM.OpenRecordset(Ssql)
If Not TBM.EOF Then
    Do While Not TBM.EOF

        TBT.AddNew
            TBT.Fields("SDate").Value = SDate
            TBT.Fields("EDate").Value = EDate
            TBT.Fields("VDate").Value = TBM.Fields("Date_Recieve").Value
            TBT.Fields("VNo").Value = TBM.Fields("Rec_No").Value
            TBT.Fields("FACTORYCode").Value = TBM.Fields("FAc_Code").Value
            TBT.Fields("FACTORYName").Value = FName
            TBT.Fields("ItemCode").Value = TBM.Fields("Cloth_Code").Value
            TBT.Fields("Itemname").Value = CName
            TBT.Fields("TANA").Value = 0
            TBT.Fields("BANA").Value = 0
            TBT.Fields("GAZANA").Value = TBM.Fields("Gazana").Value
            TBT.Fields("Amount").Value = 0
            TBT.Fields("DyingName").Value = blm.Dying(TBM.Fields("Dying_Code").Value)
            
        TBT.Update
    TBM.MoveNext
    Loop
End If
TBM.Close


'Amount
Ssql = "Select * from PaymentLoom where Date Between #" & SDate & "# and #" & EDate & "#"
Ssql = Ssql & " and FAC_Code=" & FactoryCode
Ssql = Ssql & " and Cloth_Code=" & ItemCode
Ssql = Ssql & " Order by Date,VOU_No"

Set TBM = DBM.OpenRecordset(Ssql)
If Not TBM.EOF Then
    Do While Not TBM.EOF

        TBT.AddNew
            TBT.Fields("SDate").Value = SDate
            TBT.Fields("EDate").Value = EDate
            TBT.Fields("VDate").Value = TBM.Fields("Date").Value
            TBT.Fields("VNo").Value = TBM.Fields("Vou_No").Value
            TBT.Fields("FACTORYCode").Value = TBM.Fields("FAc_Code").Value
            TBT.Fields("FACTORYName").Value = FName
            TBT.Fields("ItemCode").Value = TBM.Fields("Cloth_Code").Value
            TBT.Fields("ItemName").Value = CName
            TBT.Fields("TANA").Value = 0
            TBT.Fields("BANA").Value = 0
            TBT.Fields("GAZANA").Value = 0
            TBT.Fields("Amount").Value = TBM.Fields("AMOUNT").Value
            TBT.Fields("Description").Value = TBM.Fields("Remarks").Value & ""
            TBT.Fields("Zaty").Value = TBM.Fields("Zaty").Value & ""
            TBT.Fields("SS").Value = TBM.Fields("Social").Value & ""
            TBT.Fields("Oil").Value = TBM.Fields("Oil").Value & ""
            TBT.Fields("Misc").Value = TBM.Fields("Misc").Value & ""
            
        TBT.Update
    TBM.MoveNext
    Loop
End If
TBM.Close


TBT.Close
DBT.Close
DBM.Close

End Sub

