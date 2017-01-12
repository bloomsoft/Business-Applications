Attribute VB_Name = "bloom"
Private Blm As New bloom1
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_TAB = &H9

Global Const CPUID As String = "BFEBFBFF00000F33" 'hrblm
'Global Const CPUID As String = "BFEBFBFF00000F33" 'A & H
'Global Const CPUID As String = "0387FBFF0000068A"
'Global Const CPUID As String = "BFEBFBFF00000F34" 'Arfat Digita
Global ledgerhide As Boolean
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Type CottonIssue
    AvgRate As Double
    AvgBaleWeight As Double
End Type
Public Const EmpHead As String = "32"
Public Const BanksHead As String = "17001"
Global FStartDate As Date
Global FEndDate As Date
Global SelectedItemCode As Double
Global SelectedItemName As String
Global SelectedSHCode As Double
Global SelectedSHName As String
Global SelectedAccountCode As Double
Global SelectedAccountName As String
Global SelectedVType As Integer
Global SelectedReportTitle As String
Global YearN As Integer

Public Sub GetDates()
    Dim Ssql As String
    Dim TB As Recordset
Dim Result As CottonIssue
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)

Ssql = "Select * from FDates"
Set TB = DB.OpenRecordset(Ssql)
If Not TB.EOF Then
    FStartDate = TB.Fields("StartDate").Value
    FEndDate = TB.Fields("EndDate").Value
End If
TB.Close
DB.Close
End Sub
Public Function AvgRateAndWeight(EndDate As Date, ItemCode As Long) As CottonIssue
Dim Ssql As String
Dim RST As Recordset
Dim Result As CottonIssue
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)
Dim OpQty As Double, OpBales As Double
Dim PurchasedQty As Double, SoldQty As Double, IssuedQty As Double, SoldRQty As Double
Dim PurchasedBales As Double, SoldBales As Double, IssuedBales As Double, SoldRBales As Double
Dim PurRQty As Double, PurRBales As Double
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim GatherdPurchaseQty As Double, GatherdPurchaseBales As Double, GatherdPurchaseAmount As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim QtyStock As Double, BalesStock As Double

OpQty = Blm.Opstocks(ItemCode)
OpBales = Blm.opStockbales(ItemCode)
'Total Purchases
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Purchase where Item=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    PurchasedQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    PurchasedBales = RST.Fields("B").Value
End If
RST.Close

'Total Purchases
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from PurchaseReturn where Item=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    PurRQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    PurRBales = RST.Fields("B").Value
End If
RST.Close

'Total Sales
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Sales where Item=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    SoldQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    SoldBales = RST.Fields("B").Value
End If
RST.Close

'Total Sales Return
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from SalesReturn where Item=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    SoldRQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    SoldRBales = RST.Fields("B").Value
End If
RST.Close

'Total Issues
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Issue where ItemCode=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    IssuedQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    IssuedBales = RST.Fields("B").Value
End If
RST.Close

'Stock Till End Date
QtyStock = (OpQty + PurchasedQty + SoldRQty) - (SoldQty + IssuedQty + PurRQty)
QtyBales = (OpBales + PurchasedBales + SoldRBales) - (SoldBales + IssuedBales + PurRBales)
'MsgBox "Test"
Ssql = "Select * from Purchase where V_Date <= #" & EndDate & "# and Item=" & ItemCode & " Order by V_Date Desc"
Set RST = DB.OpenRecordset(Ssql)
'MsgBox Ssql
MsgBox QtyStock
If Not RST.EOF Then
    Do While Not RST.EOF
    
    GatherdPurchaseQty = GatherdPurchaseQty + RST.Fields("Qty")
    GatherdPurchaseBales = GatherdPurchaseBales + RST.Fields("Bales")
    GatherdPurchaseAmount = GatherdPurchaseAmount + ((RST.Fields("Qty") * RST.Fields("Rate").Value) + Val(RST.Fields("Freight").Value & ""))
    
    If GatherdPurchaseQty >= QtyStock Then
        MsgBox "Test"
        Exit Do
    End If
    RST.MoveNext
    Loop
End If
'MsgBox GatherdPurchaseBales
RST.Close
DB.Close
'MsgBox "Test2"
'Purchase Qty's AvgBale Weight and Rate / Kg
Dim AvgBaleWT As Double, AvgRate As Double
If GatherdPurchaseQty > 0 And GatherdPurchaseBales > 0 Then
AvgBaleWT = GatherdPurchaseQty / GatherdPurchaseBales
AvgRate = GatherdPurchaseAmount / GatherdPurchaseQty
End If
'MsgBox "Test"
Dim Diff As Double
If GatherdPurchaseQty > QtyStock And AvgBaleWT > 0 Then
    Diff = GatherdPurchaseQty - QtyStock
    GatherdPurchaseQty = GatherdPurchaseQty - Diff
    
    GatherdPurchaseBales = Round(GatherdPurchaseBales - (Diff / AvgBaleWT))
    GatherdPurchaseAmount = GatherdPurchaseAmount - (AvgRate * Diff)
End If
'MsgBox "Test3"
Dim FinalAvgBaleWT As Double
Dim FinalAvgRate As Double
If GatherdPurchaseQty > 0 And GatherdPurchaseBales > 0 Then
FinalAvgBaleWT = GatherdPurchaseQty / GatherdPurchaseBales
FinalAvgRate = GatherdPurchaseAmount / GatherdPurchaseQty

Result.AvgBaleWeight = FinalAvgBaleWT
Result.AvgRate = FinalAvgRate
'MsgBox "TestRate"
AvgRateAndWeight = Result
Else
AvgRateAndWeight = Result
End If

End Function
Public Function AvgRateAndWeightHeadWise(EndDate As Date, ItemCode As Long) As CottonIssue
Dim Ssql As String
Dim RST As Recordset
Dim Result As CottonIssue
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)
Dim PurchasedQty As Double, SoldQty As Double, IssuedQty As Double
Dim PurchasedBales As Double, SoldBales As Double, IssuedBales As Double
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim GatherdPurchaseQty As Double, GatherdPurchaseBales As Double, GatherdPurchaseAmount As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim QtyStock As Double, BalesStock As Double
'Total Purchases
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Purchase where Val(mid(Item,1,2))=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    PurchasedQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    PurchasedBales = RST.Fields("B").Value
End If
RST.Close

'Total Sales
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Sales where Val(mid(Item,1,2))=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    SoldQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    SoldBales = RST.Fields("B").Value
End If
RST.Close

'Total Issues
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Issue where Val(mid(ItemCode,1,2))=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    IssuedQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    IssuedBales = RST.Fields("B").Value
End If
RST.Close

'Stock Till End Date
QtyStock = PurchasedQty - (SoldQty + IssuedQty)
QtyBales = PurchasedBales - (SoldBales + IssuedBales)
'MsgBox "Test"
Ssql = "Select * from Purchase where V_Date <= #" & EndDate & "# and Val(mid(Item,1,2))=" & ItemCode & " Order by V_Date Desc"
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    Do While Not RST.EOF
    GatherdPurchaseQty = GatherdPurchaseQty + RST.Fields("Qty")
    GatherdPurchaseBales = GatherdPurchaseBales + RST.Fields("Bales")
    GatherdPurchaseAmount = GatherdPurchaseAmount + ((RST.Fields("Qty") * RST.Fields("Rate").Value) + Val(RST.Fields("Freight").Value & ""))
    'MsgBox GatherdPurchaseAmount
    If GatherdPurchaseQty >= QtyStock Then
        Exit Do
    End If
    RST.MoveNext
    Loop
End If
RST.Close
DB.Close
'MsgBox "Test2"
'Purchase Qty's AvgBale Weight and Rate / Kg
Dim AvgBaleWT As Double, AvgRate As Double
If GatherdPurchaseQty > 0 And GatherdPurchaseBales > 0 Then
AvgBaleWT = GatherdPurchaseQty / GatherdPurchaseBales
AvgRate = GatherdPurchaseAmount / GatherdPurchaseQty
End If
'MsgBox "Test"
Dim Diff As Double
If GatherdPurchaseQty > QtyStock Then
    Diff = GatherdPurchaseQty - QtyStock
    GatherdPurchaseQty = GatherdPurchaseQty - Diff
    
    GatherdPurchaseBales = Round(GatherdPurchaseBales - (Diff / AvgBaleWT))
    GatherdPurchaseAmount = GatherdPurchaseAmount - (AvgRate * Diff)
End If
'MsgBox "Test3"
Dim FinalAvgBaleWT As Double
Dim FinalAvgRate As Double
If GatherdPurchaseQty > 0 And GatherdPurchaseBales > 0 Then
FinalAvgBaleWT = GatherdPurchaseQty / GatherdPurchaseBales
FinalAvgRate = GatherdPurchaseAmount / GatherdPurchaseQty

Result.AvgBaleWeight = FinalAvgBaleWT
Result.AvgRate = FinalAvgRate
'MsgBox "TestRate"
AvgRateAndWeightHeadWise = Result
Else
AvgRateAndWeightHeadWise = Result
End If

End Function
Public Function AvgRateAndWeightSubHeadWise(EndDate As Date, ItemCode As Long) As CottonIssue
Dim Ssql As String
Dim RST As Recordset
Dim Result As CottonIssue
Dim DB As Database
Set DB = OpenDatabase(Blm.patHmain)
Dim PurchasedQty As Double, SoldQty As Double, IssuedQty As Double
Dim PurchasedBales As Double, SoldBales As Double, IssuedBales As Double
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim GatherdPurchaseQty As Double, GatherdPurchaseBales As Double, GatherdPurchaseAmount As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim QtyStock As Double, BalesStock As Double
'Total Purchases
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Purchase where Val(mid(Item,1,5))=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    PurchasedQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    PurchasedBales = RST.Fields("B").Value
End If
RST.Close

'Total Sales
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Sales where Val(mid(Item,1,5))=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    SoldQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    SoldBales = RST.Fields("B").Value
End If
RST.Close

'Total Issues
Ssql = "Select Sum(Qty) as Q,Sum(Bales) as B from Issue where Val(mid(ItemCode,1,5))=" & ItemCode & " and V_Date <= #" & EndDate & "#"
Set RST = DB.OpenRecordset(Ssql)
If Not IsNull(RST.Fields("Q").Value) Then
    IssuedQty = RST.Fields("Q").Value
End If
If Not IsNull(RST.Fields("B").Value) Then
    IssuedBales = RST.Fields("B").Value
End If
RST.Close

'Stock Till End Date
QtyStock = PurchasedQty - (SoldQty + IssuedQty)
QtyBales = PurchasedBales - (SoldBales + IssuedBales)
'MsgBox "Test"
Ssql = "Select * from Purchase where V_Date <= #" & EndDate & "# and Val(mid(Item,1,5))=" & ItemCode & " Order by V_Date Desc"
Set RST = DB.OpenRecordset(Ssql)
If Not RST.EOF Then
    Do While Not RST.EOF
    GatherdPurchaseQty = GatherdPurchaseQty + RST.Fields("Qty")
    GatherdPurchaseBales = GatherdPurchaseBales + RST.Fields("Bales")
    GatherdPurchaseAmount = GatherdPurchaseAmount + ((RST.Fields("Qty") * RST.Fields("Rate").Value) + Val(RST.Fields("Freight").Value & ""))
    'MsgBox GatherdPurchaseAmount
    If GatherdPurchaseQty >= QtyStock Then
        Exit Do
    End If
    RST.MoveNext
    Loop
End If
RST.Close
DB.Close
'MsgBox "Test2"
'Purchase Qty's AvgBale Weight and Rate / Kg
Dim AvgBaleWT As Double, AvgRate As Double
If GatherdPurchaseQty > 0 And GatherdPurchaseBales > 0 Then
AvgBaleWT = GatherdPurchaseQty / GatherdPurchaseBales
AvgRate = GatherdPurchaseAmount / GatherdPurchaseQty
End If
'MsgBox "Test"
Dim Diff As Double
If GatherdPurchaseQty > QtyStock Then
    Diff = GatherdPurchaseQty - QtyStock
    GatherdPurchaseQty = GatherdPurchaseQty - Diff
    If AvgBaleWT > 0 Then
        GatherdPurchaseBales = Round(GatherdPurchaseBales - (Diff / AvgBaleWT))
    End If
    GatherdPurchaseAmount = GatherdPurchaseAmount - (AvgRate * Diff)
End If
'MsgBox "Test3"
Dim FinalAvgBaleWT As Double
Dim FinalAvgRate As Double
If GatherdPurchaseQty > 0 And GatherdPurchaseBales > 0 Then
FinalAvgBaleWT = GatherdPurchaseQty / GatherdPurchaseBales
FinalAvgRate = GatherdPurchaseAmount / GatherdPurchaseQty

Result.AvgBaleWeight = FinalAvgBaleWT
Result.AvgRate = FinalAvgRate
'MsgBox "TestRate"
AvgRateAndWeightSubHeadWise = Result
Else
AvgRateAndWeightSubHeadWise = Result
End If

End Function

Public Function Eval(S As String) As Double
Dim p As String
Dim Flag As Integer '1=*, 2=+, 3=-,4=/
Dim FirstValue As Double
Dim SecondValue As Double
p = InStr(1, S, "*", vbBinaryCompare)
If p > 0 Then
    Flag = 1
End If
If Flag = 0 Then
p = InStr(1, S, "+", vbBinaryCompare)
If p > 0 Then
    Flag = 2
End If
End If
If Flag = 0 Then
p = InStr(1, S, "-", vbBinaryCompare)
If p > 0 Then
    Flag = 3
End If
End If
If Flag = 0 Then
p = InStr(1, S, "/", vbBinaryCompare)
If p > 0 Then
    Flag = 4
End If
End If
If Flag = 0 Then
    Eval = 0
    Exit Function
End If
FirstValue = Left(S, p - 1)
SecondValue = Mid(S, p + 1, Len(S) - p)
'MsgBox Flag & " " & FirstValue & " " & SecondValue

If Flag = 1 Then Eval = FirstValue * SecondValue
If Flag = 2 Then Eval = FirstValue + SecondValue
If Flag = 3 Then Eval = FirstValue - SecondValue
If Flag = 4 Then Eval = FirstValue / SecondValue
End Function
Private Function GetPCName() As String
Dim computer_name As String
Dim length As Long

    computer_name = Space$(256)
    length = Len(computer_name)
    GetComputerName computer_name, length
    computer_name = Left$(computer_name, length)

    GetPCName = computer_name

End Function
Sub Main()
    Dim objWMIService, colItems
        Set objWMIService = GetObject("winmgmts:\\" & GetPCName & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
        For Each objItem In colItems
        ' MsgBox "Name: " & objItem.Name & " " & objItem.DeviceID & " " & objItem.UniqueId & " " & objItem.PNPDeviceID & " " & "ProcessorId: " & objItem.ProcessorId
        If CPUID <> objItem.ProcessorId Then
            MsgBox "You are Not Licensed to Use This Software on This PC"
            End
        Else
            frmSplash.Show
        End If
        '        WScript.Echo "AddressWidth: " & objItem.AddressWidth
'        WScript.Echo "Architecture: " & objItem.Architecture
'        WScript.Echo "Availability: " & objItem.Availability
'        WScript.Echo "Caption: " & objItem.Caption
'        WScript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
'        WScript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
'        WScript.Echo "CpuStatus: " & objItem.CpuStatus
'        WScript.Echo "CreationClassName: " & objItem.CreationClassName
'        WScript.Echo "CurrentClockSpeed: " & objItem.CurrentClockSpeed
'        WScript.Echo "CurrentVoltage: " & objItem.CurrentVoltage
'        WScript.Echo "DataWidth: " & objItem.DataWidth
'        WScript.Echo "Description: " & objItem.Description
'        WScript.Echo "DeviceID: " & objItem.DeviceID
'        WScript.Echo "ErrorCleared: " & objItem.ErrorCleared
'        WScript.Echo "ErrorDescription: " & objItem.ErrorDescription
'        WScript.Echo "ExtClock: " & objItem.ExtClock
'        WScript.Echo "Family: " & objItem.Family
'        WScript.Echo "InstallDate: " & objItem.InstallDate
'        WScript.Echo "L2CacheSize: " & objItem.L2CacheSize
'        WScript.Echo "L2CacheSpeed: " & objItem.L2CacheSpeed
'        WScript.Echo "LastErrorCode: " & objItem.LastErrorCode
'        WScript.Echo "Level: " & objItem.Level
'        WScript.Echo "LoadPercentage: " & objItem.LoadPercentage
'        WScript.Echo "Manufacturer: " & objItem.Manufacturer
'        WScript.Echo "MaxClockSpeed: " & objItem.MaxClockSpeed
'        WScript.Echo "Name: " & objItem.Name
'        WScript.Echo "OtherFamilyDescription: " & objItem.OtherFamilyDescription
'        WScript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
'        WScript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
'        WScript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
'        WScript.Echo "ProcessorId: " & objItem.ProcessorId
'        WScript.Echo "ProcessorType: " & objItem.ProcessorType
'        WScript.Echo "Revision: " & objItem.Revision
'        WScript.Echo "Role: " & objItem.Role
'        WScript.Echo "SocketDesignation: " & objItem.SocketDesignation
'        WScript.Echo "Status: " & objItem.Status
'        WScript.Echo "StatusInfo: " & objItem.StatusInfo
'        WScript.Echo "Stepping: " & objItem.Stepping
'        WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
'        WScript.Echo "SystemName: " & objItem.SystemName
'        WScript.Echo "UniqueId: " & objItem.UniqueId
'        WScript.Echo "UpgradeMethod: " & objItem.UpgradeMethod
'        WScript.Echo "Version: " & objItem.Version
'        WScript.Echo "VoltageCaps: " & objItem.VoltageCaps
        Next
End Sub

Public Sub comb_fill(CNTL As Control, Ssql As String)
Dim DB As Database
Dim TB As Recordset

Set DB = OpenDatabase(Blm.patHmain)
Set TB = DB.OpenRecordset(Ssql)
'MsgBox Ssql
CNTL.clear
If Not TB.EOF Then
    Do While Not TB.EOF
        CNTL.AddItem TB.Fields("name").Value
        CNTL.ItemData(CNTL.NewIndex) = TB.Fields("code").Value
        TB.MoveNext
    Loop
CNTL.ListIndex = 0
End If
TB.Close
DB.Close

End Sub
