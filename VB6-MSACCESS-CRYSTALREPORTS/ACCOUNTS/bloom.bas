Attribute VB_Name = "bloom"
Private blm As bloom1

'Global Const CPUID As String = "BFEBFBFF00000F33" 'hrblm
Global Const CPUID As String = "3FEBFBFF00000F12" 'A & H
'Global Const CPUID As String = "0387FBFF0000068A" 'AlAzeem
'Global Const CPUID As String = "BFEBFBFF00000F34" 'Arfat Digita
'Global Const CPUID As String = "0383F9FF0000068A" ' zm eNTERPRIZES
'Global Const CPUID2 As String = "0000000000000683" 'zm Laptop
'Global Const CPUID As String = "BFEBFBFF00000F47" 'BUZZ
Global ledgerhide As Boolean
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function Eval(s As String) As Double
Dim p As String
Dim Flag As Integer '1=*, 2=+, 3=-,4=/
Dim FirstValue As Double
Dim SecondValue As Double
p = InStr(1, s, "*", vbBinaryCompare)
If p > 0 Then
    Flag = 1
End If
If Flag = 0 Then
p = InStr(1, s, "+", vbBinaryCompare)
If p > 0 Then
    Flag = 2
End If
End If
If Flag = 0 Then
p = InStr(1, s, "-", vbBinaryCompare)
If p > 0 Then
    Flag = 3
End If
End If
If Flag = 0 Then
p = InStr(1, s, "/", vbBinaryCompare)
If p > 0 Then
    Flag = 4
End If
End If
If Flag = 0 Then
    Eval = 0
    Exit Function
End If
FirstValue = Left(s, p - 1)
SecondValue = Mid(s, p + 1, Len(s) - p)
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
        If CPUID = objItem.ProcessorId Then
        frmSplash.Show
        ElseIf CPUID2 = objItem.ProcessorId Then
        frmSplash.Show
        Else
            MsgBox "You are Not Licensed to Use This Software on This PC"
            End
    
            
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

Public Sub comb_fill(cntl As Control, Ssql As String)
Dim DB As Database
Dim tb As Recordset
Set blm = New bloom1
Set DB = OpenDatabase(blm.patHmain)
Set tb = DB.OpenRecordset(Ssql)
If Not tb.EOF Then
cntl.clear
    Do While Not tb.EOF
        cntl.AddItem tb.Fields("name").Value
        cntl.ItemData(cntl.NewIndex) = tb.Fields("code").Value
        tb.MoveNext
    Loop
cntl.ListIndex = 0
End If
tb.Close
DB.Close
Set blm = Nothing
End Sub
