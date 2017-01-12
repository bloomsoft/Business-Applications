Attribute VB_Name = "Module1"
'Global Const CPUID As String = "BFEBFBFF00000F33" 'hrblm
Global Const CPUID As String = "0387FBFF00000686" 'Dilara
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Global BackupPath As String
Global SelPartyCode As Long
Global SelSerialNo As Long
Global Selpartyname As String

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
            'login2.Show
        Else
            login2.Show
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

