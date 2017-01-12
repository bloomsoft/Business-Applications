VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "This Computer's CPU ID"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "CPU ID"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
' Name: All processor information can vb
'     e retireved by WMIs Win32_Processor clas
'     s
' Description:Win32_Processor
' By: Jo Nassen
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.41907/lngWId.1/qx
'     /vb/scripts/ShowCode.htm
'for details.
'**************************************

' **************************************
'     **************************************
' This script lists the processors confi
'     guration of a remote or local computer,
'
' like processor type, bank architecture
'     , clock speed, L2 cache, manufacturer, e
'     tc.
' **************************************
'     **************************************
' Goto http://www.activxperts.com/activm
'     onitor and click on WMI Samples
' for more samples
' **************************************
'     **************************************

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Function GetPCName() As String
Dim computer_name As String
Dim length As Long

    computer_name = Space$(256)
    length = Len(computer_name)
    GetComputerName computer_name, length
    computer_name = Left$(computer_name, length)

    GetPCName = computer_name

End Function
Sub ListProcessorProperties()
    Dim objWMIService, colItems
        Set objWMIService = GetObject("winmgmts:\\" & GetPCName & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
        For Each objItem In colItems
        ' MsgBox "Name: " & objItem.Name & " " & objItem.DeviceID & " " & objItem.UniqueId & " " & objItem.PNPDeviceID & " " & "ProcessorId: " & objItem.ProcessorId
        Text1.Text = objItem.ProcessorId
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
        

Private Sub Form_Load()
ListProcessorProperties
End Sub
