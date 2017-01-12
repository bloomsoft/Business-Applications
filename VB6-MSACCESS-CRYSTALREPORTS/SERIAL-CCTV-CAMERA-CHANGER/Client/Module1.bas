Attribute VB_Name = "Module1"
Global ServerIP As String
Public FS As New Scripting.FileSystemObject
Sub Main()
Dim TS As TextStream

Set TS = FS.OpenTextFile(App.Path & "\ServerIP.txt", ForReading)
ServerIP = TS.ReadAll
TS.Close

Load Form1
Form1.Show
End Sub
