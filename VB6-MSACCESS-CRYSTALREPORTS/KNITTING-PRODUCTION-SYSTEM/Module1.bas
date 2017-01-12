Attribute VB_Name = "Module1"
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_TAB = &H9
Public CN As New ADODB.Connection
Global SelEmpName As String
Public Sub Main()
CN.CursorLocation = adUseClient
CN.Open "Provider=MSDAORA.1;Password=mlb;User ID=BLOOM;Data Source=alpha;Persist Security Info=False"
'CN.Open " Provider=MSDAORA.1;Password=mlb;User ID=bloom;Data Source=beq-local;Persist Security Info=True"
End Sub

Public Sub GOTF(ct As TextBox)
ct.BackColor = &HFFFF&
End Sub
Public Sub Lostf(ct As TextBox)
ct.BackColor = &HFFFFFF
End Sub

