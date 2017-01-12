Attribute VB_Name = "Blm_Mdl"
Global UserN As String
Public Type CTN
    Lot_no As Long
    Bobins As Long
    W_H_Code As Byte
    Qty As Currency
    Rate As Currency
    ItemCode As Long
End Type
Public Type ItemInfo
    Item As String
    Unit As String
    GroupCode As Integer
    GroupName As String
End Type
Global CTNS() As Long
Global WHs() As Byte
Global Selects As Long
Global NewStatus As Boolean
