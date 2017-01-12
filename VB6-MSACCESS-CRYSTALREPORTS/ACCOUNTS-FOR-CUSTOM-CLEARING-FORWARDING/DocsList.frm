VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form DocsList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipping Documents List"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView ListView1 
      Height          =   5385
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   9499
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Serial No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "L/C No."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Goods"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Import Value"
         Object.Width           =   3175
      EndProperty
   End
End
Attribute VB_Name = "DocsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FillList()
Dim tb As Recordset
Dim DB As Database
Dim Ssql As String
Dim LItem As ListItem
If SelPartyCode > 0 Then
    Ssql = "Select * from Docs where PartyCode=" & SelPartyCode & " and Status=0"
Else
    Ssql = "Select * from Docs"
End If
Set DB = OpenDatabase(App.Path & "\Bloom.mdb")
Set tb = DB.OpenRecordset(Ssql)
ListView1.ListItems.clear
If Not tb.EOF Then
Do While Not tb.EOF

    Set LItem = ListView1.ListItems.Add(, , tb.Fields("SrNo").Value)
    LItem.SubItems(1) = tb.Fields("LCNo").Value & ""
    LItem.SubItems(2) = tb.Fields("Goods").Value & ""
    LItem.SubItems(3) = tb.Fields("ImportValue").Value & ""
    
tb.MoveNext
Loop
End If
tb.Close
DB.Close
End Sub
Private Sub Form_Load()
FillList

End Sub

Private Sub ListView1_DblClick()
Me.Hide
Unload Me

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
SelSerialNo = Val(Item.Text)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Hide
    Unload Me
End If
End Sub
