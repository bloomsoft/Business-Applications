VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form YearClose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year Closing"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "&Transfer Balances from Current Year to New Year"
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "&Create New Year Database"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "YearClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FS As New FileSystemObject
Private Blmr As New bloom_r
Private Sub Command1_Click()
Dim Result As VbMsgBoxResult
Result = MsgBox("Do You Realy Want to Create The New Year", vbYesNo)
If Result = vbYes Then
    If FS.FolderExists(App.path & "\Years\" & YearN + 1) Then
        Result = MsgBox("It Seems You Already Have the Next Year Database, If You Proceed Your All Entries in New Year will be Disturbed, If You Still want to Continue Press (YES)", vbYesNo)
        If Result = vbNo Then
            Exit Sub
        End If
        FS.DeleteFile App.path & "\Years\" & YearN + 1 & "\Bloom.mdb", True
        DoEvents
    Else
        FS.CreateFolder App.path & "\Years\" & YearN + 1
        DoEvents
        
    End If
    FS.CopyFile App.path & "\Years\" & YearN & "\Bloom.mdb", App.path & "\Years\" & YearN + 1 & "\Bloom.mdb", True
    DoEvents
    
    Dim DB As Database
    Dim Ssql As String
    
    Set DB = OpenDatabase(App.path & "\years\" & YearN + 1 & "\Bloom.mdb")
    
    Ssql = "Delete from Alarm"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Arrivals"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Consume"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from CostSheet"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Dispatches"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from EmpAtd"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Expences"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Issue"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from IssueSH"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from OverTime"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from PL"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Production"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Purchase"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from PurchaseReturn"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from PurJob"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from SaleJob"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from Sales"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from SalesReturn"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from SalVoucher"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from SHProduction"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from VouDTL"
    DB.Execute Ssql
    DoEvents
    
    Ssql = "Delete from VouMST"
    DB.Execute Ssql
    DoEvents
    
    
    DB.Close
    
    
    
Else
    Exit Sub
End If
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
    Dim DBN As Database
    Dim DB As Database
    Dim Ssql As String
    Dim RSN As Recordset
    Dim RS As Recordset
    Set DBN = OpenDatabase(App.path & "\years\" & YearN + 1 & "\Bloom.mdb")
    Set DB = OpenDatabase(App.path & "\years\" & YearN & "\Bloom.mdb")
    
    Ssql = "Update FDates Set StartDate=#" & CDate("Jul/01/" & YearN + 1) & "#,EndDate=#" & CDate("Jun/30/" & YearN + 2) & "#"
    DBN.Execute Ssql
    DoEvents
    
    Set RSN = DBN.OpenRecordset("VouMST", dbOpenTable)
    RSN.AddNew
        RSN.Fields("v_date").Value = CDate("Jul/01/" & YearN)
        RSN.Fields("v_type").Value = 10
        RSN.Fields("v_no").Value = 1
        RSN.Fields("narration").Value = "Open Balance"
    RSN.Update
    RSN.Close
    
    Set RSN = DBN.OpenRecordset("VouDTL", dbOpenTable)
    Ssql = "Select Party,Sum(Debit)-Sum(Credit) as Bal from VouDTL Group By Party"
    Set RS = DB.OpenRecordset(Ssql)
    If Not RS.EOF Then
        Do While Not RS.EOF
        
            RSN.AddNew
                RSN.Fields("v_date").Value = CDate("Jul/01/" & YearN)
                RSN.Fields("v_type").Value = 10
                RSN.Fields("party").Value = RS.Fields("Party").Value
                RSN.Fields("Remarks").Value = "Opening Balance " & Format(CDate("Jul/01/" & YearN), "dd-MMM-yyyy")
                If RS.Fields("Bal").Value > 0 Then
                    RSN.Fields("debit").Value = RS.Fields("Bal").Value
                    RSN.Fields("credit").Value = 0
                End If
                If RS.Fields("Bal").Value < 0 Then
                    RSN.Fields("debit").Value = 0
                    RSN.Fields("credit").Value = RS.Fields("Bal").Value * -1
                End If
                
            RSN.Update
            
            Ssql = "Update Acchart Set "
            If RS.Fields("Bal") > 0 Then
                Ssql = Ssql & "Debit=" & RS.Fields("Bal").Value & ",Credit=0"
            End If
            If RS.Fields("Bal") < 0 Then
                Ssql = Ssql & "Debit=0,Credit=" & RS.Fields("Bal").Value * -1
            End If
            Ssql = Ssql & " where Code=" & RS.Fields("Party").Value
            DB.Execute Ssql
            DoEvents
        RS.MoveNext
        Loop
    End If
    RS.Close
    RSN.Close
    DB.Close
    
    'Raw Stocks
    'Blmr.OverAllRawStock CDate("DEC/31/" & YearN), ProgressBar1
    Blmr.ClosingStocks CDate("Jul/01/" & YearN - 1), CDate("Jun/30/" & YearN), ProgressBar1
    DoEvents
    Set DB = OpenDatabase(App.path & "\Book.mdb")
    Set RS = DB.OpenRecordset("Stock", dbOpenTable)
    If Not RS.EOF Then
        Do While Not RS.EOF
            
            Ssql = "Update Items Set OpWT=" & RS.Fields("ClosingQty").Value
            Ssql = Ssql & ",OpBales=" & RS.Fields("ClosingBales").Value
            Ssql = Ssql & ",Rate=" & RS.Fields("ClosingAmount").Value
            Ssql = Ssql & " where Code=" & RS.Fields("Code")
            DBN.Execute Ssql
            DoEvents
            
        RS.MoveNext
        Loop
    End If
    DBN.Close
    DB.Close
    
    MsgBox "All The Balances from the Current Year Transfered to the Next Year"
    Me.Hide
    Unload Me
End Sub
