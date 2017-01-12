VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAlarmer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   1620
      Top             =   3300
   End
   Begin MSFlexGridLib.MSFlexGrid G1 
      Height          =   5025
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8864
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAlarmer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FillGrid()
Dim Ssql As String
    Dim DB As Database
    Dim TB As Recordset
    
    Set DB = OpenDatabase(App.path & "\Years\" & YearN & "\Bloom.mdb")
    Ssql = "Select * from Alarm where VTime>=#" & Now & "# Order by VTime"
    Set TB = DB.OpenRecordset(Ssql)
    If Not TB.EOF Then
        With G1
            .Rows = 1
            .Cols = 3
            Do While Not TB.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(TB.Fields("VDate").Value, "dd-MMM-yyyy")
            .TextMatrix(.Rows - 1, 1) = Format(TB.Fields("VTime").Value, "dd-MMM-yyyy hh:nn:ss")
            .TextMatrix(.Rows - 1, 2) = TB.Fields("Note").Value & ""
            TB.MoveNext
            Loop
        End With
    End If
    TB.Close
    DB.Close
End Sub

Private Sub Timer1_Timer()
Dim R As Integer

For R = 1 To G1.Rows - 1
    With G1
        If Len(.TextMatrix(R, 1)) > 0 Then
            If CDate(.TextMatrix(R, 1)) = Now Then
                MsgBox .TextMatrix(R, 2)
                FillGrid
                DoEvents
            End If
        End If
    End With
Next R
End Sub
