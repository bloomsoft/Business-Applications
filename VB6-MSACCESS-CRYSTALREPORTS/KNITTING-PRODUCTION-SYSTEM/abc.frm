VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print  Form"
   ClientHeight    =   3330
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.ComboBox cmbPrint 
         Height          =   315
         ItemData        =   "abc.frx":0000
         Left            =   2160
         List            =   "abc.frx":0025
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview Report"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Report"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Select Form Category"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "&Enter Invoice No."
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPreview_Click()
    
    Select Case cmbPrint.Text
        Case "Invoice Form"
            If deExport.rscmdMasInv.State = adStateOpen Then
                deExport.rscmdMasInv.Close
            End If
            
                deExport.cmdMasInv Text1
                RptInvoice.Show
            
           
        Case "Paking List"
            If deExport.rscmdPaking.State = adStateOpen Then
                deExport.rscmdPaking.Close
            End If
                deExport.cmdPaking Text1
                rptPaking.Show
            
        
        Case "Visa Application"
            If deExport.rscmdPHMA.State = adStateOpen Then
                deExport.rscmdPHMA.Close
            End If
                deExport.cmdPHMA Text1
                rptPHMA.Show
            
        Case "Export License"
            If deExport.rscmdMasExpLic.State = adStateOpen Then
                deExport.rscmdMasExpLic.Close
            End If
                deExport.cmdMasExpLic Text1
                rptExpLic.Show
            
        Case "Certificate Of Origin"
            If deExport.rscmdMasCertificate.State = adStateOpen Then
                deExport.rscmdMasCertificate.Close
            End If
                deExport.cmdMasCertificate Text1
                rptCertificate.Show
        Case "Bank Form"
            If deExport.rscmdBank.State = adStateOpen Then
                deExport.rscmdBank.Close
            End If
                deExport.cmdBank Text1
                rptBank.Show
            
        Case "Release And Take"
            If deExport.rscmdRelease.State = adStateOpen Then
                deExport.rscmdRelease.Close
            End If
                deExport.cmdRelease Text1
                rptRelease.Show
            
        Case "Performa Invoice"
            If deExport.rscmdMasPerforma.State = adStateOpen Then
                deExport.rscmdMasPerforma.Close
            End If
                deExport.cmdMasPerforma Text1
                rptPerforma.Show
            
        Case "Quota Transfer"
            If deExport.rscmdQuota.State = adStateOpen Then
                deExport.rscmdQuota.Close
            End If
                deExport.cmdQuota Text1
                rptQuota.Show
            
        Case "GSP Form"
            If deExport.rscmdMasGSP.State = adStateOpen Then
                deExport.rscmdMasGSP.Close
            End If
                deExport.cmdMasGSP Text1
                rptGSP.Show
        Case "GSP Additional"
           If deExport.rscmdGSP1.State = adStateOpen Then
               deExport.rscmdGSP1.Close
          End If
               deExport.cmdGSP1 Text1
               GSP1.Show
        End Select
        
          
     
End Sub

Private Sub cmdPrint_Click()

Select Case cmbPrint.Text
        Case "Invoice Form"
            If deExport.rscmdMasInv.State = adStateOpen Then
                deExport.rscmdMasInv.Close
            End If
                deExport.cmdMasInv Text1
                RptInvoice.PrintReport True
            
           
        Case "Paking List"
            If deExport.rscmdPaking.State = adStateOpen Then
                deExport.rscmdPaking.Close
            End If
                deExport.cmdPaking Text1
                rptPaking.PrintReport True
            
        
        Case "PHMA Form"
            If deExport.rscmdPHMA.State = adStateOpen Then
                deExport.rscmdPHMA.Close
            End If
                deExport.cmdPHMA Text1
                rptPHMA.PrintReport True
            
        Case "Export License"
            If deExport.rscmdMasExpLic.State = adStateOpen Then
                deExport.rscmdMasExpLic.Close
            End If
                deExport.cmdMasExpLic Text1
                rptExpLic.PrintReport True
                
            
        Case "Certificate Of Origin"
            If deExport.rscmdMasCertificate.State = adStateOpen Then
                deExport.rscmdMasCertificate.Close
            End If
                deExport.cmdMasCertificate Text1
                rptCertificate.PrintReport True
        Case "Bank Form"
            If deExport.rscmdBank.State = adStateOpen Then
                deExport.rscmdBank.Close
            End If
                deExport.cmdBank Text1
                rptBank.PrintReport True
            
        Case "Release And Take"
            If deExport.rscmdRelease.State = adStateOpen Then
                deExport.rscmdRelease.Close
            End If
                deExport.cmdRelease Text1
                rptRelease.PrintReport True
            
        Case "Performa Invoice"
            If deExport.rscmdMasPerforma.State = adStateOpen Then
                deExport.rscmdMasPerforma.Close
            End If
                deExport.cmdMasPerforma Text1
                rptPerforma.PrintReport True
            
        Case "Quota Transfer"
            If deExport.rscmdQuota.State = adStateOpen Then
                deExport.rscmdQuota.Close
            End If
                    deExport.cmdQuota Text1
                    rptQuota.PrintReport True
                    
        Case "GSP Form"
            If deExport.rscmdMasGSP.State = adStateOpen Then
                deExport.rscmdMasGSP.Close
            End If
                deExport.cmdMasGSP Text1
                rptGSP.PrintReport True
            
        Case "GSP Additional"
           If deExport.rscmdGSP1.State = adStateOpen Then
               deExport.rscmdGSP1.Close
           End If
               deExport.cmdGSP1 Text1
               
                GSP1.PrintReport True
                
        End Select
        
        
          
End Sub
