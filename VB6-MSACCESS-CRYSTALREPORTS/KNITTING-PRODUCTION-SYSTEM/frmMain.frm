VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ALPHA kNITTING (pvt) Ltd..(Main Menu)"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   5175
      Begin VB.Label lblQuota 
         Caption         =   "                 Quota Transfer Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Tag             =   "9"
         Top             =   4680
         Width           =   3180
      End
      Begin VB.Label INv 
         BackColor       =   &H8000000B&
         Caption         =   "                  Invoice Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Tag             =   "0"
         Top             =   360
         Width           =   3180
      End
      Begin VB.Label lblpaking 
         Caption         =   "                 Paking List Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Tag             =   "1"
         Top             =   840
         Width           =   3180
      End
      Begin VB.Label lblPham 
         Caption         =   "                  Visa Application Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Tag             =   "6"
         Top             =   3240
         Width           =   3180
      End
      Begin VB.Label PerformaInv 
         Caption         =   "                  Custom Performa Invoice Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Tag             =   "5"
         Top             =   2760
         Width           =   3540
      End
      Begin VB.Label lblReleas 
         Caption         =   "                  Release And Undertake Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Tag             =   "8"
         Top             =   4200
         Width           =   3420
      End
      Begin VB.Label lblBank 
         Caption         =   "                  Bank Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Tag             =   "7"
         Top             =   3720
         Width           =   3180
      End
      Begin VB.Label lblExpLic 
         Caption         =   "                  Export License"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Tag             =   "4"
         Top             =   2280
         Width           =   3180
      End
      Begin VB.Label GSP 
         Caption         =   "                  GSP Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1800
         Width           =   3180
      End
      Begin VB.Label lblCertificate 
         Caption         =   "                  Certificate Of Origin Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1320
         Width           =   3180
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Export System ( Main Menu)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E44E27&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin VB.Menu mnuForms 
      Caption         =   "&Forms"
      Begin VB.Menu export 
         Caption         =   "&Export License"
      End
      Begin VB.Menu Invoice 
         Caption         =   "&Invoice Form"
      End
      Begin VB.Menu gspform 
         Caption         =   "&GSP Form"
      End
      Begin VB.Menu Packing 
         Caption         =   "&Packing List Form"
      End
      Begin VB.Menu Performa 
         Caption         =   "&Performa Invoice Form"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Phma 
         Caption         =   "&Visa Application Form"
      End
      Begin VB.Menu Quota 
         Caption         =   "&Qouta Transfer Form"
      End
      Begin VB.Menu Release 
         Caption         =   "&Release And Undertaking"
      End
      Begin VB.Menu bank 
         Caption         =   "&Bank Form"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu explic 
         Caption         =   "&Export License"
      End
      Begin VB.Menu InvoiceReport 
         Caption         =   "&Invoice Report"
      End
      Begin VB.Menu GSPReport 
         Caption         =   "&GSP Report"
      End
      Begin VB.Menu PakingReport 
         Caption         =   "Paking List Report"
      End
      Begin VB.Menu PerformaReport 
         Caption         =   "Performa Invoice Report"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu PhmaReport 
         Caption         =   "PHMA Report "
      End
      Begin VB.Menu QuotaReport 
         Caption         =   "Quota Transfer Report"
      End
      Begin VB.Menu ReleaseReport 
         Caption         =   "Release And UnderTake Report"
      End
      Begin VB.Menu bankreport 
         Caption         =   "Bank Report"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
      Begin VB.Menu yes 
         Caption         =   "&Yes"
         Shortcut        =   ^Y
      End
      Begin VB.Menu no 
         Caption         =   "&No"
      End
   End
   Begin VB.Menu about 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Object Variables to Create Instance of ADO

Private Sub about_Click()
    frmAbout.Show
End Sub

Private Sub bank_Click()
    frmBank.Show
End Sub

Private Sub bankreport_Click()
    frmPrint.Show
End Sub

Private Sub Certificate_Click()
    
    frmCertificate.Show
    
End Sub
 


Private Sub Combined_Click()
    frmGenSystem.Show
    
End Sub

Private Sub explic_Click()
    frmPrint.Show
End Sub

Private Sub export_Click()
    frmExpLicense.Show
    
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ChangeSetting
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ChangeSetting
End Sub

Private Sub GSP_Click()
    GSP.FontItalic = True
    frmGenSystem.Show
End Sub

Private Sub GSP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ChangeSetting
    GSP.BorderStyle = 1
   GSP.BackColor = &HFCB67C
End Sub

Private Sub gspform_Click()
    frmGenSystem.Show
End Sub

Private Sub GSPReport_Click()
frmPrint.Show
End Sub

Private Sub INv_Click()
    frmInvoice.Show
    
End Sub

Private Sub INv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   INv.BorderStyle = 1
   INv.BackColor = &HFCB67C
    
End Sub




Private Sub Invoice_Click()
    frmInvoice.Show
End Sub

Private Sub InvoiceReport_Click()
frmPrint.Show
End Sub

Private Sub lblBank_Click()
    Me.ChangeSetting
    lblBank.FontItalic = True
    frmBank.Show
    
End Sub

Private Sub lblBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBank.BorderStyle = 1
   lblBank.BackColor = &HFCB67C
End Sub

Private Sub lblCertificate_Click()
    frmCertificate.Show
    
End Sub

Private Sub lblCertificate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Me.ChangeSetting
     lblCertificate.BorderStyle = 1
   lblCertificate.BackColor = &HFCB67C
End Sub

Private Sub lblExpLic_Click()
    frmExpLicense.Show
    
End Sub

Private Sub lblExpLic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ChangeSetting
    lblExpLic.BorderStyle = 1
   lblExpLic.BackColor = &HFCB67C
End Sub

Private Sub lblpaking_Click()
    frmPacking.Show
    
End Sub

Private Sub lblpaking_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ChangeSetting
    lblpaking.BorderStyle = 1
   lblpaking.BackColor = &HFCB67C
End Sub

Private Sub lblPham_Click()
    frmPHMA.Show
    
End Sub

Private Sub lblPham_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.ChangeSetting
    lblPham.BorderStyle = 1
   lblPham.BackColor = &HFCB67C
End Sub

Private Sub lblQuota_Click()
    
    frmQuota.Show
End Sub

Private Sub lblQuota_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ChangeSetting
    lblQuota.BorderStyle = 1
    lblQuota.BackColor = &HFCB67C
End Sub

Private Sub lblReleas_Click()
    frmRelease.Show
    
End Sub

Private Sub lblReleas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ChangeSetting
    lblReleas.BorderStyle = 1
    lblReleas.BackColor = &HFCB67C
End Sub

Private Sub Packing_Click()
    frmPacking.Show
    
End Sub

Private Sub PakingReport_Click()
frmPrint.Show
End Sub

Private Sub performa_Click()
    frmPerforma.Show
End Sub

Private Sub PerformaInv_Click()
    frmPerforma.Show
    
End Sub

Private Sub PerformaInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Me.ChangeSetting
        PerformaInv.BorderStyle = 1
        PerformaInv.BackColor = &HFCB67C
End Sub

Private Sub PerformaReport_Click()
        frmPrint.Show
End Sub

Private Sub phma_Click()
    frmPHMA.Show
End Sub

Private Sub PhmaReport_Click()
    frmPrint.Show
End Sub

Private Sub quota_Click()
    frmQuota.Show
End Sub

Private Sub QuotaReport_Click()
    frmPrint.Show
End Sub

Private Sub release_Click()
    frmRelease.Show
End Sub

Private Sub ReleaseReport_Click()
    frmPrint.Show
End Sub

Private Sub yes_Click()
    Dim Res As String
    Res = MsgBox("Do You want to exit ,It will end Your application", vbYesNo)
    If Res = vbYes Then
        End
    End If
    
    
End Sub

Sub ChangeSetting()
    Dim mControl As Control
        For Each mControl In Me.Controls
            If mControl.Tag <> "" And mControl.Tag >= 0 Then
                mControl.BackColor = &H8000000B
                mControl.BorderStyle = 0
            End If
        Next
End Sub

Sub Setting()
    Dim mControl As Control
        For Each mControl In Me.Controls
            If mControl.Tag <> "" And mControl.Tag >= 0 Then
                mControl.BackColor = &HFCB67C
                mControl.BorderStyle = 1
            End If
        Next
     
End Sub

