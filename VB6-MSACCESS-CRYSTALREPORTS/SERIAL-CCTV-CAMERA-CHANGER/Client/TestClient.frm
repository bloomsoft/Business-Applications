VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "BloomSoft Remote CCTV Viewer"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5775
   Icon            =   "TestClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1920
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer1 
      Height          =   4860
      Left            =   0
      TabIndex        =   1
      Top             =   15
      Width           =   5775
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   10186
      _cy             =   8573
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Camera 4"
      Enabled         =   0   'False
      Height          =   195
      Index           =   8
      Left            =   12600
      TabIndex        =   0
      Top             =   6000
      Width           =   675
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mniStartViewing 
         Caption         =   "Start Viewing"
      End
      Begin VB.Menu mniViewCamera 
         Caption         =   "View Specific Camera"
      End
      Begin VB.Menu mniMultiMode 
         Caption         =   "Multi Mode"
      End
      Begin VB.Menu mniZoom 
         Caption         =   "Zoom"
      End
      Begin VB.Menu mniStill 
         Caption         =   "Still"
      End
      Begin VB.Menu mniAutoTimerMode 
         Caption         =   "Set Auto Timer Mode"
      End
      Begin VB.Menu mniServerSettings 
         Caption         =   "Server Settings"
      End
      Begin VB.Menu mniExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Sleep(Seconds As Integer)
Dim OLDDate As Date
OLDDate = Now
Dim D As Long

Do
D = DateDiff("s", Now, OLDDate)
If Abs(D) >= Seconds Then
    
    Exit Do
End If
DoEvents
Loop

End Sub

Private Sub ConnectAndSend(Data As String)
Dim S As String
Dim p As Integer
S = ServerIP
Winsock1.Close
Sleep 1
Winsock1.Connect S, 9000
Sleep 1
If Winsock1.State = sckConnected Then
    'MsgBox Combo1.ItemData(Combo1.ListIndex)
    Winsock1.SendData Data
End If
Sleep 1
Winsock1.Close

End Sub
Private Sub Command1_Click()
Dim S As String
Dim p As Integer
If Winsock1.State = 0 Then
p = InStr(Text1.Text, ":")
S = Left(Text1.Text, p - 1)
Winsock1.Connect S, 9000
Command2.Enabled = True
For p = 0 To 9
    Label1(p).Enabled = True
Next p
End If
MediaPlayer1.Open Text1.Text
End Sub

Private Sub Command2_Click()
Winsock1.SendData "S:" & Slider1.Value
End Sub

Private Sub Form_Resize()
MediaPlayer1.Width = Me.Width - 150
MediaPlayer1.Height = Me.Height - 150
End Sub

Private Sub SendCameraNumber(S As String)
Dim K As Long

Do While K < 1000
    K = K + 1
    DoEvents
Loop
Winsock1.SendData S
K = 0
Do While K < 1000
    K = K + 1
    DoEvents
Loop

End Sub

Private Sub Label1_Click(index As Integer)
'Dim R As Integer
'For R = 0 To 4
'    Label1(R).FontSize = 8
'    Label1(R).ForeColor = vbBlack
'Next R
'Label1(index).FontSize = 12
'Label1(index).ForeColor = vbRed
'Select Case index
'    Case 0
'        'SendCameraNumber "&H0"
'        SendCameraNumber = 1
'        Command2.Enabled = False
'    Case 1
'        'SendCameraNumber "&H8"
'        SendCameraNumber = 1
'        Command2.Enabled = False
'    Case 2
'    SendCameraNumber "&H10"
'    Command2.Enabled = False
'    Case 3
'    SendCameraNumber "&H18"
'    Command2.Enabled = False
'    Case 4
'    SendCameraNumber "T"
'    Command2.Enabled = True
'    Case 5
'    SendCameraNumber "&H20"
'    Command2.Enabled = False
'    Case 6
'    SendCameraNumber "&H28"
'    Command2.Enabled = False
'    Case 7
'    SendCameraNumber "&H30"
'    Command2.Enabled = False
'    Case 9
'    SendCameraNumber "&H38"
'    Command2.Enabled = False
'End Select
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Slider1_Scroll()
Label3.Caption = "Seconds : " & Slider1.Value
End Sub

Private Sub mniAutoTimerMode_Click()
Form3.Show vbModal
End Sub

Private Sub mniExit_Click()
End
End Sub

Private Sub mniMultiMode_Click()
ConnectAndSend "e"
End Sub

Private Sub mniServerSettings_Click()
Dim TS As TextStream
Dim J As String

J = InputBox("Please Enter Server IP Address And Add Port Too (i.e. 192.168.0.1:18666", "Server Info")

Set TS = FS.CreateTextFile(App.Path & "\ServerIP.txt", True)
TS.Write J
TS.Close
ServerIP = J
End Sub

Private Sub mniStartViewing_Click()
MediaPlayer1.URL = "http://" & ServerIP & ":1866"
MediaPlayer1.Controls.play
End Sub

Private Sub mniStill_Click()
ConnectAndSend "b"
End Sub

Private Sub mniViewCamera_Click()
Form2.Show vbModal
End Sub

Private Sub mniZoom_Click()
ConnectAndSend "c"
End Sub
